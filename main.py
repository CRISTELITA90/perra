import os
import io
import requests
import pandas as pd
import numpy as np
from fastapi import FastAPI, HTTPException
from fastapi.openapi.utils import get_openapi
from pydantic import BaseModel, Field
from typing import Optional, Any

app = FastAPI(
    title="Data Analyst API",
    version="0.3",
    description=(
        "Backend para el agente data_analyst de Copilot Studio. "
        "Carga datasets desde SharePoint y realiza análisis estadístico, EDA y forecasting. "
        "Cada respuesta incluye `text_summary`, una frase lista para mostrar al usuario."
    ),
)


# ── Microsoft Graph auth ───────────────────────────────────────────────────────

def _graph_token() -> str:
    resp = requests.post(
        f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": os.environ["CLIENT_ID"],
            "client_secret": os.environ["CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _fetch_file(token: str, file_path: str) -> bytes:
    site = os.environ["SHAREPOINT_SITE"]  # e.g. "mycompany.sharepoint.com:/sites/MySite"
    h = {"Authorization": f"Bearer {token}"}

    site_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site}", headers=h, timeout=30
    )
    site_resp.raise_for_status()
    site_id = site_resp.json()["id"]

    file_resp = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/content",
        headers=h, timeout=60,
    )
    file_resp.raise_for_status()
    return file_resp.content


def _load_df(token: str, file_path: str, sheet: Optional[str]) -> pd.DataFrame:
    raw = _fetch_file(token, file_path)
    if file_path.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(raw))
    return pd.read_excel(io.BytesIO(raw), sheet_name=sheet or 0)


def _auth_and_load(file_path: str, sheet: Optional[str]) -> pd.DataFrame:
    try:
        token = _graph_token()
    except Exception as exc:
        raise HTTPException(502, f"Graph auth failed: {exc}")
    try:
        return _load_df(token, file_path, sheet)
    except requests.HTTPError as exc:
        raise HTTPException(502, f"SharePoint fetch failed: {exc}")
    except Exception as exc:
        raise HTTPException(400, f"Could not parse file: {exc}")


# ── Schemas ────────────────────────────────────────────────────────────────────

class FileRef(BaseModel):
    file_path: str = Field(
        ...,
        description="Ruta al archivo dentro de la biblioteca de SharePoint.",
        examples=["Shared Documents/ventas.xlsx"],
    )
    sheet: Optional[str] = Field(
        None,
        description="Nombre de la hoja Excel. Dejar vacío para usar la primera hoja.",
    )


class AnalyzeRequest(FileRef):
    group_by: Optional[str] = Field(
        None,
        description="Columna categórica para filtrar (ej. 'Municipio').",
    )
    group_value: Optional[str] = Field(
        None,
        description="Valor del filtro en group_by (ej. 'Madrid').",
    )
    target_column: Optional[str] = Field(
        None,
        description="Columna numérica de interés principal: se detallan sus correlaciones y outliers.",
    )


class ForecastRequest(FileRef):
    date_column: str = Field(
        ...,
        description="Columna que contiene fechas o periodos temporales.",
        examples=["Fecha"],
    )
    value_column: str = Field(
        ...,
        description="Columna numérica a predecir.",
        examples=["Ventas"],
    )
    group_by: Optional[str] = Field(None, description="Columna para filtrar (ej. 'Municipio').")
    group_value: Optional[str] = Field(None, description="Valor del filtro.")
    forecast_periods: int = Field(
        3,
        ge=1,
        le=24,
        description="Número de periodos futuros a proyectar.",
    )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _col_profile(series: pd.Series) -> dict[str, Any]:
    dtype = str(series.dtype)
    nulls = int(series.isna().sum())
    total = len(series)
    profile: dict[str, Any] = {
        "dtype": dtype,
        "null_count": nulls,
        "null_pct": round(nulls / total * 100, 1) if total else 0,
        "unique": int(series.nunique()),
    }
    if pd.api.types.is_numeric_dtype(series):
        clean = series.dropna()
        profile.update({
            "min": round(float(clean.min()), 4),
            "max": round(float(clean.max()), 4),
            "mean": round(float(clean.mean()), 4),
            "median": round(float(clean.median()), 4),
            "std": round(float(clean.std()), 4),
            "q25": round(float(clean.quantile(0.25)), 4),
            "q75": round(float(clean.quantile(0.75)), 4),
        })
    else:
        top = series.value_counts().head(5)
        profile["top_values"] = {str(k): int(v) for k, v in top.items()}
    return profile


def _detect_outliers(series: pd.Series) -> dict[str, Any]:
    clean = series.dropna()
    q1, q3 = clean.quantile(0.25), clean.quantile(0.75)
    iqr = q3 - q1
    lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    outliers = clean[(clean < lower) | (clean > upper)]
    return {
        "count": int(len(outliers)),
        "pct": round(len(outliers) / len(clean) * 100, 1) if len(clean) else 0,
        "lower_fence": round(float(lower), 4),
        "upper_fence": round(float(upper), 4),
    }


def _correlation_matrix(df: pd.DataFrame) -> dict[str, dict[str, float]]:
    num = df.select_dtypes(include="number")
    if num.shape[1] < 2:
        return {}
    corr = num.corr().round(3)
    return {col: corr[col].dropna().to_dict() for col in corr.columns}


# ── Endpoints ──────────────────────────────────────────────────────────────────

@app.get(
    "/",
    summary="Root",
    operation_id="root",
    tags=["system"],
)
def root():
    return {"ok": True, "service": "data-analyst-api", "version": "0.3"}


@app.get(
    "/health",
    summary="Health check",
    description="Verifica que el servicio está activo. Úsalo para comprobar conectividad.",
    operation_id="healthCheck",
    tags=["system"],
)
def health():
    return {"status": "ok"}


@app.post(
    "/describe",
    summary="Auditar un dataset de SharePoint",
    description=(
        "**Paso 1 del pipeline.** Carga un archivo de SharePoint y devuelve: "
        "nombres de columnas, tipos de datos, conteo de nulos, estadísticas básicas y filas de muestra. "
        "Llama a este endpoint primero para que el agente sepa con qué columnas está trabajando."
    ),
    operation_id="describeDataset",
    tags=["analysis"],
)
def describe(req: FileRef):
    df = _auth_and_load(req.file_path, req.sheet)

    profiles = {col: _col_profile(df[col]) for col in df.columns}
    missing_map = {
        col: {"count": int(df[col].isna().sum()), "pct": round(df[col].isna().mean() * 100, 1)}
        for col in df.columns if df[col].isna().any()
    }
    sample = df.head(5).fillna("").astype(str).to_dict(orient="records")

    text = (
        f"Dataset '{req.file_path}': {df.shape[0]} filas × {df.shape[1]} columnas. "
        f"Columnas: {list(df.columns)}. "
        f"Columnas con nulos: {list(missing_map.keys()) or 'ninguna'}. "
        f"Tipos: {df.dtypes.astype(str).to_dict()}."
    )

    return {
        "status": "ok",
        "shape": {"rows": df.shape[0], "columns": df.shape[1]},
        "column_profiles": profiles,
        "missing_values": missing_map,
        "sample_rows": sample,
        "text_summary": text,
    }


@app.post(
    "/analyze",
    summary="EDA completo sobre un dataset de SharePoint",
    description=(
        "Realiza análisis exploratorio de datos (EDA) completo: estadísticas descriptivas, "
        "detección de outliers (método IQR), correlaciones de Pearson y distribuciones categóricas. "
        "Opcionalmente filtra por columna categórica y enfoca el análisis en una columna objetivo."
    ),
    operation_id="analyzeDataset",
    tags=["analysis"],
)
def analyze(req: AnalyzeRequest):
    df = _auth_and_load(req.file_path, req.sheet)

    if req.group_by and req.group_value:
        if req.group_by not in df.columns:
            raise HTTPException(400, f"Column '{req.group_by}' not found")
        df = df[df[req.group_by].astype(str) == req.group_value]
        if df.empty:
            raise HTTPException(400, "No rows match the filter")

    num_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols = df.select_dtypes(exclude="number").columns.tolist()

    stats: dict[str, Any] = {col: _col_profile(df[col]) for col in num_cols}
    outliers: dict[str, Any] = {col: _detect_outliers(df[col]) for col in num_cols}
    cat_dist: dict[str, Any] = {
        col: {"unique": int(df[col].nunique()),
              "top": {str(k): int(v) for k, v in df[col].value_counts().head(10).items()}}
        for col in cat_cols
    }
    corr = _correlation_matrix(df)

    target_summary: Optional[dict] = None
    if req.target_column:
        if req.target_column not in df.columns:
            raise HTTPException(400, f"target_column '{req.target_column}' not found")
        col_data = df[req.target_column]
        if pd.api.types.is_numeric_dtype(col_data):
            top_corr = {
                k: v for k, v in (corr.get(req.target_column) or {}).items()
                if k != req.target_column
            }
            target_summary = {
                "profile": _col_profile(col_data),
                "outliers": _detect_outliers(col_data),
                "top_correlations": dict(
                    sorted(top_corr.items(), key=lambda x: abs(x[1]), reverse=True)[:5]
                ),
            }

    filter_label = f"{req.group_by}={req.group_value}" if req.group_by else None
    text = (
        f"EDA sobre '{req.file_path}'"
        + (f" filtrado {filter_label}" if filter_label else "")
        + f": {df.shape[0]} filas, {len(num_cols)} columnas numéricas, {len(cat_cols)} categóricas. "
        f"Columnas con outliers significativos (>5%): "
        f"{[c for c, o in outliers.items() if o['pct'] > 5] or 'ninguna'}."
    )

    return {
        "status": "ok",
        "filter": filter_label,
        "shape": {"rows": df.shape[0], "columns": df.shape[1]},
        "numeric_stats": stats,
        "outliers": outliers,
        "categorical_distributions": cat_dist,
        "correlations": corr,
        "target_analysis": target_summary,
        "text_summary": text,
    }


@app.post(
    "/run-forecast",
    summary="Forecast de una serie temporal desde SharePoint",
    description=(
        "Ajusta una tendencia lineal a una columna numérica temporal y proyecta N periodos futuros. "
        "Devuelve datos históricos, dirección de tendencia (ascending/descending/flat) "
        "y valores proyectados. Opcionalmente filtra por columna de grupo."
    ),
    operation_id="runForecast",
    tags=["forecast"],
)
def run_forecast(req: ForecastRequest):
    df = _auth_and_load(req.file_path, req.sheet)

    if req.group_by and req.group_value:
        if req.group_by not in df.columns:
            raise HTTPException(400, f"Column '{req.group_by}' not found")
        df = df[df[req.group_by].astype(str) == req.group_value]

    for col in (req.date_column, req.value_column):
        if col not in df.columns:
            raise HTTPException(400, f"Column '{col}' not found")

    df[req.date_column] = pd.to_datetime(df[req.date_column], errors="coerce")
    df = df.dropna(subset=[req.date_column, req.value_column]).sort_values(req.date_column)

    if df.empty:
        raise HTTPException(400, "No valid rows after filtering")

    values = df[req.value_column].astype(float).to_numpy()
    dates = df[req.date_column].dt.strftime("%Y-%m-%d").tolist()

    summary = {
        "count":  int(len(values)),
        "total":  round(float(values.sum()), 2),
        "mean":   round(float(values.mean()), 2),
        "median": round(float(np.median(values)), 2),
        "min":    round(float(values.min()), 2),
        "max":    round(float(values.max()), 2),
        "std":    round(float(values.std()), 2),
    }

    forecast_list: list[dict] = []
    trend = "flat"
    if len(values) > 1:
        x = np.arange(len(values), dtype=float)
        coeffs = np.polyfit(x, values, 1)
        slope = float(coeffs[0])
        trend = "ascending" if slope > 0.01 else "descending" if slope < -0.01 else "flat"
        summary["slope_per_period"] = round(slope, 4)
        for i in range(1, req.forecast_periods + 1):
            forecast_list.append({
                "period": i,
                "value": round(float(np.polyval(coeffs, len(values) - 1 + i)), 2),
            })

    summary["trend"] = trend
    filter_label = f"{req.group_by}={req.group_value}" if req.group_by else None
    fcast_vals = [f["value"] for f in forecast_list]

    text = (
        f"Forecast de '{req.value_column}'"
        + (f" ({filter_label})" if filter_label else "")
        + f": {summary['count']} periodos históricos. "
        f"Media: {summary['mean']}, Tendencia: {trend}. "
        f"Proyección {req.forecast_periods} periodos: {fcast_vals}."
    )

    return {
        "status": "ok",
        "filter": filter_label,
        "summary": summary,
        "historical": [
            {"date": d, "value": round(float(v), 2)} for d, v in zip(dates, values)
        ],
        "forecast": forecast_list,
        "text_summary": text,
    }
