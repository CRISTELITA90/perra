import pandas as pd
import numpy as np
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from typing import Optional, Any, List, Dict

from betty import router as betty_router

app = FastAPI(
    title="Brain2Power API",
    version="0.5",
    description=(
        "Backend unificado: análisis de datos (data_analyst) + Betty (secretaria inteligente). "
        "Cada respuesta incluye `text_summary`, una frase lista para mostrar al usuario."
    ),
)

app.include_router(betty_router)


# ── Schemas ────────────────────────────────────────────────────────────────────

class DataInput(BaseModel):
    """Datos tabulares enviados directamente desde Copilot Studio"""
    columns: List[str] = Field(..., description="Lista de nombres de columnas")
    rows: List[Dict[str, Any]] = Field(..., description="Lista de filas como diccionarios")
    file_name: Optional[str] = Field(None, description="Nombre del archivo para referencia")


class AnalyzeDataInput(DataInput):
    group_by: Optional[str] = Field(None, description="Columna categórica para filtrar")
    group_value: Optional[str] = Field(None, description="Valor del filtro en group_by")
    target_column: Optional[str] = Field(None, description="Columna numérica de interés principal")


class ForecastDataInput(DataInput):
    date_column: str = Field(..., description="Columna que contiene fechas")
    value_column: str = Field(..., description="Columna numérica a predecir")
    group_by: Optional[str] = Field(None, description="Columna para filtrar")
    group_value: Optional[str] = Field(None, description="Valor del filtro")
    forecast_periods: int = Field(3, ge=1, le=24, description="Número de periodos futuros")


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
        if len(clean) > 0:
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
    if len(clean) < 4:
        return {"count": 0, "pct": 0.0, "lower_fence": 0.0, "upper_fence": 0.0}
    
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

@app.get("/", summary="Root", operation_id="root", tags=["system"])
def root():
    return {"ok": True, "service": "data-analyst-api", "version": "0.4"}


@app.get("/health", summary="Health check", operation_id="healthCheck", tags=["system"])
def health():
    return {"status": "ok"}


@app.post(
    "/describe-data",
    summary="Auditar un dataset recibido directamente",
    description=(
        "Recibe datos tabulares y devuelve: nombres de columnas, tipos de datos, "
        "conteo de nulos, estadísticas básicas y filas de muestra."
    ),
    operation_id="describeData",
    tags=["analysis"],
)
def describe_data(data: DataInput):
    try:
        df = pd.DataFrame(data.rows)
    except Exception as exc:
        raise HTTPException(400, f"Could not parse data: {exc}")

    if df.empty:
        raise HTTPException(400, "Dataset is empty")

    profiles = {col: _col_profile(df[col]) for col in df.columns}
    missing_map = {
        col: {"count": int(df[col].isna().sum()), "pct": round(df[col].isna().mean() * 100, 1)}
        for col in df.columns if df[col].isna().any()
    }
    sample = df.head(5).fillna("").astype(str).to_dict(orient="records")

    file_ref = data.file_name or "dataset"
    text = (
        f"Dataset '{file_ref}': {df.shape[0]} filas × {df.shape[1]} columnas. "
        f"Columnas: {list(df.columns)}. "
        f"Columnas con nulos: {list(missing_map.keys()) or 'ninguna'}."
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
    "/analyze-data",
    summary="EDA completo sobre datos recibidos",
    description=(
        "Realiza análisis exploratorio: estadísticas descriptivas, outliers (IQR), "
        "correlaciones de Pearson y distribuciones categóricas."
    ),
    operation_id="analyzeData",
    tags=["analysis"],
)
def analyze_data(data: AnalyzeDataInput):
    try:
        df = pd.DataFrame(data.rows)
    except Exception as exc:
        raise HTTPException(400, f"Could not parse data: {exc}")

    if df.empty:
        raise HTTPException(400, "Dataset is empty")

    if data.group_by and data.group_value:
        if data.group_by not in df.columns:
            raise HTTPException(400, f"Column '{data.group_by}' not found")
        df = df[df[data.group_by].astype(str) == data.group_value]
        if df.empty:
            raise HTTPException(400, "No rows match the filter")

    num_cols = df.select_dtypes(include="number").columns.tolist()
    cat_cols = df.select_dtypes(exclude="number").columns.tolist()

    stats: dict[str, Any] = {col: _col_profile(df[col]) for col in num_cols}
    outliers: dict[str, Any] = {col: _detect_outliers(df[col]) for col in num_cols}
    cat_dist: dict[str, Any] = {
        col: {
            "unique": int(df[col].nunique()),
            "top": {str(k): int(v) for k, v in df[col].value_counts().head(10).items()}
        }
        for col in cat_cols
    }
    corr = _correlation_matrix(df)

    target_summary: Optional[dict] = None
    if data.target_column:
        if data.target_column not in df.columns:
            raise HTTPException(400, f"target_column '{data.target_column}' not found")
        col_data = df[data.target_column]
        if pd.api.types.is_numeric_dtype(col_data):
            top_corr = {
                k: v for k, v in (corr.get(data.target_column) or {}).items()
                if k != data.target_column
            }
            target_summary = {
                "profile": _col_profile(col_data),
                "outliers": _detect_outliers(col_data),
                "top_correlations": dict(
                    sorted(top_corr.items(), key=lambda x: abs(x[1]), reverse=True)[:5]
                ),
            }

    filter_label = f"{data.group_by}={data.group_value}" if data.group_by else None
    file_ref = data.file_name or "dataset"
    text = (
        f"EDA sobre '{file_ref}'"
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
    "/forecast-data",
    summary="Forecast de serie temporal desde datos recibidos",
    description=(
        "Ajusta tendencia lineal y proyecta N periodos futuros. "
        "Devuelve datos históricos, tendencia y valores proyectados."
    ),
    operation_id="forecastData",
    tags=["forecast"],
)
def forecast_data(data: ForecastDataInput):
    try:
        df = pd.DataFrame(data.rows)
    except Exception as exc:
        raise HTTPException(400, f"Could not parse data: {exc}")

    if df.empty:
        raise HTTPException(400, "Dataset is empty")

    if data.group_by and data.group_value:
        if data.group_by not in df.columns:
            raise HTTPException(400, f"Column '{data.group_by}' not found")
        df = df[df[data.group_by].astype(str) == data.group_value]

    for col in (data.date_column, data.value_column):
        if col not in df.columns:
            raise HTTPException(400, f"Column '{col}' not found")

    df[data.date_column] = pd.to_datetime(df[data.date_column], errors="coerce")
    df = df.dropna(subset=[data.date_column, data.value_column]).sort_values(data.date_column)

    if df.empty:
        raise HTTPException(400, "No valid rows after filtering")

    values = df[data.value_column].astype(float).to_numpy()
    dates = df[data.date_column].dt.strftime("%Y-%m-%d").tolist()

    summary = {
        "count": int(len(values)),
        "total": round(float(values.sum()), 2),
        "mean": round(float(values.mean()), 2),
        "median": round(float(np.median(values)), 2),
        "min": round(float(values.min()), 2),
        "max": round(float(values.max()), 2),
        "std": round(float(values.std()), 2),
    }

    forecast_list: list[dict] = []
    trend = "flat"
    if len(values) > 1:
        x = np.arange(len(values), dtype=float)
        coeffs = np.polyfit(x, values, 1)
        slope = float(coeffs[0])
        trend = "ascending" if slope > 0.01 else "descending" if slope < -0.01 else "flat"
        summary["slope_per_period"] = round(slope, 4)
        for i in range(1, data.forecast_periods + 1):
            forecast_list.append({
                "period": i,
                "value": round(float(np.polyval(coeffs, len(values) - 1 + i)), 2),
            })

    summary["trend"] = trend
    filter_label = f"{data.group_by}={data.group_value}" if data.group_by else None
    fcast_vals = [f["value"] for f in forecast_list]

    file_ref = data.file_name or "dataset"
    text = (
        f"Forecast de '{data.value_column}' en '{file_ref}'"
        + (f" ({filter_label})" if filter_label else "")
        + f": {summary['count']} periodos históricos. "
        f"Media: {summary['mean']}, Tendencia: {trend}. "
        f"Proyección {data.forecast_periods} periodos: {fcast_vals}."
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
