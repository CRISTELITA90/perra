"""
Betty — Secretaria inteligente de Brain2Power
=============================================
Agente que centraliza:
  1. Oportunidades de negocio: congresos, jornadas y foros de energía / eficiencia energética.
  2. Correo Outlook: tareas y asuntos pendientes.
  3. SharePoint: documentos, entregables y actividad del equipo Brain2Power.
  4. Redes sociales: métricas Facebook (Page Insights), Instagram Business, LinkedIn.
  5. Configuración y estado de todos los servicios conectados.
  6. Briefing diario consolidado.

Variables de entorno requeridas:
  Microsoft 365 → TENANT_ID, CLIENT_ID, CLIENT_SECRET, USER_EMAIL, SHAREPOINT_SITE
  Facebook/Instagram → FB_PAGE_TOKEN, FB_PAGE_ID, IG_BUSINESS_ID
  LinkedIn → LI_ACCESS_TOKEN, LI_ORGANIZATION_ID
"""

import os
import re
import json
import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Optional

import requests
from fastapi import APIRouter, FastAPI, HTTPException, Query, Body
from pydantic import BaseModel

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("betty")

# ── Router (se incluye en main.py) ───────────────────────────────────────────
router = APIRouter(tags=["betty"])

# FastAPI standalone para desarrollo/testing independiente
app = FastAPI(
    title="Betty — Secretaria Brain2Power",
    version="1.0",
    description=(
        "Agente secretaria que detecta oportunidades de negocio en energía, "
        "gestiona tareas de Outlook y hace seguimiento de entregables en SharePoint."
    ),
)

# ── Constantes de energía ─────────────────────────────────────────────────────
ENERGY_KEYWORDS = [
    "energía", "energy", "eficiencia energética", "energy efficiency",
    "renovable", "renewable", "sostenibilidad", "sustainability",
    "brain2power", "eficiencia", "ahorro energético", "transición energética",
    "smart grid", "microgrid", "almacenamiento energía", "solar", "eólica",
    "biomasa", "cogeneración", "auditoría energética", "certificación energética",
    "electrificación", "descarbonización", "neutralidad carbono",
]

EVENT_KEYWORDS = [
    "congreso", "jornada", "foro", "conferencia", "seminario", "simposio",
    "workshop", "summit", "expo", "feria", "webinar", "encuentro",
]

# ── Microsoft Graph helpers ────────────────────────────────────────────────────

class GraphClient:
    """Cliente ligero para Microsoft Graph API con caché de token."""

    _token: Optional[str] = None
    _token_expiry: Optional[datetime] = None

    def __init__(self):
        self.tenant_id = os.environ.get("TENANT_ID", "")
        self.client_id = os.environ.get("CLIENT_ID", "")
        self.client_secret = os.environ.get("CLIENT_SECRET", "")
        self.user_email = os.environ.get("USER_EMAIL", "")
        self.sharepoint_site = os.environ.get("SHAREPOINT_SITE", "")
        self.base_url = "https://graph.microsoft.com/v1.0"

    def _get_token(self) -> str:
        now = datetime.now(timezone.utc)
        if self._token and self._token_expiry and now < self._token_expiry:
            return self._token

        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise HTTPException(
                503,
                "Faltan variables de entorno: TENANT_ID, CLIENT_ID, CLIENT_SECRET. "
                "Configúralas para conectar con Microsoft 365.",
            )

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        resp = requests.post(
            url,
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default",
            },
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
        GraphClient._token = data["access_token"]
        GraphClient._token_expiry = now + timedelta(seconds=data.get("expires_in", 3600) - 60)
        return GraphClient._token

    def get(self, path: str, params: Optional[dict] = None) -> dict:
        token = self._get_token()
        resp = requests.get(
            f"{self.base_url}{path}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            params=params or {},
            timeout=20,
        )
        if resp.status_code == 404:
            return {"value": []}
        resp.raise_for_status()
        return resp.json()


graph = GraphClient()


# ── Schemas de respuesta ───────────────────────────────────────────────────────

class TaskItem(BaseModel):
    id: str
    subject: str
    due_date: Optional[str]
    priority: Optional[str]
    status: str
    source: str  # "outlook_email" | "todo" | "planner"


class EventItem(BaseModel):
    title: str
    date: Optional[str]
    location: Optional[str]
    url: Optional[str]
    relevance_score: int
    keywords_found: list[str]
    source: str


class DeliverableItem(BaseModel):
    name: str
    modified_by: Optional[str]
    modified_date: Optional[str]
    url: Optional[str]
    folder: Optional[str]


class TeamMemberActivity(BaseModel):
    name: str
    email: Optional[str]
    last_document: Optional[str]
    last_modified: Optional[str]
    recent_emails: int


class BriefingResponse(BaseModel):
    generated_at: str
    pending_tasks: list[TaskItem]
    energy_events: list[EventItem]
    deliverables: list[DeliverableItem]
    team_activity: list[TeamMemberActivity]
    summary_text: str


# ── Helpers internos ───────────────────────────────────────────────────────────

def _relevance_score(text: str) -> tuple[int, list[str]]:
    """Puntúa la relevancia energética de un texto y devuelve keywords halladas."""
    text_lower = text.lower()
    found = [kw for kw in ENERGY_KEYWORDS if kw in text_lower]
    event_match = any(ew in text_lower for ew in EVENT_KEYWORDS)
    score = len(found) * 10 + (15 if event_match else 0)
    return score, found


def _safe_str(val: Any, default: str = "—") -> str:
    return str(val) if val is not None else default


# ── Endpoint: tareas pendientes de Outlook ──────────────────────────────────

@router.get("/betty/tasks",
    summary="Tareas pendientes en Outlook / To Do",
    operation_id="bettyTasks",
    tags=["betty"],
    response_model=list[TaskItem],
)
def get_pending_tasks(days_back: int = Query(30, ge=1, le=90)):
    """Devuelve emails sin leer importantes y tareas de Microsoft To Do."""
    tasks: list[TaskItem] = []

    # 1. Emails importantes no leídos
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        data = graph.get(
            f"/users/{graph.user_email}/messages",
            params={
                "$filter": f"isRead eq false and receivedDateTime ge {cutoff}",
                "$select": "id,subject,receivedDateTime,importance,from",
                "$orderby": "importance desc,receivedDateTime desc",
                "$top": "25",
            },
        )
        for msg in data.get("value", []):
            tasks.append(TaskItem(
                id=msg["id"],
                subject=msg.get("subject", "(sin asunto)"),
                due_date=msg.get("receivedDateTime", "")[:10],
                priority=msg.get("importance", "normal"),
                status="no leído",
                source="outlook_email",
            ))
    except Exception as exc:
        log.warning("No se pudo obtener emails: %s", exc)

    # 2. Microsoft To Do — listas de tareas
    try:
        lists_data = graph.get(f"/users/{graph.user_email}/todo/lists")
        for lst in lists_data.get("value", [])[:5]:
            list_id = lst["id"]
            tasks_data = graph.get(
                f"/users/{graph.user_email}/todo/lists/{list_id}/tasks",
                params={"$filter": "status ne 'completed'", "$top": "20"},
            )
            for t in tasks_data.get("value", []):
                due = ""
                if t.get("dueDateTime"):
                    due = t["dueDateTime"].get("dateTime", "")[:10]
                tasks.append(TaskItem(
                    id=t["id"],
                    subject=t.get("title", "(sin título)"),
                    due_date=due or None,
                    priority=t.get("importance", "normal"),
                    status=t.get("status", "notStarted"),
                    source="todo",
                ))
    except Exception as exc:
        log.warning("No se pudo obtener To Do: %s", exc)

    return tasks


# ── Endpoint: eventos de energía detectados en email/calendario ─────────────

@router.get("/betty/energy-events",
    summary="Eventos de energía detectados en Outlook",
    operation_id="bettyEnergyEvents",
    tags=["betty"],
    response_model=list[EventItem],
)
def get_energy_events(days_forward: int = Query(180, ge=7, le=365)):
    """
    Analiza el calendario y emails para detectar congresos, jornadas y foros
    relacionados con energía y eficiencia energética.
    """
    events: list[EventItem] = []
    now = datetime.now(timezone.utc)
    future = now + timedelta(days=days_forward)

    keyword_filter = " OR ".join(f'"{kw}"' for kw in ["energía", "energy", "brain2power", "eficiencia"])

    # 1. Eventos de calendario
    try:
        cal_data = graph.get(
            f"/users/{graph.user_email}/calendarView",
            params={
                "startDateTime": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "endDateTime": future.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "$select": "id,subject,start,end,location,webLink,bodyPreview",
                "$top": "50",
            },
        )
        for ev in cal_data.get("value", []):
            text = f"{ev.get('subject','')} {ev.get('bodyPreview','')} {ev.get('location',{}).get('displayName','')}"
            score, found = _relevance_score(text)
            if score > 0:
                events.append(EventItem(
                    title=ev.get("subject", "(sin título)"),
                    date=ev.get("start", {}).get("dateTime", "")[:10],
                    location=ev.get("location", {}).get("displayName"),
                    url=ev.get("webLink"),
                    relevance_score=score,
                    keywords_found=found,
                    source="outlook_calendar",
                ))
    except Exception as exc:
        log.warning("No se pudo leer calendario: %s", exc)

    # 2. Emails con términos de eventos energéticos
    cutoff = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    try:
        email_data = graph.get(
            f"/users/{graph.user_email}/messages",
            params={
                "$search": keyword_filter,
                "$select": "id,subject,receivedDateTime,from,webLink,bodyPreview",
                "$top": "25",
            },
        )
        for msg in email_data.get("value", []):
            text = f"{msg.get('subject','')} {msg.get('bodyPreview','')}"
            score, found = _relevance_score(text)
            event_hit = any(ew in text.lower() for ew in EVENT_KEYWORDS)
            if score > 0 and event_hit:
                events.append(EventItem(
                    title=f"[EMAIL] {msg.get('subject', '(sin asunto)')}",
                    date=msg.get("receivedDateTime", "")[:10],
                    location=None,
                    url=msg.get("webLink"),
                    relevance_score=score,
                    keywords_found=found,
                    source="outlook_email",
                ))
    except Exception as exc:
        log.warning("No se pudo buscar emails de eventos: %s", exc)

    events.sort(key=lambda e: e.relevance_score, reverse=True)
    return events


# ── Endpoint: entregables de SharePoint ─────────────────────────────────────

@router.get("/betty/deliverables",
    summary="Entregables del proyecto Brain2Power en SharePoint",
    operation_id="bettyDeliverables",
    tags=["betty"],
    response_model=list[DeliverableItem],
)
def get_deliverables(folder: str = Query("brain2power", description="Carpeta raíz en SharePoint")):
    """Lista documentos recientes en la carpeta del proyecto en SharePoint."""
    deliverables: list[DeliverableItem] = []
    site = graph.sharepoint_site or os.environ.get("SHAREPOINT_SITE", "")

    if not site:
        raise HTTPException(503, "Variable SHAREPOINT_SITE no configurada.")

    try:
        # Resolver site ID
        site_data = graph.get(f"/sites/{site}")
        site_id = site_data.get("id", "")

        # Buscar drive raíz
        drives = graph.get(f"/sites/{site_id}/drives")
        drive_id = drives.get("value", [{}])[0].get("id", "") if drives.get("value") else ""

        if not drive_id:
            raise HTTPException(404, "No se encontró ningún drive en el sitio SharePoint.")

        # Buscar carpeta del proyecto
        items = graph.get(
            f"/drives/{drive_id}/root:/{folder}:/children",
            params={"$select": "id,name,lastModifiedBy,lastModifiedDateTime,webUrl,folder", "$top": "50"},
        )
        for item in items.get("value", []):
            mod_by = item.get("lastModifiedBy", {}).get("user", {}).get("displayName")
            deliverables.append(DeliverableItem(
                name=item.get("name", ""),
                modified_by=mod_by,
                modified_date=item.get("lastModifiedDateTime", "")[:10],
                url=item.get("webUrl"),
                folder=folder,
            ))
    except HTTPException:
        raise
    except Exception as exc:
        log.error("Error accediendo SharePoint: %s", exc)
        raise HTTPException(502, f"No se pudo acceder a SharePoint: {exc}")

    deliverables.sort(key=lambda d: d.modified_date or "", reverse=True)
    return deliverables


# ── Endpoint: actividad del equipo ─────────────────────────────────────────

@router.get("/betty/team-activity",
    summary="Actividad reciente del equipo Brain2Power",
    operation_id="bettyTeamActivity",
    tags=["betty"],
    response_model=list[TeamMemberActivity],
)
def get_team_activity(
    team_emails: str = Query(
        "",
        description="Lista de emails del equipo separados por coma. Ej: ana@empresa.com,luis@empresa.com",
    ),
    days_back: int = Query(14, ge=1, le=60),
):
    """Resume la actividad reciente de cada integrante del equipo en emails y SharePoint."""
    if not team_emails.strip():
        return []

    members = [e.strip() for e in team_emails.split(",") if e.strip()]
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    activities: list[TeamMemberActivity] = []

    for email in members:
        recent_emails = 0
        last_doc = None
        last_mod = None

        # Contar emails recibidos de este miembro
        try:
            msgs = graph.get(
                f"/users/{graph.user_email}/messages",
                params={
                    "$filter": f"from/emailAddress/address eq '{email}' and receivedDateTime ge {cutoff}",
                    "$select": "id,subject,receivedDateTime",
                    "$top": "25",
                },
            )
            recent_emails = len(msgs.get("value", []))
        except Exception as exc:
            log.warning("No se pudo obtener emails de %s: %s", email, exc)

        # Último documento modificado en SharePoint por este usuario
        site = graph.sharepoint_site or os.environ.get("SHAREPOINT_SITE", "")
        if site:
            try:
                site_data = graph.get(f"/sites/{site}")
                site_id = site_data.get("id", "")
                drives = graph.get(f"/sites/{site_id}/drives")
                drive_id = drives.get("value", [{}])[0].get("id", "") if drives.get("value") else ""
                if drive_id:
                    search_res = graph.get(
                        f"/drives/{drive_id}/root/search(q='brain2power')",
                        params={
                            "$select": "name,lastModifiedBy,lastModifiedDateTime,webUrl",
                            "$top": "50",
                        },
                    )
                    for item in search_res.get("value", []):
                        mod_email = (
                            item.get("lastModifiedBy", {})
                            .get("user", {})
                            .get("email", "")
                        )
                        if mod_email.lower() == email.lower():
                            mod_date = item.get("lastModifiedDateTime", "")
                            if not last_mod or mod_date > last_mod:
                                last_doc = item.get("name")
                                last_mod = mod_date[:10]
            except Exception as exc:
                log.warning("No se pudo buscar docs de %s: %s", email, exc)

        name = email.split("@")[0].replace(".", " ").title()
        activities.append(TeamMemberActivity(
            name=name,
            email=email,
            last_document=last_doc,
            last_modified=last_mod,
            recent_emails=recent_emails,
        ))

    return activities


# ── Endpoint: briefing diario completo ─────────────────────────────────────

@router.get("/betty/briefing",
    summary="Briefing diario completo de Betty",
    operation_id="bettyBriefing",
    tags=["betty"],
    response_model=BriefingResponse,
)
def get_briefing(
    team_emails: str = Query("", description="Emails del equipo separados por coma"),
    sharepoint_folder: str = Query("brain2power", description="Carpeta raíz en SharePoint"),
):
    """
    Resumen ejecutivo que Betty prepara cada mañana:
    tareas pendientes, eventos de energía, entregables del proyecto y actividad del equipo.
    """
    now_str = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    # Recopilar todo en paralelo (secuencial por simplicidad, suficiente para una secretaria)
    tasks: list[TaskItem] = []
    energy_events: list[EventItem] = []
    deliverables: list[DeliverableItem] = []
    team_activity: list[TeamMemberActivity] = []
    errors: list[str] = []

    try:
        tasks = get_pending_tasks()
    except HTTPException as e:
        errors.append(f"Tareas: {e.detail}")

    try:
        energy_events = get_energy_events()
    except HTTPException as e:
        errors.append(f"Eventos: {e.detail}")

    try:
        deliverables = get_deliverables(folder=sharepoint_folder)
    except HTTPException as e:
        errors.append(f"Entregables: {e.detail}")

    try:
        if team_emails.strip():
            team_activity = get_team_activity(team_emails=team_emails)
    except HTTPException as e:
        errors.append(f"Equipo: {e.detail}")

    # Resumen textual
    high_tasks = [t for t in tasks if t.priority in ("high", "urgent")]
    upcoming_events = energy_events[:3]
    recent_docs = deliverables[:5]

    summary_lines = [
        f"Buenos días. Briefing Brain2Power — {now_str}.",
        "",
        f"📋 TAREAS PENDIENTES: {len(tasks)} total"
        + (f", {len(high_tasks)} de alta prioridad." if high_tasks else "."),
    ]

    if high_tasks:
        for t in high_tasks[:3]:
            summary_lines.append(f"  • [{t.source}] {t.subject} (vence: {t.due_date or 'sin fecha'})")

    summary_lines += [
        "",
        f"⚡ OPORTUNIDADES ENERGÉTICAS: {len(energy_events)} eventos detectados.",
    ]
    for ev in upcoming_events:
        kw_str = ", ".join(ev.keywords_found[:3])
        summary_lines.append(f"  • {ev.title} | {ev.date or '?'} | Keywords: {kw_str}")

    summary_lines += [
        "",
        f"📁 ENTREGABLES RECIENTES en SharePoint/{sharepoint_folder}: {len(deliverables)} documentos.",
    ]
    for doc in recent_docs:
        summary_lines.append(f"  • {doc.name} — modificado por {doc.modified_by or '?'} el {doc.modified_date or '?'}")

    if team_activity:
        summary_lines += ["", f"👥 ACTIVIDAD DEL EQUIPO ({len(team_activity)} integrantes):"]
        for m in team_activity:
            summary_lines.append(
                f"  • {m.name}: {m.recent_emails} emails enviados"
                + (f", último doc: {m.last_document} ({m.last_modified})" if m.last_document else ", sin docs recientes.")
            )

    if errors:
        summary_lines += ["", "⚠️ Algunos servicios no respondieron:", *[f"  - {e}" for e in errors]]

    summary_text = "\n".join(summary_lines)

    return BriefingResponse(
        generated_at=now_str,
        pending_tasks=tasks,
        energy_events=energy_events,
        deliverables=deliverables,
        team_activity=team_activity,
        summary_text=summary_text,
    )


# ── Endpoint raíz de Betty ─────────────────────────────────────────────────

@router.get("/betty", summary="Presentación de Betty", operation_id="bettyHello", tags=["betty"])
def betty_hello():
    return {
        "agent": "Betty",
        "role": "Secretaria inteligente de Brain2Power",
        "version": "1.0",
        "capabilities": [
            "Detectar oportunidades de negocio (congresos, jornadas, foros de energía)",
            "Revisar tareas pendientes en Outlook y Microsoft To Do",
            "Seguimiento de entregables en SharePoint",
            "Monitorizar actividad del equipo Brain2Power",
            "Generar briefing diario consolidado",
        ],
        "endpoints": {
            "GET /betty/tasks": "Tareas pendientes en Outlook / To Do",
            "GET /betty/energy-events": "Eventos de energía detectados en email y calendario",
            "GET /betty/deliverables": "Entregables del proyecto en SharePoint",
            "GET /betty/team-activity": "Actividad reciente del equipo",
            "GET /betty/briefing": "Briefing diario completo",
        },
        "env_required": ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "USER_EMAIL", "SHAREPOINT_SITE"],
        "env_social": ["FB_PAGE_TOKEN", "FB_PAGE_ID", "IG_BUSINESS_ID", "LI_ACCESS_TOKEN", "LI_ORGANIZATION_ID"],
    }


# ═══════════════════════════════════════════════════════════════════════════════
# BLOQUE REDES SOCIALES — Facebook, Instagram Business, LinkedIn
# ═══════════════════════════════════════════════════════════════════════════════

FB_GRAPH = "https://graph.facebook.com/v19.0"
LI_API   = "https://api.linkedin.com/v2"


class SocialMetrics(BaseModel):
    platform: str
    status: str          # "ok" | "token_expired" | "not_configured" | "error"
    token_valid: bool
    token_expires: Optional[str]
    followers: Optional[int]
    impressions_30d: Optional[int]
    engagement_rate: Optional[float]
    reach_30d: Optional[int]
    top_post: Optional[str]
    raw: Optional[dict]


class FbTokenExchange(BaseModel):
    short_lived_token: str   # token de ~1h obtenido en Meta for Developers


# ── Facebook helpers ──────────────────────────────────────────────────────────

def _fb_get(path: str, params: dict | None = None) -> dict:
    token = os.environ.get("FB_PAGE_TOKEN", "")
    if not token:
        raise HTTPException(503, "FB_PAGE_TOKEN no configurado. Ver /betty/config para instrucciones.")
    r = requests.get(
        f"{FB_GRAPH}/{path}",
        params={"access_token": token, **(params or {})},
        timeout=15,
    )
    data = r.json()
    if "error" in data:
        raise HTTPException(502, f"Facebook API: {data['error'].get('message', data['error'])}")
    return data


def _fb_check_token() -> dict:
    """Devuelve info del token de página actual via debug_token."""
    app_token = f"{os.environ.get('FB_APP_ID','')}|{os.environ.get('FB_APP_SECRET','')}"
    token = os.environ.get("FB_PAGE_TOKEN", "")
    if not token:
        return {"valid": False, "reason": "FB_PAGE_TOKEN no configurado"}
    if not app_token.startswith("|") and "|" in app_token and len(app_token) > 2:
        r = requests.get(
            f"{FB_GRAPH}/debug_token",
            params={"input_token": token, "access_token": app_token},
            timeout=10,
        )
        info = r.json().get("data", {})
        return {
            "valid": info.get("is_valid", False),
            "expires_at": datetime.fromtimestamp(info.get("expires_at", 0), tz=timezone.utc).isoformat() if info.get("expires_at") else "never",
            "scopes": info.get("scopes", []),
            "app_id": info.get("app_id"),
        }
    # Sin app secret, hacer llamada mínima
    try:
        _fb_get("me", {"fields": "id,name"})
        return {"valid": True, "expires_at": "unknown — configura FB_APP_ID y FB_APP_SECRET para detalles"}
    except HTTPException:
        return {"valid": False, "reason": "Token inválido o expirado"}


# ── LinkedIn helpers ──────────────────────────────────────────────────────────

def _li_get(path: str, params: dict | None = None) -> dict:
    token = os.environ.get("LI_ACCESS_TOKEN", "")
    if not token:
        raise HTTPException(503, "LI_ACCESS_TOKEN no configurado. Ver /betty/config.")
    r = requests.get(
        f"{LI_API}/{path}",
        headers={
            "Authorization": f"Bearer {token}",
            "LinkedIn-Version": "202401",
            "X-Restli-Protocol-Version": "2.0.0",
        },
        params=params or {},
        timeout=15,
    )
    if r.status_code == 401:
        raise HTTPException(401, "LinkedIn token inválido o expirado. Renovar en LinkedIn Developer Portal.")
    if not r.ok:
        raise HTTPException(r.status_code, f"LinkedIn API error {r.status_code}: {r.text[:200]}")
    return r.json()


# ── Endpoint: métricas Facebook Page Insights ─────────────────────────────────

@router.get(
    "/betty/social/facebook",
    summary="Métricas Facebook Page Insights (Brain2Power)",
    operation_id="bettyFacebook",
    tags=["betty-social"],
    response_model=SocialMetrics,
)
def get_facebook_metrics():
    """
    Lee métricas reales de la página de Facebook vía Graph API.
    Requiere FB_PAGE_TOKEN (token de página de larga duración) y FB_PAGE_ID.
    """
    page_id = os.environ.get("FB_PAGE_ID", "")
    if not page_id:
        return SocialMetrics(
            platform="facebook", status="not_configured", token_valid=False,
            token_expires=None, followers=None, impressions_30d=None,
            engagement_rate=None, reach_30d=None, top_post=None, raw=None,
        )

    token_info = _fb_check_token()
    if not token_info.get("valid", False):
        return SocialMetrics(
            platform="facebook", status="token_expired", token_valid=False,
            token_expires=None, followers=None, impressions_30d=None,
            engagement_rate=None, reach_30d=None, top_post=None,
            raw={"token_info": token_info, "fix": "Usa POST /betty/social/facebook/refresh-token"},
        )

    try:
        # Seguidores y fans
        page_data = _fb_get("me", {"fields": "name,fan_count,followers_count"})
        followers = page_data.get("followers_count") or page_data.get("fan_count")

        # Insights: impresiones y alcance últimos 30 días
        since = int((datetime.now(timezone.utc) - timedelta(days=30)).timestamp())
        until = int(datetime.now(timezone.utc).timestamp())
        insights = _fb_get(
            "me/insights",
            {
                "metric": "page_impressions,page_reach,page_engaged_users,page_post_engagements",
                "period": "month",
                "since": since,
                "until": until,
            },
        )
        metrics_map: dict[str, int] = {}
        for item in insights.get("data", []):
            vals = item.get("values", [])
            total = sum(v.get("value", 0) for v in vals if isinstance(v.get("value"), (int, float)))
            metrics_map[item["name"]] = int(total)

        impressions = metrics_map.get("page_impressions")
        reach       = metrics_map.get("page_reach")
        engaged     = metrics_map.get("page_engaged_users")
        eng_rate    = round(engaged / reach * 100, 2) if reach and engaged else None

        # Post más reciente
        posts = _fb_get("me/posts", {"fields": "message,created_time", "$top": "1"})
        top_post = None
        if posts.get("data"):
            p = posts["data"][0]
            top_post = f"{p.get('created_time','')[:10]}: {p.get('message','')[:80]}"

        return SocialMetrics(
            platform="facebook", status="ok", token_valid=True,
            token_expires=token_info.get("expires_at"),
            followers=followers, impressions_30d=impressions,
            engagement_rate=eng_rate, reach_30d=reach,
            top_post=top_post, raw=metrics_map,
        )
    except HTTPException:
        raise
    except Exception as exc:
        log.error("Facebook metrics error: %s", exc)
        raise HTTPException(502, str(exc))


# ── Endpoint: renovar token de Facebook ──────────────────────────────────────

@router.post(
    "/betty/social/facebook/refresh-token",
    summary="Renovar token de Facebook (short → long-lived de 60 días)",
    operation_id="bettyFacebookRefreshToken",
    tags=["betty-social"],
)
def facebook_refresh_token(body: FbTokenExchange):
    """
    Intercambia un token de corta duración (~1h) por uno de larga duración (60 días).

    **Pasos para obtener el short_lived_token:**
    1. Ve a https://developers.facebook.com/tools/explorer/
    2. Selecciona tu App y tu Página Brain2Power
    3. Genera un token con permisos: pages_show_list, pages_read_engagement,
       pages_read_user_content, read_insights
    4. Pega el token en este endpoint — Betty lo convierte a 60 días
    5. Copia el long_lived_token resultante y ponlo en la variable FB_PAGE_TOKEN

    Requiere: FB_APP_ID y FB_APP_SECRET configurados.
    """
    app_id     = os.environ.get("FB_APP_ID", "")
    app_secret = os.environ.get("FB_APP_SECRET", "")

    if not app_id or not app_secret:
        raise HTTPException(
            503,
            "Faltan FB_APP_ID y FB_APP_SECRET. "
            "Encuéntralos en https://developers.facebook.com → tu app → Configuración básica.",
        )

    r = requests.get(
        f"{FB_GRAPH}/oauth/access_token",
        params={
            "grant_type": "fb_exchange_token",
            "client_id": app_id,
            "client_secret": app_secret,
            "fb_exchange_token": body.short_lived_token,
        },
        timeout=15,
    )
    data = r.json()
    if "error" in data:
        raise HTTPException(400, f"Facebook rechazó el token: {data['error'].get('message')}")

    long_token = data.get("access_token", "")
    expires_in = data.get("expires_in", 0)
    expires_dt = (datetime.now(timezone.utc) + timedelta(seconds=expires_in)).isoformat() if expires_in else "~60 días"

    # Obtener token de página (el long-lived user token → page token)
    pages_r = requests.get(
        f"{FB_GRAPH}/me/accounts",
        params={"access_token": long_token},
        timeout=15,
    ).json()

    page_tokens = [
        {"page_id": p["id"], "page_name": p["name"], "page_token": p["access_token"]}
        for p in pages_r.get("data", [])
    ]

    return {
        "status": "ok",
        "long_lived_user_token": long_token,
        "expires": expires_dt,
        "page_tokens": page_tokens,
        "next_steps": [
            "1. Copia el page_token de Brain2Power de la lista page_tokens",
            "2. Guárdalo en la variable de entorno FB_PAGE_TOKEN de tu servidor",
            "3. Guarda el page_id en FB_PAGE_ID",
            "4. Repite este proceso en ~55 días (antes de que expire)",
        ],
    }


# ── Endpoint: métricas Instagram Business ────────────────────────────────────

@router.get(
    "/betty/social/instagram",
    summary="Métricas Instagram Business (Brain2Power)",
    operation_id="bettyInstagram",
    tags=["betty-social"],
    response_model=SocialMetrics,
)
def get_instagram_metrics():
    """
    Lee métricas reales de la cuenta de Instagram Business vía Facebook Graph API.
    Requiere FB_PAGE_TOKEN (mismo token de página) y IG_BUSINESS_ID.
    """
    ig_id = os.environ.get("IG_BUSINESS_ID", "")
    if not ig_id:
        return SocialMetrics(
            platform="instagram", status="not_configured", token_valid=False,
            token_expires=None, followers=None, impressions_30d=None,
            engagement_rate=None, reach_30d=None, top_post=None,
            raw={"fix": "Configura IG_BUSINESS_ID. Puedes encontrarlo en FB_PAGE_ID → /instagram_accounts"},
        )

    token_info = _fb_check_token()
    if not token_info.get("valid", False):
        return SocialMetrics(
            platform="instagram", status="token_expired", token_valid=False,
            token_expires=None, followers=None, impressions_30d=None,
            engagement_rate=None, reach_30d=None, top_post=None,
            raw={"fix": "Renueva el token con POST /betty/social/facebook/refresh-token"},
        )

    try:
        # Datos básicos cuenta
        ig_data = _fb_get(ig_id, {"fields": "username,followers_count,media_count,biography"})
        followers = ig_data.get("followers_count")

        # Insights de cuenta (últimos 30 días)
        since = int((datetime.now(timezone.utc) - timedelta(days=30)).timestamp())
        until = int(datetime.now(timezone.utc).timestamp())
        insights = _fb_get(
            f"{ig_id}/insights",
            {
                "metric": "impressions,reach,profile_views,follower_count",
                "period": "day",
                "since": since,
                "until": until,
            },
        )
        metrics_map: dict[str, int] = {}
        for item in insights.get("data", []):
            total = sum(v.get("value", 0) for v in item.get("values", []) if isinstance(v.get("value"), (int, float)))
            metrics_map[item["name"]] = int(total)

        impressions = metrics_map.get("impressions")
        reach       = metrics_map.get("reach")

        # Top media reciente (engagement)
        media = _fb_get(
            f"{ig_id}/media",
            {"fields": "caption,timestamp,like_count,comments_count,media_type", "limit": "5"},
        )
        top_post = None
        best_eng = 0
        for m in media.get("data", []):
            eng = (m.get("like_count", 0) or 0) + (m.get("comments_count", 0) or 0)
            if eng > best_eng:
                best_eng = eng
                cap = (m.get("caption") or "")[:60]
                top_post = f"{m.get('timestamp','')[:10]} [{m.get('media_type','')}] {cap} — {eng} interacciones"

        eng_rate = round(best_eng / followers * 100, 2) if followers and best_eng else None

        return SocialMetrics(
            platform="instagram", status="ok", token_valid=True,
            token_expires=token_info.get("expires_at"),
            followers=followers, impressions_30d=impressions,
            engagement_rate=eng_rate, reach_30d=reach,
            top_post=top_post, raw=metrics_map,
        )
    except HTTPException:
        raise
    except Exception as exc:
        log.error("Instagram metrics error: %s", exc)
        raise HTTPException(502, str(exc))


# ── Endpoint: métricas LinkedIn Organization ─────────────────────────────────

@router.get(
    "/betty/social/linkedin",
    summary="Métricas LinkedIn Organization (Brain2Power)",
    operation_id="bettyLinkedin",
    tags=["betty-social"],
    response_model=SocialMetrics,
)
def get_linkedin_metrics():
    """
    Lee seguidores y estadísticas de la página de organización en LinkedIn.
    Requiere LI_ACCESS_TOKEN (OAuth 2.0) y LI_ORGANIZATION_ID (urn:li:organization:XXXXX).
    """
    org_id = os.environ.get("LI_ORGANIZATION_ID", "")
    token  = os.environ.get("LI_ACCESS_TOKEN", "")

    if not org_id or not token:
        return SocialMetrics(
            platform="linkedin", status="not_configured", token_valid=False,
            token_expires=None, followers=None, impressions_30d=None,
            engagement_rate=None, reach_30d=None, top_post=None,
            raw={
                "fix": [
                    "1. Ve a https://www.linkedin.com/developers/apps → crea/selecciona tu app",
                    "2. Permisos necesarios: r_organization_social, rw_organization_admin",
                    "3. Copia el Access Token a LI_ACCESS_TOKEN",
                    "4. El Organization ID está en la URL de tu página LinkedIn: linkedin.com/company/XXXXX → LI_ORGANIZATION_ID=urn:li:organization:XXXXX",
                ]
            },
        )

    try:
        # Verificar token
        me = _li_get("me", {"projection": "(id)"})
        token_valid = "id" in me

        # Seguidores de la organización
        org_urn = org_id if org_id.startswith("urn:li:") else f"urn:li:organization:{org_id}"
        followers_data = _li_get(
            "organizationalEntityFollowerStatistics",
            {"q": "organizationalEntity", "organizationalEntity": org_urn},
        )
        followers = None
        for elem in followers_data.get("elements", []):
            followers = elem.get("followerCountsByAssociationType", [{}])[0].get("followerCounts", {}).get("organicFollowerCount")
            break

        # Share statistics (últimas 4 semanas)
        now_ms   = int(datetime.now(timezone.utc).timestamp() * 1000)
        start_ms = int((datetime.now(timezone.utc) - timedelta(days=30)).timestamp() * 1000)
        share_stats = _li_get(
            "organizationalEntityShareStatistics",
            {
                "q": "organizationalEntity",
                "organizationalEntity": org_urn,
                "timeIntervals.timeGranularityType": "MONTH",
                "timeIntervals.timeRange.start": start_ms,
                "timeIntervals.timeRange.end": now_ms,
            },
        )
        impressions = 0
        clicks      = 0
        for elem in share_stats.get("elements", []):
            stats = elem.get("totalShareStatistics", {})
            impressions += stats.get("impressionCount", 0)
            clicks      += stats.get("clickCount", 0)

        eng_rate = round(clicks / impressions * 100, 2) if impressions and clicks else None

        return SocialMetrics(
            platform="linkedin", status="ok", token_valid=token_valid,
            token_expires="~60 días desde emisión — revisar en LinkedIn Developer Portal",
            followers=followers, impressions_30d=impressions or None,
            engagement_rate=eng_rate, reach_30d=None,
            top_post=None,
            raw={"clicks_30d": clicks, "impressions_30d": impressions},
        )
    except HTTPException:
        raise
    except Exception as exc:
        log.error("LinkedIn metrics error: %s", exc)
        raise HTTPException(502, str(exc))


# ── Endpoint: dashboard social completo ──────────────────────────────────────

@router.get(
    "/betty/social",
    summary="Dashboard completo de redes sociales Brain2Power",
    operation_id="bettySocialDashboard",
    tags=["betty-social"],
)
def social_dashboard():
    """Resumen unificado de Facebook, Instagram y LinkedIn con alertas de token."""
    results = {}
    alerts  = []

    for platform, fn in [("facebook", get_facebook_metrics), ("instagram", get_instagram_metrics), ("linkedin", get_linkedin_metrics)]:
        try:
            m = fn()
            results[platform] = m.model_dump()
            if m.status == "token_expired":
                alerts.append(f"🔴 {platform.upper()}: token expirado — renovar urgente")
            elif m.status == "not_configured":
                alerts.append(f"⚠️  {platform.upper()}: credenciales no configuradas")
        except Exception as exc:
            results[platform] = {"status": "error", "detail": str(exc)}
            alerts.append(f"🔴 {platform.upper()}: error — {exc}")

    summary = {
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "alerts": alerts,
        "platforms": results,
        "kpis_vs_plan": {
            "objetivo_publicaciones": "15 total (Plan Difusión v5)",
            "objetivo_impresiones": ">7.500",
            "objetivo_engagement": ">3%",
            "instagram_engagement_real": results.get("instagram", {}).get("engagement_rate"),
            "linkedin_seguidores_real": results.get("linkedin", {}).get("followers"),
            "nota": "Instagram engagement histórico 8% (27/05/2026) — muy por encima del objetivo",
        },
    }
    return summary


# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE BETTY — estado de credenciales y guía de setup
# ═══════════════════════════════════════════════════════════════════════════════

_ENV_VARS = {
    "Microsoft 365 / Graph API": {
        "TENANT_ID":       "ID del tenant de Azure AD (portal.azure.com → Azure Active Directory)",
        "CLIENT_ID":       "App registration Client ID (Azure AD → App registrations)",
        "CLIENT_SECRET":   "Client secret de la app (Azure AD → App registrations → Certificates & secrets)",
        "USER_EMAIL":      "Email del usuario cuyo buzón y calendario leerá Betty",
        "SHAREPOINT_SITE": "Host+ruta del sitio SharePoint (ej: oceanicanarias.sharepoint.com:/sites/brain2power)",
    },
    "Facebook / Instagram Business": {
        "FB_APP_ID":       "ID de tu app Meta (developers.facebook.com → tu app → Configuración básica)",
        "FB_APP_SECRET":   "App Secret Meta (misma pantalla que FB_APP_ID)",
        "FB_PAGE_TOKEN":   "Token de página de larga duración (60 días). Obtener con POST /betty/social/facebook/refresh-token",
        "FB_PAGE_ID":      "ID numérico de la página Facebook de Brain2Power (Configuración página → Información de la página)",
        "IG_BUSINESS_ID":  "ID de cuenta de Instagram Business. Encontrar con GET /{FB_PAGE_ID}?fields=instagram_accounts",
    },
    "LinkedIn": {
        "LI_ACCESS_TOKEN":    "OAuth 2.0 Access Token (linkedin.com/developers → tu app → OAuth)",
        "LI_ORGANIZATION_ID": "ID de organización LinkedIn (ej: urn:li:organization:12345678 — URL de la página de empresa)",
    },
}


@router.get(
    "/betty/config",
    summary="Estado de configuración de Betty — todas las credenciales",
    operation_id="bettyConfig",
    tags=["betty-config"],
)
def get_config_status():
    """
    Muestra qué variables de entorno están configuradas (sin revelar los valores)
    y proporciona instrucciones exactas para las que faltan.
    """
    status: dict[str, Any] = {}
    missing: list[str] = []
    configured: list[str] = []

    for group, vars_dict in _ENV_VARS.items():
        group_status: dict[str, Any] = {}
        for var, description in vars_dict.items():
            val = os.environ.get(var, "")
            ok  = bool(val)
            group_status[var] = {
                "configured": ok,
                "preview": f"{val[:4]}***{val[-4:]}" if len(val) > 8 else ("✅ set" if ok else "❌ missing"),
                "description": description,
            }
            if ok:
                configured.append(var)
            else:
                missing.append(var)
        status[group] = group_status

    fb_token_check = {}
    if os.environ.get("FB_PAGE_TOKEN"):
        try:
            fb_token_check = _fb_check_token()
        except Exception:
            fb_token_check = {"valid": False, "reason": "Error al verificar"}

    return {
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "summary": {
            "total_vars": len(configured) + len(missing),
            "configured": len(configured),
            "missing": len(missing),
            "health": "🟢 COMPLETO" if not missing else f"🟡 PARCIAL — faltan {len(missing)} variables" if len(missing) < 5 else "🔴 INCOMPLETO",
        },
        "facebook_token_status": fb_token_check,
        "variables_by_group": status,
        "missing_vars": missing,
        "configured_vars": configured,
        "quick_setup_guide": {
            "step_1": "Añade las variables de entorno a tu Azure App Service: Portal Azure → tu App Service → Configuración → Configuración de la aplicación",
            "step_2": "Para Facebook: ve a https://developers.facebook.com/tools/explorer/ y genera un token, luego llama POST /betty/social/facebook/refresh-token",
            "step_3": "Para Instagram Business ID: GET https://graph.facebook.com/v19.0/{FB_PAGE_ID}?fields=instagram_accounts&access_token={FB_PAGE_TOKEN}",
            "step_4": "Para LinkedIn: https://www.linkedin.com/developers/apps → genera token con permisos r_organization_social",
            "step_5": "Verifica con GET /betty/config — todos los campos deben aparecer en configured_vars",
        },
    }


# Standalone: solo registrar si se ejecuta betty.py directamente
if __name__ == "__main__":
    app.include_router(router)
