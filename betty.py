"""
Betty — Secretaria inteligente de Brain2Power
=============================================
Agente que centraliza:
  1. Oportunidades de negocio: congresos, jornadas y foros de energía / eficiencia energética.
  2. Correo Outlook: tareas y asuntos pendientes.
  3. SharePoint: documentos, entregables y actividad del equipo Brain2Power.
  4. Briefing diario consolidado.

Integración Microsoft 365 vía MSAL (client_credentials).
Variables de entorno requeridas:
  TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_SITE, USER_EMAIL
"""

import os
import re
import json
import logging
from datetime import datetime, timedelta, timezone
from typing import Any, Optional

import requests
from fastapi import APIRouter, FastAPI, HTTPException, Query
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

@router.get("/tasks",
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

@router.get("/energy-events",
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

@router.get("/deliverables",
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

@router.get("/team-activity",
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

@router.get("/briefing",
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
    }


# Register router in standalone app
app.include_router(router)
