import os
from fastapi import FastAPI

app = FastAPI(title="Forecast API", version="0.1")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/run-forecast")
def run_forecast(payload: dict | None = None):
    return {
        "status": "received",
        "env_check": {
            "TENANT_ID": bool(os.getenv("TENANT_ID")),
            "CLIENT_ID": bool(os.getenv("CLIENT_ID")),
            "CLIENT_SECRET": bool(os.getenv("CLIENT_SECRET")),
            "SHAREPOINT_SITE": os.getenv("SHAREPOINT_SITE"),
        }
    }
