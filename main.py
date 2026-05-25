"""
main.py — API FastAPI para Bot Conexia SEROS
Corre en PikaPods, expone endpoints para el frontend en Vercel.
"""
import os, subprocess, sys, threading, uuid
from pathlib import Path
from datetime import datetime
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

app = FastAPI(title="Conexia SEROS API")

# ── CORS: permitir llamadas desde Vercel ──────────────────────────────────────
# Reemplazar con tu URL real de Vercel antes de deployar
VERCEL_URL = os.getenv("VERCEL_URL", "*")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[VERCEL_URL, "http://localhost:3000", "http://localhost:5500"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Rutas ─────────────────────────────────────────────────────────────────────
UPLOAD_DIR = Path("/app/uploads")
OUTPUT_DIR = Path("/app/output")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
EXCEL_PATH = UPLOAD_DIR / "credenciales_seros.xlsx"

# ── Estado global de ejecución ────────────────────────────────────────────────
estado = {
    "status":    "idle",      # idle | running | done | error
    "inicio":    None,
    "fin":       None,
    "log":       [],
    "ok":        0,
    "fallidos":  0,
    "ultima_carpeta": None,
}


def _run_bot():
    """Corre el bot en un thread separado y actualiza el estado global."""
    estado["status"]  = "running"
    estado["inicio"]  = datetime.now().isoformat()
    estado["log"]     = []
    estado["ok"]      = 0
    estado["fallidos"]= 0

    env = os.environ.copy()
    env["EXCEL_PATH"] = str(EXCEL_PATH)
    env["OUTPUT_DIR"] = str(OUTPUT_DIR)
    env["HEADLESS"]   = "true"

    try:
        proc = subprocess.Popen(
            [sys.executable, "conexia_bot.py"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace",
            env=env, bufsize=1
        )
        for line in proc.stdout:
            line = line.rstrip()
            estado["log"].append(line)
            if len(estado["log"]) > 500:      # evitar memoria infinita
                estado["log"] = estado["log"][-500:]
            if "✓" in line and ".html" in line:
                estado["ok"] += 1
            if "Credenciales inválidas" in line:
                estado["fallidos"] += 1

        proc.wait()
        estado["status"] = "done" if proc.returncode == 0 else "error"
    except Exception as e:
        estado["log"].append(f"ERROR CRÍTICO: {e}")
        estado["status"] = "error"

    estado["fin"] = datetime.now().isoformat()

    # Buscar subcarpeta más reciente para los archivos de descarga
    subcarpetas = sorted(
        [d for d in OUTPUT_DIR.iterdir() if d.is_dir() and d.name[:8].isdigit()],
        reverse=True
    ) if OUTPUT_DIR.exists() else []
    estado["ultima_carpeta"] = str(subcarpetas[0]) if subcarpetas else None


# ── Endpoints ─────────────────────────────────────────────────────────────────

@app.get("/")
def root():
    return {"app": "Conexia SEROS API", "version": "1.0"}


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    """Sube el Excel de credenciales al servidor."""
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(400, "Solo se aceptan archivos .xlsx")
    contenido = await file.read()
    EXCEL_PATH.write_bytes(contenido)
    # Contar registros
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
        ws = wb.active
        filas = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if any(c for c in r))
    except Exception:
        filas = "?"
    return {"ok": True, "registros": filas}


@app.post("/run")
def run():
    """Inicia la extracción. Solo una ejecución a la vez."""
    if estado["status"] == "running":
        raise HTTPException(409, "Ya hay una ejecución en curso")
    if not EXCEL_PATH.exists():
        raise HTTPException(400, "Primero subí el Excel de credenciales")
    t = threading.Thread(target=_run_bot, daemon=True)
    t.start()
    return {"ok": True, "mensaje": "Ejecución iniciada"}


@app.get("/status")
def get_status():
    """Estado actual de la ejecución (polling desde el frontend)."""
    return {
        "status":   estado["status"],
        "inicio":   estado["inicio"],
        "fin":      estado["fin"],
        "ok":       estado["ok"],
        "fallidos": estado["fallidos"],
        "log_tail": estado["log"][-50:],   # últimas 50 líneas
        "excel_ok": EXCEL_PATH.exists(),
    }


@app.get("/log")
def get_log():
    """Log completo de la última ejecución."""
    return {"log": estado["log"]}


def _get_ultima(patron: str) -> Path | None:
    if not estado["ultima_carpeta"]:
        return None
    carpeta = Path(estado["ultima_carpeta"])
    archivos = sorted(carpeta.glob(patron), reverse=True)
    return archivos[0] if archivos else None


@app.get("/download/excel")
def download_excel():
    f = _get_ultima("Consolidado_*.xlsx")
    if not f:
        raise HTTPException(404, "Excel no disponible aún")
    return FileResponse(str(f), filename=f.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/download/pdfs")
def download_pdfs():
    f = _get_ultima("PDFs_*.zip")
    if not f:
        raise HTTPException(404, "ZIP no disponible aún")
    return FileResponse(str(f), filename=f.name, media_type="application/zip")


@app.get("/download/fallidos")
def download_fallidos():
    f = _get_ultima("FALLIDOS_*.txt")
    if not f:
        raise HTTPException(404, "Archivo de fallidos no disponible")
    return FileResponse(str(f), filename=f.name, media_type="text/plain")
