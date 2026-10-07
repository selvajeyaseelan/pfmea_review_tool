import os
import json
import asyncio
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pfmea_auditor import (
    audit_pfmea_generator, 
    AIAG_SEVERITY_STANDARDS, 
    AIAG_OCCURRENCE_STANDARDS, 
    AIAG_DETECTION_STANDARDS
)

app = FastAPI(title="AI based Intelligent PFMEA Auditor")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
STATIC_DIR = "static"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

async def stream_audit_sse(file_path: str, target_priority: int = 1, det_threshold: float = 4.0):
    """
    Asynchronous event generator reading from audit_pfmea_generator and yielding SSE chunks.
    """
    try:
        for event in audit_pfmea_generator(file_path, target_priority=target_priority, det_threshold=det_threshold):
            payload = f"data: {json.dumps(event)}\n\n"
            yield payload
            await asyncio.sleep(0.01)
    except Exception as e:
        error_event = {"type": "error", "message": f"Server processing error: {str(e)}"}
        yield f"data: {json.dumps(error_event)}\n\n"

@app.post("/api/audit-upload")
async def audit_upload(
    file: UploadFile = File(...), 
    target_priority: int = Form(1), 
    det_threshold: float = Form(4.0)
):
    if not file.filename.lower().endswith(('.xlsx', '.xlsm', '.xltx')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx) are supported.")
    
    dest_path = os.path.join(UPLOAD_DIR, f"upload_{file.filename}")
    content = await file.read()
    with open(dest_path, "wb") as f:
        f.write(content)
        
    return StreamingResponse(
        stream_audit_sse(dest_path, target_priority=target_priority, det_threshold=det_threshold),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )

@app.get("/api/audit-sample")
async def audit_sample(target_priority: int = 1, det_threshold: float = 4.0):
    sample_file = "pfmea_org.xlsx"
    if not os.path.exists(sample_file):
        raise HTTPException(status_code=404, detail="Sample pfmea_org.xlsx not found on server.")
    return StreamingResponse(
        stream_audit_sse(sample_file, target_priority=target_priority, det_threshold=det_threshold),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )

@app.get("/api/download/excel")
async def download_excel():
    path = "pfmea_audited.xlsx"
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Audited Excel file not found. Please run an audit first.")
    return FileResponse(
        path, 
        filename="pfmea_audited.xlsx", 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/api/download/html")
async def download_html():
    path = "audited_pfmea.html"
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Audited HTML report not found. Please run an audit first.")
    return FileResponse(path, filename="audited_pfmea.html", media_type="text/html")

@app.get("/api/standards")
async def get_standards():
    return {
        "severity": AIAG_SEVERITY_STANDARDS,
        "occurrence": AIAG_OCCURRENCE_STANDARDS,
        "detection": AIAG_DETECTION_STANDARDS
    }

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>AI based Intelligent PFMEA Auditor - Initializing UI...</h1>")

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
