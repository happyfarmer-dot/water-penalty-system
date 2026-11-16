#!/usr/bin/env python3
"""
台灣自來水公司承攬商違規缺失辨識與罰單生成系統 - Web API 伺服器
"""

import base64
import datetime
import re
from pathlib import Path
from typing import Optional
from urllib.parse import quote

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from excel_parser import parse_deficiency_excel
from matcher import load_clauses_data, match_deficiencies
from ticket_generator import (
    calculate_total,
    format_violation_line,
    generate_ticket_docx,
    int_to_chinese_currency,
)

BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
SAMPLE_EXCEL = Path("/home/antigravity/督導紀錄表-1151101(烏日南屯北段).xlsx")

app = FastAPI(title="承攬商違規缺失辨識與罰單生成系統")
app.mount("/data", StaticFiles(directory=str(BASE_DIR / "data")), name="data")
app.mount("/templates", StaticFiles(directory=str(BASE_DIR / "templates")), name="templates")

CLAUSES_CACHE = load_clauses_data()


class AnalyzeRequest(BaseModel):
    filename: str
    content_base64: str


class TicketItem(BaseModel):
    clause_id: str = ""
    clause_text: str = ""
    unit_fine: int = 0
    count: int = 1
    repeat_level: int = 0
    selected: bool = True
    custom_line: Optional[str] = None


class TicketRequest(BaseModel):
    project_name: str = ""
    project_code: str = ""
    contractor: str = ""
    safety_officer: str = ""
    occur_time: str = ""
    occur_location: str = ""
    violation_summary: str = ""
    evidence_note: str = ""
    items: list[TicketItem] = []
    override_total: Optional[int] = None
    output_filename: Optional[str] = None


@app.get("/")
def index():
    return FileResponse(str(BASE_DIR / "index.html"))


@app.get("/style.css")
def serve_css():
    return FileResponse(str(BASE_DIR / "style.css"), media_type="text/css")


@app.get("/app.js")
def serve_app_js():
    return FileResponse(str(BASE_DIR / "app.js"), media_type="application/javascript")


@app.get("/engine.js")
def serve_engine_js():
    return FileResponse(str(BASE_DIR / "engine.js"), media_type="application/javascript")


@app.get("/api/clauses")
def get_clauses():
    return CLAUSES_CACHE


@app.post("/api/analyze")
def analyze_excel(req: AnalyzeRequest):
    if not req.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="請上傳 .xlsx 格式的督導紀錄表")
    try:
        raw_bytes = base64.b64decode(req.content_base64)
        parsed = parse_deficiency_excel(raw_bytes)
        result = match_deficiencies(parsed, CLAUSES_CACHE)
        result["source_filename"] = req.filename
        return result
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"解析缺失單失敗：{e}")


@app.post("/api/analyze-sample")
def analyze_sample():
    if not SAMPLE_EXCEL.exists():
        raise HTTPException(status_code=404, detail="找不到範例督導紀錄表檔案")
    parsed = parse_deficiency_excel(SAMPLE_EXCEL)
    result = match_deficiencies(parsed, CLAUSES_CACHE)
    result["source_filename"] = SAMPLE_EXCEL.name
    return result


def _to_dict(model):
    return model.model_dump() if hasattr(model, "model_dump") else model.dict()


@app.post("/api/preview")
def preview_ticket(req: TicketRequest):
    items_dict = [_to_dict(it) for it in req.items]
    selected = [it for it in items_dict if it.get("selected", True)]
    total = req.override_total if req.override_total is not None else calculate_total(selected)
    chinese_total = int_to_chinese_currency(int(total))
    lines = [format_violation_line(it) for it in selected]
    return {
        "total": int(total),
        "chinese_total": chinese_total,
        "lines": lines,
        "selected_count": len(selected),
    }


@app.post("/api/generate")
def generate_ticket(req: TicketRequest):
    payload = _to_dict(req)
    payload["items"] = [_to_dict(it) for it in req.items]
    docx_bytes = generate_ticket_docx(payload)

    selected = [it for it in payload["items"] if it.get("selected", True)]
    total = req.override_total if req.override_total is not None else calculate_total(selected)
    chinese_total = int_to_chinese_currency(int(total))

    # 檔名整理
    base_name = req.output_filename
    if not base_name:
        m_date = re.search(r"(\d{2,3})年(\d{1,2})月(\d{1,2})日", req.occur_time)
        date_tag = f"{m_date.group(1)}{int(m_date.group(2)):02d}{int(m_date.group(3)):02d}" if m_date else datetime.datetime.now().strftime("%Y%m%d")
        base_name = f"(_罰單_)_違規稽查表{date_tag}.docx"
    if not base_name.lower().endswith(".docx"):
        base_name += ".docx"

    safe_name = re.sub(r'[\\/:*?"<>|]', "_", base_name)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUTPUT_DIR / safe_name
    out_path.write_bytes(docx_bytes)

    return JSONResponse({
        "filename": safe_name,
        "saved_path": str(out_path),
        "total": int(total),
        "chinese_total": chinese_total,
        "content_base64": base64.b64encode(docx_bytes).decode("ascii"),
    })


@app.get("/api/download/{filename}")
def download_generated(filename: str):
    safe_name = Path(filename).name
    target = OUTPUT_DIR / safe_name
    if not target.exists():
        raise HTTPException(status_code=404, detail="檔案不存在")
    encoded = quote(safe_name)
    return Response(
        content=target.read_bytes(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"},
    )
