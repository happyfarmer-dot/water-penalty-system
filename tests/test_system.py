#!/usr/bin/env python3
import base64
import io
import sys
from pathlib import Path
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import AnalyzeRequest, TicketItem, TicketRequest, analyze_excel, analyze_sample, generate_ticket, get_clauses, index, preview_ticket
from excel_parser import parse_deficiency_excel
from matcher import load_clauses_data, match_deficiencies
from ticket_generator import calculate_total, generate_ticket_docx, int_to_chinese_currency

SAMPLE_EXCEL = Path("/home/antigravity/督導紀錄表-1151101(烏日南屯北段).xlsx")


def test_clauses_loaded():
    data = load_clauses_data()
    assert len(data["clauses"]) >= 190
    ids = {c["id"] for c in data["clauses"]}
    for must_have in ["1.1", "1.2", "2.18", "2.35", "2.53", "4.7", "4.8", "5.1", "5.4", "5.8", "6.2", "6.5", "7.4", "12.6", "13.2", "13.10", "天條1", "天條4"]:
        assert must_have in ids, f"Missing clause {must_have}"


def test_chinese_currency():
    assert int_to_chinese_currency(199000) == "壹拾玖萬玖仟元整"
    assert int_to_chinese_currency(205000) == "貳拾萬伍仟元整"
    assert int_to_chinese_currency(200500) == "貳拾萬零伍佰元整"
    assert int_to_chinese_currency(5000) == "伍仟元整"
    assert int_to_chinese_currency(300000) == "參拾萬元整"
    assert int_to_chinese_currency(0) == "零元整"


def test_excel_parse_and_match():
    parsed = parse_deficiency_excel(SAMPLE_EXCEL)
    meta = parsed["meta"]
    assert "烏日至南屯送水管" in meta["project_name"]
    assert "TP-113-0401-0002" in meta["project_code"]
    assert "TD-113-0401-0002" in meta["project_code"]
    assert meta["contractor"] == "宥林工程有限公司"
    assert meta["inspection_date"]["roc_year"] == "115"
    assert meta["inspection_date"]["is_weekend"] is True
    assert len(parsed["deficiencies"]) == 22

    matched = match_deficiencies(parsed)
    items_by_seq = {it["seq"]: it for it in matched["items"]}
    assert items_by_seq["1"]["clause_id"] == "13.10"
    assert items_by_seq["2"]["clause_id"] == "13.2"
    assert items_by_seq["3"]["clause_id"] == "2.35"
    assert items_by_seq["3"]["count"] == 2
    assert items_by_seq["4"]["clause_id"] == "2.35"
    assert items_by_seq["5"]["clause_id"] == "12.6"
    assert items_by_seq["6"]["clause_id"] == "12.1"
    assert items_by_seq["7"]["clause_id"] == "2.18"
    assert items_by_seq["7"]["count"] == 3
    assert items_by_seq["8"]["clause_id"] == "4.8"
    assert items_by_seq["9"]["clause_id"] == "5.8"
    assert items_by_seq["10"]["clause_id"] == "2.53"
    assert items_by_seq["11"]["clause_id"] == "4.7"
    assert items_by_seq["12"]["clause_id"] == "7.4"
    assert items_by_seq["13"]["clause_id"] == "6.2"
    assert items_by_seq["15"]["clause_id"] == "6.5"
    assert items_by_seq["17"]["clause_id"] == "5.1"
    assert any(s["clause_id"] == "1.2" for s in matched["suggestions"])


def test_generate_docx_roundtrip():
    payload = {
        "project_name": "臺中雲林水源調度-烏日至南屯送水管(北段)",
        "project_code": "TP-113-0401-0002\nTD-113-0401-0002",
        "contractor": "宥林工程有限公司",
        "safety_officer": "陳儀芝、張嘉勇",
        "occur_time": "中華民國115年11月01日10時00分",
        "occur_location": "工地(節點17-1)",
        "violation_summary": "",
        "evidence_note": "督導紀錄表-1151101",
        "items": [
            {"clause_id": "2.18", "clause_text": "承攬商於上工前，未依規定辦理1.工作場所簽到4.危害因素告知5.工作場所簽退。", "unit_fine": 3000, "count": 3, "repeat_level": 0, "selected": True},
            {"clause_id": "4.8", "clause_text": "未設置過捲預防裝置或警報裝置、制動器、離合器及其他安全裝置無法使用。", "unit_fine": 20000, "count": 1, "repeat_level": 1, "selected": True},
        ],
    }
    assert calculate_total(payload["items"]) == 9000 + 40000
    docx_bytes = generate_ticket_docx(payload)
    doc = Document(io.BytesIO(docx_bytes))
    tbl = doc.tables[0]
    assert "臺中雲林水源調度" in tbl.rows[0].cells[1].text
    assert "TP-113-0401-0002" in tbl.rows[0].cells[3].text
    assert "49,000元" in tbl.rows[2].cells[1].text
    assert "重複2倍=40,000元" in tbl.rows[2].cells[1].text
    assert "肆萬玖仟元整" in tbl.rows[3].cells[1].text
    assert "督導紀錄表-1151101" in tbl.rows[4].cells[1].text


def test_api_handlers():
    assert Path(index().path).exists()
    assert len(get_clauses()["clauses"]) >= 190

    data = analyze_sample()
    assert len(data["items"]) == 22

    b64 = base64.b64encode(SAMPLE_EXCEL.read_bytes()).decode("ascii")
    up_data = analyze_excel(AnalyzeRequest(filename="test.xlsx", content_base64=b64))
    assert len(up_data["items"]) == 22

    req = TicketRequest(
        project_name=data["meta"]["project_name"],
        project_code=data["meta"]["project_code"],
        contractor=data["meta"]["contractor"],
        safety_officer=data["meta"]["safety_officer"],
        occur_time="中華民國115年11月01日10時00分",
        occur_location="工地(節點17-1)",
        items=[
            TicketItem(
                clause_id=it["clause_id"],
                clause_text=it["clause_text"],
                unit_fine=it["unit_fine"],
                count=it["count"],
                repeat_level=it["repeat_level"],
                selected=True,
            )
            for it in data["items"]
        ],
        output_filename="test_output_ticket.docx",
    )
    pv = preview_ticket(req)
    assert pv["total"] > 0

    import json
    gen_resp = generate_ticket(req)
    gen_json = json.loads(gen_resp.body.decode("utf-8"))
    assert gen_json["filename"] == "test_output_ticket.docx"
    assert Path(gen_json["saved_path"]).exists()
    Path(gen_json["saved_path"]).unlink(missing_ok=True)


if __name__ == "__main__":
    test_clauses_loaded()
    test_chinese_currency()
    test_excel_parse_and_match()
    test_generate_docx_roundtrip()
    test_api_handlers()
    print("ALL TESTS PASSED!")
