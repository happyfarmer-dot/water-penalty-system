#!/usr/bin/env python3
"""
依台灣自來水股份有限公司「承攬商違反職業安全衛生規定罰扣款通知單（TS-02）」範本，
生成完整格式一致之 .docx 罰單。
"""

import copy
import io
from pathlib import Path
from typing import Optional
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

TEMPLATE_PATH = Path(__file__).parent / "templates" / "ticket_template.docx"

DIGITS = "零壹貳參肆伍陸柒捌玖"
SMALL_UNITS = ["", "拾", "佰", "仟"]
BIG_UNITS = ["", "萬", "億", "兆"]


def int_to_chinese_currency(amount: int) -> str:
    """將整數金額轉換為新台幣國字大寫，例如 199000 -> 壹拾玖萬玖仟元整"""
    if amount <= 0:
        return "零元整"

    s = str(int(amount))
    # 由低位每4位切成一組
    groups = []
    while s:
        groups.append(s[-4:])
        s = s[:-4]

    parts = []
    need_zero_between_groups = False

    for g_idx in range(len(groups) - 1, -1, -1):
        grp = groups[g_idx]
        g_val = int(grp)
        if g_val == 0:
            if parts:
                need_zero_between_groups = True
            continue

        if need_zero_between_groups or (parts and len(grp) == 4 and grp[0] == "0"):
            parts.append("零")
            need_zero_between_groups = False

        grp_str = ""
        zero_flag = False
        n = len(grp)
        for i, ch in enumerate(grp):
            d = int(ch)
            u_idx = n - 1 - i
            if d == 0:
                if grp_str:
                    zero_flag = True
            else:
                if zero_flag:
                    grp_str += "零"
                    zero_flag = False
                grp_str += DIGITS[d] + SMALL_UNITS[u_idx]

        parts.append(grp_str + BIG_UNITS[g_idx])

    return "".join(parts) + "元整"


def _set_run_font(run, font_name: str = "標楷體", size_pt: int = 14, bold: Optional[bool] = None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)


def _replace_cell_paragraphs(cell, lines: list[str], size_pt: int = 14, line_twips: Optional[int] = None, line_rule: str = "atLeast"):
    """保留第一段之 pPr 樣式，清空儲存格並寫入多行段落"""
    tc = cell._tc
    p_elems = tc.findall(qn("w:p"))
    template_pPr = copy.deepcopy(p_elems[0].find(qn("w:pPr"))) if p_elems and p_elems[0].find(qn("w:pPr")) is not None else None

    for p in p_elems:
        tc.remove(p)

    if not lines:
        lines = [""]

    for text in lines:
        p = OxmlElement("w:p")
        if template_pPr is not None:
            pPr = copy.deepcopy(template_pPr)
            if line_twips is not None:
                spacing = pPr.find(qn("w:spacing"))
                if spacing is None:
                    spacing = OxmlElement("w:spacing")
                    pPr.append(spacing)
                spacing.set(qn("w:line"), str(line_twips))
                spacing.set(qn("w:lineRule"), line_rule)
            p.append(pPr)
        tc.append(p)
        para = cell.paragraphs[-1]
        if text:
            run = para.add_run(text)
            _set_run_font(run, "標楷體", size_pt)


def format_violation_line(item: dict) -> str:
    """格式化單筆違規明細行"""
    if item.get("custom_line"):
        return item["custom_line"].strip()

    cid = str(item.get("clause_id", "")).strip()
    text = str(item.get("clause_text", "")).strip()
    unit_fine = int(item.get("unit_fine") or 0)
    count = max(1, int(item.get("count") or 1))
    repeat_level = int(item.get("repeat_level") or 0)
    multiplier = {0: 1, 1: 2, 2: 3, 3: 4}.get(repeat_level, 1)
    subtotal = unit_fine * count * multiplier

    # 確保條文開頭不重複帶條號
    if cid and text.startswith(cid):
        prefix_body = text
    elif cid:
        # 工安天條顯示「工安天條第X項：...」或直接「條號+內容」
        if cid.startswith("天條"):
            prefix_body = f"職安保命條款({cid})：{text}"
        else:
            prefix_body = f"{cid}{text}"
    else:
        prefix_body = text

    if unit_fine <= 0:
        return f"{prefix_body}(依契約規定辦理)"

    if count == 1 and multiplier == 1:
        fee_part = f"({unit_fine:,}元)"
    elif multiplier == 1:
        fee_part = f"({unit_fine:,}元×{count}={subtotal:,}元)"
    elif count == 1:
        fee_part = f"({unit_fine:,}元×重複{multiplier}倍={subtotal:,}元)"
    else:
        fee_part = f"({unit_fine:,}元×{count}×重複{multiplier}倍={subtotal:,}元)"

    return f"{prefix_body}{fee_part}"


def calculate_total(items: list[dict]) -> int:
    total = 0
    for it in items:
        if not it.get("selected", True):
            continue
        unit_fine = int(it.get("unit_fine") or 0)
        count = max(1, int(it.get("count") or 1))
        repeat_level = int(it.get("repeat_level") or 0)
        multiplier = {0: 1, 1: 2, 2: 3, 3: 4}.get(repeat_level, 1)
        total += unit_fine * count * multiplier
    return total


def generate_ticket_docx(
    ticket_data: dict,
    template_path: Path = TEMPLATE_PATH,
    output_path: Optional[Path] = None,
) -> bytes:
    """
    ticket_data 欄位：
      - project_name: str
      - project_code: str (可含換行)
      - contractor: str
      - safety_officer: str
      - occur_time: str (例如：中華民國115年11月01日10時00分)
      - occur_location: str (例如：工地(節點17-1))
      - violation_summary: str (接在 [違規情形]： 後方的總述，可留空)
      - evidence_note: str (接在 [佐證資料] 下方的備註，可留空)
      - items: list[dict] (已確認之違規條文清單)
      - override_total: Optional[int]
    """
    doc = Document(str(template_path))
    tbl = doc.tables[0]

    # Row 0: 工程名稱 / 工程編號
    _replace_cell_paragraphs(tbl.rows[0].cells[1], [ticket_data.get("project_name", "")], size_pt=14)
    code_lines = [ln.strip() for ln in str(ticket_data.get("project_code", "")).splitlines() if ln.strip()]
    _replace_cell_paragraphs(tbl.rows[0].cells[3], code_lines or [""], size_pt=14)

    # Row 1: 廠商名稱 / 職安人員
    _replace_cell_paragraphs(tbl.rows[1].cells[1], [ticket_data.get("contractor", "")], size_pt=14)
    _replace_cell_paragraphs(tbl.rows[1].cells[3], [ticket_data.get("safety_officer", "")], size_pt=14)

    # Row 2: 違規事項
    selected_items = [it for it in ticket_data.get("items", []) if it.get("selected", True)]
    total = ticket_data.get("override_total")
    if total is None:
        total = calculate_total(selected_items)
    total = int(total)

    occur_time = ticket_data.get("occur_time", "").strip()
    occur_location = ticket_data.get("occur_location", "").strip()
    summary = ticket_data.get("violation_summary", "").strip()

    v_lines = [
        f"[發生時間]：{occur_time}",
        f"[發生地點]：{occur_location}",
        f"[違規情形]：{summary}" if summary else "[違規情形]：",
        f"[罰款金額]：{total:,}元",
    ]
    for it in selected_items:
        v_lines.append(format_violation_line(it))

    _replace_cell_paragraphs(tbl.rows[2].cells[1], v_lines, size_pt=12, line_twips=260, line_rule="exact")

    # Row 3: 罰款總額（大寫）
    chinese_total = int_to_chinese_currency(total)
    _replace_cell_paragraphs(tbl.rows[3].cells[1], [f"新台幣(大寫)：{chinese_total}"], size_pt=14)

    # Row 4: 備註 / 佐證資料
    evidence = str(ticket_data.get("evidence_note", "")).strip()
    note_lines = ["[佐證資料] "]
    if evidence:
        note_lines.extend(evidence.splitlines())
    else:
        note_lines.append("")
    _replace_cell_paragraphs(tbl.rows[4].cells[1], note_lines, size_pt=14)

    buf = io.BytesIO()
    doc.save(buf)
    data = buf.getvalue()

    if output_path is not None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(data)

    return data


if __name__ == "__main__":
    # 快速測試國字大寫
    for val in [199000, 5000, 100500, 1234567, 300000]:
        print(val, "->", int_to_chinese_currency(val))
