#!/usr/bin/env python3
"""
解析稽核單位開立之督導紀錄表（.xlsx），萃取出：
- 工程基本資料（工程名稱、工程編號、督導日期、監造單位、承攬商、工地負責人、職安人員、契約金額）
- 壹、書面資料 缺失明細
- 貳、現場作業 缺失明細
"""

import datetime
import io
import re
from pathlib import Path
from typing import Any
import openpyxl


def _format_project_code(raw_code: str) -> str:
    """將例如 TP11304010002TD11304010002 轉為 TP-113-0401-0002\\nTD-113-0401-0002"""
    codes = re.findall(r"[A-Z]{2}\d{11}", raw_code.upper())
    if codes:
        formatted = [f"{c[:2]}-{c[2:5]}-{c[5:9]}-{c[9:13]}" for c in codes]
        return "\n".join(formatted)
    # 若本身已有橫線分隔
    dash_codes = re.findall(r"[A-Z]{2}-\d{3}-\d{4}-\d{4}", raw_code.upper())
    if dash_codes:
        return "\n".join(dash_codes)
    return raw_code.strip()


def _parse_date(val: Any) -> dict:
    """解析西元或民國日期，回傳民國年月日與星期資訊"""
    dt = None
    if isinstance(val, datetime.datetime):
        dt = val.date()
    elif isinstance(val, datetime.date):
        dt = val
    elif val is not None:
        s = str(val).strip()
        # 西元 YYYY-MM-DD / YYYY/MM/DD
        m = re.search(r"(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})", s)
        if m:
            dt = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        else:
            # 民國 115/11/01 or 115年11月01日
            m2 = re.search(r"(\d{2,3})[-/年.](\d{1,2})[-/月.](\d{1,2})", s)
            if m2:
                roc_y = int(m2.group(1))
                dt = datetime.date(roc_y + 1911, int(m2.group(2)), int(m2.group(3)))

    if dt is None:
        return {
            "raw": str(val or "").strip(),
            "roc_year": "",
            "month": "",
            "day": "",
            "roc_date_text": "",
            "roc_datetime_default": "中華民國   年   月   日   時   分",
            "weekday": None,
            "is_weekend": False,
        }

    roc_y = dt.year - 1911
    weekday_names = ["一", "二", "三", "四", "五", "六", "日"]
    wd_idx = dt.weekday()
    return {
        "raw": dt.isoformat(),
        "roc_year": str(roc_y),
        "month": f"{dt.month:02d}",
        "day": f"{dt.day:02d}",
        "roc_date_text": f"中華民國{roc_y}年{dt.month:02d}月{dt.day:02d}日",
        "roc_datetime_default": f"中華民國{roc_y}年{dt.month:02d}月{dt.day:02d}日10時00分",
        "weekday": f"星期{weekday_names[wd_idx]}",
        "is_weekend": wd_idx >= 5,
    }


def _find_right_value(ws, row: int, label_col: int) -> str:
    """在同一列 label_col 右側找第一個非空儲存格值"""
    for c in range(label_col + 1, ws.max_column + 1):
        v = ws.cell(row, c).value
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""


def parse_deficiency_excel(source: Path | bytes) -> dict:
    if isinstance(source, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(source), data_only=True)
    else:
        wb = openpyxl.load_workbook(str(source), data_only=True)

    # 選擇優先工作表
    ws = wb[wb.sheetnames[0]]
    for name in wb.sheetnames:
        if "督導" in name or "查核" in name or "缺失" in name:
            ws = wb[name]
            break

    meta = {
        "title": "",
        "project_name": "",
        "project_code": "",
        "inspection_date": None,
        "supervision_unit": "",
        "contractor": "",
        "project_summary": "",
        "contract_amount": "",
        "site_manager": "",
        "safety_officer": "",
    }

    # 偵測標題與基本資料（掃描前15列）
    for r in range(1, min(16, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is None:
                continue
            s = str(val).strip()
            if r == 1 and not meta["title"]:
                meta["title"] = s
            clean_label = re.sub(r"\s+", "", s)
            if clean_label in ("工程名稱", "工程名稱："):
                raw_proj = _find_right_value(ws, r, c)
                lines = [ln.strip() for ln in raw_proj.splitlines() if ln.strip()]
                if lines:
                    meta["project_name"] = lines[0]
                    if len(lines) > 1:
                        meta["project_code"] = _format_project_code(" ".join(lines[1:]))
                    else:
                        # 嘗試從單行括號萃取編號
                        m = re.search(r"[（(]([^）)]+)[）)]", lines[0])
                        if m and re.search(r"[A-Z]{2}\d", m.group(1)):
                            meta["project_code"] = _format_project_code(m.group(1))
                            meta["project_name"] = lines[0][: m.start()].strip()
            elif clean_label in ("工程編號", "工程編號："):
                meta["project_code"] = _format_project_code(_find_right_value(ws, r, c))
            elif clean_label in ("督導日期", "查核日期", "稽查日期", "檢查日期"):
                for nc in range(c + 1, ws.max_column + 1):
                    nv = ws.cell(r, nc).value
                    if nv is not None and str(nv).strip() != "":
                        meta["inspection_date"] = _parse_date(nv)
                        break
            elif clean_label in ("監造單位", "監造單位："):
                meta["supervision_unit"] = _find_right_value(ws, r, c)
            elif clean_label in ("承攬商", "承攬廠商", "廠商名稱"):
                meta["contractor"] = _find_right_value(ws, r, c)
            elif clean_label in ("契約金額", "工程金額"):
                meta["contract_amount"] = _find_right_value(ws, r, c)
            elif clean_label in ("工作場所負責人", "工地負責人"):
                meta["site_manager"] = _find_right_value(ws, r, c)
            elif clean_label in ("職安人員", "職業安全衛生人員"):
                meta["safety_officer"] = _find_right_value(ws, r, c)

    if meta["inspection_date"] is None:
        meta["inspection_date"] = _parse_date(None)

    # 解析「壹、書面資料」與「貳、現場作業」缺失列
    deficiencies = []
    current_section = ""
    in_table = False
    col_seq = 1
    col_cat = 2
    col_desc = 7

    for r in range(1, ws.max_row + 1):
        row_texts = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        joined = "".join(row_texts)

        if not joined:
            in_table = False
            continue

        # 偵測簽名區即停止
        if any(k in joined for k in ("承攬商簽名", "監造單位簽名", "受查單位", "督導委員簽名")):
            break

        # 偵測章節標題
        if "書面資料" in joined and ("壹" in joined or len(joined) < 20):
            current_section = "壹、書面資料"
            in_table = False
            continue
        if "現場作業" in joined and ("貳" in joined or len(joined) < 20):
            current_section = "貳、現場作業"
            in_table = False
            continue

        # 偵測表頭列（序號 / 違反事項 / 說明）
        if "序號" in joined and ("違反事項" in joined or "缺失事項" in joined or "說明" in joined):
            in_table = True
            for c in range(1, ws.max_column + 1):
                hv = re.sub(r"\s+", "", str(ws.cell(r, c).value or ""))
                if hv == "序號":
                    col_seq = c
                elif hv in ("違反事項", "缺失事項", "項目"):
                    col_cat = c
                elif hv in ("說明", "缺失說明", "缺失情形", "內容"):
                    col_desc = c
            continue

        if in_table or (current_section and ws.cell(r, col_seq).value is not None):
            seq_val = ws.cell(r, col_seq).value
            cat_val = str(ws.cell(r, col_cat).value or "").strip()
            desc_val = str(ws.cell(r, col_desc).value or "").strip()

            # 若說明欄位是合併在其他欄，嘗試收集 col_cat 之後第一個非空值
            if not desc_val:
                for c in range(col_cat + 1, ws.max_column + 1):
                    tv = str(ws.cell(r, c).value or "").strip()
                    if tv:
                        desc_val = tv
                        break

            if not cat_val and not desc_val:
                continue

            seq_str = str(seq_val).strip() if seq_val is not None else str(len(deficiencies) + 1)
            deficiencies.append({
                "index": len(deficiencies) + 1,
                "section": current_section or "稽查缺失",
                "seq": seq_str,
                "category": cat_val,
                "description": desc_val,
            })

    return {
        "meta": meta,
        "deficiencies": deficiencies,
    }


if __name__ == "__main__":
    import json
    res = parse_deficiency_excel(Path("/home/antigravity/督導紀錄表-1151101(烏日南屯北段).xlsx"))
    print(json.dumps(res, ensure_ascii=False, indent=2))
