#!/usr/bin/env python3
"""
從契約 DOCX（06_施工說明書總則及有關規定等）擷取罰則條文：
- 承攬商安全衛生違規稽查表（TS-01，112.12修訂）
- 台灣自來水公司職安保命條款（工安天條）
- 相關計罰備註規則
"""

import json
import re
from pathlib import Path
from docx import Document

DEFAULT_CONTRACT_PATH = Path("/home/antigravity/06_施工說明書總則及有關規定等(113年元月版).docx")
OUTPUT_PATH = Path(__file__).parent / "data" / "clauses.json"

# 契約原文因嵌字遺漏「1.5公尺」之條款補全（與罰單範例 5.8 一致）
TEXT_FIXES = {
    "5.5": "開挖垂直深度未達1.5公尺開挖作業如有崩塌或土石飛落之虞，未設置適當之擋土支撐設施。",
    "5.8": "作業地點高差1.5公尺以上未設置勞工能安全上下之設備。",
    "5.12": "施工人員在開挖垂直深度達1.5公尺以上，未設置擋土支撐設施處任意活動。",
    "5.13": "開挖垂直深度達1.5公尺以上未設置擋土支撐，且未經具有專業工程技師人員簽認其安全性者。",
}

# TS-01 與 工安天條 競合對應（依契約備註：相同缺失事項取大值開罰）
LIFESAVING_OVERLAPS = {
    "2.45": ["天條1"],
    "2.53": ["天條1"],
    "5.4": ["天條1"],
    "1.1": ["天條3", "天條4"],
    "6.4": ["天條4"],
    "6.5": ["天條4"],
    "6.6": ["天條4"],
    "7.2": ["天條5"],
    "2.29": ["天條6"],
    "2.44": ["天條6"],
    "12.11": ["天條3"],
    "12.12": ["天條3"],
}


def parse_fine_amount(fine_str: str) -> dict:
    s = re.sub(r"\s+", " ", fine_str).strip()
    if not s:
        return {"fine": None, "fine_text": "依契約規定（如遭處罰鍰總額之50%等）", "per_subitem": False}

    # Check "違反其中1項罰扣3,000"
    m_sub = re.search(r"違反其中\s*1\s*項罰扣\s*([\d,]+)", s)
    if m_sub:
        val = int(m_sub.group(1).replace(",", ""))
        return {"fine": val, "fine_text": s, "per_subitem": True}

    # Plain number like "20,000" or "3,000"
    m_num = re.fullmatch(r"[\d,]+", s)
    if m_num:
        return {"fine": int(s.replace(",", "")), "fine_text": s, "per_subitem": False}

    # Complex text (e.g. 14.6 or 天條8)
    return {"fine": None, "fine_text": s, "per_subitem": False}


def clean_category(cat: str) -> str:
    # Collapse excess spaces/newlines in vertical/spaced Chinese headers
    cat = re.sub(r"\s+", "", cat).strip()
    # Fix "高架作業）或高樓作業（" artifact from vertical text
    cat = cat.replace("高架作業）或高樓作業（", "高架（或高樓）作業")
    # Deduplicate repeated category names caused by merged cells
    if len(cat) >= 6 and len(cat) % 2 == 0:
        half = len(cat) // 2
        if cat[:half] == cat[half:]:
            cat = cat[:half]
    return cat


def extract_ts01_clauses(table) -> list[dict]:
    clauses = []
    # Pattern matches item codes like "1.1 □...", "2.10□...", "10.19□..."
    code_pat = re.compile(r"(?<!\d)(\d{1,2}\.\d{1,2})\s*□\s*")

    for ri, row in enumerate(table.rows):
        if ri == 0:
            continue
        seen = set()
        cells = []
        for c in row.cells:
            if id(c._tc) not in seen:
                seen.add(id(c._tc))
                cells.append(c.text.strip())
        if len(cells) < 3:
            continue

        category = clean_category(cells[0])
        items_text = cells[1]
        fine_raw = cells[2]
        remark = cells[3] if len(cells) > 3 else ""
        fine_info = parse_fine_amount(fine_raw)

        # Split items_text by item codes
        matches = list(code_pat.finditer(items_text))
        for idx, m in enumerate(matches):
            code = m.group(1)
            start = m.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(items_text)
            body = re.sub(r"\s+", " ", items_text[start:end]).strip()
            # Clean stray spaces inside full-width text like "承攬商 勞工" -> keep readable
            if code in TEXT_FIXES:
                body = TEXT_FIXES[code]

            clauses.append({
                "id": code,
                "source": "TS-01",
                "category": category,
                "text": body,
                "fine": fine_info["fine"],
                "fine_text": fine_info["fine_text"],
                "per_subitem": fine_info["per_subitem"],
                "remark": re.sub(r"\s+", " ", remark).strip(),
                "overlaps": LIFESAVING_OVERLAPS.get(code, []),
            })
    return clauses


def extract_lifesaving_clauses(table) -> list[dict]:
    clauses = []
    for ri, row in enumerate(table.rows):
        if ri <= 1:
            continue
        seen = set()
        cells = []
        for c in row.cells:
            if id(c._tc) not in seen:
                seen.add(id(c._tc))
                cells.append(c.text.strip())
        if not cells or not cells[0].isdigit():
            continue
        seq = cells[0].strip()
        body = re.sub(r"\s+", " ", cells[1]).strip()
        fine_raw = cells[2] if len(cells) > 2 else ""
        fine_info = parse_fine_amount(fine_raw)
        clauses.append({
            "id": f"天條{seq}",
            "source": "職安保命條款(工安天條)",
            "category": "職安保命條款(工安天條)",
            "text": body,
            "fine": fine_info["fine"],
            "fine_text": fine_info["fine_text"],
            "per_subitem": False,
            "remark": "違反者該勞工1個月內禁止進入工區；半年累計3次撤換工地主任及職安人員",
            "overlaps": [],
        })
    return clauses


def find_contract_tables(doc: Document):
    ts01_tbl = None
    ls_tbl = None
    for t in doc.tables:
        head = " ".join(c.text.strip() for r in t.rows[:2] for c in r.cells)
        if "稽     查     缺     失    事    項" in head or "未申請工作許可證擅自作業" in head:
            ts01_tbl = t
        elif "職安保命條款" in head or "工安天條" in head:
            ls_tbl = t
    return ts01_tbl, ls_tbl


def main(contract_path: Path = DEFAULT_CONTRACT_PATH, out_path: Path = OUTPUT_PATH):
    doc = Document(str(contract_path))
    ts01_tbl, ls_tbl = find_contract_tables(doc)
    if ts01_tbl is None or ls_tbl is None:
        raise RuntimeError("無法在契約文件內找到 TS-01 稽查表或職安保命條款表格")

    ts01 = extract_ts01_clauses(ts01_tbl)
    lifesaving = extract_lifesaving_clauses(ls_tbl)

    payload = {
        "contract_file": contract_path.name,
        "version": "TS-01 (112.12修訂) & 職安保命條款(工安天條)",
        "rules": {
            "repeat_multipliers": {
                "0": {"label": "首次違規 (1倍)", "multiplier": 1},
                "1": {"label": "第1次重複違反 (2倍)", "multiplier": 2},
                "2": {"label": "第2次重複違反 (3倍)", "multiplier": 3},
                "3": {"label": "第3次以上重複違反 (4倍)", "multiplier": 4},
            },
            "overlap_rule": "違反「承攬商安全衛生違規稽查表」或「台灣自來水公司職安保命條款(工安天條)」之相同缺失事項者，其罰款則取大值開罰。",
            "objection_days": 10,
        },
        "clauses": ts01 + lifesaving,
    }

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Extracted {len(ts01)} TS-01 clauses + {len(lifesaving)} 工安天條 clauses -> {out_path}")


if __name__ == "__main__":
    main()
