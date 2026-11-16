#!/usr/bin/env python3
"""
契約條文自動辨識引擎：
結合「領域規則（精準對應TS-01與工安天條）」與「字詞相似度排序」，
將督導紀錄表的各項缺失對應到最符合的契約罰則條款，並提供候選條款與競合/倍數提醒。
"""

import json
import math
import re
from pathlib import Path
from typing import Optional

CLAUSES_PATH = Path(__file__).parent / "data" / "clauses.json"


def load_clauses_data(path: Path = CLAUSES_PATH) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


# 高精準領域規則：(regex_pattern, [(clause_id, score_boost, reason), ...])
DOMAIN_RULES: list[tuple[str, list[tuple[str, float, str]]]] = [
    # 1. 作業申請 / 假日施工
    (r"(未申請.*許可|未張貼.*許可|無.*工作許可)", [
        ("1.1", 0.90, "未申請或未具備工作許可證"),
        ("6.4", 0.65, "進入局限空間未有同意文件"),
    ]),
    (r"(假日|夜間).*施工", [("1.2", 0.95, "假日或夜間施工")]),

    # 2. 承攬商安全衛生管理
    (r"(簽到|簽退|危害告知|生理協調|額溫)", [
        ("2.18", 0.96, "未依規定辦理簽到、危害告知、簽退等上工前規定"),
        ("2.38", 0.72, "每日作業前未對勞工施以危害告知"),
        ("2.47", 0.65, "未依規定上傳簽到及危害因素告知紀錄表"),
    ]),
    (r"(教育訓練|職安卡|資格資料|受訓)", [
        ("2.3", 0.90, "承攬商未提供勞工安全衛生教育訓練資料"),
        ("2.55", 0.80, "作業勞工未參加一般安全衛生教育訓練取得合格證明"),
    ]),
    (r"(自動檢查|定期檢查|作業檢點|當日.*檢查)", [
        ("2.11", 0.85, "未辦理當日車輛、機械、設備等自動檢查"),
    ]),
    (r"(滅火器).*(定期檢查|壓力|有效期限|過期|逾期)", [
        ("12.1", 0.95, "滅火器壓力不正常或逾有效期限/未定期檢查"),
        ("2.14", 0.60, "工程車輛未配置有效期限內滅火器"),
    ]),
    (r"(安全帽|頤帶)", [("2.26", 0.95, "未配戴安全帽或頤帶未扣")]),
    (r"(背心|反光|識別證)", [
        ("2.31", 0.85, "未穿戴識別背心或識別證"),
        ("2.32", 0.82, "反光背心或安全帽未貼反光帶"),
    ]),
    (r"(工地主任|工作場所負責人|職安人員|作業主管).*(未見|未在場|不在場|未於現場|常駐)", [
        ("2.35", 0.98, "工地負責人、職安人員或作業主管未於現場監督指揮"),
    ]),
    (r"(安全帶|防墜器|垂直母索|安全網)", [
        ("2.45", 0.92, "應使用而未使用安全帶、防墜器或安全網"),
        ("天條1", 0.90, "工安天條：高處作業未確實勾掛安全帶"),
    ]),
    (r"(攝影機|即時影像|監視系統|斷訊)", [
        ("2.52", 0.82, "即時影像監視系統相關規定"),
        ("2.50", 0.72, "未依規定上傳鄰接道路作業影像"),
        ("2.51", 0.72, "未依規定上傳局限空間作業照片"),
    ]),
    (r"(2m|2公尺).*(開口|護欄|護蓋|圍欄|覆蓋)", [
        ("2.53", 0.97, "高差2公尺以上工作場所邊緣及開口未設防護措施"),
        ("天條1", 0.92, "工安天條：高處(開口)作業預防墜落措施（與2.53同額30,000元）"),
        ("5.4", 0.82, "人孔等地面開口未設圍欄、覆蓋等安全警告標示"),
    ]),
    (r"(交通錐|連桿|拒馬|改道|交維|交通維持)", [("2.41", 0.92, "鄰接道路作業未落實交通維持設施")]),
    (r"(交通引導|指揮旗手|電動旗手)", [("2.42", 0.92, "未設置交通引導人員或電動旗手")]),
    (r"(職業安全衛生告示牌|危險作業告示|未張貼.*告示|警告標示)", [
        ("2.30", 0.86, "現場各項危險作業告示不足"),
        ("2.9", 0.80, "未於工作場所設置警告標示"),
        ("2.10", 0.78, "現場未張貼各項高風險作業警告標示牌"),
    ]),

    # 3. 高架/施工架作業
    (r"(施工架|鷹架).*(踏板|滿舖)", [("3.2", 0.94, "施工架踏板未依規定滿舖")]),
    (r"(施工架|鷹架).*(荷重)", [("3.3", 0.94, "施工架未標示最大荷重")]),
    (r"(施工架|鷹架).*(護欄|扶手|插鞘|壁拉桿)", [("3.11", 0.94, "施工架無護欄、扶手或插鞘壁拉桿")]),
    (r"(施工架|鷹架)", [("3.1", 0.75, "施工架未符合CNS4750規定")]),

    # 4. 起重機作業
    (r"(吊升荷重.*三公噸以下|荷重.*安定性試驗)", [("4.3", 0.95, "三公噸以下起重機未有荷重及安定性試驗證明")]),
    (r"(吊升荷重.*三公噸以上|起重機合格證|操作人員合格證)", [("4.4", 0.95, "三公噸以上起重機及操作人員資格證明")]),
    (r"(吊鉤|防滑舌片|防脫裝置)", [("4.7", 0.98, "吊鉤無防滑舌片裝置或故障無法使用")]),
    (r"(過捲揚|過捲預防|過捲|制動器|離合器)", [("4.8", 0.98, "未設置過捲預防裝置或安全裝置無法使用")]),
    (r"(鋼索|吊鏈).*(損傷|變形)", [("4.5", 0.94, "鋼索及吊鏈有明顯損傷或變形")]),
    (r"(吊運作業半徑|吊掛物.*上方|作業半徑.*嚴禁)", [("4.1", 0.90, "未防止人員進入吊掛作業半徑")]),

    # 5. 露天開挖作業
    (r"(旋轉警示燈|倒車.*警示燈|蜂鳴器)", [("5.1", 0.98, "營建車輛機械倒車或旋轉警示燈與蜂鳴器未裝置或失效")]),
    (r"(迴轉半徑|操作半徑).*禁止進入", [("5.2", 0.93, "未設置禁止人員進入機械操作半徑設施及警示")]),
    (r"(人孔|地面開口|工作井.*開口)", [
        ("5.4", 0.88, "人孔等地面開口未設圍欄、覆蓋及警告標示"),
        ("2.53", 0.86, "高差2公尺以上開口未設圍欄、覆蓋防護"),
    ]),
    (r"(高低落差|高差.*1\.5|上下設備|爬梯|護籠|轉折平台)", [
        ("5.8", 0.96, "作業地點高差1.5公尺以上未設置安全上下設備"),
        ("2.40", 0.80, "安全爬梯不符合規定"),
    ]),
    (r"(開挖.*土石|土石堆置|坡肩|開口上端.*堆置)", [("5.10", 0.86, "開挖面或開口上方堆置土石物料")]),
    (r"(擋土支撐)", [("5.5", 0.88, "開挖作業未設置適當擋土支撐設施")]),

    # 6. 局限空間作業
    (r"(局限空間|工作井).*(告示牌|注意事項|門禁|上鎖管制)", [
        ("6.2", 0.96, "未於局限空間作業場所入口公告注意事項或門禁管制"),
        ("1.1", 0.75, "未張貼/申請局限空間作業許可"),
    ]),
    (r"(局限空間).*(監視人員)", [("6.3", 0.95, "局限空間作業未指派專職現場監視人員")]),
    (r"(氣體.*偵測|連續監測|氧氣|有害氣體|危害物質濃度|通風)", [
        ("6.5", 0.96, "作業前未依規定實施通風、測定氧氣及有害氣體濃度"),
        ("6.6", 0.92, "作業時未持續通風與連續測定氧氣及有害氣體濃度"),
    ]),
    (r"(緊急救援設備|空氣呼吸器|救生索)", [("6.1", 0.95, "未設置緊急救援設備")]),

    # 7. 用電安全
    (r"(電焊機|發電機|電氣設備).*(接地|非帶電金屬)", [("7.4", 0.98, "電焊機或電氣設備非帶電金屬部分未接地")]),
    (r"(漏電斷路器)", [("7.5", 0.95, "電氣設備每一分路未正確設置漏電斷路器")]),
    (r"(臨時配線|移動電線|裸線|破皮|絕緣不良)", [
        ("7.7", 0.88, "使用破皮或絕緣不良之電線"),
        ("7.1", 0.85, "道路放置未防護之臨時配線"),
    ]),
    (r"(交流電焊機|自動電擊防止)", [("7.10", 0.96, "交流電焊機未設置自動電擊防止裝置")]),

    # 12. 火災預防與危險物
    (r"(鋼瓶).*(直立|固定|遮陽)", [("12.6", 0.98, "鋼瓶貯存未直立並固定")]),
    (r"(GHS|SDS|安全資料表|柴油桶|易燃物|可燃性液體)", [
        ("12.10", 0.88, "易燃物或可燃性液體貯存場所無安全防護/標示措施"),
        ("15.2", 0.72, "違反其他職業安全衛生法令相關事項（GHS/SDS標示）"),
    ]),

    # 13. 整理整頓
    (r"(行走動線|通道|走道|逃生動線|堆置材料)", [("13.2", 0.98, "通道、走道堆積材料物料雜物影響通行及逃生動線")]),
    (r"(整理整頓|雜亂|清潔|收工後)", [("13.10", 0.94, "其他違反整理、整頓、整潔、清掃及紀律事項")]),
]


def _char_ngrams(text: str) -> set[str]:
    clean = re.sub(r"[\s\W_]+", "", text)
    grams = set()
    for n in (2, 3):
        for i in range(max(0, len(clean) - n + 1)):
            grams.add(clean[i : i + n])
    if not grams and clean:
        grams = set(clean)
    return grams


def _similarity_score(query: str, clause_text: str) -> float:
    q_grams = _char_ngrams(query)
    c_grams = _char_ngrams(clause_text)
    if not q_grams or not c_grams:
        return 0.0
    inter = len(q_grams & c_grams)
    return inter / math.sqrt(len(q_grams) * len(c_grams))


def _infer_subitem_details(clause_id: str, combined_text: str) -> tuple[int, str]:
    """針對 2.18 / 2.35 等可計次條款，自動推算違反子項數量或人員數"""
    if clause_id == "2.18":
        subitems = [
            ("1.工作場所簽到", r"簽到"),
            ("2.生理協調平衡檢測", r"生理協調"),
            ("3.額溫量測(上工前)", r"額溫"),
            ("4.危害因素告知", r"危害告知|危害因素"),
            ("5.工作場所簽退", r"簽退|簽到退"),
            ("6.各項作業拍攝照片", r"拍攝照片|拍照"),
            ("7.做成紀錄表", r"紀錄表"),
        ]
        hit = [label for label, pat in subitems if re.search(pat, combined_text)]
        if hit:
            note = "未依規定辦理" + "".join(hit)
            return len(hit), note
    elif clause_id == "2.35":
        roles = [
            ("工地負責人(工地主任)", r"工地主任|工作場所負責人|工地負責人"),
            ("職安人員", r"職安人員|職業安全衛生管理人員"),
            ("急救人員", r"急救人員"),
            ("作業主管", r"作業主管"),
        ]
        hit_count = sum(1 for _, pat in roles if re.search(pat, combined_text))
        if hit_count >= 1:
            return hit_count, f"缺席角色數：{hit_count}位"
    return 1, ""


def match_deficiencies(parsed_excel: dict, clauses_data: Optional[dict] = None) -> dict:
    if clauses_data is None:
        clauses_data = load_clauses_data()

    clauses_by_id = {c["id"]: c for c in clauses_data["clauses"]}
    deficiencies = parsed_excel.get("deficiencies", [])
    meta = parsed_excel.get("meta", {})

    matched_items = []
    for item in deficiencies:
        combined = f"{item.get('category', '')} {item.get('description', '')}".strip()
        scores: dict[str, tuple[float, list[str]]] = {}

        # 1. 套用領域規則
        for pat, targets in DOMAIN_RULES:
            if re.search(pat, combined, re.IGNORECASE):
                for cid, boost, reason in targets:
                    if cid not in clauses_by_id:
                        continue
                    prev_score, prev_reasons = scores.get(cid, (0.0, []))
                    if boost > prev_score:
                        scores[cid] = (boost, prev_reasons + [reason])
                    elif reason not in prev_reasons:
                        scores[cid] = (prev_score, prev_reasons + [reason])

        # 2. 輔以字詞相似度補足候選清單
        for c in clauses_data["clauses"]:
            cid = c["id"]
            sim = _similarity_score(combined, f"{c['category']} {c['text']}")
            prev_score, prev_reasons = scores.get(cid, (0.0, []))
            blended = max(prev_score, sim * 0.75) if prev_score == 0 else min(0.99, prev_score + sim * 0.05)
            scores[cid] = (blended, prev_reasons)

        ranked = sorted(scores.items(), key=lambda x: x[1][0], reverse=True)
        top_candidates = []
        for cid, (sc, reasons) in ranked[:6]:
            cl = clauses_by_id[cid]
            # 檢查工安天條競合
            overlap_warning = None
            if cl.get("overlaps"):
                better = []
                for oid in cl["overlaps"]:
                    ocl = clauses_by_id.get(oid)
                    if ocl and ocl.get("fine") is not None and cl.get("fine") is not None:
                        if ocl["fine"] > cl["fine"]:
                            better.append(f"{oid}（{ocl['fine']:,}元 > {cl['fine']:,}元，依備註取大值）")
                        elif ocl["fine"] == cl["fine"]:
                            better.append(f"{oid}（同為{ocl['fine']:,}元，附註另有1個月禁入原則）")
                if better:
                    overlap_warning = "涉職安保命條款：" + "；".join(better)

            top_candidates.append({
                "clause_id": cid,
                "source": cl["source"],
                "category": cl["category"],
                "text": cl["text"],
                "fine": cl["fine"],
                "fine_text": cl["fine_text"],
                "per_subitem": cl["per_subitem"],
                "confidence": round(sc, 3),
                "reasons": reasons,
                "overlap_warning": overlap_warning,
            })

        best = top_candidates[0]
        count, subitem_note = _infer_subitem_details(best["clause_id"], combined)

        # 若無確切罰款（例如 15.3 或 14.6），標記 fine=0 讓使用者手填
        unit_fine = best["fine"] if best["fine"] is not None else 0

        # 預設寫入罰單的違規條文文字
        if best["clause_id"] == "2.18" and subitem_note:
            ticket_text = f"承攬商於上工前，{subitem_note}。"
        else:
            ticket_text = best["text"]

        matched_items.append({
            "deficiency_index": item["index"],
            "section": item["section"],
            "seq": item["seq"],
            "raw_category": item["category"],
            "raw_description": item["description"],
            "selected": True,
            "clause_id": best["clause_id"],
            "source": best["source"],
            "clause_text": ticket_text,
            "unit_fine": unit_fine,
            "count": count,
            "repeat_level": 0,  # 0->1x, 1->2x, 2->3x, 3->4x
            "confidence": best["confidence"],
            "reasons": best["reasons"],
            "overlap_warning": best["overlap_warning"],
            "candidates": top_candidates,
        })

    # 額外契約提醒（例如：督導日期為週六日 -> 建議核對 1.2 假日施工）
    suggestions = []
    insp_date = meta.get("inspection_date", {})
    if insp_date.get("is_weekend"):
        cl_12 = clauses_by_id.get("1.2")
        if cl_12:
            suggestions.append({
                "clause_id": "1.2",
                "title": f"督導日期（{insp_date.get('roc_date_text')}）為{insp_date.get('weekday')}",
                "message": "若承攬商未經同意於假日擅自施工，可增列「1.2 未經同意假日或夜間擅自施工者」(20,000元)。",
                "clause": cl_12,
            })

    return {
        "meta": meta,
        "items": matched_items,
        "suggestions": suggestions,
    }


if __name__ == "__main__":
    from excel_parser import parse_deficiency_excel
    parsed = parse_deficiency_excel(Path("/home/antigravity/督導紀錄表-1151101(烏日南屯北段).xlsx"))
    res = match_deficiencies(parsed)
    for it in res["items"]:
        print(f"#{it['seq']} [{it['raw_category']}] {it['raw_description'][:32]}...")
        print(f"   => {it['clause_id']} ({it['unit_fine']}元 x {it['count']}) conf={it['confidence']} | {it['reasons']}")
        if it["overlap_warning"]:
            print(f"   !! {it['overlap_warning']}")
    print("Suggestions:", [s["title"] for s in res["suggestions"]])
