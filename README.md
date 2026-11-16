# 台灣自來水公司｜承攬商違規缺失辨識與罰單生成系統

依據契約 [06_施工說明書總則及有關規定等(113年元月版).docx](../06_%E6%96%BD%E5%B7%A5%E8%AA%AA%E6%98%8E%E6%9B%B8%E7%B8%BD%E5%89%87%E5%8F%8A%E6%9C%89%E9%97%9C%E8%A6%8F%E5%AE%9A%E7%AD%89(113%E5%B9%B4%E5%85%83%E6%9C%88%E7%89%88).docx) 第 118～128 頁之「承攬商安全衛生違規稽查表（TS-01，112.12修訂）」及「台灣自來水公司職安保命條款（工安天條）」，將稽核單位開立之 `督導紀錄表*.xlsx` 自動辨識並對應契約罰則，提供完整手動修改與即時預覽，確認後輸出與 `(_罰單_)_違規稽查表1151101.docx` 格式一致之 TS-02 Word 罰扣款通知單。

## 目錄結構
- `extract_clauses.py`：從契約 `.docx` 擷取全部 185 條 TS-01 與 10 條職安保命條款至 `data/clauses.json`。
- `excel_parser.py`：解析 `督導紀錄表*.xlsx` 基本資料及「壹、書面資料」「貳、現場作業」缺失清單。
- `matcher.py`：結合領域規則與字詞相似度進行條文比對、2.18/2.35 計次推算、假日施工提醒及工安天條競合（取大值）警示。
- `ticket_generator.py`：套用 `templates/ticket_template.docx` 生成維持原 `標楷體` 排版與新台幣大寫金額之 `.docx` 罰單。
- `app.py` + `static/`：互動式 Web 系統（上傳 → 人工修改與確認 → 即時預覽 → 生成下載）。
- `tests/test_system.py`：自動化測試程式。

## 本機執行方式
```bash
pip install fastapi uvicorn openpyxl python-docx
python3 -m uvicorn app:app --host 0.0.0.0 --port 8765
```
