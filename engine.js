/**
 * 純前端引擎（支援 GitHub Pages 靜態網頁直接運行）：
 * 1. 使用 SheetJS (XLSX) 解析督導紀錄表 .xlsx
 * 2. 執行 TS-01 + 職安保命條款(工安天條) 領域規則與字詞相似度比對
 * 3. 新台幣國字大寫轉換
 * 4. 使用 JSZip + XMLSerializer 讀寫 templates/ticket_template.docx 生成完整排版 Word 罰單
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DOMAIN_RULES = [
  [/(未申請.*許可|未張貼.*許可|無.*工作許可)/i, [
    ["1.1", 0.90, "未申請或未具備工作許可證"],
    ["6.4", 0.65, "進入局限空間未有同意文件"],
  ]],
  [/(假日|夜間).*施工/i, [["1.2", 0.95, "假日或夜間施工"]]],
  [/(簽到|簽退|危害告知|生理協調|額溫)/i, [
    ["2.18", 0.96, "未依規定辦理簽到、危害告知、簽退等上工前規定"],
    ["2.38", 0.72, "每日作業前未對勞工施以危害告知"],
    ["2.47", 0.65, "未依規定上傳簽到及危害因素告知紀錄表"],
  ]],
  [/(教育訓練|職安卡|資格資料|受訓)/i, [
    ["2.3", 0.90, "承攬商未提供勞工安全衛生教育訓練資料"],
    ["2.55", 0.80, "作業勞工未參加一般安全衛生教育訓練取得合格證明"],
  ]],
  [/(自動檢查|定期檢查|作業檢點|當日.*檢查)/i, [
    ["2.11", 0.85, "未辦理當日車輛、機械、設備等自動檢查"],
  ]],
  [/(滅火器).*(定期檢查|壓力|有效期限|過期|逾期)/i, [
    ["12.1", 0.95, "滅火器壓力不正常或逾有效期限/未定期檢查"],
    ["2.14", 0.60, "工程車輛未配置有效期限內滅火器"],
  ]],
  [/(安全帽|頤帶)/i, [["2.26", 0.95, "未配戴安全帽或頤帶未扣"]]],
  [/(背心|反光|識別證)/i, [
    ["2.31", 0.85, "未穿戴識別背心或識別證"],
    ["2.32", 0.82, "反光背心或安全帽未貼反光帶"],
  ]],
  [/(工地主任|工作場所負責人|職安人員|作業主管).*(未見|未在場|不在場|未於現場|常駐)/i, [
    ["2.35", 0.98, "工地負責人、職安人員或作業主管未於現場監督指揮"],
  ]],
  [/(安全帶|防墜器|垂直母索|安全網)/i, [
    ["2.45", 0.92, "應使用而未使用安全帶、防墜器或安全網"],
    ["天條1", 0.90, "工安天條：高處作業未確實勾掛安全帶"],
  ]],
  [/(攝影機|即時影像|監視系統|斷訊)/i, [
    ["2.52", 0.82, "即時影像監視系統相關規定"],
    ["2.50", 0.72, "未依規定上傳鄰接道路作業影像"],
    ["2.51", 0.72, "未依規定上傳局限空間作業照片"],
  ]],
  [/(2m|2公尺).*(開口|護欄|護蓋|圍欄|覆蓋)/i, [
    ["2.53", 0.97, "高差2公尺以上工作場所邊緣及開口未設防護措施"],
    ["天條1", 0.92, "工安天條：高處(開口)作業預防墜落措施（與2.53同額30,000元）"],
    ["5.4", 0.82, "人孔等地面開口未設圍欄、覆蓋等安全警告標示"],
  ]],
  [/(交通錐|連桿|拒馬|改道|交維|交通維持)/i, [["2.41", 0.92, "鄰接道路作業未落實交通維持設施"]]],
  [/(交通引導|指揮旗手|電動旗手)/i, [["2.42", 0.92, "未設置交通引導人員或電動旗手"]]],
  [/(職業安全衛生告示牌|危險作業告示|未張貼.*告示|警告標示)/i, [
    ["2.30", 0.86, "現場各項危險作業告示不足"],
    ["2.9", 0.80, "未於工作場所設置警告標示"],
    ["2.10", 0.78, "現場未張貼各項高風險作業警告標示牌"],
  ]],
  [/(施工架|鷹架).*(踏板|滿舖)/i, [["3.2", 0.94, "施工架踏板未依規定滿舖"]]],
  [/(施工架|鷹架).*(荷重)/i, [["3.3", 0.94, "施工架未標示最大荷重"]]],
  [/(施工架|鷹架).*(護欄|扶手|插鞘|壁拉桿)/i, [["3.11", 0.94, "施工架無護欄、扶手或插鞘壁拉桿"]]],
  [/(施工架|鷹架)/i, [["3.1", 0.75, "施工架未符合CNS4750規定"]]],
  [/(吊升荷重.*三公噸以下|荷重.*安定性試驗)/i, [["4.3", 0.95, "三公噸以下起重機未有荷重及安定性試驗證明"]]],
  [/(吊升荷重.*三公噸以上|起重機合格證|操作人員合格證)/i, [["4.4", 0.95, "三公噸以上起重機及操作人員資格證明"]]],
  [/(吊鉤|防滑舌片|防脫裝置)/i, [["4.7", 0.98, "吊鉤無防滑舌片裝置或故障無法使用"]]],
  [/(過捲揚|過捲預防|過捲|制動器|離合器)/i, [["4.8", 0.98, "未設置過捲預防裝置或安全裝置無法使用"]]],
  [/(鋼索|吊鏈).*(損傷|變形)/i, [["4.5", 0.94, "鋼索及吊鏈有明顯損傷或變形"]]],
  [/(吊運作業半徑|吊掛物.*上方|作業半徑.*嚴禁)/i, [["4.1", 0.90, "未防止人員進入吊掛作業半徑"]]],
  [/(旋轉警示燈|倒車.*警示燈|蜂鳴器)/i, [["5.1", 0.98, "營建車輛機械倒車或旋轉警示燈與蜂鳴器未裝置或失效"]]],
  [/(迴轉半徑|操作半徑).*禁止進入/i, [["5.2", 0.93, "未設置禁止人員進入機械操作半徑設施及警示"]]],
  [/(人孔|地面開口|工作井.*開口)/i, [
    ["5.4", 0.88, "人孔等地面開口未設圍欄、覆蓋及警告標示"],
    ["2.53", 0.86, "高差2公尺以上開口未設圍欄、覆蓋防護"],
  ]],
  [/(高低落差|高差.*1\.5|上下設備|爬梯|護籠|轉折平台)/i, [
    ["5.8", 0.96, "作業地點高差1.5公尺以上未設置安全上下設備"],
    ["2.40", 0.80, "安全爬梯不符合規定"],
  ]],
  [/(開挖.*土石|土石堆置|坡肩|開口上端.*堆置)/i, [["5.10", 0.86, "開挖面或開口上方堆置土石物料"]]],
  [/(擋土支撐)/i, [["5.5", 0.88, "開挖作業未設置適當擋土支撐設施"]]],
  [/(局限空間|工作井).*(告示牌|注意事項|門禁|上鎖管制)/i, [
    ["6.2", 0.96, "未於局限空間作業場所入口公告注意事項或門禁管制"],
    ["1.1", 0.75, "未張貼/申請局限空間作業許可"],
  ]],
  [/(局限空間).*(監視人員)/i, [["6.3", 0.95, "局限空間作業未指派專職現場監視人員"]]],
  [/(氣體.*偵測|連續監測|氧氣|有害氣體|危害物質濃度|通風)/i, [
    ["6.5", 0.96, "作業前未依規定實施通風、測定氧氣及有害氣體濃度"],
    ["6.6", 0.92, "作業時未持續通風與連續測定氧氣及有害氣體濃度"],
  ]],
  [/(緊急救援設備|空氣呼吸器|救生索)/i, [["6.1", 0.95, "未設置緊急救援設備"]]],
  [/(電焊機|發電機|電氣設備).*(接地|非帶電金屬)/i, [["7.4", 0.98, "電焊機或電氣設備非帶電金屬部分未接地"]]],
  [/(漏電斷路器)/i, [["7.5", 0.95, "電氣設備每一分路未正確設置漏電斷路器"]]],
  [/(臨時配線|移動電線|裸線|破皮|絕緣不良)/i, [
    ["7.7", 0.88, "使用破皮或絕緣不良之電線"],
    ["7.1", 0.85, "道路放置未防護之臨時配線"],
  ]],
  [/(交流電焊機|自動電擊防止)/i, [["7.10", 0.96, "交流電焊機未設置自動電擊防止裝置"]]],
  [/(鋼瓶).*(直立|固定|遮陽)/i, [["12.6", 0.98, "鋼瓶貯存未直立並固定"]]],
  [/(GHS|SDS|安全資料表|柴油桶|易燃物|可燃性液體)/i, [
    ["12.10", 0.88, "易燃物或可燃性液體貯存場所無安全防護/標示措施"],
    ["15.2", 0.72, "違反其他職業安全衛生法令相關事項（GHS/SDS標示）"],
  ]],
  [/(行走動線|通道|走道|逃生動線|堆置材料)/i, [["13.2", 0.98, "通道、走道堆積材料物料雜物影響通行及逃生動線"]]],
  [/(整理整頓|雜亂|清潔|收工後)/i, [["13.10", 0.94, "其他違反整理、整頓、整潔、清掃及紀律事項"]]],
];

export function intToChineseCurrency(amount) {
  const n = Math.floor(Number(amount) || 0);
  if (n <= 0) return "零元整";
  const digits = "零壹貳參肆伍陸柒捌玖";
  const smallUnits = ["", "拾", "佰", "仟"];
  const bigUnits = ["", "萬", "億", "兆"];

  let s = String(n);
  const groups = [];
  while (s.length > 0) {
    groups.push(s.slice(-4));
    s = s.slice(0, -4);
  }

  const parts = [];
  let needZeroBetween = false;
  for (let gIdx = groups.length - 1; gIdx >= 0; gIdx--) {
    const grp = groups[gIdx];
    const gVal = parseInt(grp, 10);
    if (gVal === 0) {
      if (parts.length > 0) needZeroBetween = true;
      continue;
    }
    if (needZeroBetween || (parts.length > 0 && grp.length === 4 && grp[0] === "0")) {
      parts.push("零");
      needZeroBetween = false;
    }
    let grpStr = "";
    let zeroFlag = false;
    for (let i = 0; i < grp.length; i++) {
      const d = parseInt(grp[i], 10);
      const uIdx = grp.length - 1 - i;
      if (d === 0) {
        if (grpStr) zeroFlag = true;
      } else {
        if (zeroFlag) {
          grpStr += "零";
          zeroFlag = false;
        }
        grpStr += digits[d] + smallUnits[uIdx];
      }
    }
    parts.push(grpStr + bigUnits[gIdx]);
  }
  return parts.join("") + "元整";
}

function formatProjectCode(raw) {
  const upper = String(raw || "").toUpperCase();
  const codes = upper.match(/[A-Z]{2}\d{11}/g);
  if (codes && codes.length > 0) {
    return codes.map((c) => `${c.slice(0, 2)}-${c.slice(2, 5)}-${c.slice(5, 9)}-${c.slice(9, 13)}`).join("\n");
  }
  const dashCodes = upper.match(/[A-Z]{2}-\d{3}-\d{4}-\d{4}/g);
  if (dashCodes && dashCodes.length > 0) return dashCodes.join("\n");
  return String(raw || "").trim();
}

function parseDateValue(val) {
  let dt = null;
  if (val instanceof Date && !isNaN(val)) {
    dt = new Date(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate()));
  } else if (typeof val === "number" && val > 20000 && val < 80000) {
    // Excel serial date
    const utcDays = Math.floor(val - 25569);
    dt = new Date(utcDays * 86400 * 1000);
  } else if (val != null) {
    const s = String(val).trim();
    const m1 = s.match(/(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})/);
    if (m1) {
      dt = new Date(Date.UTC(Number(m1[1]), Number(m1[2]) - 1, Number(m1[3])));
    } else {
      const m2 = s.match(/(\d{2,3})[-/年.](\d{1,2})[-/月.](\d{1,2})/);
      if (m2) {
        dt = new Date(Date.UTC(Number(m2[1]) + 1911, Number(m2[2]) - 1, Number(m2[3])));
      }
    }
  }

  if (!dt || isNaN(dt)) {
    return {
      raw: String(val || ""),
      roc_year: "",
      month: "",
      day: "",
      roc_date_text: "",
      roc_datetime_default: "中華民國   年   月   日   時   分",
      weekday: null,
      is_weekend: false,
    };
  }

  const y = dt.getUTCFullYear();
  const m = dt.getUTCMonth() + 1;
  const d = dt.getUTCDate();
  const rocY = y - 1911;
  const mm = String(m).padStart(2, "0");
  const dd = String(d).padStart(2, "0");
  const wd = dt.getUTCDay(); // 0=Sun, 6=Sat
  const wdNames = ["日", "一", "二", "三", "四", "五", "六"];

  return {
    raw: `${y}-${mm}-${dd}`,
    roc_year: String(rocY),
    month: mm,
    day: dd,
    roc_date_text: `中華民國${rocY}年${mm}月${dd}日`,
    roc_datetime_default: `中華民國${rocY}年${mm}月${dd}日10時00分`,
    weekday: `星期${wdNames[wd]}`,
    is_weekend: wd === 0 || wd === 6,
  };
}

export function parseDeficiencyWorkbook(arrayBuffer) {
  const wb = window.XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  let sheetName = wb.SheetNames[0];
  for (const name of wb.SheetNames) {
    if (name.includes("督導") || name.includes("查核") || name.includes("缺失")) {
      sheetName = name;
      break;
    }
  }
  const ws = wb.Sheets[sheetName];
  const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const meta = {
    title: "",
    project_name: "",
    project_code: "",
    inspection_date: null,
    supervision_unit: "",
    contractor: "",
    contract_amount: "",
    site_manager: "",
    safety_officer: "",
  };

  const findRight = (rowArr, colIdx) => {
    for (let c = colIdx + 1; c < rowArr.length; c++) {
      if (rowArr[c] != null && String(rowArr[c]).trim() !== "") {
        return rowArr[c];
      }
    }
    return null;
  };

  for (let r = 0; r < Math.min(16, rows.length); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] == null) continue;
      const s = String(row[c]).trim();
      if (r === 0 && !meta.title) meta.title = s;
      const lbl = s.replace(/\s+/g, "");
      if (lbl === "工程名稱" || lbl === "工程名稱：") {
        const rawProj = String(findRight(row, c) || "");
        const lines = rawProj.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
        if (lines.length > 0) {
          meta.project_name = lines[0];
          if (lines.length > 1) {
            meta.project_code = formatProjectCode(lines.slice(1).join(" "));
          } else {
            const m = lines[0].match(/[（(]([^）)]+)[）)]/);
            if (m && /[A-Z]{2}\d/.test(m[1])) {
              meta.project_code = formatProjectCode(m[1]);
              meta.project_name = lines[0].slice(0, m.index).trim();
            }
          }
        }
      } else if (lbl === "工程編號" || lbl === "工程編號：") {
        meta.project_code = formatProjectCode(String(findRight(row, c) || ""));
      } else if (["督導日期", "查核日期", "稽查日期", "檢查日期"].includes(lbl)) {
        meta.inspection_date = parseDateValue(findRight(row, c));
      } else if (lbl === "監造單位" || lbl === "監造單位：") {
        meta.supervision_unit = String(findRight(row, c) || "").trim();
      } else if (["承攬商", "承攬廠商", "廠商名稱"].includes(lbl)) {
        meta.contractor = String(findRight(row, c) || "").trim();
      } else if (["契約金額", "工程金額"].includes(lbl)) {
        meta.contract_amount = String(findRight(row, c) || "").trim();
      } else if (["工作場所負責人", "工地負責人"].includes(lbl)) {
        meta.site_manager = String(findRight(row, c) || "").trim();
      } else if (["職安人員", "職業安全衛生人員"].includes(lbl)) {
        meta.safety_officer = String(findRight(row, c) || "").trim();
      }
    }
  }

  if (!meta.inspection_date) meta.inspection_date = parseDateValue(null);

  const deficiencies = [];
  let currentSection = "";
  let inTable = false;
  let colSeq = 0;
  let colCat = 1;
  let colDesc = 6;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const texts = row.map((v) => (v == null ? "" : String(v).trim()));
    const joined = texts.join("");
    if (!joined) {
      inTable = false;
      continue;
    }
    if (["承攬商簽名", "監造單位簽名", "受查單位", "督導委員簽名"].some((k) => joined.includes(k))) {
      break;
    }
    if (joined.includes("書面資料") && (joined.includes("壹") || joined.length < 20)) {
      currentSection = "壹、書面資料";
      inTable = false;
      continue;
    }
    if (joined.includes("現場作業") && (joined.includes("貳") || joined.length < 20)) {
      currentSection = "貳、現場作業";
      inTable = false;
      continue;
    }
    if (joined.includes("序號") && (joined.includes("違反事項") || joined.includes("缺失事項") || joined.includes("說明"))) {
      inTable = true;
      for (let c = 0; c < row.length; c++) {
        const hv = texts[c].replace(/\s+/g, "");
        if (hv === "序號") colSeq = c;
        else if (["違反事項", "缺失事項", "項目"].includes(hv)) colCat = c;
        else if (["說明", "缺失說明", "缺失情形", "內容"].includes(hv)) colDesc = c;
      }
      continue;
    }

    if (inTable || (currentSection && row[colSeq] != null)) {
      const seqVal = row[colSeq];
      const catVal = texts[colCat] || "";
      let descVal = texts[colDesc] || "";
      if (!descVal) {
        for (let c = colCat + 1; c < texts.length; c++) {
          if (texts[c]) {
            descVal = texts[c];
            break;
          }
        }
      }
      if (!catVal && !descVal) continue;
      deficiencies.push({
        index: deficiencies.length + 1,
        section: currentSection || "稽查缺失",
        seq: seqVal != null ? String(seqVal).trim() : String(deficiencies.length + 1),
        category: catVal,
        description: descVal,
      });
    }
  }

  return { meta, deficiencies };
}

function charNgrams(text) {
  const clean = String(text || "").replace(/[\s\W_]+/g, "");
  const grams = new Set();
  for (const n of [2, 3]) {
    for (let i = 0; i <= clean.length - n; i++) {
      grams.add(clean.slice(i, i + n));
    }
  }
  if (grams.size === 0 && clean) {
    for (const ch of clean) grams.add(ch);
  }
  return grams;
}

function similarityScore(q, cText) {
  const qg = charNgrams(q);
  const cg = charNgrams(cText);
  if (qg.size === 0 || cg.size === 0) return 0;
  let inter = 0;
  for (const g of qg) if (cg.has(g)) inter++;
  return inter / Math.sqrt(qg.size * cg.size);
}

function inferSubitemDetails(clauseId, combined) {
  if (clauseId === "2.18") {
    const subs = [
      ["1.工作場所簽到", /簽到/],
      ["2.生理協調平衡檢測", /生理協調/],
      ["3.額溫量測(上工前)", /額溫/],
      ["4.危害因素告知", /危害告知|危害因素/],
      ["5.工作場所簽退", /簽退|簽到退/],
      ["6.各項作業拍攝照片", /拍攝照片|拍照/],
      ["7.做成紀錄表", /紀錄表/],
    ];
    const hit = subs.filter(([, re]) => re.test(combined)).map(([lbl]) => lbl);
    if (hit.length > 0) return [hit.length, `未依規定辦理${hit.join("")}`];
  } else if (clauseId === "2.35") {
    const roles = [
      /工地主任|工作場所負責人|工地負責人/,
      /職安人員|職業安全衛生管理人員/,
      /急救人員/,
      /作業主管/,
    ];
    const cnt = roles.reduce((acc, re) => acc + (re.test(combined) ? 1 : 0), 0);
    if (cnt >= 1) return [cnt, `缺席角色數：${cnt}位`];
  }
  return [1, ""];
}

export function matchDeficienciesClient(parsed, clausesData) {
  const clauses = clausesData.clauses || [];
  const byId = Object.fromEntries(clauses.map((c) => [c.id, c]));
  const meta = parsed.meta || {};
  const deficiencies = parsed.deficiencies || [];

  const items = deficiencies.map((item) => {
    const combined = `${item.category || ""} ${item.description || ""}`.trim();
    const scores = new Map();

    for (const [pat, targets] of DOMAIN_RULES) {
      if (pat.test(combined)) {
        for (const [cid, boost, reason] of targets) {
          if (!byId[cid]) continue;
          const prev = scores.get(cid) || [0, []];
          const nextScore = boost > prev[0] ? boost : prev[0];
          const nextReasons = prev[1].includes(reason) ? prev[1] : [...prev[1], reason];
          scores.set(cid, [nextScore, nextReasons]);
        }
      }
    }

    for (const c of clauses) {
      const sim = similarityScore(combined, `${c.category} ${c.text}`);
      const prev = scores.get(c.id) || [0, []];
      const blended = prev[0] === 0 ? sim * 0.75 : Math.min(0.99, prev[0] + sim * 0.05);
      scores.set(c.id, [blended, prev[1]]);
    }

    const ranked = Array.from(scores.entries()).sort((a, b) => b[1][0] - a[1][0]);
    const candidates = ranked.slice(0, 6).map(([cid, [sc, reasons]]) => {
      const cl = byId[cid];
      let overlapWarning = null;
      if (cl.overlaps && cl.overlaps.length > 0) {
        const msgs = [];
        for (const oid of cl.overlaps) {
          const ocl = byId[oid];
          if (ocl && ocl.fine != null && cl.fine != null) {
            if (ocl.fine > cl.fine) {
              msgs.push(`${oid}（${ocl.fine.toLocaleString()}元 > ${cl.fine.toLocaleString()}元，依備註取大值）`);
            } else if (ocl.fine === cl.fine) {
              msgs.push(`${oid}（同為${ocl.fine.toLocaleString()}元，附註另有1個月禁入原則）`);
            }
          }
        }
        if (msgs.length > 0) overlapWarning = "涉職安保命條款：" + msgs.join("；");
      }
      return {
        clause_id: cid,
        source: cl.source,
        category: cl.category,
        text: cl.text,
        fine: cl.fine,
        fine_text: cl.fine_text,
        per_subitem: cl.per_subitem,
        confidence: Math.round(sc * 1000) / 1000,
        reasons,
        overlap_warning: overlapWarning,
      };
    });

    const best = candidates[0];
    const [count, subNote] = inferSubitemDetails(best.clause_id, combined);
    const unitFine = best.fine != null ? best.fine : 0;
    const ticketText = best.clause_id === "2.18" && subNote ? `承攬商於上工前，${subNote}。` : best.text;

    return {
      deficiency_index: item.index,
      section: item.section,
      seq: item.seq,
      raw_category: item.category,
      raw_description: item.description,
      selected: true,
      clause_id: best.clause_id,
      source: best.source,
      clause_text: ticketText,
      unit_fine: unitFine,
      count,
      repeat_level: 0,
      confidence: best.confidence,
      reasons: best.reasons,
      overlap_warning: best.overlap_warning,
      candidates,
    };
  });

  const suggestions = [];
  if (meta.inspection_date?.is_weekend && byId["1.2"]) {
    suggestions.push({
      clause_id: "1.2",
      title: `督導日期（${meta.inspection_date.roc_date_text}）為${meta.inspection_date.weekday}`,
      message: "若承攬商未經同意於假日擅自施工，可增列「1.2 未經同意假日或夜間擅自施工者」(20,000元)。",
      clause: byId["1.2"],
    });
  }

  return { meta, items, suggestions };
}

export function formatViolationLine(it) {
  if (it.custom_line && it.custom_line.trim()) return it.custom_line.trim();
  const cid = String(it.clause_id || "").trim();
  const text = String(it.clause_text || "").trim();
  const unitFine = Number(it.unit_fine) || 0;
  const count = Math.max(1, Number(it.count) || 1);
  const mult = { 0: 1, 1: 2, 2: 3, 3: 4 }[Number(it.repeat_level) || 0] || 1;
  const subtotal = unitFine * count * mult;

  let body = text;
  if (cid && !text.startsWith(cid)) {
    body = cid.startsWith("天條") ? `職安保命條款(${cid})：${text}` : `${cid}${text}`;
  }
  if (unitFine <= 0) return `${body}(依契約規定辦理)`;
  if (count === 1 && mult === 1) return `${body}(${unitFine.toLocaleString()}元)`;
  if (mult === 1) return `${body}(${unitFine.toLocaleString()}元×${count}=${subtotal.toLocaleString()}元)`;
  if (count === 1) return `${body}(${unitFine.toLocaleString()}元×重複${mult}倍=${subtotal.toLocaleString()}元)`;
  return `${body}(${unitFine.toLocaleString()}元×${count}×重複${mult}倍=${subtotal.toLocaleString()}元)`;
}

export function calculateTotal(items) {
  return items.reduce((sum, it) => {
    if (!it.selected) return sum;
    const fine = Number(it.unit_fine) || 0;
    const cnt = Math.max(1, Number(it.count) || 1);
    const mult = { 0: 1, 1: 2, 2: 3, 3: 4 }[Number(it.repeat_level) || 0] || 1;
    return sum + fine * cnt * mult;
  }, 0);
}

function getUniqueRowCells(tr) {
  return Array.from(tr.children).filter((el) => el.localName === "tc");
}

function replaceCellLines(xmlDoc, tc, lines, sizePt, lineTwips = null, lineRule = "atLeast") {
  const existingPs = Array.from(tc.children).filter((el) => el.localName === "p");
  let templatePPr = null;
  if (existingPs.length > 0) {
    const pPr = Array.from(existingPs[0].children).find((el) => el.localName === "pPr");
    if (pPr) templatePPr = pPr.cloneNode(true);
  }
  for (const p of existingPs) tc.removeChild(p);

  const safeLines = lines.length > 0 ? lines : [""];
  const szHalfPt = String(sizePt * 2);

  for (const text of safeLines) {
    const p = xmlDoc.createElementNS(W_NS, "w:p");
    if (templatePPr) {
      const pPr = templatePPr.cloneNode(true);
      if (lineTwips !== null) {
        let spacing = Array.from(pPr.children).find((el) => el.localName === "spacing");
        if (!spacing) {
          spacing = xmlDoc.createElementNS(W_NS, "w:spacing");
          pPr.appendChild(spacing);
        }
        spacing.setAttributeNS(W_NS, "w:line", String(lineTwips));
        spacing.setAttributeNS(W_NS, "w:lineRule", lineRule);
      }
      p.appendChild(pPr);
    }

    if (text) {
      const r = xmlDoc.createElementNS(W_NS, "w:r");
      const rPr = xmlDoc.createElementNS(W_NS, "w:rPr");
      const rFonts = xmlDoc.createElementNS(W_NS, "w:rFonts");
      rFonts.setAttributeNS(W_NS, "w:ascii", "標楷體");
      rFonts.setAttributeNS(W_NS, "w:eastAsia", "標楷體");
      rFonts.setAttributeNS(W_NS, "w:hAnsi", "標楷體");
      rFonts.setAttributeNS(W_NS, "w:cs", "標楷體");
      const sz = xmlDoc.createElementNS(W_NS, "w:sz");
      sz.setAttributeNS(W_NS, "w:val", szHalfPt);
      const szCs = xmlDoc.createElementNS(W_NS, "w:szCs");
      szCs.setAttributeNS(W_NS, "w:val", szHalfPt);
      rPr.appendChild(rFonts);
      rPr.appendChild(sz);
      rPr.appendChild(szCs);

      const t = xmlDoc.createElementNS(W_NS, "w:t");
      t.setAttribute("xml:space", "preserve");
      t.textContent = text;

      r.appendChild(rPr);
      r.appendChild(t);
      p.appendChild(r);
    }
    tc.appendChild(p);
  }
}

export async function generateTicketDocxBlob(ticketData, templateArrayBuffer) {
  const zip = await window.JSZip.loadAsync(templateArrayBuffer);
  const docXmlStr = await zip.file("word/document.xml").async("string");
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(docXmlStr, "application/xml");

  const tbl = xmlDoc.getElementsByTagNameNS(W_NS, "tbl")[0];
  const trs = Array.from(tbl.children).filter((el) => el.localName === "tr");

  const r0 = getUniqueRowCells(trs[0]);
  const r1 = getUniqueRowCells(trs[1]);
  const r2 = getUniqueRowCells(trs[2]);
  const r3 = getUniqueRowCells(trs[3]);
  const r4 = getUniqueRowCells(trs[4]);

  replaceCellLines(xmlDoc, r0[1], [ticketData.project_name || ""], 14);
  const codeLines = String(ticketData.project_code || "").split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  replaceCellLines(xmlDoc, r0[3], codeLines.length ? codeLines : [""], 14);

  replaceCellLines(xmlDoc, r1[1], [ticketData.contractor || ""], 14);
  replaceCellLines(xmlDoc, r1[3], [ticketData.safety_officer || ""], 14);

  const selected = (ticketData.items || []).filter((it) => it.selected !== false);
  const total = ticketData.override_total != null ? Number(ticketData.override_total) : calculateTotal(selected);

  const vLines = [
    `[發生時間]：${(ticketData.occur_time || "").trim()}`,
    `[發生地點]：${(ticketData.occur_location || "").trim()}`,
    ticketData.violation_summary?.trim() ? `[違規情形]：${ticketData.violation_summary.trim()}` : "[違規情形]：",
    `[罰款金額]：${total.toLocaleString()}元`,
    ...selected.map(formatViolationLine),
  ];
  replaceCellLines(xmlDoc, r2[1], vLines, 12, 260, "exact");

  const chineseTotal = intToChineseCurrency(total);
  replaceCellLines(xmlDoc, r3[1], [`新台幣(大寫)：${chineseTotal}`], 14);

  const ev = String(ticketData.evidence_note || "").trim();
  const noteLines = ev ? ["[佐證資料] ", ...ev.split(/\r?\n/)] : ["[佐證資料] ", ""];
  replaceCellLines(xmlDoc, r4[1], noteLines, 14);

  const serializer = new XMLSerializer();
  zip.file("word/document.xml", serializer.serializeToString(xmlDoc));
  return await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
