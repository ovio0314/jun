/*
 * 전표 검토 도우미 — 브라우저판 엔진 (Python voucher/ 패키지의 규칙을 그대로 옮김)
 * 브라우저와 Node 양쪽에서 동작한다. 외부 네트워크를 쓰지 않는다.
 * 의존: Big (big.js, 정확한 십진 계산), ExcelJS (xlsx 읽기/쓰기)
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("big.js"));
  } else {
    root.VoucherEngine = factory(root.Big);
  }
})(typeof self !== "undefined" ? self : this, function (Big) {
  "use strict";

  // ------------------------------------------------------------------ 열 정의 ----
  const EXPECTED_COLUMNS = [
    "회계단위", "회계일", "전표분개유형", "전표관리단위", "기표번호", "전표기표번호", "행번호",
    "기표자", "기표부서", "기표일", "승인자", "승인부서", "승인번호", "계정과목코드", "계정과목",
    "비용구분", "차변금액", "대변금액", "외화차변금액", "외화대변금액", "적요", "통화", "환율",
    "귀속부서", "활동센터", "예산부서", "증빙", "관계회사", "전표종류그룹", "원가항목", "원가항목명",
    "거래처사업자번호", "전자결재진행상태", "출납예정일", "출납방법",
    "관리항목1", "관리항목2", "관리항목3", "관리항목4", "관리항목5",
    "관리항목6", "관리항목7", "관리항목8", "관리항목9", "관리항목10", "파일첨부여부",
  ];
  const REQUIRED_COLUMNS = ["회계단위", "전표관리단위", "기표번호", "계정과목코드", "계정과목", "차변금액", "대변금액"];
  const AMOUNT_COLUMNS = ["차변금액", "대변금액", "외화차변금액", "외화대변금액", "환율"];
  const DATE_COLUMNS = ["회계일", "기표일", "출납예정일"];
  const FLAG_COLUMNS = ["파일첨부여부"];
  const MANAGEMENT_ITEM_COLUMNS = Array.from({ length: 10 }, (_, i) => `관리항목${i + 1}`);
  const MIN_HEADER_MATCH = 5;
  const HEADER_SCAN_ROWS = 30;

  const ZERO = new Big(0);
  const SIDE_DEBIT = "차변", SIDE_CREDIT = "대변", SIDE_BOTH = "차대동시", SIDE_NONE = "금액없음", SIDE_ERROR = "금액오류";
  const PASS = "통과", ERROR = "오류", REVIEW = "확인 필요", UNCHECKED = "미검사", NA = "해당 없음";
  const SUMMARY_PASS = "설정된 검사 통과";
  const SUMMARY_ORDER = [ERROR, REVIEW, UNCHECKED, SUMMARY_PASS];
  const AREA_DATA = "데이터 점검", AREA_AMOUNT = "금액 검사", AREA_ACCOUNT = "계정 검사", AREA_MEMO = "적요 검사", AREA_EVIDENCE = "증빙 검사";
  const AREAS = [AREA_DATA, AREA_AMOUNT, AREA_ACCOUNT, AREA_MEMO, AREA_EVIDENCE];
  const PRIORITY = { [ERROR]: 0, [REVIEW]: 1, [UNCHECKED]: 2, [PASS]: 3, [NA]: 4 };
  const ROUNDING = { "버림": 0, "반올림": 1, "올림": 3 }; // big.js rounding modes
  const APPROVED_STATES = new Set(["승인", "결재완료", "완료", "승인완료"]);
  const ERP_FLAG_NOTE = "ERP 첨부 표시 FALSE, 다우오피스 증빙 미확인";
  const REVIEW_STATES = ["미검토", "검토중", "검토완료"];

  class VoucherFileError extends Error {}

  // ------------------------------------------------------------------ 문자열 ----
  const WS_RE = /[\s 　​﻿]+/g;
  function pad(n) { return String(n).padStart(2, "0"); }
  function isoDate(d) { return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }

  /** 셀 값을 원본에 가깝게 문자열로 (숫자로 저장된 코드 포함). */
  function rawText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v instanceof Date) {
      const t = v.getUTCHours() + v.getUTCMinutes() + v.getUTCSeconds();
      return t === 0 ? isoDate(v) : `${isoDate(v)} ${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}:${pad(v.getUTCSeconds())}`;
    }
    if (typeof v === "number") return Number.isInteger(v) ? v.toFixed(0) : String(v);
    return String(v);
  }
  function cleanText(v) {
    if (v === null || v === undefined) return "";
    return rawText(v).replace(WS_RE, " ").trim();
  }
  function norm(t) { return String(t == null ? "" : t).replace(/\s+/g, "").toLowerCase(); }
  function normSpace(t) { return String(t == null ? "" : t).replace(/\s+/g, " ").trim(); }

  /** ExcelJS 셀 값 → 원시 값 (문자열/숫자/불리언/Date/null) */
  function cellValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date || typeof v !== "object") return v;
    if (Array.isArray(v.richText)) return v.richText.map((p) => p.text).join("");
    if ("result" in v) return cellValue(v.result);
    if ("text" in v) return v.text;
    if ("error" in v) return String(v.error);
    return String(v);
  }

  // ------------------------------------------------------------------ 금액 ----
  /** 금액 해석 → {amt: Big|null, st: "정상"|"빈값"|"오류"}. 해석 불가를 0 으로 바꾸지 않는다. */
  function parseAmount(v) {
    if (v === null || v === undefined) return { amt: null, st: "빈값" };
    if (typeof v === "boolean" || v instanceof Date) return { amt: null, st: "오류" };
    if (v instanceof Big) return { amt: v, st: "정상" };
    if (typeof v === "number") {
      if (!Number.isFinite(v)) return { amt: null, st: "오류" };
      return { amt: new Big(String(v)), st: "정상" };
    }
    let text = String(v).replace(WS_RE, "");
    if (text === "") return { amt: null, st: "빈값" };
    text = text.replace(/,/g, "");
    if (text.endsWith("원")) text = text.slice(0, -1);
    let negative = false;
    if (text.startsWith("(") && text.endsWith(")")) { negative = true; text = text.slice(1, -1); }
    for (const mark of ["△", "▲", "-"]) {
      if (text.startsWith(mark)) { negative = !negative; text = text.slice(mark.length); break; }
    }
    if (text.startsWith("+")) text = text.slice(1);
    if (!/^\d+(\.\d+)?$/.test(text)) return { amt: null, st: "오류" };
    const amt = new Big(text);
    return { amt: negative ? amt.times(-1) : amt, st: "정상" };
  }

  function classifyFlag(v) {
    if (v === null || v === undefined) return "빈값";
    if (typeof v === "boolean") return v ? "불리언 TRUE" : "불리언 FALSE";
    const t = String(v).trim();
    if (t === "") return "빈값";
    if (t.toUpperCase() === "TRUE") return "문자열 TRUE";
    if (t.toUpperCase() === "FALSE") return "문자열 FALSE";
    return `기타값(${t})`;
  }

  /** 날짜 → "YYYY-MM-DD" 또는 null */
  function parseDate(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return isoDate(v);
    const t = cleanText(v);
    if (!t) return null;
    const d = t.replace(/[^\d]/g, "");
    if (d.length === 8) {
      const y = +d.slice(0, 4), m = +d.slice(4, 6), day = +d.slice(6);
      const dt = new Date(Date.UTC(y, m - 1, day));
      if (dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === day) return isoDate(dt);
    }
    return null;
  }

  function fmt(v) {
    if (v === null || v === undefined) return "";
    if (v instanceof Big) {
      const integral = v.eq(v.round(0, 0));
      const s = integral ? v.toFixed(0) : v.toString();
      const neg = s.startsWith("-");
      const [i, f] = (neg ? s.slice(1) : s).split(".");
      const withComma = i.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      return (neg ? "-" : "") + withComma + (f ? "." + f : "");
    }
    return String(v);
  }
  const amt = (v) => (v instanceof Big ? v : ZERO);
  const sum = (arr) => arr.reduce((a, b) => a.plus(b), ZERO);

  // ------------------------------------------------------------------ 엑셀 읽기 ----
  async function loadWorkbook(ExcelJS, buffer) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      return wb;
    } catch (e) {
      throw new VoucherFileError("엑셀 파일을 읽을 수 없습니다. .xlsx 형식인지, 암호가 걸려 있지 않은지 확인해 주세요. (상세: " + (e && e.message ? e.message.slice(0, 80) : e) + ")");
    }
  }

  function sheetRows(ws) {
    const rows = [];
    const ncol = ws.actualColumnCount ? Math.max(ws.columnCount, ws.actualColumnCount) : ws.columnCount;
    ws.eachRow({ includeEmpty: true }, (row, n) => {
      const vals = [];
      for (let c = 1; c <= ncol; c++) vals.push(cellValue(row.getCell(c).value));
      rows[n - 1] = vals;
    });
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  /** 모든 시트에서 헤더 후보를 찾는다. 반환: [{sheet, headerRow, matched, title, headers, preamble, rows}] */
  function findHeaderCandidates(wb) {
    const expected = new Set(EXPECTED_COLUMNS);
    const out = [];
    wb.eachSheet((ws) => {
      const rows = sheetRows(ws);
      let best = null, title = "";
      const preamble = [];
      for (let i = 0; i < Math.min(HEADER_SCAN_ROWS, rows.length); i++) {
        const headers = rows[i].map(cleanText);
        const matched = new Set(headers.filter((h) => h && expected.has(h))).size;
        if (matched >= MIN_HEADER_MATCH && (!best || matched > best.matched)) {
          best = { sheet: ws.name, headerRow: i + 1, matched, title, headers, preamble: preamble.slice(), rows };
        }
        if (!best && matched < MIN_HEADER_MATCH) {
          const text = headers.filter(Boolean).join(" ");
          if (text) { preamble.push(text); title = title || headers.find(Boolean); }
        }
      }
      if (best) out.push(best);
    });
    out.sort((a, b) => b.matched - a.matched);
    return out;
  }

  function autoMapping(headers) {
    const m = {};
    for (const h of headers) if (EXPECTED_COLUMNS.includes(h) && !(h in m)) m[h] = h;
    return m;
  }

  function duplicateHeaders(headers) {
    const seen = new Set(), dups = [];
    for (const h of headers) {
      if (!h) continue;
      if (seen.has(h) && !dups.includes(h)) dups.push(h);
      seen.add(h);
    }
    return dups;
  }

  function voucherKey(r) {
    const no = r["기표번호"] || `(기표번호없음-${r._엑셀행}행)`;
    return `${r["회계단위"] || ""}|${r["전표관리단위"] || ""}|${no}`;
  }

  /** 선택한 후보 기준으로 데이터를 읽는다. 반환: {raw, rows, mapping, sheet, headerRow, title, preamble, errors, warnings} */
  function loadVouchers(cand, mapping) {
    const headers = cand.headers;
    mapping = Object.assign({}, mapping || autoMapping(headers));
    const errors = [], warnings = [];
    for (const d of duplicateHeaders(headers)) {
      if (Object.values(mapping).includes(d) && REQUIRED_COLUMNS.includes(d)) errors.push(`필수 열 '${d}' 헤더가 중복되어 있습니다. 원본에서 중복 열을 확인해 주세요.`);
      else warnings.push(`헤더 '${d}' 가 중복되어 있어 첫 번째 열만 사용합니다.`);
    }
    const missing = REQUIRED_COLUMNS.filter((c) => !(c in mapping));
    if (missing.length) errors.push("필수 열이 없습니다: " + missing.join(", ") + ". 열 매핑을 확인해 주세요.");
    const optMissing = EXPECTED_COLUMNS.filter((c) => !(c in mapping) && !REQUIRED_COLUMNS.includes(c));
    if (optMissing.length) warnings.push("선택 열 누락(해당 정보는 표시되지 않음): " + optMissing.join(", "));
    const colIndex = {};
    for (const [std, src] of Object.entries(mapping)) { const i = headers.indexOf(src); if (i >= 0) colIndex[std] = i; }

    const raw = [], rows = [];
    for (let i = cand.headerRow; i < cand.rows.length; i++) {
      const row = cand.rows[i] || [];
      if (row.every((v) => v === null || v === undefined || cleanText(v) === "")) continue;
      const excelRow = i + 1;
      const rr = { _엑셀행: excelRow }, nr = { _엑셀행: excelRow };
      for (const std of EXPECTED_COLUMNS) {
        const idx = colIndex[std];
        const v = idx !== undefined && idx < row.length ? row[idx] : null;
        rr[std] = rawText(v);
        if (AMOUNT_COLUMNS.includes(std)) { const p = parseAmount(v); nr[std] = p.amt; nr[std + "_상태"] = p.st; }
        else if (DATE_COLUMNS.includes(std)) nr[std] = parseDate(v);
        else if (FLAG_COLUMNS.includes(std)) { nr[std] = cleanText(v); nr[std + "_구분"] = classifyFlag(v); }
        else nr[std] = cleanText(v);
      }
      raw.push(rr); rows.push(nr);
    }
    if (!rows.length) errors.push("헤더 아래에 데이터 행이 없습니다.");
    for (const r of rows) r.전표키 = voucherKey(r);
    const blank = rows.filter((r) => r["기표번호"] === "").length;
    if (blank) warnings.push(`기표번호가 비어 있는 행 ${blank}건 — 다른 행과 묶지 않고 행별로 따로 표시합니다.`);
    const bad = rows.filter((r) => r["차변금액_상태"] === "오류" || r["대변금액_상태"] === "오류").length;
    if (bad) warnings.push(`원화 금액을 해석할 수 없는 행 ${bad}건 — 0 으로 대체하지 않고 합계에서 제외했습니다.`);
    return { raw, rows, mapping, sheet: cand.sheet, headerRow: cand.headerRow, title: cand.title, preamble: cand.preamble, errors, warnings };
  }

  // ------------------------------------------------------------------ 차대 구분 ----
  function classifySide(r) {
    if (r["차변금액_상태"] === "오류" || r["대변금액_상태"] === "오류") return SIDE_ERROR;
    const d = amt(r["차변금액"]), c = amt(r["대변금액"]);
    if (!d.eq(0) && !c.eq(0)) return SIDE_BOTH;
    if (!d.eq(0)) return SIDE_DEBIT;
    if (!c.eq(0)) return SIDE_CREDIT;
    return SIDE_NONE;
  }
  function withSide(rows) { return rows.map((r, i) => Object.assign({}, r, { 차대구분: classifySide(r), _i: i })); }
  function groupBy(rows, keyFn) {
    const m = new Map();
    for (const r of rows) { const k = keyFn(r); if (!m.has(k)) m.set(k, []); m.get(k).push(r); }
    return m;
  }
  const groupRecords = (rows) => groupBy(rows, (r) => r.전표키);
  function totals(recs) {
    const ok = recs.filter((r) => r.차대구분 !== SIDE_ERROR);
    return { d: sum(ok.map((r) => amt(r["차변금액"]))), c: sum(ok.map((r) => amt(r["대변금액"]))), allOk: ok.length === recs.length };
  }
  function top(values, n = 3) {
    const c = new Map();
    for (const v of values) if (v) c.set(v, (c.get(v) || 0) + 1);
    return [...c.entries()].sort((a, b) => b[1] - a[1]).slice(0, n).map(([k, v]) => `${k} (${v})`).join(" / ");
  }

  // ------------------------------------------------------------------ 계정과목 현황 ----
  function accountInventory(rows) {
    const data = withSide(rows);
    const g = groupBy(data, (r) => r["계정과목코드"] + "\u0000" + r["계정과목"]);
    const out = [];
    for (const [k, recs] of [...g.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1))) {
      const [code, name] = k.split("\u0000");
      const cnt = (s) => recs.filter((r) => r.차대구분 === s).length;
      const nd = cnt(SIDE_DEBIT), nc = cnt(SIDE_CREDIT);
      let usage = "판정 불가";
      if (nd && !nc) usage = "차변 전용"; else if (nc && !nd) usage = "대변 전용";
      else if (nd && nc) usage = nd >= nc * 3 ? "차변 위주" : nc >= nd * 3 ? "대변 위주" : "양쪽 사용";
      out.push({
        계정과목코드: code, 계정과목: name, 사용구분: usage, 차변행수: nd, 대변행수: nc,
        차대동시행수: cnt(SIDE_BOTH), 금액없음행수: cnt(SIDE_NONE), 금액오류행수: cnt(SIDE_ERROR),
        차변합계: sum(recs.filter((r) => r.차대구분 === SIDE_DEBIT || r.차대구분 === SIDE_BOTH).map((r) => amt(r["차변금액"]))),
        대변합계: sum(recs.filter((r) => r.차대구분 === SIDE_CREDIT || r.차대구분 === SIDE_BOTH).map((r) => amt(r["대변금액"]))),
        사용전표수: new Set(recs.map((r) => r.전표키)).size,
        비용구분: top(recs.map((r) => r["비용구분"])), "기표부서(상위)": top(recs.map((r) => r["기표부서"])),
        "적요예시(상위)": top(recs.map((r) => r["적요"])),
      });
    }
    return out;
  }

  function codeNameConflicts(rows) {
    const out = [];
    const byCode = groupBy(rows, (r) => r["계정과목코드"]);
    for (const [code, recs] of byCode) { const n = [...new Set(recs.map((r) => r["계정과목"]))].sort(); if (n.length > 1) out.push({ 유형: "코드 하나에 계정명 여러 개", 기준: code, 값: n.join(", ") }); }
    const byName = groupBy(rows, (r) => r["계정과목"]);
    for (const [name, recs] of byName) { const c = [...new Set(recs.map((r) => r["계정과목코드"]))].sort(); if (c.length > 1) out.push({ 유형: "계정명 하나에 코드 여러 개", 기준: name, 값: c.join(", ") }); }
    return out;
  }

  function counterpartAccounts(rows, code, side) {
    const data = rows[0] && rows[0].차대구분 ? rows : withSide(rows);
    const opposite = side === SIDE_DEBIT ? SIDE_CREDIT : SIDE_DEBIT;
    const keys = new Set(data.filter((r) => r["계정과목코드"] === code && r.차대구분 === side).map((r) => r.전표키));
    const sub = data.filter((r) => keys.has(r.전표키) && r.차대구분 === opposite && r["계정과목코드"] !== code);
    const g = groupBy(sub, (r) => r["계정과목코드"] + "\u0000" + r["계정과목"]);
    return [...g.entries()].map(([k, recs]) => { const [c, n] = k.split("\u0000"); return { 계정과목코드: c, 계정과목: n, 함께쓴전표수: new Set(recs.map((r) => r.전표키)).size }; })
      .sort((a, b) => b.함께쓴전표수 - a.함께쓴전표수 || (a.계정과목코드 < b.계정과목코드 ? -1 : 1));
  }

  function managementItemProfile(rows, code) {
    const g = rows.filter((r) => r["계정과목코드"] === code);
    const out = [];
    for (const col of MANAGEMENT_ITEM_COLUMNS) {
      const filled = g.filter((r) => String(r[col] || "").length > 0);
      if (!filled.length) continue;
      out.push({ 관리항목: col, 입력행수: filled.length, 전체행수: g.length, "예시값(상위)": top(filled.map((r) => r[col])) });
    }
    return out;
  }

  function fileOverview(rows) {
    const dates = rows.map((r) => r["회계일"]).filter(Boolean).sort();
    const inv = accountInventory(rows);
    const g = groupRecords(withSide(rows));
    let mismatch = 0;
    for (const recs of g.values()) { const t = totals(recs); if (!t.d.eq(t.c)) mismatch++; }
    const flags = {};
    for (const r of rows) flags[r["파일첨부여부_구분"]] = (flags[r["파일첨부여부_구분"]] || 0) + 1;
    return {
      원본행수: rows.length, 전표수: g.size, 기표번호수: new Set(rows.map((r) => r["기표번호"])).size,
      회계일_시작: dates[0] || null, 회계일_종료: dates[dates.length - 1] || null,
      계정과목수: inv.length, 차변사용계정수: inv.filter((r) => r.차변행수 > 0).length, 대변사용계정수: inv.filter((r) => r.대변행수 > 0).length,
      차대변불일치전표수: mismatch, 적요빈칸행수: rows.filter((r) => r["적요"] === "").length, 파일첨부여부구분: flags,
    };
  }

  // ------------------------------------------------------------------ 설정 ----
  const DEFAULT_SETTINGS = {
    vat: {
      account_codes: [], name_keywords: ["부가세", "부가가치세"], supply_columns: {}, general_evidence_values: [],
      excluded_evidence_keywords: ["불공제", "영세", "면세", "수정", "분할", "계산서(면세)", "수입"],
      rate: "0.1", rounding: "", tolerance: "0",
    },
    memo: { blank_status: "오류", rules: [] },
    account_rules: [],
    evidence: { require_tax_invoice_for_vat: true, require_approval_doc: true },
    required_fields: ["기표번호", "회계일", "계정과목코드", "계정과목"],
  };
  const clone = (o) => JSON.parse(JSON.stringify(o));
  function defaultSettings() { return clone(DEFAULT_SETTINGS); }
  function mergeDefaults(data) {
    const out = defaultSettings();
    for (const [k, v] of Object.entries(data || {})) {
      if (v && typeof v === "object" && !Array.isArray(v) && out[k] && typeof out[k] === "object" && !Array.isArray(out[k])) Object.assign(out[k], v);
      else out[k] = v;
    }
    return out;
  }
  function stableStringify(o) {
    if (o === null || typeof o !== "object") return JSON.stringify(o);
    if (Array.isArray(o)) return "[" + o.map(stableStringify).join(",") + "]";
    return "{" + Object.keys(o).sort().map((k) => JSON.stringify(k) + ":" + stableStringify(o[k])).join(",") + "}";
  }
  function settingsVersion(s) { return "S-" + sha256(stableStringify(s)).slice(0, 10); }

  // ------------------------------------------------------------------ 검사 ----
  function CR(key, area, code, status, actual = "", compare = "", diff = "", reason = "", ver = "", rows = "") {
    return { 전표키: key, 검사영역: area, 검사코드: code, 상태: status, 실제값: actual, 비교값: compare, 차이: diff, 근거: reason, 적용규칙버전: ver, 원본행번호: rows };
  }
  const rowsOf = (recs) => recs.map((r) => r._엑셀행).join(",");
  function areaStatus(sts) { if (!sts.length) return UNCHECKED; return sts.reduce((a, b) => (PRIORITY[b] < PRIORITY[a] ? b : a)); }
  function summaryStatus(sts) {
    if (sts.includes(ERROR)) return ERROR;
    if (sts.includes(REVIEW)) return REVIEW;
    if (sts.includes(UNCHECKED)) return UNCHECKED;
    return SUMMARY_PASS;
  }

  function checkData(data, groups, settings, ver) {
    const out = [];
    const required = settings.required_fields || [];
    const identCount = new Map(), fullCount = new Map();
    const ident = (r) => `${r["회계단위"]}|${r["전표관리단위"]}|${r["전표기표번호"]}`;
    const full = (r) => EXPECTED_COLUMNS.map((c) => { const v = r[c]; return v instanceof Big ? v.toString() : v == null ? "" : String(v); }).join("\u0001");
    for (const r of data) { identCount.set(ident(r), (identCount.get(ident(r)) || 0) + 1); fullCount.set(full(r), (fullCount.get(full(r)) || 0) + 1); }
    for (const [key, recs] of groups) {
      const res = [];
      for (const r of recs) {
        const row = String(r._엑셀행);
        const missing = required.filter((c) => r[c] === null || r[c] === undefined || r[c] === "");
        if (missing.length) res.push(CR(key, AREA_DATA, "D01", ERROR, missing.join(", "), "", "", "필수 값 누락", ver, row));
        const side = r.차대구분;
        if (side === SIDE_ERROR) {
          const bad = ["차변금액", "대변금액"].filter((c) => r[c + "_상태"] === "오류");
          res.push(CR(key, AREA_DATA, "D05", ERROR, bad.join(", "), "", "", "금액을 해석할 수 없음 (0 으로 대체하지 않음)", ver, row));
        } else if (side === SIDE_BOTH) {
          res.push(CR(key, AREA_DATA, "D04", ERROR, `차변 ${fmt(r["차변금액"])} / 대변 ${fmt(r["대변금액"])}`, "", "", "한 행에 차변·대변 동시 입력", ver, row));
        } else if (side === SIDE_NONE) {
          res.push(CR(key, AREA_DATA, "D06", REVIEW, "차변·대변 모두 0 또는 빈값", "", "", "금액 없는 분개 행", ver, row));
        }
      }
      const di = recs.filter((r) => r["전표기표번호"] !== "" && identCount.get(ident(r)) > 1);
      if (di.length) res.push(CR(key, AREA_DATA, "D02", REVIEW, [...new Set(di.map((r) => r["전표기표번호"]))].sort().join(", "), "", "", "동일 분개 식별자(전표기표번호) 중복 — 자동 삭제하지 않음", ver, rowsOf(di)));
      const df = recs.filter((r) => fullCount.get(full(r)) > 1);
      if (df.length) res.push(CR(key, AREA_DATA, "D03", REVIEW, `${df.length}행`, "", "", "모든 열이 같은 중복 행 — 추출 중복 여부 확인 (자동 삭제하지 않음)", ver, rowsOf(df)));
      if (!res.length) res.push(CR(key, AREA_DATA, "D00", PASS, `${recs.length}행`, "", "", "데이터 점검 이상 없음", ver, rowsOf(recs)));
      out.push(...res);
    }
    return out;
  }

  function isVatRow(r, vat) {
    if ((vat.account_codes || []).includes(r["계정과목코드"])) return true;
    const name = String(r["계정과목"] || "");
    return (vat.name_keywords || []).some((k) => k && name.includes(k));
  }
  const vatRowAmount = (r) => (r.차대구분 === SIDE_DEBIT ? amt(r["차변금액"]) : amt(r["대변금액"]));

  function checkVatRow(key, r, vat, ver) {
    const row = String(r._엑셀행);
    const code = r["계정과목코드"] || "";
    const ev = normSpace(r["증빙"]);
    const tax = vatRowAmount(r);
    const res = (st, reason, actual = "", compare = "", diff = "") => CR(key, AREA_AMOUNT, "V01", st, actual, compare, diff, `[전표 내부 계산 점검] ${reason}`, ver, row);
    if ((vat.excluded_evidence_keywords || []).some((k) => k && ev.includes(k))) return res(NA, `특수 증빙 유형('${ev}') — 일반과세 규칙으로 판정하지 않음`, fmt(tax));
    if (r.차대구분 !== SIDE_DEBIT && r.차대구분 !== SIDE_CREDIT) return res(UNCHECKED, "부가세 행 금액 판정 불가 (0/빈값/차대동시/오류)", fmt(tax));
    const missing = [];
    const col = (vat.supply_columns || {})[code];
    if (!col) missing.push("공급가액 열 매핑");
    const general = vat.general_evidence_values || [];
    if (!general.length) missing.push("일반과세 증빙 유형");
    if (!vat.rounding) missing.push("반올림 기준");
    if (missing.length) return res(UNCHECKED, "설정 필요: " + missing.join(", ") + " (관리항목2는 공급가액 '후보'일 뿐 확인 전 사용 안 함)", fmt(tax));
    if (!general.includes(ev)) return res(UNCHECKED, `증빙 유형 '${ev || "(빈값)"}' 이 일반과세 대상으로 설정되지 않음`, fmt(tax));
    const p = parseAmount(r[col]);
    if (p.st === "빈값") return res(UNCHECKED, `공급가액 열(${col}) 값 없음`, fmt(tax));
    if (p.st === "오류") return res(REVIEW, `공급가액 열(${col}) 값 해석 불가: ${r[col]}`, fmt(tax));
    let rate, tol;
    try { rate = new Big(String(vat.rate || "0.1")); tol = new Big(String(vat.tolerance || "0")); } catch (e) { return res(UNCHECKED, "세율/허용차이 설정 값 오류", fmt(tax)); }
    const expected = p.amt.times(rate).round(0, ROUNDING[vat.rounding]);
    const diff = tax.minus(expected);
    return res(diff.abs().lte(tol) ? PASS : REVIEW, `공급가액(${col}) ${fmt(p.amt)} × ${rate.toString()} (${vat.rounding}) · 허용차이 ${fmt(tol)}원`, fmt(tax), fmt(expected), fmt(diff));
  }

  function checkAmount(groups, settings, ver) {
    const out = [];
    const vat = settings.vat || {};
    for (const [key, recs] of groups) {
      const t = totals(recs);
      if (!t.allOk) out.push(CR(key, AREA_AMOUNT, "A01", UNCHECKED, `차변 ${fmt(t.d)} / 대변 ${fmt(t.c)}`, "", "", "해석 불가 금액 행이 있어 차대 합계를 확정할 수 없음", ver, rowsOf(recs)));
      else if (t.d.eq(t.c)) out.push(CR(key, AREA_AMOUNT, "A01", PASS, fmt(t.d), fmt(t.c), "0", "차변합계 = 대변합계 (원화)", ver, rowsOf(recs)));
      else out.push(CR(key, AREA_AMOUNT, "A01", ERROR, fmt(t.d), fmt(t.c), fmt(t.d.minus(t.c)), "차대변 불일치 또는 추출 누락 확인 (필터된 불완전 전표일 수 있음)", ver, rowsOf(recs)));
      const vr = recs.filter((r) => isVatRow(r, vat));
      if (!vr.length) out.push(CR(key, AREA_AMOUNT, "V00", NA, "", "", "", "부가세 행 없음", ver, ""));
      for (const r of vr) out.push(checkVatRow(key, r, vat, ver));
    }
    return out;
  }

  function ruleApplies(rule, r) {
    if (!rule.활성) return false;
    const kws = (rule.적요조건 || []).filter(Boolean);
    if (kws.length && !kws.some((k) => normSpace(r["적요"]).includes(normSpace(k)))) return false;
    const types = (rule.적용전표유형 || []).filter(Boolean);
    if (types.length && !types.includes(r["전표분개유형"])) return false;
    const side = rule.차대구분 || "전체";
    if (side !== "전체" && r.차대구분 !== side) return false;
    const day = r["회계일"];
    const start = /^\d{4}-\d{2}-\d{2}$/.test(rule.적용시작일 || "") ? rule.적용시작일 : null;
    const end = /^\d{4}-\d{2}-\d{2}$/.test(rule.적용종료일 || "") ? rule.적용종료일 : null;
    if (day && start && day < start) return false;
    if (day && end && day > end) return false;
    return true;
  }
  const ruleVer = (rule) => `${rule.규칙ID || ""}@v${rule.버전 || ""}`;

  function checkAccount(groups, settings, ver) {
    const out = [];
    const rules = (settings.account_rules || []).filter((r) => r.활성);
    const sided = (r) => r.차대구분 === SIDE_DEBIT || r.차대구분 === SIDE_CREDIT;
    for (const [key, recs] of groups) {
      const res = [], uncovered = [];
      for (const r of recs) {
        if (!sided(r)) continue;
        const hits = rules.filter((rule) => ruleApplies(rule, r));
        if (!hits.length) { uncovered.push(r); continue; }
        const allowedBy = hits.filter((h) => (h.허용계정코드 || []).includes(r["계정과목코드"]));
        const vers = hits.map(ruleVer).join(", ");
        if (allowedBy.length) res.push(CR(key, AREA_ACCOUNT, "C01", PASS, `${r["계정과목코드"]} ${r["계정과목"]}`, (allowedBy[0].허용계정코드 || []).join(", "), "", allowedBy[0].근거 || "", `${ver} / ${vers}`, String(r._엑셀행)));
        else {
          const allowed = [...new Set(hits.flatMap((h) => h.허용계정코드 || []))].sort();
          res.push(CR(key, AREA_ACCOUNT, "C01", REVIEW, `${r["계정과목코드"]} ${r["계정과목"]}`, allowed.join(", "), "", "회사 규칙의 허용 계정이 아님: " + hits.map((h) => h.근거 || "").join(" / "), `${ver} / ${vers}`, String(r._엑셀행)));
        }
      }
      for (const rule of rules) {
        const need = new Set(rule.필수상대계정코드 || []);
        const matched = recs.filter((r) => sided(r) && ruleApplies(rule, r));
        if (!need.size || !matched.length) continue;
        const codes = new Set(recs.map((r) => r["계정과목코드"]));
        const inter = [...codes].filter((c) => need.has(c)).sort();
        const mrows = matched.map((r) => r._엑셀행).join(",");
        if (inter.length) res.push(CR(key, AREA_ACCOUNT, "C02", PASS, inter.join(", "), [...need].sort().join(", "), "", `필수 계정 조합 충족: ${rule.근거 || ""}`, `${ver} / ${ruleVer(rule)}`, mrows));
        else res.push(CR(key, AREA_ACCOUNT, "C02", REVIEW, [...codes].sort().join(", "), [...need].sort().join(", "), "", `필수 계정 조합 없음: ${rule.근거 || ""}`, `${ver} / ${ruleVer(rule)}`, mrows));
      }
      if (uncovered.length) res.push(CR(key, AREA_ACCOUNT, "C00", UNCHECKED, `${uncovered.length}행`, "", "", "적용되는 회사 승인 계정 규칙 없음", ver, rowsOf(uncovered)));
      if (!res.length) res.push(CR(key, AREA_ACCOUNT, "C00", UNCHECKED, "", "", "", "검사 가능한 분개 행 없음", ver, rowsOf(recs)));
      out.push(...res);
    }
    return out;
  }

  function checkMemo(groups, settings, ver) {
    const cfg = settings.memo || {};
    const blank = cfg.blank_status || ERROR;
    const rules = (cfg.rules || []).filter((r) => r.활성);
    const out = [];
    for (const [key, recs] of groups) {
      const res = [];
      for (const r of recs) {
        const memo = normSpace(r["적요"]);
        if (!memo) { res.push(CR(key, AREA_MEMO, "M01", blank, "(빈칸)", "", "", "적요 빈칸", ver, String(r._엑셀행))); continue; }
        for (const rule of rules) {
          const codes = rule.대상계정코드 || [];
          if (codes.length && !codes.includes(r["계정과목코드"])) continue;
          const cond = (rule.적요조건 || []).filter(Boolean);
          if (cond.length && !cond.some((k) => memo.includes(normSpace(k)))) continue;
          const need = (rule.필수포함 || []).filter(Boolean);
          const lack = need.filter((k) => !memo.includes(normSpace(k)));
          res.push(CR(key, AREA_MEMO, "M02", lack.length ? REVIEW : PASS, memo, need.join(", "), lack.length ? `누락: ${lack.join(", ")}` : "", rule.근거 || "", `${ver} / ${ruleVer(rule)}`, String(r._엑셀행)));
        }
      }
      if (!res.length) res.push(CR(key, AREA_MEMO, "M00", PASS, "", "", "", "적요 빈칸 없음" + (rules.length ? "" : " (거래유형별 필수 적요 기준 미설정 — 적정성은 판단하지 않음)"), ver, rowsOf(recs)));
      out.push(...res);
    }
    return out;
  }

  // ------------------------------------------------------------------ 증빙 (인터페이스) ----
  const TAX_INVOICE = "세금계산서", APPROVAL_DOC = "품의서", ACCOUNT_NOTE = "계정해설";
  const DOC_FIELDS = {
    [TAX_INVOICE]: ["승인번호", "공급자사업자번호", "작성일", "공급가액", "세액", "합계", "출처"],
    [APPROVAL_DOC]: ["문서번호", "승인상태", "목적", "금액", "기간시작", "기간종료"],
    [ACCOUNT_NOTE]: ["계정코드", "적용기준", "문서근거"],
  };
  const CHECK_FIELDS = { [TAX_INVOICE]: ["승인번호", "공급가액", "세액"], [APPROVAL_DOC]: ["문서번호", "승인상태"], [ACCOUNT_NOTE]: ["계정코드", "적용기준"] };
  const LINK_CONFIRMED = "확정", LINK_CANDIDATE = "후보", LINK_FAILED = "연결실패";
  const CONFIRM_PENDING = "미확인", CONFIRM_OK = "확인", CONFIRM_FAILED = "인식실패";

  const nowIso = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
  const fieldUsable = (f) => f && f.value !== null && f.value !== undefined && f.value !== "" && f.confirm_status !== CONFIRM_PENDING && f.confirm_status !== CONFIRM_FAILED;

  /** 수기 증빙 저장소. 2차에서 다우오피스/OCR 공급자도 같은 {docs, links} 형태로 붙인다. */
  class EvidenceStore {
    constructor(data) { this.docs = (data && data.docs) || {}; this.links = (data && data.links) || []; }
    addDoc(type, values, { user = "", fileName = "", approvalDocId = "", source = "수기입력" } = {}) {
      if (!DOC_FIELDS[type]) throw new Error("알 수 없는 증빙 유형: " + type);
      const fields = {};
      for (const [k, v] of Object.entries(values)) {
        fields[k] = v && typeof v === "object" && "value" in v ? v : {
          value: v instanceof Big ? v.toString() : v, source, page: null, region: null, confidence: null,
          confirm_status: source === "수기입력" ? CONFIRM_OK : CONFIRM_PENDING, confirmed_by: source === "수기입력" ? user : "", confirmed_at: source === "수기입력" ? nowIso() : "",
        };
      }
      const id = "EV-" + Math.random().toString(16).slice(2, 12);
      this.docs[id] = { doc_id: id, doc_type: type, fields, source, file_name: fileName, approval_doc_id: approvalDocId, registered_by: user, registered_at: nowIso(), recognition_failed: false };
      return this.docs[id];
    }
    link(voucherKey, docId, method = "수동", { user = "", note = "" } = {}) {
      const status = method === "거래처+금액(후보 전용)" ? LINK_CANDIDATE : this.docs[docId] ? LINK_CONFIRMED : LINK_FAILED;
      this.links = this.links.filter((l) => !(l.voucher_key === voucherKey && l.doc_id === docId));
      const l = { voucher_key: voucherKey, doc_id: docId, method, status, linked_by: user, linked_at: nowIso(), note };
      this.links.push(l);
      return l;
    }
    unlink(voucherKey, docId) { this.links = this.links.filter((l) => !(l.voucher_key === voucherKey && l.doc_id === docId)); }
    linksFor(key) { return this.links.filter((l) => l.voucher_key === key); }
    value(doc, name) { const f = doc.fields[name]; return fieldUsable(f) ? f.value : null; }
    unusable(doc, names) { return names.filter((n) => !fieldUsable(doc.fields[n])); }
    toJSON() { return { docs: this.docs, links: this.links }; }
  }

  function checkEvidence(groups, settings, ver, store) {
    const cfg = settings.evidence || {};
    const vat = settings.vat || {};
    store = store || new EvidenceStore();
    const out = [];
    for (const [key, recs] of groups) {
      const flags = new Set(recs.map((r) => r["파일첨부여부_구분"]));
      const flagNote = flags.has("문자열 FALSE") || flags.has("불리언 FALSE") ? ERP_FLAG_NOTE : "다우오피스 증빙 미확인";
      const vr = recs.filter((r) => isVatRow(r, vat));
      const required = [];
      if (cfg.require_tax_invoice_for_vat !== false && vr.length) required.push(TAX_INVOICE);
      if (cfg.require_approval_doc !== false) required.push(APPROVAL_DOC);
      if (!required.length) { out.push(CR(key, AREA_EVIDENCE, "E00", NA, "", "", "", "설정상 요구 증빙 없음", ver, "")); continue; }
      const links = store.linksFor(key);
      for (const type of required) {
        const code = type === TAX_INVOICE ? "E01" : "E02";
        const typed = links.map((l) => [l, store.docs[l.doc_id]]).filter(([, d]) => !d || d.doc_type === type);
        const confirmed = typed.filter(([l, d]) => l.status === LINK_CONFIRMED && d).map(([, d]) => d);
        if (!confirmed.length) {
          let reason = `${type} 미제공/미연결 · ${flagNote}`;
          if (typed.some(([l]) => l.status === LINK_CANDIDATE)) reason = `${type} 후보만 있음 — 수동 연결 확정 필요`;
          else if (typed.some(([l, d]) => l.status === LINK_FAILED || !d)) reason = `${type} 연결 실패`;
          out.push(CR(key, AREA_EVIDENCE, code, UNCHECKED, "", "", "", reason, ver, ""));
          continue;
        }
        const failed = confirmed.filter((d) => d.recognition_failed);
        const pending = confirmed.filter((d) => !d.recognition_failed).map((d) => [d, store.unusable(d, CHECK_FIELDS[type])]);
        const usable = pending.filter(([, m]) => !m.length).map(([d]) => d);
        if (!usable.length) {
          const why = failed.length ? "인식 실패" : "필수 값 수동 확인 전 또는 누락: " + [...new Set(pending.flatMap(([, m]) => m))].sort().join(", ");
          out.push(CR(key, AREA_EVIDENCE, code, UNCHECKED, "", "", "", `${type} ${why}`, ver, ""));
          continue;
        }
        if (type === TAX_INVOICE) out.push(...compareTaxInvoice(key, usable, vr, vat, ver, store));
        else for (const d of usable) {
          const state = normSpace(store.value(d, "승인상태"));
          out.push(CR(key, AREA_EVIDENCE, code, APPROVED_STATES.has(state) ? PASS : REVIEW, `${store.value(d, "문서번호")} / ${state}`, "승인 상태", "", `[실제 증빙 대조] 품의서 승인 상태 확인 (목적: ${store.value(d, "목적") || "-"})`, ver, ""));
        }
      }
    }
    return out;
  }

  function compareTaxInvoice(key, docs, vr, vat, ver, store) {
    const invTax = sum(docs.map((d) => new Big(String(store.value(d, "세액")))));
    const rows = vr.map((r) => r._엑셀행).join(",");
    const ids = docs.map((d) => store.value(d, "승인번호")).join(", ");
    if (!vr.length) return [CR(key, AREA_EVIDENCE, "E01", REVIEW, fmt(invTax), "", "", `[실제 증빙 대조] 세금계산서(${ids})는 연결됐으나 전표에 부가세 행이 없음`, ver, "")];
    const out = [];
    const vTax = sum(vr.map(vatRowAmount));
    const diff = vTax.minus(invTax);
    out.push(CR(key, AREA_EVIDENCE, "E01", diff.eq(0) ? PASS : REVIEW, fmt(vTax), fmt(invTax), fmt(diff), `[실제 증빙 대조] 전표 부가세 합계 vs 세금계산서 세액 합계 (${ids})`, ver, rows));
    const cols = vr.map((r) => (vat.supply_columns || {})[r["계정과목코드"]]);
    if (cols.some((c) => !c)) out.push(CR(key, AREA_EVIDENCE, "E03", UNCHECKED, "", "", "", "[실제 증빙 대조] 공급가액 열 매핑 미설정 — 공급가액 대조 안 함", ver, rows));
    else {
      const sup = vr.map((r, i) => parseAmount(r[cols[i]]));
      if (sup.some((p) => p.st !== "정상")) out.push(CR(key, AREA_EVIDENCE, "E03", UNCHECKED, "", "", "", "[실제 증빙 대조] 전표 공급가액 값 없음/해석 불가", ver, rows));
      else {
        const vs = sum(sup.map((p) => p.amt)), is = sum(docs.map((d) => new Big(String(store.value(d, "공급가액")))));
        out.push(CR(key, AREA_EVIDENCE, "E03", vs.eq(is) ? PASS : REVIEW, fmt(vs), fmt(is), fmt(vs.minus(is)), "[실제 증빙 대조] 전표 공급가액 vs 세금계산서 공급가액", ver, rows));
      }
    }
    return out;
  }

  function runChecks(rows, settings, store) {
    const data = withSide(rows);
    const groups = groupRecords(data);
    const ver = settingsVersion(settings);
    return [...checkData(data, groups, settings, ver), ...checkAmount(groups, settings, ver), ...checkAccount(groups, settings, ver),
      ...checkMemo(groups, settings, ver), ...checkEvidence(groups, settings, ver, store)];
  }

  function voucherSummary(rows, results) {
    const groups = groupRecords(withSide(rows));
    const byKey = groupBy(results, (r) => r.전표키);
    const first = (recs, c) => { const r = recs.find((x) => x[c] !== null && x[c] !== undefined && x[c] !== ""); return r ? r[c] : ""; };
    const out = [];
    for (const [key, recs] of groups) {
      const t = totals(recs);
      const res = byKey.get(key) || [];
      const areas = {};
      for (const a of AREAS) areas[a] = areaStatus(res.filter((r) => r.검사영역 === a).map((r) => r.상태));
      const reasons = res.filter((r) => [ERROR, REVIEW, UNCHECKED].includes(r.상태)).sort((a, b) => PRIORITY[a.상태] - PRIORITY[b.상태]);
      const memos = recs.map((r) => r["적요"]).filter(Boolean);
      const cnt = new Map(); memos.forEach((m) => cnt.set(m, (cnt.get(m) || 0) + 1));
      let rep = ""; let best = 0; for (const [m, c] of cnt) if (c > best) { best = c; rep = m; }
      out.push(Object.assign({
        전표키: key, 회계단위: first(recs, "회계단위"), 전표관리단위: first(recs, "전표관리단위"), 기표번호: first(recs, "기표번호"),
        회계일: first(recs, "회계일") || null, 기표자: first(recs, "기표자"), 기표부서: first(recs, "기표부서"), 결재상태: first(recs, "전자결재진행상태"),
        전표유형: first(recs, "전표분개유형"), 대표적요: rep, 행수: recs.length, 차변합계: t.d, 대변합계: t.c, 차이: t.d.minus(t.c),
      }, areas, {
        검사결과: summaryStatus(AREAS.map((a) => areas[a])),
        사유: [...new Set(reasons.map((r) => `[${r.검사영역}] ${r.근거}`))].join(" / ").slice(0, 500),
      }));
    }
    return out;
  }

  function summaryCounts(summary) {
    const c = { 전체: summary.length };
    for (const s of SUMMARY_ORDER) c[s] = summary.filter((r) => r.검사결과 === s).length;
    return c;
  }

  function sortErrorFirst(summary) {
    const o = Object.fromEntries(SUMMARY_ORDER.map((s, i) => [s, i]));
    return summary.slice().sort((a, b) => o[a.검사결과] - o[b.검사결과] || String(a.회계일 || "￿").localeCompare(String(b.회계일 || "￿")) || String(a.기표번호).localeCompare(String(b.기표번호)));
  }

  /** 전표 단위 필터 (중복 집계 없음). f: {date:[s,e], 기표부서:[...], ...} */
  function applyFilters(summary, reviews, f) {
    let v = summary.map((s) => Object.assign({}, s, { 검토상태: (reviews[s.전표키] || {}).검토상태 || "미검토", 변경감지: !!(reviews[s.전표키] || {}).변경감지 }));
    if (f.date) v = v.filter((s) => s.회계일 && s.회계일 >= f.date[0] && s.회계일 <= f.date[1]);
    for (const col of ["기표부서", "기표자", "결재상태", "전표유형", "검사결과", "검토상태"]) if (f[col] && f[col].length) v = v.filter((s) => f[col].includes(s[col]));
    return v;
  }

  // ------------------------------------------------------------------ 검토 상태 ----
  const HASH_COLUMNS = EXPECTED_COLUMNS.filter((c) => c !== "승인번호" && c !== "전자결재진행상태");
  function cellStr(v) { if (v === null || v === undefined) return ""; if (v instanceof Big) return v.toString(); return String(v); }
  function contentHashes(rows) {
    const lines = new Map();
    for (const r of rows) { if (!lines.has(r.전표키)) lines.set(r.전표키, []); lines.get(r.전표키).push(HASH_COLUMNS.map((c) => cellStr(r[c])).join("\x1f")); }
    const out = {};
    for (const [k, v] of lines) out[k] = sha256(v.sort().join("\x1e")).slice(0, 16);
    return out;
  }
  function initReviews(hashes, previous) {
    previous = previous || {};
    const out = {};
    for (const [key, h] of Object.entries(hashes)) {
      const p = previous[key];
      if (!p) out[key] = { 전표키: key, 검토상태: "미검토", 메모: "", 검토자: "", 검토일시: "", 내용해시: h, 변경감지: false, 이전검토상태: "" };
      else if (p.내용해시 === h) out[key] = Object.assign({}, p);
      else out[key] = { 전표키: key, 검토상태: "미검토", 메모: p.메모, 검토자: p.검토자, 검토일시: p.검토일시, 내용해시: h, 변경감지: true, 이전검토상태: p.검토상태 };
    }
    return out;
  }
  function updateReview(reviews, key, { status, memo, reviewer } = {}) {
    const r = reviews[key];
    if (status !== undefined) {
      if (!REVIEW_STATES.includes(status)) throw new Error("알 수 없는 검토상태: " + status);
      r.검토상태 = status;
      if (status === "검토완료") r.변경감지 = false;
    }
    if (memo !== undefined) r.메모 = memo;
    if (reviewer !== undefined) r.검토자 = reviewer;
    r.검토일시 = nowIso();
    return r;
  }

  // ------------------------------------------------------------------ 저장 ----
  const SAVE_FORMAT = "voucher-review-browser-save/1";
  function serializeSession(s) {
    const enc = (v) => (v instanceof Big ? { __d: v.toString() } : v);
    return JSON.stringify({
      format: SAVE_FORMAT, saved_at: nowIso(), source: s.source, raw: s.raw,
      rows: s.rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, enc(v)]))),
      reviews: s.reviews, evidence: s.evidence, settings: s.settings,
    });
  }
  function deserializeSession(text) {
    const d = JSON.parse(text);
    if (d.format !== SAVE_FORMAT) throw new Error("지원하지 않는 저장 파일 형식입니다.");
    d.rows = d.rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, v && typeof v === "object" && "__d" in v ? new Big(v.__d) : v])));
    return d;
  }

  // ------------------------------------------------------------------ 엑셀 출력 ----
  const FORMULA_PREFIX = ["=", "+", "-", "@", "\t", "\r", "\n"];
  function safeCell(v) {
    if (typeof v === "string" && FORMULA_PREFIX.some((p) => v.startsWith(p))) return "'" + v;
    if (v instanceof Big) { const i = v.eq(v.round(0, 0)); if (i && Math.abs(Number(v)) <= Number.MAX_SAFE_INTEGER) return Number(v.toFixed(0)); return v.toString(); }
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    return v === null || v === undefined ? "" : v;
  }
  function buildResultSheets(summary, results, raw, rows, reviews, criteria) {
    const rvCols = ["검토상태", "메모", "검토자", "검토일시", "변경감지", "이전검토상태"];
    const s = summary.map((x) => { const r = reviews[x.전표키] || {}; const o = Object.assign({}, x); for (const c of rvCols) o[c] = r[c] === undefined ? (c === "검토상태" ? "미검토" : "") : r[c]; return o; });
    const original = raw.map((r, i) => Object.assign({ 전표키: rows[i].전표키 }, r));
    return { 전표요약: s, 검사상세: results, 원본분개: original, 적용기준: criteria };
  }
  function checkExportCounts(sheets, nRows, nVouchers) {
    const p = [];
    if (sheets.원본분개.length !== nRows) p.push(`원본분개 ${sheets.원본분개.length}행 ≠ 원본 ${nRows}행`);
    if (sheets.전표요약.length !== nVouchers) p.push(`전표요약 ${sheets.전표요약.length}건 ≠ 전표 ${nVouchers}건`);
    const keys = sheets.전표요약.map((r) => r.전표키);
    if (new Set(keys).size !== keys.length) p.push("전표요약에 중복 전표키");
    const ks = new Set(keys);
    if (sheets.검사상세.some((r) => !ks.has(r.전표키))) p.push("검사상세에 요약에 없는 전표키");
    return p;
  }
  async function toExcelBuffer(ExcelJS, sheets) {
    const wb = new ExcelJS.Workbook();
    for (const [name, data] of Object.entries(sheets)) {
      const ws = wb.addWorksheet(name.slice(0, 31));
      const cols = data.length ? Object.keys(data[0]) : [];
      ws.addRow(cols.map(safeCell));
      for (const r of data) ws.addRow(cols.map((c) => safeCell(r[c])));
    }
    return wb.xlsx.writeBuffer();
  }

  // ------------------------------------------------------------------ 계정 질문 ----
  const STOPWORDS = new Set(["차변", "대변", "계정", "계정과목", "과목", "전표", "분개", "처리", "입력", "작성", "제조업", "우리", "회사",
    "어떤", "어떻게", "무엇", "뭐", "뭘", "뭔가요", "써요", "쓰나요", "써야", "쓰면", "하나요", "해야", "하면",
    "알려줘", "알려주세요", "경우", "할때", "때", "관련", "질문", "건", "것", "좀", "무슨", "되나요", "돼요"]);
  const GENERIC = new Set(["지급", "결제", "출금", "송금", "입금", "매출", "판매", "수입", "투입", "정산", "카드", "교육", "개발",
    "연구", "보수", "수리", "전화", "도서", "임시", "가수", "선수", "회비", "상각", "폐기", "불량", "선적"]);
  const JOSA = ["에서는", "으로는", "에서", "으로", "에는", "하고", "이고", "하면", "할때", "을", "를", "이", "가", "은", "는", "에", "의", "로", "도", "와", "과", "만"];
  const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  function baseName(name) { let n = norm(name); n = n.replace(/[\(\[（].*?[\)\]）]$/, ""); return n.split(/[-_/]/)[0]; }

  class AccountAdvisor {
    constructor(guide, rows, inventory) {
      this.guide = guide;
      this.rows = rows && rows.length ? withSide(rows) : null;
      this.inventory = inventory || null;
      this.aliases = {};
      for (const [k, v] of Object.entries(guide.aliases || {})) this.aliases[norm(k)] = new Set(v.map(norm));
      const names = new Set();
      for (const s of guide.scenarios) for (const e of s.분개) for (const i of [...e.차변, ...e.대변]) names.add(norm(i.계정));
      for (const [k, v] of Object.entries(this.aliases)) { names.add(k); v.forEach((a) => names.add(a)); }
      if (this.inventory) for (const r of this.inventory) { names.add(norm(r.계정과목)); names.add(baseName(r.계정과목)); }
      this.accountNames = [...names].filter((n) => n.length >= 2).sort((a, b) => b.length - a.length);
    }
    detectContext(q) {
      const ctx = this.guide.context_keywords || {};
      const hits = Object.fromEntries(Object.entries(ctx).map(([k, ws]) => [k, ws.filter((w) => q.includes(norm(w))).length]));
      const mx = Math.max(0, ...Object.values(hits));
      const best = Object.keys(hits).filter((k) => hits[k] && hits[k] === mx);
      return best.length === 1 ? best[0] : null;
    }
    static detectSide(q) { const d = q.includes("차변"), c = q.includes("대변"); return d && !c ? SIDE_DEBIT : c && !d ? SIDE_CREDIT : null; }
    spans(q) { const s = []; for (const n of this.accountNames) { let i = q.indexOf(n); while (i >= 0) { s.push([i, i + n.length]); i = q.indexOf(n, i + 1); } } return s; }
    keywordIn(q, kw, spans) {
      const k = norm(kw);
      if (!k) return false;
      if (/^[a-z0-9/&]+$/.test(k)) return new RegExp(`(?<![a-z0-9])${escRe(k)}(?![a-z0-9])`).test(q);
      let i = q.indexOf(k);
      while (i >= 0) {
        const a = i, b = i + k.length;
        if (!spans.some(([s, e]) => s <= a && b <= e && e - s > b - a)) return true;
        i = q.indexOf(k, i + 1);
      }
      return false;
    }
    matchScenarios(question, topN = 3) {
      const q = norm(question), context = this.detectContext(q), spans = this.spans(q);
      let res = [];
      for (const s of this.guide.scenarios) {
        let hits = s.키워드.filter((k) => this.keywordIn(q, k, spans));
        hits = hits.filter((k) => !hits.some((o) => norm(k) !== norm(o) && norm(o).includes(norm(k))));
        if (!hits.length) continue;
        let score = hits.reduce((a, k) => a + norm(k).length * (GENERIC.has(k) ? 0.5 : 1), 0);
        if (q.includes(norm(s.거래유형))) score += 5;
        const preferred = context ? s.분개.map((e, i) => ((e.조건 || "").includes(context) ? i : -1)).filter((i) => i >= 0) : [];
        res.push({ scenario: s, score, matched: hits, preferred });
      }
      res.sort((a, b) => b.score - a.score || (a.scenario.id < b.scenario.id ? -1 : 1));
      if (res.length) { const cut = res[0].score * 0.4; res = res.filter((m) => m.score >= cut); }
      return res.slice(0, topN);
    }
    queryTerms(question, matches) {
      const terms = [];
      for (const m of matches) terms.push(...m.matched);
      for (const word of question.split(/[\s,.?!·/()"']+/)) {
        let w = word.trim();
        for (const j of JOSA) if (w.endsWith(j) && w.length - j.length >= 2) { w = w.slice(0, -j.length); break; }
        if (w.length >= 2 && !STOPWORDS.has(norm(w)) && !terms.includes(w)) terms.push(w);
      }
      return [...new Set(terms.filter((t) => !STOPWORDS.has(norm(t))))];
    }
    companyCandidates(account) {
      if (!this.inventory) return [];
      const t = norm(account), alias = this.aliases[t] || new Set();
      const out = [];
      for (const r of this.inventory) {
        const cn = norm(r.계정과목), cb = baseName(r.계정과목);
        let kind = null;
        if (cb === t || cn === t) kind = "일치"; else if (alias.has(cb) || alias.has(cn)) kind = "동의어"; else if (t.length >= 3 && cn.includes(t)) kind = "유사";
        if (kind) out.push({ 일치구분: kind, 계정과목코드: r.계정과목코드, 계정과목: r.계정과목, 사용구분: r.사용구분, 차변행수: r.차변행수, 대변행수: r.대변행수, 사용전표수: r.사용전표수, 비용구분: r.비용구분 });
      }
      const o = { 일치: 0, 동의어: 1, 유사: 2 };
      return out.sort((a, b) => o[a.일치구분] - o[b.일치구분] || b.사용전표수 - a.사용전표수);
    }
    similarMemo(terms, limit = 15) {
      if (!this.rows || !terms.length) return [];
      const hits = this.rows.map((r) => terms.filter((t) => norm(r["적요"]).includes(norm(t))).length);
      const mx = Math.max(0, ...hits);
      if (!mx) return [];
      const mask = new Set(this.rows.filter((r, i) => hits[i] === mx).map((r) => r._i));
      const keys = new Set(this.rows.filter((r) => mask.has(r._i)).map((r) => r.전표키));
      const sub = this.rows.filter((r) => keys.has(r.전표키) && (r.차대구분 === SIDE_DEBIT || r.차대구분 === SIDE_CREDIT));
      const g = groupBy(sub, (r) => [r.차대구분, r["계정과목코드"], r["계정과목"]].join("\u0000"));
      const out = [];
      for (const [k, recs] of g) {
        const [side, code, name] = k.split("\u0000");
        let ex = recs.filter((r) => mask.has(r._i)).map((r) => r["적요"]);
        if (!ex.length) ex = recs.map((r) => r["적요"]);
        out.push({ 차대구분: side, 계정과목코드: code, 계정과목: name, 전표수: new Set(recs.map((r) => r.전표키)).size, 적요예시: [...new Set(ex)].slice(0, 2).join(" / ") });
      }
      const so = { [SIDE_DEBIT]: 0, [SIDE_CREDIT]: 1 };
      out.sort((a, b) => so[a.차대구분] - so[b.차대구분] || b.전표수 - a.전표수 || (a.계정과목코드 < b.계정과목코드 ? -1 : 1));
      const cnt = {}; return out.filter((r) => (cnt[r.차대구분] = (cnt[r.차대구분] || 0) + 1) <= limit);
    }
    mentionedAccounts(question) {
      if (!this.inventory) return [];
      const q = norm(question), seen = [], out = [];
      const inv = this.inventory.slice().sort((a, b) => norm(b.계정과목).length - norm(a.계정과목).length);
      for (const r of inv) {
        const b = baseName(r.계정과목);
        if (b.length < 2 || (!q.includes(norm(r.계정과목)) && !q.includes(b))) continue;
        if (seen.some((s) => s.includes(b))) continue;
        seen.push(b);
        const p = Object.assign({}, r);
        if (this.rows) { p.차변일때_상대계정 = counterpartAccounts(this.rows, r.계정과목코드, SIDE_DEBIT).slice(0, 5); p.대변일때_상대계정 = counterpartAccounts(this.rows, r.계정과목코드, SIDE_CREDIT).slice(0, 5); }
        out.push(p);
      }
      return out.slice(0, 5);
    }
    ask(question) {
      question = (question || "").trim();
      const q = norm(question);
      const ans = { question, context: this.detectContext(q), side: AccountAdvisor.detectSide(q), matches: [], company: {}, similar: [], terms: [], mentioned: [], notes: [] };
      if (!question) { ans.notes.push("질문을 입력해 주세요. 예: '공장 전기요금 낼 때 차변 대변 뭐 써요?'"); return ans; }
      ans.matches = this.matchScenarios(question);
      for (const m of ans.matches) for (const e of m.scenario.분개) for (const it of [...e.차변, ...e.대변]) if (!(it.계정 in ans.company) && !it.계정.startsWith("해당 ")) ans.company[it.계정] = this.companyCandidates(it.계정);
      ans.terms = this.queryTerms(question, ans.matches);
      ans.similar = this.similarMemo(ans.terms);
      ans.mentioned = this.mentionedAccounts(question);
      if (!this.inventory) ans.notes.push("전표 파일을 불러오지 않아 회사 실제 사용 계정은 표시하지 않았습니다.");
      if (ans.matches.some((m) => m.scenario.분개.length > 1) && !ans.context) ans.notes.push("제조(공장·생산) 비용인지 판관(본사·영업·관리) 비용인지에 따라 계정이 달라질 수 있습니다. 질문에 '공장' 또는 '본사' 등을 넣으면 해당 분개를 먼저 보여 드립니다.");
      if (!ans.matches.length && !ans.mentioned.length && !ans.similar.length) ans.notes.push("일치하는 거래유형이나 과거 적요를 찾지 못했습니다. 거래 내용을 구체적으로 적어 주세요 (예: '생산라인 외주 도금 가공비', '거래처 명절 선물'). 회계팀 확인을 권장합니다.");
      return ans;
    }
  }
  function validateGuide(d) {
    if (!d || !Array.isArray(d.scenarios)) throw new Error("가이드 JSON 에 scenarios 목록이 없습니다.");
    d.scenarios.forEach((s, i) => {
      for (const k of ["id", "거래유형", "키워드", "분개"]) if (!(k in s)) throw new Error(`${i + 1}번째 거래유형에 '${k}' 항목이 없습니다.`);
      for (const e of s.분개) if (!e.차변 || !e.차변.length || !e.대변 || !e.대변.length) throw new Error(`거래유형 '${s.거래유형}' 의 분개에 차변/대변이 모두 있어야 합니다.`);
    });
  }

  // ------------------------------------------------------------------ SHA-256 (동기, 의존성 없음) ----
  function sha256(str) {
    const bytes = typeof TextEncoder !== "undefined" ? new TextEncoder().encode(str) : Buffer.from(str, "utf8");
    const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
    const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    const l = bytes.length, nBlocks = ((l + 9 + 63) >> 6);
    const w = new Uint32Array(nBlocks * 16);
    for (let i = 0; i < l; i++) w[i >> 2] |= bytes[i] << (24 - (i % 4) * 8);
    w[l >> 2] |= 0x80 << (24 - (l % 4) * 8);
    w[nBlocks * 16 - 1] = l * 8; w[nBlocks * 16 - 2] = Math.floor((l * 8) / 0x100000000);
    const W = new Uint32Array(64);
    for (let b = 0; b < nBlocks; b++) {
      for (let t = 0; t < 16; t++) W[t] = w[b * 16 + t];
      for (let t = 16; t < 64; t++) {
        const x = W[t - 15], y = W[t - 2];
        const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
        const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
        W[t] = (W[t - 16] + s0 + W[t - 7] + s1) >>> 0;
      }
      let [a, bb, c, d, e, f, g, h] = H;
      for (let t = 0; t < 64; t++) {
        const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
        const ch = (e & f) ^ (~e & g);
        const t1 = (h + S1 + ch + K[t] + W[t]) >>> 0;
        const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
        const maj = (a & bb) ^ (a & c) ^ (bb & c);
        const t2 = (S0 + maj) >>> 0;
        h = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = bb; bb = a; a = (t1 + t2) >>> 0;
      }
      H[0] = (H[0] + a) >>> 0; H[1] = (H[1] + bb) >>> 0; H[2] = (H[2] + c) >>> 0; H[3] = (H[3] + d) >>> 0;
      H[4] = (H[4] + e) >>> 0; H[5] = (H[5] + f) >>> 0; H[6] = (H[6] + g) >>> 0; H[7] = (H[7] + h) >>> 0;
    }
    return H.map((x) => x.toString(16).padStart(8, "0")).join("");
  }

  return {
    Big, EXPECTED_COLUMNS, REQUIRED_COLUMNS, MANAGEMENT_ITEM_COLUMNS, AREAS, SUMMARY_ORDER, REVIEW_STATES, STATUSES: [PASS, ERROR, REVIEW, UNCHECKED, NA],
    SIDE_DEBIT, SIDE_CREDIT, PASS, ERROR, REVIEW, UNCHECKED, NA, SUMMARY_PASS, TAX_INVOICE, APPROVAL_DOC, DOC_FIELDS,
    LINK_CONFIRMED, LINK_CANDIDATE, LINK_FAILED, CONFIRM_PENDING, CONFIRM_OK,
    VoucherFileError, rawText, cleanText, parseAmount, classifyFlag, parseDate, fmt, norm,
    loadWorkbook, findHeaderCandidates, autoMapping, loadVouchers, voucherKey,
    classifySide, withSide, accountInventory, codeNameConflicts, counterpartAccounts, managementItemProfile, fileOverview,
    defaultSettings, mergeDefaults, settingsVersion, isVatRow,
    runChecks, voucherSummary, summaryCounts, sortErrorFirst, applyFilters, areaStatus, summaryStatus,
    EvidenceStore, contentHashes, initReviews, updateReview, serializeSession, deserializeSession,
    safeCell, buildResultSheets, checkExportCounts, toExcelBuffer,
    AccountAdvisor, validateGuide, baseName, sha256, stableStringify,
  };
});
