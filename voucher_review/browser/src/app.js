/* 전표 검토 도우미 — 브라우저판 화면. 모든 처리는 이 브라우저 안에서만 이루어진다. */
(function () {
  "use strict";
  const E = window.VoucherEngine;
  const ExcelJS = window.ExcelJS;
  const DEFAULT_GUIDE = window.VOUCHER_GUIDE;
  const EXAMPLE_RULES = window.VOUCHER_EXAMPLE_RULES;
  const LS_SETTINGS = "voucher_review.settings", LS_GUIDE = "voucher_review.guide";

  // ------------------------------------------------------------ 상태 ----
  const S = {
    tab: "review", wb: null, sourceName: "", sourceSha: "", candidates: [], candIdx: 0, mapping: null,
    load: null, inventory: null, results: null, summary: null, checkedVersion: "", checkedAt: "",
    reviews: {}, evidence: new E.EvidenceStore(), reviewer: "", selectedKey: null, invCode: null,
    filters: {}, sort: "오류 우선", question: "", flash: null,
    settings: lsGet(LS_SETTINGS) ? E.mergeDefaults(lsGet(LS_SETTINGS)) : E.defaultSettings(),
    guide: lsGet(LS_GUIDE) || DEFAULT_GUIDE, draft: null,
  };
  function lsGet(k) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
  function lsDel(k) { try { localStorage.removeItem(k); } catch (e) { /* 무시 */ } }

  // ------------------------------------------------------------ 도우미 ----
  const $ = (sel, el = document) => el.querySelector(sel);
  const esc = (v) => String(v == null ? "" : v).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const show = (v) => (v instanceof E.Big ? E.fmt(v) : v === true ? "예" : v === false ? "" : v == null ? "" : String(v));
  const ready = () => S.load && !S.load.errors.length;
  function badge(s) { const k = String(s || "").split(" ")[0]; return s ? `<span class="badge b-${esc(k)}">${esc(s)}</span>` : ""; }
  const STATUS_COLS = new Set([...E.AREAS, "검사결과", "상태"]);
  const NUM_COLS = new Set(["차변합계", "대변합계", "차이", "차변금액", "대변금액", "차변행수", "대변행수", "사용전표수", "행수", "전표수", "함께쓴전표수", "입력행수", "전체행수"]);
  function table(rows, cols, opts = {}) {
    if (!rows.length) return `<p class="muted">${esc(opts.empty || "표시할 내용이 없습니다.")}</p>`;
    cols = cols || Object.keys(rows[0]);
    const wrapCols = new Set(opts.wrap || ["사유", "근거", "적요예시(상위)", "적요예시", "예시값(상위)"]);
    const head = cols.map((c) => `<th>${esc(c)}</th>`).join("");
    const body = rows.map((r) => {
      const attrs = opts.rowKey ? ` class="clickable${opts.selected === r[opts.rowKey] ? " selected" : ""}" data-action="${esc(opts.action)}" data-key="${esc(r[opts.rowKey])}"` : "";
      return `<tr${attrs}>` + cols.map((c) => {
        const v = r[c];
        if (STATUS_COLS.has(c)) return `<td>${badge(v)}</td>`;
        const cls = [NUM_COLS.has(c) ? "num" : "", wrapCols.has(c) ? "wrap" : ""].filter(Boolean).join(" ");
        return `<td${cls ? ` class="${cls}"` : ""}>${esc(show(v))}</td>`;
      }).join("") + "</tr>";
    }).join("");
    return `<div class="tablewrap"${opts.height ? ` style="max-height:${opts.height}px"` : ""}><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  function multiSelect(name, options, selected) {
    selected = selected || [];
    const label = selected.length ? selected.join(", ") : "전체";
    return `<details class="ms"><summary>${esc(label)}</summary><div class="ms-box">` +
      options.map((o) => `<label><input type="checkbox" data-ms="${esc(name)}" value="${esc(o)}"${selected.includes(o) ? " checked" : ""}> ${esc(o)}</label>`).join("") +
      "</div></details>";
  }
  function flash(kind, text) { S.flash = { kind, text }; }
  function download(name, blob) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function busy(on) { $("#loading").classList.toggle("on", !!on); }
  async function sha256Buffer(buf) {
    try { if (crypto && crypto.subtle) { const h = await crypto.subtle.digest("SHA-256", buf); return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, "0")).join(""); } } catch (e) { /* 아래로 */ }
    return "";
  }
  const today = () => new Date().toISOString().slice(0, 10).replace(/-/g, "");

  // ------------------------------------------------------------ 동작 ----
  function setLoaded(load, name, sha, previousReviews) {
    S.load = load; S.sourceName = name; S.sourceSha = sha;
    S.inventory = load.errors.length ? null : E.accountInventory(load.rows);
    S.results = null; S.summary = null; S.selectedKey = null; S.invCode = null; S.filters = {};
    S.reviews = load.errors.length ? {} : E.initReviews(E.contentHashes(load.rows), previousReviews);
  }
  function runChecks() {
    if (!ready()) return;
    S.results = E.runChecks(S.load.rows, S.settings, S.evidence);
    S.summary = E.voucherSummary(S.load.rows, S.results);
    S.checkedVersion = E.settingsVersion(S.settings);
    S.checkedAt = new Date().toLocaleString("ko-KR");
  }

  async function onFile(file) {
    busy(true);
    try {
      const buf = await file.arrayBuffer();
      S.wb = await E.loadWorkbook(ExcelJS, buf);
      S.candidates = E.findHeaderCandidates(S.wb);
      S.candIdx = 0; S.sourceName = file.name; S.sourceSha = await sha256Buffer(buf);
      S.mapping = S.candidates.length ? E.autoMapping(S.candidates[0].headers) : null;
      if (!S.candidates.length) flash("error", "전표 헤더(회계단위·기표번호·계정과목 등)가 있는 시트를 찾지 못했습니다.");
      else flash("info", `파일을 읽었습니다: ${file.name}. 시트와 열 매핑을 확인한 뒤 '불러오기'를 눌러 주세요.`);
    } catch (e) {
      S.wb = null; S.candidates = [];
      flash("error", e instanceof E.VoucherFileError ? e.message : "파일을 여는 중 오류가 발생했습니다: " + e.message);
    } finally { busy(false); render(); }
  }

  // ------------------------------------------------------------ 사이드바 ----
  function renderSidebar() {
    const c = S.candidates[S.candIdx];
    let html = `<h2>1. 전표 파일</h2>
      <p class="muted">영림원 ERP 전표조건검색 엑셀(.xlsx). 이 PC의 브라우저 안에서만 처리하며 어디에도 전송하지 않습니다. 원본 파일은 수정하지 않습니다.</p>
      <label class="filebtn"><input type="file" id="file" accept=".xlsx" hidden><span>📂 엑셀 파일 선택</span></label>
      <p class="muted">${S.sourceName ? "선택한 파일: " + esc(S.sourceName) : "선택한 파일 없음"}</p>`;
    if (S.candidates.length) {
      html += `<label>2. 시트 / 헤더 행</label><select id="cand">${S.candidates.map((x, i) => `<option value="${i}"${i === S.candIdx ? " selected" : ""}>${esc(x.sheet)} / ${x.headerRow}행 (일치 열 ${x.matched}개)</option>`).join("")}</select>`;
      if (c.preamble.length) html += `<p class="muted">헤더 위 내용: ${esc(c.preamble.join(" | "))}</p>`;
      const opts = ["(없음)", ...c.headers.filter(Boolean)];
      html += `<details><summary class="small">3. 열 매핑 확인 (자동 인식)</summary>` + E.EXPECTED_COLUMNS.map((col) => {
        const cur = (S.mapping || {})[col] || "(없음)";
        return `<label>${esc(col)}${E.REQUIRED_COLUMNS.includes(col) ? " *" : ""}</label><select data-map="${esc(col)}">${opts.map((o) => `<option${o === cur ? " selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
      }).join("") + `</details><button class="block" data-action="load">불러오기</button>`;
    }
    html += `<h2>4. 검사</h2><button class="block primary" data-action="run"${ready() ? "" : " disabled"}>검사 실행</button>
      <label>검토자 이름</label><input type="text" id="reviewer" value="${esc(S.reviewer)}">
      <hr style="border:0;border-top:1px solid var(--line);margin:16px 0">
      <button class="block" data-action="reset">초기화 (현재 화면만 비우기)</button>
      <p class="muted">초기화는 이 화면의 작업만 비웁니다. 원본 파일과 브라우저에 저장한 작업은 지우지 않습니다.</p>`;
    $("aside").innerHTML = html;
  }

  // ------------------------------------------------------------ 전표 검사 ----
  function criteriaRows() {
    const s = S.settings, v = s.vat;
    const rows = [
      ["원본 파일", S.sourceName], ["원본 SHA-256", (S.sourceSha || "").slice(0, 16)], ["시트/헤더행", `${S.load.sheet} / ${S.load.headerRow}`],
      ["헤더 위 조회 조건", (S.load.preamble || []).join(" | ")], ["설정 버전", E.settingsVersion(s)], ["검사 실행 시각", S.checkedAt],
      ["요약 기준", "오류 > 확인 필요 > 미검사 > 설정된 검사 통과 (검사 통과는 최종 회계 적정성 보증 아님)"],
      ["부가세 계정코드", v.account_codes.join(", ")], ["부가세 계정명 키워드", v.name_keywords.join(", ")],
      ["공급가액 열 매핑", Object.entries(v.supply_columns).map(([a, b]) => `${a}→${b}`).join(", ") || "미설정"],
      ["일반과세 증빙 유형", v.general_evidence_values.join(", ") || "미설정"], ["특수 증빙 제외 키워드", v.excluded_evidence_keywords.join(", ")],
      ["반올림 기준", v.rounding || "미설정"], ["허용 차이(원)", v.tolerance], ["적요 빈칸 상태", s.memo.blank_status],
      ["증빙 요구", `부가세 전표 세금계산서=${s.evidence.require_tax_invoice_for_vat}, 품의서=${s.evidence.require_approval_doc}`],
      ["파일첨부여부", "증빙 판정에 사용하지 않음 (ERP 첨부 표시 FALSE, 다우오피스 증빙 미확인)"], ["실행 환경", "브라우저판 (외부 전송 없음)"],
    ];
    for (const r of s.account_rules) rows.push([`계정규칙 ${r.규칙ID}@v${r.버전}`, `활성=${r.활성} · ${r.이름 || ""} · 근거: ${r.근거 || ""}`]);
    for (const r of s.memo.rules) rows.push([`적요규칙 ${r.규칙ID}@v${r.버전}`, `활성=${r.활성} · ${r.이름 || ""} · 근거: ${r.근거 || ""}`]);
    return rows.map(([a, b]) => ({ 항목: a, 값: b }));
  }

  function currentView() {
    let v = E.applyFilters(S.summary, S.reviews, S.filters);
    if (S.sort === "오류 우선") v = E.sortErrorFirst(v);
    else v.sort((a, b) => String(a[S.sort] || "￿").localeCompare(String(b[S.sort] || "￿")));
    return v;
  }

  function renderReview() {
    if (!ready()) return `<div class="msg info">왼쪽에서 전표 파일을 선택하고 <b>불러오기</b> → <b>검사 실행</b>을 눌러 주세요.</div>`;
    if (!S.summary) return `<div class="msg info">파일을 불러왔습니다. 왼쪽의 <b>검사 실행</b>을 눌러 주세요.</div>`;
    const counts = E.summaryCounts(S.summary);
    const dates = S.load.rows.map((r) => r["회계일"]).filter(Boolean).sort();
    const changed = Object.values(S.reviews).filter((r) => r.변경감지).length;
    let h = "";
    if (S.checkedVersion !== E.settingsVersion(S.settings)) h += `<div class="msg warn">검사 설정이 바뀌었습니다. 최신 설정으로 보려면 <b>검사 실행</b>을 다시 눌러 주세요.</div>`;
    h += `<p class="muted">대상 기간(파일 기준): ${esc(dates[0] || "-")} ~ ${esc(dates[dates.length - 1] || "-")} · 헤더 위 조회 조건: ${esc((S.load.preamble || []).join(" | ") || "(없음)")} · 설정 버전 ${esc(S.checkedVersion)}. 이 파일의 전표 수를 업무상 검토 건수(약 750~800건, 집계 기간 미확인)와 같은 모집단으로 가정하지 않습니다.</p>`;
    h += `<div class="cards">` + [["원본 행 수", S.load.rows.length], ["전표 수", counts.전체], ["🔴 오류", counts[E.ERROR]], ["🟠 확인 필요", counts[E.REVIEW]], ["⚪ 미검사", counts[E.UNCHECKED]], ["🟢 설정된 검사 통과", counts[E.SUMMARY_PASS]]]
      .map(([k, v]) => `<div class="card" data-metric="${esc(k)}"><div class="k">${esc(k)}</div><div class="v">${Number(v).toLocaleString()}</div></div>`).join("") + `</div>`;
    if (changed) h += `<div class="msg warn">이전 저장 이후 분개 내용이 바뀐 전표 ${changed}건 — 기존 검토완료를 승계하지 않고 '미검토'로 되돌렸습니다.</div>`;

    const opt = (col) => [...new Set(S.summary.map((s) => s[col]).filter(Boolean))].sort();
    const f = S.filters;
    h += `<div class="panel"><div class="row">
      <div><label>회계일 시작</label><input type="date" data-filter-date="0" value="${esc(f.date ? f.date[0] : dates[0] || "")}" min="${esc(dates[0] || "")}" max="${esc(dates[dates.length - 1] || "")}"></div>
      <div><label>회계일 종료</label><input type="date" data-filter-date="1" value="${esc(f.date ? f.date[1] : dates[dates.length - 1] || "")}" min="${esc(dates[0] || "")}" max="${esc(dates[dates.length - 1] || "")}"></div>
      ${["기표부서", "기표자", "결재상태", "전표유형"].map((c) => `<div><label>${c}</label>${multiSelect(c, opt(c), f[c])}</div>`).join("")}
    </div><div class="row">
      <div><label>검사결과</label>${multiSelect("검사결과", E.SUMMARY_ORDER, f.검사결과)}</div>
      <div><label>검토상태</label>${multiSelect("검토상태", E.REVIEW_STATES, f.검토상태)}</div>
      <div><label>정렬</label><select id="sort">${["오류 우선", "회계일", "기표번호"].map((o) => `<option${o === S.sort ? " selected" : ""}>${o}</option>`).join("")}</select></div>
      <div><button data-action="clear-filters">필터 초기화 (전체)</button></div>
    </div></div>`;
    const view = currentView();
    const fc = E.SUMMARY_ORDER.map((s) => `${s} ${view.filter((v) => v.검사결과 === s).length}`).join(" + ");
    h += `<p class="muted">필터 결과 ${view.length}건 = ${esc(fc)} (미검사 전표도 숨기지 않음). 행을 누르면 아래에 상세가 열립니다.</p>`;
    h += table(view, ["기표번호", "회계일", "기표자", "기표부서", "대표적요", "차변합계", "대변합계", "차이", ...E.AREAS, "검사결과", "사유", "검토상태", "변경감지", "회계단위", "전표관리단위"],
      { rowKey: "전표키", action: "select", selected: S.selectedKey, height: 420 });

    const sheets = E.buildResultSheets(S.summary, S.results, S.load.raw, S.load.rows, S.reviews, criteriaRows());
    const problems = E.checkExportCounts(sheets, S.load.rows.length, counts.전체);
    h += problems.length ? `<div class="msg error">결과 건수 대조 실패: ${esc(problems.join("; "))}</div>`
      : `<p class="muted" id="count-check">다운로드 건수 대조: 원본분개 ${sheets.원본분개.length}행 = 원본 ${S.load.rows.length}행, 전표요약 ${sheets.전표요약.length}건 = 전표 ${counts.전체}건</p>`;
    h += `<button data-action="export-results">검사 결과 엑셀 다운로드 (전표요약·검사상세·원본분개·적용기준)</button>`;
    if (S.selectedKey && S.summary.some((s) => s.전표키 === S.selectedKey)) h += renderDetail(S.selectedKey);
    return h;
  }

  function renderDetail(key) {
    const lines = S.load.rows.map((r, i) => [r, S.load.raw[i]]).filter(([r]) => r.전표키 === key);
    const s = S.summary.find((x) => x.전표키 === key);
    const rv = S.reviews[key];
    const cols = ["_엑셀행", "전표기표번호", "행번호", "계정과목코드", "계정과목", "차변금액", "대변금액", "적요", "증빙", "거래처사업자번호", "비용구분", "귀속부서", "파일첨부여부", "파일첨부여부_구분", ...E.MANAGEMENT_ITEM_COLUMNS];
    const lineRows = lines.map(([r, raw]) => { const o = {}; for (const c of cols) o[c] = E.MANAGEMENT_ITEM_COLUMNS.includes(c) ? raw[c] : r[c]; return o; });
    let h = `<div class="panel" id="detail"><h2>전표 상세 — ${esc(s.기표번호)}</h2>
      <p>전표 키 <code>${esc(key)}</code> · 요약 ${badge(s.검사결과)} · ${E.AREAS.map((a) => `${esc(a)} ${badge(s[a])}`).join(" · ")}</p>
      <p class="muted">검사 통과는 설정된 기준에 대한 결과이며 최종 회계 적정성 보증이 아닙니다.</p>
      <h3>분개 전체 (관리항목은 원본 값 그대로)</h3>${table(lineRows, cols, { height: 260 })}
      <h3>검사 근거·해설</h3>${table(S.results.filter((r) => r.전표키 === key), ["검사영역", "검사코드", "상태", "실제값", "비교값", "차이", "근거", "적용규칙버전", "원본행번호"], { height: 260 })}
      <div class="grid2"><div class="panel"><h3>검토 (자동 검사와 별도로 저장)</h3>`;
    if (rv.변경감지) h += `<div class="msg warn">분개 내용 변경 감지 — 이전 검토상태 '${esc(rv.이전검토상태)}' 는 승계하지 않았습니다.</div>`;
    h += `<label>검토상태</label><select id="rv-status">${E.REVIEW_STATES.map((o) => `<option${o === rv.검토상태 ? " selected" : ""}>${o}</option>`).join("")}</select>
      <label>검토자</label><input type="text" id="rv-reviewer" value="${esc(rv.검토자 || S.reviewer)}">
      <label>메모</label><textarea id="rv-memo">${esc(rv.메모)}</textarea>
      <p class="muted">마지막 검토일시: ${esc(rv.검토일시 || "-")}</p>
      <button class="primary" data-action="save-review" data-key="${esc(key)}">검토 저장</button></div>`;
    h += `<div class="panel"><h3>증빙 연결 (수기 입력 · 2차에서 다우오피스/OCR 연동 예정)</h3>
      <p class="muted">전표 키 기준 수동 연결입니다. 파일명·ERP 첨부 표시·'증빙' 열 텍스트로 적정성을 판단하지 않습니다.</p>`;
    for (const l of S.evidence.linksFor(key)) {
      const d = S.evidence.docs[l.doc_id];
      const desc = d ? `${d.doc_type} · ` + E.DOC_FIELDS[d.doc_type].slice(0, 3).filter((k) => d.fields[k]).map((k) => `${k}=${d.fields[k].value}`).join(", ") : "(문서 없음)";
      h += `<p class="small">[${esc(l.status)}/${esc(l.method)}] ${esc(desc)} <button data-action="unlink" data-key="${esc(key)}" data-doc="${esc(l.doc_id)}">해제</button></p>`;
    }
    const kind = S.evKind || E.TAX_INVOICE;
    h += `<div class="row"><label><input type="radio" name="evkind" value="${E.TAX_INVOICE}"${kind === E.TAX_INVOICE ? " checked" : ""}> 세금계산서</label>
      <label><input type="radio" name="evkind" value="${E.APPROVAL_DOC}"${kind === E.APPROVAL_DOC ? " checked" : ""}> 품의서</label></div>`;
    const fld = (id, label) => `<label>${esc(label)}</label><input type="text" data-ev="${esc(id)}">`;
    if (kind === E.TAX_INVOICE) h += fld("승인번호", "승인번호") + fld("공급자사업자번호", "공급자 사업자번호") + fld("작성일", "작성일 (YYYY-MM-DD)") + fld("공급가액", "공급가액") + fld("세액", "세액") + fld("합계", "합계");
    else h += fld("문서번호", "문서번호") + `<label>승인상태</label><select data-ev="승인상태">${["승인", "결재완료", "진행중", "반려", "미확인"].map((o) => `<option>${o}</option>`).join("")}</select>` + fld("목적", "목적") + fld("금액", "금액") + fld("기간시작", "기간 시작 (YYYY-MM-DD)") + fld("기간종료", "기간 종료 (YYYY-MM-DD)");
    h += fld("__approval", "결재문서 ID (있으면)") + fld("__file", "원본 파일명 (참고용)");
    h += `<button class="block" data-action="add-evidence" data-key="${esc(key)}">수기 등록 후 이 전표에 연결</button></div></div></div>`;
    return h;
  }

  // ------------------------------------------------------------ 계정 질문하기 ----
  const EXAMPLES = ["공장 전기요금 낼 때 차변 대변 뭐 써요?", "원재료 외상으로 샀어요", "생산직 급여 계정 알려줘", "거래처 접대 식사 법인카드",
    "외주 가공비 세금계산서 받았어요", "제품 수출 매출 영세율", "기계장치 설비 구입", "제품 판매 택배 운반비"];
  function renderQA() {
    let h = `<h2>제조업 전표, 차변·대변에 어떤 계정을 쓰나요?</h2>`;
    if (!ready()) h += `<p class="muted">전표 파일을 불러오면 회사가 실제로 쓰는 계정코드와 과거 비슷한 적요의 처리 사례도 함께 보여 드립니다.</p>`;
    h += `<div class="chips">${EXAMPLES.map((e) => `<button data-action="example" data-q="${esc(e)}">${esc(e)}</button>`).join("")}</div>
      <div class="row"><div style="flex:6"><input type="text" id="question" placeholder="예: 공장 전기요금 낼 때 차변 대변 뭐 써요?" value="${esc(S.question)}"></div><div style="flex:1"><button class="primary" data-action="ask">질문</button></div></div>`;
    if (!S.question) return h;
    const adv = new E.AccountAdvisor(S.guide, ready() ? S.load.rows : null, ready() ? S.inventory : null);
    const a = adv.ask(S.question);
    const chips = [a.context ? `비용 구분 추정: <b>${esc(a.context)}</b>` : "", a.side ? `질문한 쪽: <b>${esc(a.side)}</b>` : ""].filter(Boolean);
    if (chips.length) h += `<p class="muted">${chips.join(" · ")}</p>`;
    for (const n of a.notes) h += `<div class="msg info">${esc(n)}</div>`;
    a.matches.forEach((m, i) => {
      const s = m.scenario;
      h += `<div class="panel qa-answer"><h2>${i + 1}. ${esc(s.거래유형)}</h2><p class="muted">일치 키워드: ${esc(m.matched.join(", "))} · 가이드 버전 ${esc(S.guide.version || "")}</p>`;
      const order = [...m.preferred, ...s.분개.map((_, j) => j).filter((j) => !m.preferred.includes(j))];
      for (const j of order) {
        const e = s.분개[j];
        const sides = [["차변", e.차변], ["대변", e.대변]];
        if (a.side === E.SIDE_CREDIT) sides.reverse();
        h += `<div class="entry${m.preferred.includes(j) ? " pref" : ""}"><b>▸ ${esc(e.조건 || "")}</b>${m.preferred.includes(j) ? " ⭐ 질문 맥락과 일치" : ""}<div class="grid2">` +
          sides.map(([side, items]) => `<div><b>${side}</b>${a.side === side ? " ← 질문" : ""}<ul>${items.map((it) => `<li>${esc(it.계정)}${it.비고 ? ` — <i>${esc(it.비고)}</i>` : ""}</li>`).join("")}</ul></div>`).join("") + `</div></div>`;
      }
      if (s.해설) h += `<p>📘 ${esc(s.해설)}</p>`;
      for (const c of s.주의 || []) h += `<p>⚠️ ${esc(c)}</p>`;
      if (ready()) {
        const names = [];
        for (const e of s.분개) for (const it of [...e.차변, ...e.대변]) if (it.계정 in a.company && !names.includes(it.계정)) names.push(it.계정);
        const rows = [];
        for (const n of names) { const c = a.company[n]; if (!c.length) rows.push({ "가이드 계정": n, 일치구분: "회사 데이터에서 미발견" }); for (const r of c) rows.push(Object.assign({ "가이드 계정": n }, r)); }
        h += `<h3>🏢 이 파일에서 회사가 실제로 쓴 계정코드</h3>${table(rows, ["가이드 계정", "일치구분", "계정과목코드", "계정과목", "사용구분", "차변행수", "대변행수", "사용전표수", "비용구분"], { height: 260 })}
          <p class="muted">일치=계정명 동일, 동의어=가이드 동의어, 유사=계정명 포함. 코드 확정은 회계팀 기준을 따르세요.</p>`;
      }
      h += `</div>`;
    });
    if (ready() && a.similar.length) {
      h += `<div class="panel"><h2>🔎 과거 전표에서 비슷한 적요의 처리 사례</h2><p class="muted">적요 검색어: ${esc(a.terms.join(", "))} · 과거 처리가 곧 정답은 아닙니다.</p><div class="grid2">` +
        [E.SIDE_DEBIT, E.SIDE_CREDIT].map((side) => `<div><b>${side}</b>${table(a.similar.filter((r) => r.차대구분 === side), ["계정과목코드", "계정과목", "전표수", "적요예시"])}</div>`).join("") + `</div></div>`;
    }
    for (const p of a.mentioned) {
      h += `<div class="panel"><h2>📌 ${esc(p.계정과목)} (${esc(p.계정과목코드)}) — 이 파일의 사용 현황</h2>
        <p>사용구분 <b>${esc(p.사용구분)}</b> · 차변 ${p.차변행수}행 / 대변 ${p.대변행수}행 · 사용 전표 ${p.사용전표수}건 · 비용구분 ${esc(p.비용구분 || "-")}</p>
        ${p["적요예시(상위)"] ? `<p class="muted">적요 예시: ${esc(p["적요예시(상위)"])}</p>` : ""}
        <div class="grid2"><div><b>차변에 쓸 때 대변 상대계정</b>${table(p.차변일때_상대계정 || [])}</div><div><b>대변에 쓸 때 차변 상대계정</b>${table(p.대변일때_상대계정 || [])}</div></div></div>`;
    }
    return h;
  }

  // ------------------------------------------------------------ 계정과목 현황 ----
  function renderInventory() {
    if (!ready()) return `<div class="msg info">왼쪽에서 전표 파일을 불러오면 차변·대변에 쓰인 계정과목 전체 목록이 표시됩니다.</div>`;
    const ov = E.fileOverview(S.load.rows), inv = S.inventory;
    let h = `<div class="cards">` + [["사용 계정과목", ov.계정과목수], ["차변에 쓰인 계정", ov.차변사용계정수], ["대변에 쓰인 계정", ov.대변사용계정수], ["양쪽 모두 쓰인 계정", inv.filter((r) => r.차변행수 > 0 && r.대변행수 > 0).length]]
      .map(([k, v]) => `<div class="card"><div class="k">${k}</div><div class="v">${v}</div></div>`).join("") + `</div>`;
    const kw = S.invKeyword || "";
    const view = inv.filter((r) => !kw || r.계정과목코드.includes(kw) || r.계정과목.includes(kw) || r["적요예시(상위)"].includes(kw));
    h += `<label>계정코드/계정명/적요 검색</label><input type="text" id="inv-kw" value="${esc(kw)}">`;
    h += `<p class="muted">행을 누르면 아래에 계정 상세가 열립니다.</p>` + table(view, null, { rowKey: "계정과목코드", action: "inv-select", selected: S.invCode, height: 420 });
    const conflicts = E.codeNameConflicts(S.load.rows);
    if (conflicts.length) h += `<div class="msg warn">계정코드와 계정명이 1:1 이 아닌 경우가 있습니다. 확인이 필요합니다.</div>` + table(conflicts);
    h += `<p><button data-action="export-inventory">계정과목 현황 엑셀 다운로드</button></p>`;
    if (S.invCode) {
      const r = inv.find((x) => x.계정과목코드 === S.invCode);
      if (r) {
        const lines = E.withSide(S.load.rows.filter((x) => x.계정과목코드 === S.invCode));
        h += `<div class="panel"><h2>${esc(r.계정과목코드)} ${esc(r.계정과목)}</h2><div class="grid2">
          <div><b>차변으로 쓴 전표의 대변 상대계정</b>${table(E.counterpartAccounts(S.load.rows, S.invCode, E.SIDE_DEBIT))}</div>
          <div><b>대변으로 쓴 전표의 차변 상대계정</b>${table(E.counterpartAccounts(S.load.rows, S.invCode, E.SIDE_CREDIT))}</div></div>
          <h3>관리항목 사용 현황 (같은 관리항목 열이라도 계정마다 의미가 다릅니다)</h3>${table(E.managementItemProfile(S.load.rows, S.invCode))}
          <h3>분개 행 (원본 엑셀 행번호 포함)</h3>${table(lines, ["_엑셀행", "기표번호", "행번호", "회계일", "기표부서", "기표자", "차대구분", "차변금액", "대변금액", "적요", "비용구분"], { height: 300 })}</div>`;
      }
    }
    return h;
  }

  // ------------------------------------------------------------ 검사 설정 ----
  const ACC_FIELDS = ["규칙ID", "이름", "적요조건", "적용전표유형", "차대구분", "허용계정코드", "필수상대계정코드", "근거", "적용시작일", "적용종료일", "활성", "버전", "예시"];
  const MEMO_FIELDS = ["규칙ID", "이름", "대상계정코드", "적요조건", "필수포함", "근거", "활성", "버전", "예시"];
  const LIST_FIELDS = new Set(["적요조건", "적용전표유형", "허용계정코드", "필수상대계정코드", "대상계정코드", "필수포함"]);
  const splitList = (t) => String(t || "").split(",").map((x) => x.trim()).filter(Boolean);
  function rulesEditor(kind, rules, fields) {
    const head = fields.map((f) => `<th>${esc(f)}</th>`).join("") + "<th></th>";
    const body = rules.map((r, i) => "<tr>" + fields.map((f) => {
      const id = `data-rule="${kind}" data-i="${i}" data-f="${esc(f)}"`;
      if (f === "활성" || f === "예시") return `<td><input type="checkbox" ${id}${r[f] ? " checked" : ""}></td>`;
      if (f === "차대구분") return `<td><select ${id}>${["전체", "차변", "대변"].map((o) => `<option${(r[f] || "전체") === o ? " selected" : ""}>${o}</option>`).join("")}</select></td>`;
      const v = LIST_FIELDS.has(f) ? (r[f] || []).join(", ") : r[f] || "";
      return `<td><input type="text" ${id} value="${esc(v)}"></td>`;
    }).join("") + `<td><button data-action="del-rule" data-kind="${kind}" data-i="${i}">삭제</button></td></tr>`).join("");
    return `<div class="tablewrap rules"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div><button data-action="add-rule" data-kind="${kind}">규칙 추가</button>`;
  }
  function renderSettings() {
    if (!S.draft) S.draft = JSON.parse(JSON.stringify(S.settings));
    const d = S.draft, v = d.vat;
    let h = `<p><b>현재 적용 설정 버전:</b> <code>${esc(E.settingsVersion(S.settings))}</code> — 결과의 '적용규칙버전'에 기록됩니다.</p>
      <p class="muted">보수적 기본값: 회사가 기준을 설정하기 전에는 계정 검사·부가세 계산 점검·거래유형별 적요 검사를 '미검사'로 둡니다. 아래를 고친 뒤 <b>설정 적용</b>을 눌러야 반영됩니다.</p>
      <div class="panel"><h2>부가세 (전표 내부 계산 점검)</h2>
      <label>부가세 계정코드 (쉼표 구분, 비우면 계정명 키워드로만 식별)</label><input type="text" id="vat-codes" value="${esc(v.account_codes.join(", "))}">
      <label>부가세 계정명 키워드 (쉼표 구분)</label><input type="text" id="vat-names" value="${esc(v.name_keywords.join(", "))}">`;
    if (ready()) {
      const vatCodes = [...new Set(S.load.rows.filter((r) => E.isVatRow(r, v)).map((r) => r.계정과목코드))].sort();
      h += `<h3>공급가액 열 매핑 — 계정별로 관리항목 의미가 다릅니다. 확인한 열만 지정하세요.</h3>`;
      for (const code of vatCodes) {
        const sub = S.load.rows.filter((r) => r.계정과목코드 === code);
        const samples = E.MANAGEMENT_ITEM_COLUMNS.map((c) => [c, sub.map((r) => r[c]).filter(Boolean).slice(0, 3)]).filter(([, s]) => s.length);
        const cur = v.supply_columns[code] || "";
        h += `<label>${esc(code)} ${esc(sub[0].계정과목)} 공급가액 열</label><select data-supply="${esc(code)}"><option value="">(미설정)</option>${E.MANAGEMENT_ITEM_COLUMNS.map((c) => `<option${c === cur ? " selected" : ""}>${c}</option>`).join("")}</select>
          <p class="muted">값 예시: ${esc(samples.map(([c, s]) => `${c}=${s.join(", ")}`).join(" · ") || "(관리항목 값 없음)")}${samples.some(([c]) => c === "관리항목2") ? " (관리항목2는 공급가액 <b>후보</b>일 뿐 — 값을 보고 확인 후 지정)" : ""}</p>`;
      }
      const evs = [...new Set(S.load.rows.map((r) => r.증빙).filter(Boolean))].sort();
      h += `<label>일반과세 내부 계산 대상 '증빙' 값 (선택 안 함 → 부가세 계산 점검 미검사)</label>${multiSelect("general-ev", [...new Set([...evs, ...v.general_evidence_values])], v.general_evidence_values)}`;
    } else h += `<p class="muted">전표 파일을 불러오면 부가세 계정별 공급가액 열과 증빙 유형을 지정할 수 있습니다.</p>`;
    h += `<label>특수 증빙 제외 키워드 (불공제·영세율·면세·수정·분할 등, 쉼표 구분)</label><input type="text" id="vat-excl" value="${esc(v.excluded_evidence_keywords.join(", "))}">
      <div class="row"><div><label>반올림 기준</label><select id="vat-round">${["", "버림", "반올림", "올림"].map((o) => `<option value="${o}"${o === v.rounding ? " selected" : ""}>${o || "(미설정 → 미검사)"}</option>`).join("")}</select></div>
      <div><label>세율</label><input type="text" id="vat-rate" value="${esc(v.rate)}"></div><div><label>허용 차이(원)</label><input type="text" id="vat-tol" value="${esc(v.tolerance)}"></div></div></div>
      <div class="panel"><h2>회사 승인 계정 규칙</h2><p class="muted">규칙이 적용되지 않는 분개는 계정 검사 '미검사'. 거래처만으로 계정을 확정하지 않습니다. 목록 값은 쉼표로 구분, 날짜는 YYYY-MM-DD.</p>
      <button data-action="add-examples">합성 데이터용 예시 규칙 추가 (비활성 상태로)</button>${rulesEditor("acc", d.account_rules, ACC_FIELDS)}</div>
      <div class="panel"><h2>적요 검사</h2><label>적요 빈칸 판정</label><select id="memo-blank">${["오류", "확인 필요"].map((o) => `<option${o === d.memo.blank_status ? " selected" : ""}>${o}</option>`).join("")}</select>
      <p class="muted">거래유형별 필수 적요 정보: 대상계정코드/적요조건에 해당하는 행의 적요에 '필수포함' 단어가 모두 있어야 합니다. 글자 수만으로 적정성을 판단하지 않습니다.</p>${rulesEditor("memo", d.memo.rules, MEMO_FIELDS)}</div>
      <div class="panel"><h2>증빙 요구 기준</h2>
      <label><input type="checkbox" id="ev-tax"${d.evidence.require_tax_invoice_for_vat ? " checked" : ""}> 부가세 행이 있는 전표는 세금계산서 필요</label>
      <label><input type="checkbox" id="ev-appr"${d.evidence.require_approval_doc ? " checked" : ""}> 모든 전표에 품의서 필요</label>
      <p class="muted">증빙이 연결·확인되지 않으면 증빙 검사는 '미검사'로 남고, 전체 요약도 '설정된 검사 통과'가 되지 않습니다.</p></div>
      <div class="row"><button class="primary" data-action="apply-settings">설정 적용 (이번 화면)</button><button data-action="save-settings">설정 적용 + 이 브라우저에 저장</button>
      <button data-action="default-settings">기본값으로 되돌리기</button><button data-action="export-settings">설정 파일 내보내기</button>
      <label style="flex:1 1 200px">설정 파일 가져오기<input type="file" id="import-settings" accept=".json"></label></div>
      <div class="panel"><h2>계정 질문 가이드 (제조업 참고 분개)</h2><p class="muted">버전 ${esc(S.guide.version || "-")} · ${esc(S.guide.기준 || "")}</p>
      <div class="row"><button data-action="export-guide">가이드 JSON 다운로드</button><label style="flex:1 1 200px">회사 기준으로 수정한 가이드 JSON 올리기<input type="file" id="import-guide" accept=".json"></label>
      ${lsGet(LS_GUIDE) ? `<button data-action="reset-guide">저장된 회사 가이드 삭제 (기본 가이드로)</button>` : ""}</div></div>`;
    return h;
  }
  function readDraftFromForm() {
    const d = S.draft, v = d.vat;
    const val = (id) => ($("#" + id) ? $("#" + id).value : null);
    if (val("vat-codes") !== null) v.account_codes = splitList(val("vat-codes"));
    if (val("vat-names") !== null) v.name_keywords = splitList(val("vat-names"));
    if (val("vat-excl") !== null) v.excluded_evidence_keywords = splitList(val("vat-excl"));
    if (val("vat-round") !== null) v.rounding = val("vat-round");
    if (val("vat-rate") !== null) v.rate = val("vat-rate").trim() || "0.1";
    if (val("vat-tol") !== null) v.tolerance = val("vat-tol").trim() || "0";
    document.querySelectorAll("[data-supply]").forEach((el) => { if (el.value) v.supply_columns[el.dataset.supply] = el.value; else delete v.supply_columns[el.dataset.supply]; });
    const gev = [...document.querySelectorAll('[data-ms="general-ev"]')];
    if (gev.length) v.general_evidence_values = gev.filter((x) => x.checked).map((x) => x.value);
    if (val("memo-blank") !== null) d.memo.blank_status = val("memo-blank");
    if ($("#ev-tax")) d.evidence.require_tax_invoice_for_vat = $("#ev-tax").checked;
    if ($("#ev-appr")) d.evidence.require_approval_doc = $("#ev-appr").checked;
    document.querySelectorAll("[data-rule]").forEach((el) => {
      const list = el.dataset.rule === "acc" ? d.account_rules : d.memo.rules;
      const r = list[+el.dataset.i], f = el.dataset.f;
      if (!r) return;
      if (el.type === "checkbox") r[f] = el.checked;
      else if (LIST_FIELDS.has(f)) r[f] = splitList(el.value);
      else r[f] = el.value;
    });
    d.account_rules = d.account_rules.filter((r) => String(r.규칙ID || "").trim() || r._new);
    d.memo.rules = d.memo.rules.filter((r) => String(r.규칙ID || "").trim() || r._new);
  }
  function cleanRules(list) { return list.filter((r) => String(r.규칙ID || "").trim()).map((r) => { const o = Object.assign({}, r); delete o._new; o.버전 = String(o.버전 || "1"); return o; }); }
  function applyDraft() {
    readDraftFromForm();
    const s = JSON.parse(JSON.stringify(S.draft));
    s.account_rules = cleanRules(s.account_rules); s.memo.rules = cleanRules(s.memo.rules);
    S.settings = E.mergeDefaults(s); S.draft = null;
    runChecks();
  }

  // ------------------------------------------------------------ 저장·불러오기 ----
  const DB_NAME = "voucher_review", STORE = "saves";
  function db() {
    return new Promise((res, rej) => {
      if (!window.indexedDB) return rej(new Error("이 브라우저는 저장 기능(IndexedDB)을 지원하지 않습니다."));
      const r = indexedDB.open(DB_NAME, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: "name" });
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  }
  async function dbOp(mode, fn) { const d = await db(); return new Promise((res, rej) => { const tx = d.transaction(STORE, mode); const st = tx.objectStore(STORE); const req = fn(st); tx.oncomplete = () => res(req && req.result); tx.onerror = () => rej(tx.error); }); }
  const dbList = () => dbOp("readonly", (st) => st.getAll());
  const dbPut = (rec) => dbOp("readwrite", (st) => st.put(rec));
  const dbDel = (name) => dbOp("readwrite", (st) => st.delete(name));
  function sessionText() {
    return E.serializeSession({ source: { name: S.sourceName, sha256: S.sourceSha, sheet: S.load.sheet, header_row: S.load.headerRow, preamble: S.load.preamble },
      raw: S.load.raw, rows: S.load.rows, reviews: S.reviews, evidence: S.evidence.toJSON(), settings: S.settings });
  }
  function reopen(text) {
    const d = E.deserializeSession(text);
    const load = { raw: d.raw, rows: d.rows, mapping: {}, sheet: d.source.sheet, headerRow: d.source.header_row, title: "", preamble: d.source.preamble || [], errors: [], warnings: [] };
    setLoaded(load, d.source.name, d.source.sha256, d.reviews);
    S.evidence = new E.EvidenceStore(d.evidence); S.settings = E.mergeDefaults(d.settings); S.draft = null; S.candidates = []; S.wb = null;
    runChecks();
  }
  function carryOver(text) {
    const d = E.deserializeSession(text);
    S.reviews = E.initReviews(E.contentHashes(S.load.rows), d.reviews);
    const keys = new Set(S.load.rows.map((r) => r.전표키));
    const st = new E.EvidenceStore(d.evidence); st.links = st.links.filter((l) => keys.has(l.voucher_key)); S.evidence = st;
    runChecks();
    return Object.values(S.reviews).filter((r) => r.변경감지).length;
  }
  let saveList = null;
  async function refreshSaves() { try { saveList = (await dbList()).sort((a, b) => (a.saved_at < b.saved_at ? 1 : -1)); } catch (e) { saveList = []; S.dbError = e.message; } }
  function renderStorage() {
    let h = `<p class="muted">작업은 <b>이 PC의 이 브라우저</b>에 저장됩니다(IndexedDB). 브라우저 사용 기록을 지우면 함께 지워질 수 있으니, 중요한 작업은 <b>작업 파일로 내보내기</b>로 따로 보관하세요. 저장에는 전표 데이터·검토 메모·수기 증빙이 들어 있습니다.</p>`;
    if (S.dbError) h += `<div class="msg warn">${esc(S.dbError)} — 작업 파일 내보내기/가져오기를 사용하세요.</div>`;
    if (ready()) h += `<div class="panel"><h2>현재 작업 저장</h2><label>저장 이름</label><input type="text" id="save-name" value="${esc(today() + "_" + S.sourceName.replace(/\.[^.]+$/, ""))}">
      <div class="row"><button class="primary" data-action="save-db">브라우저에 저장</button><button data-action="save-file">작업 파일로 내보내기 (.json)</button></div></div>`;
    h += `<div class="panel"><h2>작업 파일 가져오기</h2><label>저장해 둔 작업 파일(.json)</label><input type="file" id="import-session" accept=".json">
      <div class="row"><label><input type="radio" name="imp-mode" value="reopen" checked> 재열기 (저장 당시 데이터로)</label><label><input type="radio" name="imp-mode" value="carry"${ready() ? "" : " disabled"}> 현재 불러온 파일에 검토상태·증빙 이어받기</label></div></div>`;
    if (saveList === null) { refreshSaves().then(render); return h + `<p class="muted">저장 목록을 읽는 중…</p>`; }
    if (!saveList.length) return h + `<div class="msg info">브라우저에 저장된 작업이 없습니다.</div>`;
    h += `<div class="panel"><h2>브라우저에 저장된 작업</h2><select id="save-pick">${saveList.map((s) => `<option value="${esc(s.name)}">${esc(s.name)} (${esc(s.saved_at)})</option>`).join("")}</select>
      <div class="row"><button data-action="reopen-db">재열기 (저장 당시 데이터로)</button><button data-action="carry-db"${ready() ? "" : " disabled"}>현재 불러온 파일에 검토상태·증빙 이어받기</button><button data-action="export-db">이 저장을 파일로 내보내기</button></div></div>
      <div class="panel"><h2>저장 데이터 삭제</h2><p class="muted">원본 엑셀 파일은 삭제하지 않습니다. 선택한 저장 1개만 삭제합니다.</p>
      <select id="del-pick">${saveList.map((s) => `<option value="${esc(s.name)}">${esc(s.name)}</option>`).join("")}</select>
      <label>확인을 위해 삭제할 저장 이름을 그대로 입력하세요</label><input type="text" id="del-confirm"><button data-action="delete-db">저장 삭제</button></div>`;
    return h;
  }

  // ------------------------------------------------------------ 렌더 ----
  const TABS = [["review", "🔍 전표 검사"], ["qa", "💬 계정 질문하기"], ["inventory", "📋 계정과목 현황"], ["settings", "⚙️ 검사 설정"], ["storage", "💾 저장·불러오기"]];
  function render() {
    renderSidebar();
    let h = "";
    if (S.flash) h += `<div class="msg ${esc(S.flash.kind)}" id="flash">${esc(S.flash.text)}</div>`; // 다음 사용자 동작 때 지움
    if (S.load) { for (const e of S.load.errors) h += `<div class="msg error">${esc(e)}</div>`; for (const w of S.load.warnings) h += `<div class="msg warn">${esc(w)}</div>`; }
    if (ready()) { const ov = E.fileOverview(S.load.rows); h += `<div class="msg info">파일: <b>${esc(S.sourceName)}</b> · 시트 ${esc(S.load.sheet)} (${S.load.headerRow}행 헤더) · 회계일 ${esc(ov.회계일_시작)} ~ ${esc(ov.회계일_종료)} · 원본 ${ov.원본행수.toLocaleString()}행 / 전표 ${ov.전표수.toLocaleString()}건</div>`; }
    h += `<div class="tabs">${TABS.map(([k, l]) => `<button data-tab="${k}" class="${S.tab === k ? "active" : ""}">${l}</button>`).join("")}</div>`;
    h += ({ review: renderReview, qa: renderQA, inventory: renderInventory, settings: renderSettings, storage: renderStorage })[S.tab]();
    $("main").innerHTML = h;
  }

  // ------------------------------------------------------------ 이벤트 ----
  document.addEventListener("change", async (ev) => {
    const t = ev.target;
    if (t.type === "file" || t.id === "cand") S.flash = null;
    if (t.id === "file" && t.files[0]) return onFile(t.files[0]);
    if (t.id === "cand") { S.candIdx = +t.value; S.mapping = E.autoMapping(S.candidates[S.candIdx].headers); return render(); }
    if (t.dataset.map) { if (t.value === "(없음)") delete S.mapping[t.dataset.map]; else S.mapping[t.dataset.map] = t.value; return; }
    if (t.id === "reviewer") { S.reviewer = t.value; return; }
    if (t.dataset.filterDate !== undefined) {
      const dates = S.load.rows.map((r) => r["회계일"]).filter(Boolean).sort();
      const cur = S.filters.date ? S.filters.date.slice() : [dates[0], dates[dates.length - 1]];
      cur[+t.dataset.filterDate] = t.value || (t.dataset.filterDate === "0" ? dates[0] : dates[dates.length - 1]);
      S.filters.date = cur[0] === dates[0] && cur[1] === dates[dates.length - 1] ? undefined : cur; // 전체 기간이면 필터 안 함
      return render();
    }
    if (t.dataset.ms && S.tab === "review") {
      const name = t.dataset.ms;
      S.filters[name] = [...document.querySelectorAll(`[data-ms="${CSS.escape(name)}"]`)].filter((x) => x.checked).map((x) => x.value);
      const open = t.closest("details"); render(); if (open) { const d = document.querySelector(`[data-ms="${CSS.escape(name)}"]`); if (d) d.closest("details").open = true; }
      return;
    }
    if (t.id === "sort") { S.sort = t.value; return render(); }
    if (t.name === "evkind") { S.evKind = t.value; return render(); }
    if (t.id === "inv-kw") { S.invKeyword = t.value.trim(); return render(); }
    if (t.id === "question") { S.question = t.value; return render(); }
    if (t.id === "import-settings" && t.files[0]) {
      try { S.draft = E.mergeDefaults(JSON.parse(await t.files[0].text())); flash("info", "설정 파일을 읽었습니다. 내용을 확인한 뒤 '설정 적용'을 눌러 주세요."); } catch (e) { flash("error", "설정 파일을 읽을 수 없습니다: " + e.message); }
      return render();
    }
    if (t.id === "import-guide" && t.files[0]) {
      try { const g = JSON.parse(await t.files[0].text()); E.validateGuide(g); S.guide = g; lsSet(LS_GUIDE, g); flash("ok", "회사 가이드를 적용하고 이 브라우저에 저장했습니다."); } catch (e) { flash("error", "가이드 JSON 을 읽을 수 없습니다: " + e.message); }
      return render();
    }
    if (t.id === "import-session" && t.files[0]) {
      const mode = (document.querySelector('[name="imp-mode"]:checked') || {}).value;
      try {
        const text = await t.files[0].text();
        if (mode === "carry" && ready()) flash("ok", `이어받았습니다. 분개가 바뀐 전표 ${carryOver(text)}건은 검토완료를 승계하지 않고 '미검토'로 표시했습니다.`);
        else { reopen(text); flash("ok", "작업 파일을 재열었습니다 (저장 당시 설정으로 다시 검사)."); }
      } catch (e) { flash("error", "작업 파일을 읽을 수 없습니다: " + e.message); }
      return render();
    }
  });
  document.addEventListener("keydown", (ev) => { if (ev.key === "Enter" && ev.target.id === "question") { S.question = ev.target.value; render(); } });

  document.addEventListener("click", async (ev) => {
    const tab = ev.target.closest("[data-tab]");
    if (tab || ev.target.closest("[data-action]")) S.flash = null;
    if (tab) { if (S.tab === "settings" && S.draft) readDraftFromForm(); S.tab = tab.dataset.tab; if (S.tab === "storage") saveList = null; return render(); }
    const el = ev.target.closest("[data-action]");
    if (!el) return;
    const a = el.dataset.action;
    try {
      if (a === "load") {
        const res = E.loadVouchers(S.candidates[S.candIdx], S.mapping);
        setLoaded(res, S.sourceName, S.sourceSha);
        flash(res.errors.length ? "error" : "info", res.errors.length ? "불러오지 못했습니다. 아래 안내를 확인해 주세요." : `불러왔습니다: ${S.sourceName}. '검사 실행'을 눌러 주세요.`);
      } else if (a === "run") { runChecks(); flash("ok", "검사를 실행했습니다. 검토상태는 바뀌지 않습니다."); S.tab = "review"; }
      else if (a === "reset") {
        Object.assign(S, { wb: null, candidates: [], load: null, inventory: null, results: null, summary: null, reviews: {}, evidence: new E.EvidenceStore(), selectedKey: null, invCode: null, filters: {}, question: "", draft: null, sourceName: "", sourceSha: "" });
        const f = $("#file"); if (f) f.value = "";
        flash("ok", "현재 화면의 작업을 비웠습니다. 원본 파일과 브라우저에 저장한 작업은 삭제하지 않았습니다.");
      } else if (a === "select") { S.selectedKey = el.dataset.key; render(); const d = $("#detail"); if (d) d.scrollIntoView({ behavior: "smooth" }); return; }
      else if (a === "clear-filters") { S.filters = {}; }
      else if (a === "save-review") {
        E.updateReview(S.reviews, el.dataset.key, { status: $("#rv-status").value, memo: $("#rv-memo").value, reviewer: $("#rv-reviewer").value });
        flash("ok", "검토 내용을 반영했습니다. (보관하려면 '저장·불러오기' 탭에서 저장)");
      } else if (a === "unlink") { S.evidence.unlink(el.dataset.key, el.dataset.doc); runChecks(); flash("ok", "증빙 연결을 해제하고 다시 검사했습니다."); }
      else if (a === "add-evidence") {
        const kind = S.evKind || E.TAX_INVOICE, vals = {}, errs = [];
        document.querySelectorAll("[data-ev]").forEach((x) => { vals[x.dataset.ev] = x.value.trim(); });
        for (const k of ["공급가액", "세액", "합계", "금액"]) if (k in vals && vals[k] !== "") { const p = E.parseAmount(vals[k]); if (p.st === "오류") errs.push(`${k} 값을 해석할 수 없습니다: ${vals[k]}`); else vals[k] = p.amt; }
        if (errs.length) { flash("error", errs.join(" / ")); return render(); }
        const approval = vals.__approval, file = vals.__file; delete vals.__approval; delete vals.__file;
        vals.출처 = "수기입력";
        const clean = Object.fromEntries(Object.entries(vals).filter(([, v]) => v !== "" && v !== null));
        const d = S.evidence.addDoc(kind, clean, { user: S.reviewer, fileName: file, approvalDocId: approval });
        S.evidence.link(el.dataset.key, d.doc_id, "수동", { user: S.reviewer });
        runChecks(); flash("ok", `${kind}를 등록·연결하고 다시 검사했습니다.`);
      } else if (a === "export-results") {
        busy(true);
        const counts = E.summaryCounts(S.summary);
        const sheets = E.buildResultSheets(S.summary, S.results, S.load.raw, S.load.rows, S.reviews, criteriaRows());
        const p = E.checkExportCounts(sheets, S.load.rows.length, counts.전체);
        if (p.length) { busy(false); flash("error", "건수 대조 실패로 다운로드하지 않았습니다: " + p.join("; ")); return render(); }
        const buf = await E.toExcelBuffer(ExcelJS, sheets);
        busy(false);
        download(`전표검사결과_${today()}.xlsx`, new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
        return;
      } else if (a === "example") { S.question = el.dataset.q; }
      else if (a === "ask") { S.question = ($("#question") || {}).value || ""; }
      else if (a === "inv-select") { S.invCode = el.dataset.key; }
      else if (a === "export-inventory") {
        const buf = await E.toExcelBuffer(ExcelJS, { 계정과목현황: S.inventory, 코드명불일치: E.codeNameConflicts(S.load.rows), 적용기준: [
          { 항목: "원본 파일", 값: S.sourceName }, { 항목: "시트/헤더행", 값: `${S.load.sheet} / ${S.load.headerRow}` },
          { 항목: "차대구분", 값: "차변금액만 있으면 차변, 대변금액만 있으면 대변, 둘 다 있으면 차대동시" }, { 항목: "빈 금액", 값: "0 으로 합산(상태는 빈값으로 보존)" },
          { 항목: "해석 불가 금액", 값: "0 대체 없이 금액오류로 분리, 합계 제외" }, { 항목: "외화", 값: "원화 합계에 포함하지 않음" }] });
        download("계정과목_현황.xlsx", new Blob([buf])); return;
      } else if (a === "add-rule" || a === "del-rule" || a === "add-examples") {
        readDraftFromForm();
        const list = el.dataset.kind === "memo" ? S.draft.memo.rules : S.draft.account_rules;
        if (a === "add-rule") list.push({ 규칙ID: "", 활성: false, 버전: "1", 차대구분: "전체", _new: true });
        else if (a === "del-rule") list.splice(+el.dataset.i, 1);
        else {
          const have = new Set(S.draft.account_rules.map((r) => r.규칙ID)); S.draft.account_rules.push(...EXAMPLE_RULES.account_rules.filter((r) => !have.has(r.규칙ID)));
          const hm = new Set(S.draft.memo.rules.map((r) => r.규칙ID)); S.draft.memo.rules.push(...EXAMPLE_RULES.memo_rules.filter((r) => !hm.has(r.규칙ID)));
          flash("info", "예시 규칙을 비활성 상태로 추가했습니다. 합성 데이터 시연용이며 실제 데이터에는 쓰지 마세요.");
        }
      } else if (a === "apply-settings") { applyDraft(); flash("ok", `적용했습니다. 설정 버전 ${E.settingsVersion(S.settings)} 로 다시 검사했습니다.`); }
      else if (a === "save-settings") {
        applyDraft();
        const ok = lsSet(LS_SETTINGS, S.settings);
        flash(ok ? "ok" : "warn", ok ? `적용하고 이 브라우저에 저장했습니다 (설정 버전 ${E.settingsVersion(S.settings)}).` : "적용했지만 브라우저 저장에 실패했습니다. 설정 파일 내보내기를 사용하세요.");
      }
      else if (a === "default-settings") { S.settings = E.defaultSettings(); S.draft = null; lsDel(LS_SETTINGS); flash("info", "기본 설정으로 되돌렸습니다. 검사 실행을 다시 눌러 주세요."); }
      else if (a === "export-settings") { readDraftFromForm(); download("검사설정.json", new Blob([JSON.stringify(S.settings, null, 2)], { type: "application/json" })); return; }
      else if (a === "export-guide") { download("account_guide.json", new Blob([JSON.stringify(S.guide, null, 2)], { type: "application/json" })); return; }
      else if (a === "reset-guide") { if (!confirm("저장된 회사 가이드를 삭제하고 기본 가이드로 되돌릴까요?")) return; lsDel(LS_GUIDE); S.guide = DEFAULT_GUIDE; flash("ok", "기본 가이드로 되돌렸습니다."); }
      else if (a === "save-db") {
        const name = ($("#save-name").value || "").trim() || today();
        await dbPut({ name, saved_at: new Date().toLocaleString("ko-KR"), text: sessionText() }); saveList = null;
        flash("ok", `브라우저에 저장했습니다: ${name}`);
      } else if (a === "save-file") { const name = ($("#save-name").value || "").trim() || today(); download(`${name}.json`, new Blob([sessionText()], { type: "application/json" })); return; }
      else if (a === "reopen-db" || a === "carry-db" || a === "export-db") {
        const rec = saveList.find((s) => s.name === $("#save-pick").value);
        if (a === "export-db") { download(`${rec.name}.json`, new Blob([rec.text], { type: "application/json" })); return; }
        if (a === "reopen-db") { reopen(rec.text); flash("ok", `재열었습니다: ${rec.name} (저장 당시 설정으로 다시 검사)`); }
        else flash("ok", `이어받았습니다. 분개가 바뀐 전표 ${carryOver(rec.text)}건은 검토완료를 승계하지 않고 '미검토'로 표시했습니다.`);
      } else if (a === "delete-db") {
        const target = $("#del-pick").value;
        if ($("#del-confirm").value !== target) { flash("error", "확인용 이름이 일치하지 않아 삭제하지 않았습니다."); return render(); }
        await dbDel(target); saveList = null; flash("ok", `삭제했습니다: ${target}`);
      }
    } catch (e) { busy(false); flash("error", "처리 중 오류: " + e.message); console.error(e); }
    render();
  });

  window.addEventListener("error", (e) => { const m = $("main"); if (m) m.insertAdjacentHTML("afterbegin", `<div class="msg error">화면 오류: ${esc(e.message)}</div>`); });
  window.__voucherState = S; // 테스트용
  render();
})();
