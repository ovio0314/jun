// 파이썬 엔진과의 동등성 검사용: xlsx + 설정(JSON) → 검사 결과 JSON (stdout)
const fs = require("fs");
const path = require("path");
const E = require("../src/engine.js");
const ExcelJS = require("exceljs");
(async () => {
  const [xlsx, settingsPath, evidencePath] = process.argv.slice(2);
  const wb = await E.loadWorkbook(ExcelJS, fs.readFileSync(xlsx));
  const cand = E.findHeaderCandidates(wb)[0];
  const load = E.loadVouchers(cand);
  const settings = settingsPath ? E.mergeDefaults(JSON.parse(fs.readFileSync(settingsPath, "utf8"))) : E.defaultSettings();
  let store;
  if (evidencePath) {
    const ev = JSON.parse(fs.readFileSync(evidencePath, "utf8"));
    store = new E.EvidenceStore();
    for (const d of ev) { const doc = store.addDoc(d.type, d.values); store.link(d.key, doc.doc_id, "수동"); }
  }
  const results = E.runChecks(load.rows, settings, store);
  const summary = E.voucherSummary(load.rows, results);
  const s = (v) => (v instanceof E.Big ? v.toString() : v);
  const guide = JSON.parse(fs.readFileSync(path.join(__dirname, "../../voucher/data/manufacturing_guide.json"), "utf8"));
  const inv = E.accountInventory(load.rows);
  const adv = new E.AccountAdvisor(guide, load.rows, inv);
  const questions = ["공장 전기요금 낼 때 차변 대변 뭐 써요?", "원재료 외상으로 샀어요", "외상매입금 대변에 쓰나요", "거래처 접대 식사 법인카드", "본사 사무실 전기요금", "우주선 발사"];
  console.log(JSON.stringify({
    errors: load.errors, warnings: load.warnings, preamble: load.preamble, header_row: load.headerRow,
    keys: load.rows.map((r) => r.전표키),
    results: results.map((r) => { const o = Object.assign({}, r); delete o.적용규칙버전; return o; }),
    summary: summary.map((x) => Object.fromEntries(Object.entries(x).map(([k, v]) => [k, s(v)]))),
    overview: E.fileOverview(load.rows),
    inventory: inv.map((x) => Object.fromEntries(Object.entries(x).map(([k, v]) => [k, s(v)]))),
    qa: questions.map((q) => { const a = adv.ask(q); return { context: a.context, side: a.side, ids: a.matches.map((m) => m.scenario.id), preferred: a.matches.map((m) => m.preferred), terms: a.terms, similar: a.similar, mentioned: a.mentioned.map((m) => m.계정과목코드) }; }),
    hashes: E.contentHashes(load.rows),
  }));
})().catch((e) => { console.error(e); process.exit(1); });
