// 브라우저판 빌드: 라이브러리·데이터·화면을 HTML 파일 하나로 합친다 (외부 네트워크 불필요).
// 실행: npm install && npm run build   →  dist/전표검토.html
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const read = (p) => readFileSync(join(here, p), "utf8");
const pkg = (name) => JSON.parse(read(`node_modules/${name}/package.json`)).version;
// 인라인 스크립트 안의 </script 를 끊어 HTML 파서가 조기 종료하지 않게 한다
const inline = (js) => js.replace(/<\/script/gi, "<\\/script").replace(/<!--/g, "<\\!--");
const json = (obj) => JSON.stringify(obj).replace(/</g, "\\u003c");

const guide = JSON.parse(read("../voucher/data/manufacturing_guide.json"));
const rules = JSON.parse(read("../voucher/data/example_rules.json"));
const versions = { "big.js": pkg("big.js"), exceljs: pkg("exceljs") };
const built = new Date().toISOString().slice(0, 10);

const html = `<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<!-- 네트워크 차단: 이 페이지는 어떤 외부 주소로도 데이터를 보낼 수 없다 -->
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; connect-src 'none'; form-action 'none'; base-uri 'none'">
<title>전표 검토 도우미</title>
<style>${read("src/style.css")}</style>
</head>
<body>
<header><h1>📒 전표 검토 도우미</h1>
<p>브라우저판 · 설치 없음 · 인터넷 연결 없이 이 PC 안에서만 처리 · 결재·반려·ERP 수정 기능 없음 · 검사 통과는 최종 회계 적정성 보증이 아닙니다 · 빌드 ${built}</p></header>
<div class="layout"><aside></aside><main></main></div>
<div class="loading" id="loading">처리 중…</div>
<noscript>이 파일은 JavaScript 가 켜진 브라우저(크롬·엣지)에서 열어 주세요.</noscript>
<script>/* big.js ${versions["big.js"]} (MIT) */\n${inline(read("node_modules/big.js/big.js"))}</script>
<script>/* exceljs ${versions.exceljs} (MIT) */\n${inline(read("node_modules/exceljs/dist/exceljs.min.js"))}</script>
<script>window.VOUCHER_GUIDE=${json(guide)};window.VOUCHER_EXAMPLE_RULES=${json(rules)};window.VOUCHER_BUILD=${json({ built, versions })};</script>
<script>${inline(read("src/engine.js"))}</script>
<script>${inline(read("src/app.js"))}</script>
</body>
</html>
`;
mkdirSync(join(here, "dist"), { recursive: true });
const out = join(here, "dist", "전표검토.html");
writeFileSync(out, html);
console.log(`built ${out} (${(html.length / 1024 / 1024).toFixed(2)} MB)`, versions);
