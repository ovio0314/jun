const test = require("node:test");
const assert = require("node:assert");
const E = require("../src/engine.js");

test("sha256 표준 벡터", () => {
  assert.strictEqual(E.sha256("abc"), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
  assert.strictEqual(E.sha256(""), "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
  assert.strictEqual(E.sha256("전표"), require("crypto").createHash("sha256").update("전표").digest("hex"));
  const long = "가".repeat(1000);
  assert.strictEqual(E.sha256(long), require("crypto").createHash("sha256").update(long).digest("hex"));
});

test("금액 해석: 0 으로 대체하지 않음", () => {
  const p = (v) => { const r = E.parseAmount(v); return [r.amt === null ? null : r.amt.toString(), r.st]; };
  assert.deepStrictEqual(p("1,234,000"), ["1234000", "정상"]);
  assert.deepStrictEqual(p("(1,000)"), ["-1000", "정상"]);
  assert.deepStrictEqual(p("△500"), ["-500", "정상"]);
  assert.deepStrictEqual(p(0.1), ["0.1", "정상"]);
  assert.deepStrictEqual(p(""), [null, "빈값"]);
  assert.deepStrictEqual(p("abc"), [null, "오류"]);
  assert.deepStrictEqual(p(true), [null, "오류"]);
});

test("수식 주입 방지", () => {
  assert.strictEqual(E.safeCell("=SUM(A1)"), "'=SUM(A1)");
  assert.strictEqual(E.safeCell("@x"), "'@x");
  assert.strictEqual(E.safeCell("정상"), "정상");
  assert.strictEqual(E.safeCell(new E.Big("1000")), 1000);
  assert.strictEqual(E.safeCell(new E.Big("0.5")), "0.5");
});

test("플래그 구분", () => {
  assert.strictEqual(E.classifyFlag("FALSE"), "문자열 FALSE");
  assert.strictEqual(E.classifyFlag(false), "불리언 FALSE");
  assert.strictEqual(E.classifyFlag(null), "빈값");
});
