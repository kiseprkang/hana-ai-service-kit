// 출시 전 점검: 필수 문서 누락 여부와 게이트 판정 상태를 확인합니다.
// 사용법: node scripts/check-gates.mjs
// 종료 코드: 0 = 모두 통과, 1 = 누락 또는 미검토·보류 게이트 있음

import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const REQUIRED_DOCS = [
  "00_과제정의서.md",
  "01_고객_리서치.md",
  "02_PRD.md",
  "03_요구사항정의서.md",
  "04_기능명세서.md",
  "05_화면설계서.md",
  "08_검증계획.md",
  "09_의사결정기록.md",
  "10_출시점검표.md",
];
const PASSING_STATUSES = ["통과", "조건부 통과", "해당 없음"];
const STATUS_PATTERN = /\*\*판정:\*\*\s*(.+)/;

function checkDocs() {
  const missing = REQUIRED_DOCS.filter((name) => !existsSync(join(ROOT, "docs", name)));
  return missing;
}

function checkGates() {
  const gateDir = join(ROOT, "gates");
  const files = readdirSync(gateDir).filter((name) => name.endsWith(".md"));
  return files.map((name) => {
    const text = readFileSync(join(gateDir, name), "utf8");
    const match = text.match(STATUS_PATTERN);
    const status = match ? match[1].trim() : "판정 줄 없음";
    const passed = PASSING_STATUSES.includes(status);
    return { name, status, passed };
  });
}

function main() {
  const missingDocs = checkDocs();
  const gates = checkGates();

  console.log("== 필수 문서 ==");
  if (missingDocs.length === 0) {
    console.log("모두 존재합니다.");
  } else {
    missingDocs.forEach((name) => console.log(`  [누락] docs/${name}`));
  }

  console.log("\n== 게이트 판정 ==");
  gates.forEach((gate) => {
    const mark = gate.passed ? "[통과]" : "[미완료]";
    console.log(`  ${mark} ${gate.name}: ${gate.status}`);
  });

  const blocked = gates.filter((gate) => !gate.passed);
  const ok = missingDocs.length === 0 && blocked.length === 0;
  console.log(
    ok
      ? "\n결과: 출시 점검 조건을 충족합니다. (최종 판단은 사람이 합니다)"
      : `\n결과: 미완료 항목 ${missingDocs.length + blocked.length}건이 있습니다.`,
  );
  process.exit(ok ? 0 : 1);
}

main();
