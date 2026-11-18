// 연결 확인 점검: docs/03의 요구사항(REQ-)과 비기능 요구(NFR-)가 docs/08 검증 계획에 모두 연결됐는지 확인합니다.
// 사용법: node scripts/check-traceability.mjs
// 종료 코드: 0 = 모두 연결됨, 1 = 연결되지 않은 ID 있음 또는 파일 없음

import { existsSync, readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const REQUIREMENTS = join(ROOT, "docs", "03_요구사항정의서.md");
const TEST_PLAN = join(ROOT, "docs", "08_검증계획.md");
const ID_PATTERN = /\b(?:REQ|NFR)-\d{3}\b/g;

function readIds(filePath) {
  const text = readFileSync(filePath, "utf8");
  return [...new Set(text.match(ID_PATTERN) ?? [])];
}

function main() {
  const missingFiles = [REQUIREMENTS, TEST_PLAN].filter((file) => !existsSync(file));
  if (missingFiles.length > 0) {
    missingFiles.forEach((file) => console.log(`[파일 없음] ${file}`));
    process.exit(1);
  }

  const requirementIds = readIds(REQUIREMENTS);
  const testPlanText = readFileSync(TEST_PLAN, "utf8");

  if (requirementIds.length === 0) {
    console.log("03_요구사항정의서.md에 REQ-, NFR- 형식의 ID가 없습니다. ID를 먼저 부여하세요.");
    process.exit(1);
  }

  const unlinked = requirementIds.filter((id) => !testPlanText.includes(id));

  console.log("== 요구사항 ID 연결 확인 ==");
  requirementIds.forEach((id) => {
    const linked = !unlinked.includes(id);
    console.log(`  ${linked ? "[연결]" : "[미연결]"} ${id}`);
  });

  if (unlinked.length === 0) {
    console.log("\n결과: 모든 요구사항이 검증 계획에 연결되어 있습니다.");
    process.exit(0);
  }
  console.log(`\n결과: 연결되지 않은 요구사항 ${unlinked.length}건. docs/08의 시나리오 '연결 ID'에 추가하세요.`);
  process.exit(1);
}

main();
