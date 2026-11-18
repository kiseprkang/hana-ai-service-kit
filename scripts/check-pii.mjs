// Claude Code PreToolUse 자동 점검 장치: Write/Edit 내용에 개인식별 번호 형식이 있으면 저장을 차단합니다.
// 입력: 표준 입력으로 전달되는 JSON (tool_input.content, new_string, edits 등)
// 차단 시 종료 코드 2 (stderr 메시지가 AI 작업 도구에게 전달됨)

const PATTERNS = [
  { label: "주민등록번호 형식", regex: /\b\d{6}-?[1-4]\d{6}\b/ },
  { label: "카드번호 형식", regex: /\b\d{4}[- ]\d{4}[- ]\d{4}[- ]\d{4}\b/ },
];

function collectText(toolInput = {}) {
  const parts = [];
  if (typeof toolInput.content === "string") parts.push(toolInput.content);
  if (typeof toolInput.new_string === "string") parts.push(toolInput.new_string);
  if (Array.isArray(toolInput.edits)) {
    toolInput.edits.forEach((edit) => {
      if (typeof edit.new_string === "string") parts.push(edit.new_string);
    });
  }
  return parts.join("\n");
}

function findViolations(text) {
  return PATTERNS.filter(({ regex }) => regex.test(text)).map(({ label }) => label);
}

function readStdin() {
  return new Promise((resolve, reject) => {
    let data = "";
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (chunk) => (data += chunk));
    process.stdin.on("end", () => resolve(data));
    process.stdin.on("error", reject);
  });
}

async function main() {
  const raw = await readStdin();
  let payload;
  try {
    // 일부 셸은 입력 앞에 BOM(﻿)을 붙이므로 제거합니다.
    payload = JSON.parse((raw || "{}").replace(/^﻿/, ""));
  } catch {
    // 입력을 해석할 수 없으면 차단하지 않고 통과시킵니다 (자동 점검 장치 오류로 작업이 멈추지 않도록).
    process.exit(0);
  }

  const violations = findViolations(collectText(payload.tool_input));
  if (violations.length === 0) process.exit(0);

  console.error(
    `[보안 자동 점검 장치] 저장이 차단되었습니다: ${violations.join(", ")}이(가) 감지되었습니다.`,
  );
  console.error("실제 번호는 마스킹하세요. 예: 900101-1****** , 1234-****-****-5678");
  process.exit(2);
}

main();
