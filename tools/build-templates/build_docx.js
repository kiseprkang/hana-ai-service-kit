// 문서형 템플릿(.docx) 생성 스크립트
// 실행: NODE_PATH=<전역 npm 모듈 경로> node tools/build-templates/build_docx.js
// 출력: templates/docx/*.docx

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, HeadingLevel, AlignmentType, Footer, PageNumber,
} = require("docx");

const OUT_DIR = path.join(__dirname, "..", "..", "templates", "docx");
const FONT = "맑은 고딕";
const CONTENT_WIDTH = 9638; // A4 여백 2cm 기준 본문 폭 (DXA)
const GRAY = "7F7F7F";
const HEADER_FILL = "D9E2F3";
const LABEL_FILL = "F2F2F2";
const GATES = [
  ["G1 전자금융거래법", ""],
  ["G2 개인신용정보·개인정보", ""],
  ["G3 금융소비자보호", ""],
  ["G4 특금법(AML)", ""],
  ["G5 전자금융 보안", ""],
  ["G6 내부통제", ""],
  ["G7 AI 거버넌스", ""],
];

const run = (text, opts = {}) =>
  new TextRun({
    text,
    font: FONT,
    size: opts.size || 20,
    bold: opts.bold,
    italics: opts.italics,
    color: opts.color,
  });

const para = (text, opts = {}) =>
  new Paragraph({
    alignment: opts.align,
    spacing: { after: opts.after ?? 120 },
    children: [run(text, opts)],
  });

const title = (text) =>
  new Paragraph({
    spacing: { after: 160 },
    children: [run(text, { size: 36, bold: true, color: "1F3864" })],
  });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 280, after: 120 },
    children: [run(text, { size: 28, bold: true, color: "1F3864" })],
  });

const guide = (text) =>
  new Paragraph({
    spacing: { after: 120 },
    children: [run(`작성 안내: ${text}`, { italics: true, color: GRAY, size: 18 })],
  });

const cell = (text, width, opts = {}) =>
  new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill
      ? { fill: opts.fill, type: ShadingType.CLEAR, color: "auto" }
      : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: String(text)
      .split("\n")
      .map((line) => new Paragraph({ children: [run(line, { bold: opts.bold, size: 18 })] })),
  });

// 표: 헤더 행 + 데이터 행. widths 합은 CONTENT_WIDTH와 같아야 합니다.
const grid = (headers, rows, widths) => {
  if (widths.reduce((a, b) => a + b, 0) !== CONTENT_WIDTH) {
    throw new Error(`표 폭 합계가 ${CONTENT_WIDTH}이 아님: ${headers.join("|")}`);
  }
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => cell(h, widths[i], { bold: true, fill: HEADER_FILL })),
      }),
      ...rows.map(
        (row) =>
          new TableRow({ children: row.map((c, i) => cell(c, widths[i])) }),
      ),
    ],
  });
};

// 항목/내용 2열 표 (왼쪽 항목 칸은 회색 배경)
const fields = (pairs) => {
  const left = 2600;
  return grid(
    ["항목", "내용"],
    pairs.map(([k, v]) => [k, v ?? ""]),
    [left, CONTENT_WIDTH - left],
  );
};

const spacer = () => para("", { after: 80 });

const metaTable = (docTitle) =>
  fields([
    ["문서명", docTitle],
    ["과제명", ""],
    ["작성자 / 소속", ""],
    ["작성일", ""],
    ["버전 / 상태", "v0.1 / 초안"],
  ]);

const checkbox = (text) => para(`☐ ${text}`, { after: 80 });

// ---------- 문서별 본문 ----------

function doc00() {
  return [
    title("00. 과제 정의서"),
    para("교육용 템플릿입니다. 실제 수치와 정책은 담당 부서 확인 후 기입하세요.", { color: GRAY, size: 18 }),
    metaTable("00. 과제 정의서"),
    spacer(),
    h1("1. 배경과 문제"),
    guide("현재 상황, 해결하려는 문제, 문제를 겪는 사람을 적습니다."),
    fields([["현재 상황", ""], ["해결하려는 문제", ""], ["문제를 겪는 사람", ""]]),
    h1("2. 목표와 성공 지표 (KPI)"),
    guide("측정 가능한 지표만 적습니다. 현재값이 없으면 [확인 필요]로 둡니다."),
    grid(
      ["지표", "현재값", "목표값", "측정 방법", "리뷰일"],
      [["", "", "", "", ""], ["", "", "", "", ""]],
      [2400, 1500, 1500, 2638, 1600],
    ),
    h1("3. 가설"),
    guide("우리는 (대상)이 (상황)에서 (해결책)을 쓰면 (결과)가 나타날 것이라고 믿습니다."),
    fields([["가설", ""], ["검증 방법", ""]]),
    h1("4. 범위"),
    fields([["포함", ""], ["제외", ""]]),
    h1("5. 규제 영향 1차 판단"),
    guide("해당 여부는 '예 / 아니오 / 확인 필요' 중 하나로 표시하고 근거를 적습니다. 상세 검토는 gates/ 문서에서 합니다."),
    grid(
      ["게이트", "해당 여부", "근거"],
      GATES.map(([g]) => [g, "", ""]),
      [3400, 2000, 4238],
    ),
    h1("6. 일정과 승인"),
    fields([["주요 일정", ""], ["승인자", ""]]),
  ];
}

function doc01() {
  return [
    title("01. 고객 리서치"),
    metaTable("01. 고객 리서치"),
    spacer(),
    h1("1. 리서치 목적과 방법"),
    guide("무엇을 알고 싶은지, 어떤 방법(인터뷰, 설문, 행동 데이터, 민원 분석)으로 조사하는지 적습니다."),
    fields([["조사 질문", ""], ["방법", ""], ["표본과 기간", ""]]),
    h1("2. 세그먼트"),
    grid(
      ["세그먼트", "특성", "주요 니즈", "불편 요인", "규모"],
      [["", "", "", "", ""], ["", "", "", "", ""]],
      [1700, 2000, 2200, 2238, 1500],
    ),
    h1("3. 고객 여정 (AS-IS)"),
    guide("현재 고객이 겪는 단계별 행동과 불편 지점을 적습니다."),
    grid(
      ["단계", "행동", "접점(화면·채널)", "불편·이탈 지점", "감정"],
      [["", "", "", "", ""], ["", "", "", "", ""], ["", "", "", "", ""]],
      [1500, 2100, 2000, 2638, 1400],
    ),
    h1("4. 인사이트 요약"),
    fields([["인사이트 1", ""], ["인사이트 2", ""], ["인사이트 3", ""]]),
    h1("5. 고객 보호 관점 점검"),
    checkbox("고객이 이해하기 어려운 용어나 조건이 있는가?"),
    checkbox("연령, 디지털 역량 차이를 고려했는가?"),
    checkbox("고객에게 불리하게 작용할 수 있는 부분이 있는가? (→ G3 검토)"),
  ];
}

function doc02() {
  return [
    title("02. PRD (제품 요구사항 문서)"),
    metaTable("02. PRD"),
    spacer(),
    h1("1. 제품 개요"),
    fields([["한 줄 정의", ""], ["해결하는 문제 (00 참조)", ""], ["대상 사용자", ""]]),
    h1("2. 목표와 비목표"),
    fields([["목표", ""], ["비목표 (이번 릴리스에서 하지 않는 것)", ""]]),
    h1("3. 핵심 사용자 시나리오"),
    fields([["시나리오 1", ""], ["시나리오 2", ""]]),
    h1("4. 기능 범위와 릴리스 단계"),
    grid(
      ["릴리스", "포함 기능", "출시 기준", "목표 시기"],
      [["R1 (MVP)", "", "", ""], ["R2", "", "", ""]],
      [1800, 3400, 3038, 1400],
    ),
    h1("5. 금융 규제 요건 연결"),
    guide("각 게이트에서 요구하는 내용을 PRD에 어떻게 반영했는지 적습니다."),
    grid(
      ["게이트", "요건 반영 내용", "현재 상태"],
      [
        ["G1 전자금융거래법", "", ""],
        ["G2 개인신용정보·개인정보 (수집 항목, 동의, 보유 기간)", "", ""],
        ["G3 금융소비자보호 (설명 의무, 적합성·적정성)", "", ""],
        ["G4 특금법 (고객확인, 의심거래 연계)", "", ""],
        ["G5 전자금융 보안 (인증, 접근 통제, 암호화)", "", ""],
        ["G6 내부통제 (승인, 변경 이력, 감사 로그)", "", ""],
        ["G7 AI 거버넌스 (해당 시 06 참조)", "", ""],
      ],
      [3600, 4038, 2000],
    ),
    h1("6. 성공 지표와 리스크"),
    fields([["성공 지표", "00 과제 정의서의 KPI 참조"], ["주요 리스크", ""], ["완화 방안", ""]]),
    h1("7. 미결 사항"),
    guide("확정되지 않은 내용은 [확인 필요]로 표시하고 담당자와 기한을 적습니다."),
    fields([["[확인 필요] 항목", ""], ["담당자 / 기한", ""]]),
  ];
}

function featureBlock(n) {
  return [
    h1(`기능 ${n}: (기능명)   관련 요구사항: REQ-___`),
    h1("1) 사용자 스토리"),
    guide("사용자로서, (상황)에서, (목적)을 위해 (행동)을 하고 싶다."),
    fields([["사용자 스토리", ""]]),
    h1("2) 전제 조건 · 기본 흐름"),
    fields([["로그인 / 인증 수준", ""], ["가입·동의 상태", ""], ["기본 흐름 1", ""], ["기본 흐름 2", ""], ["기본 흐름 3", ""]]),
    h1("3) 예외 흐름"),
    grid(
      ["상황", "처리 방식", "사용자 안내 문구"],
      [["", "", ""], ["", "", ""]],
      [3000, 3500, 3138],
    ),
    h1("4) 수용 기준 (Acceptance Criteria)"),
    checkbox(""),
    checkbox(""),
    h1("5) 데이터 항목"),
    grid(
      ["항목", "출처", "암호화 여부", "보유 기간", "개인정보 해당 (G2)"],
      [["", "", "", "", ""]],
      [2000, 2000, 1600, 1800, 2238],
    ),
    h1("6) 규제 확인 포인트"),
    fields([["G1 전자금융 거래 조건", ""], ["G3 판매 원칙 적용", ""], ["G6 감사 로그", ""]]),
  ];
}

function doc04() {
  return [
    title("04. 기능 명세서"),
    metaTable("04. 기능 명세서"),
    guide("개발팀이 바로 구현할 수 있을 정도로 구체적으로 적습니다. 기능마다 아래 블록을 복제해 사용하세요. 화면 상세는 05를 따릅니다."),
    ...featureBlock(1),
    ...featureBlock(2),
  ];
}

function doc06() {
  return [
    title("06. AI 기능 설계서"),
    para("AI 기능이 없는 과제는 '해당 없음'으로 표시하고 넘어갑니다.", { color: GRAY, size: 18 }),
    metaTable("06. AI 기능 설계서"),
    spacer(),
    h1("1. AI 기능 개요"),
    fields([["기능명", ""], ["사용 목적과 자동화 범위", ""], ["사람의 최종 판단 지점 (Human-in-the-loop)", ""]]),
    h1("2. 모델과 데이터"),
    fields([
      ["모델 / 서비스", "[확인 필요]"],
      ["호스팅 위치 (사내 / 외부)", ""],
      ["입력 데이터와 개인정보 포함 여부 (G2)", ""],
      ["학습·파인튜닝 데이터 출처", ""],
      ["외부 전송 여부와 차단 장치", ""],
    ]),
    h1("3. 평가 기준"),
    grid(
      ["항목", "기준값", "측정 방법", "리뷰일"],
      [["정확도", "[확인 필요]", "", ""], ["편향 (집단별 차이)", "", "", ""], ["환각·오답 시 처리", "", "", ""]],
      [2400, 2400, 3200, 1638],
    ),
    h1("4. 설명 가능성과 고객 안내"),
    fields([["AI 사용 안내 문구", ""], ["결과 근거 제공 방식", ""], ["고객 이의 제기 경로", ""]]),
    h1("5. 운영과 모니터링"),
    fields([["로그 항목과 보관 기간", ""], ["이상 탐지 기준", ""], ["중단(킬 스위치) 조건과 담당자", ""]]),
    h1("6. 게이트 연결"),
    fields([["G7 AI 거버넌스 검토 결과", ""], ["G2 개인정보 검토 결과", ""]]),
  ];
}

// ---------- 빌드 ----------

const DOCS = [
  ["00_과제정의서.docx", doc00],
  ["01_고객_리서치.docx", doc01],
  ["02_PRD.docx", doc02],
  ["04_기능명세서.docx", doc04],
  ["06_AI기능설계서.docx", doc06],
];

const footer = () =>
  new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ children: ["- ", PageNumber.CURRENT, " -"], font: FONT, size: 16, color: GRAY }),
        ],
      }),
    ],
  });

async function build() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  for (const [fileName, makeChildren] of DOCS) {
    const doc = new Document({
      styles: {
        default: {
          document: { run: { font: FONT, size: 20 } },
        },
      },
      sections: [
        {
          properties: {
            page: {
              size: { width: 11906, height: 16838 }, // A4
              margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
            },
          },
          footers: { default: footer() },
          children: makeChildren(),
        },
      ],
    });
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(path.join(OUT_DIR, fileName), buffer);
    console.log("작성:", fileName);
  }
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
