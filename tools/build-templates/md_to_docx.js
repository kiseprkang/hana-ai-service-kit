// 실습 가이드 Markdown을 Word(.docx)로 변환합니다. md 내용을 그대로 옮기므로 md를 고치고 이 스크립트만 다시 실행하면 됩니다.
// 지원: #/##/### 제목, 문단, 불릿(-, *), 체크(- [ ]), 번호 문단, 인용(>), 표(|), 코드 블록(```), **굵게**, `코드`
// 실행: NODE_PATH=<전역 npm 모듈 경로> node tools/build-templates/md_to_docx.js
// 출력: 3회차 강의자료/수강생_실습_가이드.docx

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, HeadingLevel, AlignmentType, Footer, PageNumber, LevelFormat,
} = require("docx");

const SRC = path.join(__dirname, "..", "..", "practice", "실습_가이드.md");
// 출력 파일 이름은 인자로 바꿀 수 있습니다. 기본값은 3회차 강의자료 폴더의 "수강생_실습_가이드_전체판.docx"입니다.
const OUT = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(__dirname, "..", "..", "..", "수강생_실습_가이드_전체판.docx");
const FONT = "맑은 고딕";
const CODE_FONT = "Consolas";
const CONTENT_WIDTH = 9638;
const GRAY = "7F7F7F";
const TEAL = "028090";
const RED = "C00000";
const HEADER_FILL = "D9E2F3";
const CODE_FILL = "F2F2F2";

function run(text, opts = {}) {
  return new TextRun({
    text,
    font: opts.font || FONT,
    size: opts.size || 21,
    bold: opts.bold,
    italics: opts.italics,
    color: opts.color,
  });
}

// **굵게**와 `코드`를 해석해서 TextRun 배열을 만듭니다.
function inlineRuns(text, base = {}) {
  const parts = text.split(/(\*\*[^*]+\*\*|`[^`]+`)/).filter((s) => s.length > 0);
  return parts.map((part) => {
    if (part.startsWith("**") && part.endsWith("**") && part.length > 4) {
      return run(part.slice(2, -2), { ...base, bold: true });
    }
    if (part.startsWith("`") && part.endsWith("`") && part.length > 2) {
      return run(part.slice(1, -1), { ...base, font: CODE_FONT, size: (base.size || 21) - 2 });
    }
    return run(part, base);
  });
}

const paragraph = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: opts.after ?? 120 },
    indent: opts.indent,
    children: inlineRuns(text, opts),
  });

const titleBlock = (text) =>
  new Paragraph({
    spacing: { after: 200 },
    children: inlineRuns(text, { size: 40, bold: true, color: "1F3864" }),
  });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 320, after: 140 },
    children: inlineRuns(text, { size: 30, bold: true, color: "1F3864" }),
  });

const h2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 220, after: 100 },
    children: inlineRuns(text, { size: 24, bold: true, color: "262626" }),
  });

const h3 = (text) =>
  new Paragraph({
    keepNext: true,
    spacing: { before: 160, after: 80 },
    children: inlineRuns(text, { size: 22, bold: true, color: "404040" }),
  });

function quote(text) {
  let color = GRAY;
  if (text.startsWith("💡")) color = TEAL;
  if (text.startsWith("⚠")) color = RED;
  return new Paragraph({
    spacing: { after: 120 },
    indent: { left: 360 },
    children: inlineRuns(text, { color }),
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: inlineRuns(text),
  });
}

function codeBlock(lines) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_WIDTH, type: WidthType.DXA },
            shading: { fill: CODE_FILL, type: ShadingType.CLEAR, color: "auto" },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: (lines.length ? lines : [""]).map(
              (line) => new Paragraph({ children: [run(line, { font: CODE_FONT, size: 19 })] }),
            ),
          }),
        ],
      }),
    ],
  });
}

// 표 열 너비는 글자 수에 비례해 정합니다 (최소 4, 최대 30 글자 기준).
function tableBlock(rawLines) {
  const rows = rawLines
    .filter((r) => !/^\|\s*:?-{2,}/.test(r.trim()))
    .map((r) => r.trim().replace(/^\|/, "").replace(/\|$/, "").split("|").map((c) => c.trim()));
  const cols = Math.max(...rows.map((r) => r.length));
  const weights = Array.from({ length: cols }, (_, c) =>
    Math.min(30, Math.max(4, ...rows.map((r) => (r[c] || "").length))),
  );
  const total = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((CONTENT_WIDTH * w) / total));
  widths[cols - 1] += CONTENT_WIDTH - widths.reduce((a, b) => a + b, 0);

  const makeCell = (text, width, header) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: header ? { fill: HEADER_FILL, type: ShadingType.CLEAR, color: "auto" } : undefined,
      margins: { top: 70, bottom: 70, left: 110, right: 110 },
      children: [new Paragraph({ children: inlineRuns(text, { size: 19, bold: header }) })],
    });

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(
      (cells, rowIndex) =>
        new TableRow({
          tableHeader: rowIndex === 0,
          children: widths.map((w, c) => makeCell(cells[c] || "", w, rowIndex === 0)),
        }),
    ),
  });
}

// md 한 파일을 docx 블록 배열로 바꿉니다.
function markdownToBlocks(markdown) {
  const lines = markdown.split(/\r?\n/);
  const blocks = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    if (trimmed.startsWith("```")) {
      const code = [];
      i += 1;
      while (i < lines.length && !lines[i].trim().startsWith("```")) {
        code.push(lines[i]);
        i += 1;
      }
      blocks.push(codeBlock(code));
      blocks.push(paragraph("", { after: 60 }));
      i += 1;
      continue;
    }

    if (trimmed.startsWith("|")) {
      const tableLines = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) {
        tableLines.push(lines[i]);
        i += 1;
      }
      blocks.push(tableBlock(tableLines));
      blocks.push(paragraph("", { after: 60 }));
      continue;
    }

    if (trimmed === "" || trimmed === "---") {
      i += 1;
      continue;
    }
    if (trimmed.startsWith("#### ")) blocks.push(h3(trimmed.slice(5)));
    else if (trimmed.startsWith("### ")) blocks.push(h2(trimmed.slice(4)));
    else if (trimmed.startsWith("## ")) blocks.push(h1(trimmed.slice(3)));
    else if (trimmed.startsWith("# ")) blocks.push(titleBlock(trimmed.slice(2)));
    else if (trimmed.startsWith("> ")) blocks.push(quote(trimmed.slice(2)));
    else if (/^[-*] \[[ xX]\] /.test(trimmed)) {
      const checked = /^[-*] \[[xX]\]/.test(trimmed);
      const text = trimmed.replace(/^[-*] \[[ xX]\] /, "");
      blocks.push(paragraph(`${checked ? "☑" : "☐"} ${text}`, { indent: { left: 360 }, after: 60 }));
    } else if (/^[-*] /.test(trimmed)) blocks.push(bullet(trimmed.slice(2)));
    else if (/^\d+\. /.test(trimmed) || /^[①-⑨] /.test(trimmed)) {
      blocks.push(paragraph(trimmed, { indent: { left: 360 }, after: 80 }));
    } else blocks.push(paragraph(trimmed));

    i += 1;
  }
  return blocks;
}

const footer = () =>
  new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ["- ", PageNumber.CURRENT, " -"], font: FONT, size: 16, color: GRAY })],
      }),
    ],
  });

function build() {
  const markdown = fs.readFileSync(SRC, "utf8");
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 21 } } } },
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 540, hanging: 270 } } },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
          },
        },
        footers: { default: footer() },
        children: markdownToBlocks(markdown),
      },
    ],
  });

  return Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(OUT, buffer);
    console.log("작성:", OUT);
  });
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
