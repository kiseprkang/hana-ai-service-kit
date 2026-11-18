// 화면 설계서(.pptx) 템플릿 만들기 프로그램
// 실행: NODE_PATH=<전역 npm 모듈 경로> node tools/build-templates/build_pptx.js
// 출력: templates/pptx/05_화면설계서.pptx

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const OUT_DIR = path.join(__dirname, "..", "..", "templates", "pptx");
const FONT = "Malgun Gothic"; // 한글 표시를 위해 Windows 기본 글꼴 사용
const COLOR = {
  navy: "1F3864",
  teal: "028090",
  ink: "262626",
  muted: "7F7F7F",
  panel: "F2F2F2",
  line: "BFBFBF",
  white: "FFFFFF",
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33" x 7.5"
pres.author = "교육용 템플릿";
pres.title = "05. 화면 설계서 템플릿";

const headerBlock = (slide, title, subtitle) => {
  slide.addText(title, {
    x: 0.6, y: 0.4, w: 12.1, h: 0.7,
    fontFace: FONT, fontSize: 28, bold: true, color: COLOR.navy, margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6, y: 1.1, w: 12.1, h: 0.4,
      fontFace: FONT, fontSize: 13, color: COLOR.muted, margin: 0,
    });
  }
};

const footerNote = (slide) => {
  slide.addText("교육용 템플릿 · 실제 화면은 담당 부서 가이드와 사내 디자인 시스템을 따릅니다", {
    x: 0.6, y: 6.95, w: 12.1, h: 0.3,
    fontFace: FONT, fontSize: 10, color: COLOR.muted, margin: 0,
  });
};

// 1. 표지
{
  const slide = pres.addSlide();
  slide.background = { color: COLOR.navy };
  slide.addText("05. 화면 설계서", {
    x: 0.9, y: 2.4, w: 11.5, h: 1.0,
    fontFace: FONT, fontSize: 44, bold: true, color: COLOR.white, margin: 0,
  });
  slide.addText("과제명, 버전, 작성자, 작성일을 입력하세요", {
    x: 0.9, y: 3.5, w: 11.5, h: 0.5,
    fontFace: FONT, fontSize: 18, color: "CADCFC", margin: 0,
  });
}

// 2. 작성 안내
{
  const slide = pres.addSlide();
  headerBlock(slide, "작성 안내", "화면 한 장당 한 페이지를 씁니다");
  const steps = [
    ["1", "화면 목록에 모든 화면을 등록하고 화면 ID를 부여합니다 (SCR-001부터)."],
    ["2", "화면마다 와이어프레임을 넣고, 아래 항목을 빠짐없이 채웁니다."],
    ["3", "안내 문구는 금융소비자보호법 설명 의무 대상이면 반드시 표기합니다 (G3)."],
    ["4", "모든 화면은 04 기능 명세서의 기능 ID와 연결합니다."],
  ];
  steps.forEach(([num, text], i) => {
    const y = 1.9 + i * 1.1;
    slide.addShape(pres.ShapeType.ellipse, {
      x: 0.7, y, w: 0.6, h: 0.6,
      fill: { color: COLOR.teal }, line: { color: COLOR.teal },
    });
    slide.addText(num, {
      x: 0.7, y, w: 0.6, h: 0.6, align: "center", valign: "middle",
      fontFace: FONT, fontSize: 18, bold: true, color: COLOR.white, margin: 0,
    });
    slide.addText(text, {
      x: 1.6, y, w: 11.0, h: 0.6, valign: "middle",
      fontFace: FONT, fontSize: 18, color: COLOR.ink, margin: 0, isTextBox: true,
    });
  });
  footerNote(slide);
}

// 3. 화면 목록 (표)
{
  const slide = pres.addSlide();
  headerBlock(slide, "화면 목록", "화면 ID, 진입 경로, 연결 기능을 한눈에 봅니다");
  const header = ["화면 ID", "화면명", "진입 경로", "주요 구성 요소", "연결 기능 (04)", "상태"]
    .map((text) => ({ text, options: { bold: true, fill: { color: "D9E2F3" }, color: COLOR.ink } }));
  const example = ["SCR-001", "자산 요약", "홈 > 자산 탭", "요약 카드, 계좌 목록", "FN-01", "초안"];
  const blank = ["", "", "", "", "", ""];
  const rows = [header, example, blank, blank, blank].map((row) =>
    row.map((text) => (typeof text === "string"
      ? { text, options: { color: COLOR.ink, fontFace: FONT } }
      : text)),
  );
  rows[1] = example.map((text) => ({ text, options: { italic: true, color: COLOR.muted, fontFace: FONT } }));
  slide.addTable(rows, {
    x: 0.6, y: 1.8, w: 12.1,
    colW: [1.4, 2.2, 2.4, 3.1, 1.9, 1.1],
    fontSize: 13, border: { type: "solid", pt: 0.75, color: COLOR.line },
    rowH: 0.5, valign: "middle",
  });
  footerNote(slide);
}

// 4. 화면 상세 템플릿 (화면마다 복제해서 사용)
function screenSlide(id, name) {
  const slide = pres.addSlide();
  headerBlock(slide, `${id}  ${name}`, "화면 상세 · 와이어프레임과 설명을 함께 적습니다");

  // 와이어프레임 자리
  slide.addShape(pres.ShapeType.rect, {
    x: 0.6, y: 1.8, w: 5.6, h: 4.9,
    fill: { color: COLOR.panel }, line: { color: COLOR.line, dashType: "dash", pt: 1 },
  });
  slide.addText("와이어프레임 삽입 위치\n(화면 캡처 또는 목업 이미지)", {
    x: 0.6, y: 1.8, w: 5.6, h: 4.9, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 14, color: COLOR.muted, margin: 0,
  });

  const labelOpt = { bold: true, fill: { color: "D9E2F3" }, color: COLOR.ink, fontFace: FONT };
  const rows = [
    ["목적", ""],
    ["입력 항목과 검증 규칙", ""],
    ["버튼과 동작", ""],
    ["안내 문구 (설명 의무 대상 시 표기)", ""],
    ["오류 상태와 메시지", ""],
    ["접근성 고려 사항", ""],
    ["연결 기능 (04)", ""],
  ].map(([label, value]) => [
    { text: label, options: labelOpt },
    { text: value, options: { color: COLOR.ink, fontFace: FONT } },
  ]);
  slide.addTable(rows, {
    x: 6.6, y: 1.8, w: 6.1,
    colW: [2.2, 3.9],
    fontSize: 13, border: { type: "solid", pt: 0.75, color: COLOR.line },
    rowH: 0.62, valign: "middle",
  });
  footerNote(slide);
  return slide;
}

screenSlide("SCR-001", "(화면명)");

fs.mkdirSync(OUT_DIR, { recursive: true });
pres.writeFile({ fileName: path.join(OUT_DIR, "05_화면설계서.pptx") }).then((file) => {
  console.log("작성:", file);
});
