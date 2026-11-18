"""표형 템플릿(.xlsx) 생성 스크립트.

실행: python tools/build-templates/build_xlsx.py
출력: templates/xlsx/*.xlsx, templates/gates/게이트_검토표.xlsx
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[2]
XLSX_DIR = ROOT / "templates" / "xlsx"
GATE_DIR = ROOT / "templates" / "gates"

FONT = "맑은 고딕"
HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLANK_ROWS = 12

JUDGMENT = ["미검토", "검토중", "통과", "조건부 통과", "보류", "해당 없음"]
YES_NO_NA = ["예", "아니오", "해당 없음"]
SEVERITY = ["치명", "중대", "경미"]
TEST_RESULT = ["미실행", "통과", "실패"]
DEFECT_STATUS = ["신규", "수정중", "완료", "보류"]


def add_dropdown(ws, column, first_row, last_row, options):
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(options) + '"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    dv.add(f"{column}{first_row}:{column}{last_row}")


def new_sheet(wb, name, title, guide, headers, widths, example=None, first=False):
    """제목(1행), 작성 안내(2행), 헤더(3행), 예시 행(선택), 빈 입력 행으로 구성된 시트를 만듭니다."""
    ws = wb.active if first else wb.create_sheet()
    ws.title = name

    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    ws["A2"] = guide
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="7F7F7F")

    for col, (header, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = Font(name=FONT, bold=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = width

    row = 4
    if example:
        for col, value in enumerate(example, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.font = Font(name=FONT, italic=True, color="7F7F7F")
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    first_input_row = row
    for _ in range(BLANK_ROWS):
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = Font(name=FONT)
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    ws.freeze_panes = "A4"
    ws.sheet_view.zoomScale = 110
    return ws, first_input_row, row - 1


def save(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    print("작성:", path.relative_to(ROOT))


def build_requirements():
    wb = Workbook()
    ws, first, last = new_sheet(
        wb, "기능요구사항", "03. 요구사항 정의서 — 기능 요구사항",
        "기능 영역 단위로 나눕니다. 세부 동작은 04 기능 명세서에 씁니다. 예시 행은 참고용이며 지우고 사용하세요.",
        ["ID", "기능 영역", "요구사항", "우선순위", "관련 PRD 항목", "관련 게이트", "상태"],
        [11, 16, 40, 12, 16, 14, 12],
        example=["REQ-001", "계좌 조회", "로그인 고객의 계좌별 잔액을 한 화면에 표시한다", "필수", "PRD 3번", "G2, G5", "초안"],
    )
    add_dropdown(ws, "D", first, last, ["필수", "권장", "선택"])
    add_dropdown(ws, "G", first, last, ["초안", "검토중", "확정", "폐기"])

    ws2, first, last = new_sheet(
        wb, "비기능요구사항", "03. 요구사항 정의서 — 비기능 요구사항",
        "성능, 보안, 가용성, 감사 로그 등 기준값이 필요한 요구를 적습니다. 기준값을 모르면 [확인 필요]로 둡니다.",
        ["ID", "구분", "요구사항", "기준값", "확인 방법"],
        [11, 14, 44, 20, 26],
        example=["NFR-001", "성능", "조회 화면 응답 시간", "[확인 필요]", "부하 테스트"],
    )
    add_dropdown(ws2, "B", first, last, ["성능", "보안", "가용성", "감사 로그", "기타"])
    save(wb, XLSX_DIR / "03_요구사항정의서.xlsx")


def build_process():
    wb = Workbook()
    headers = ["단계", "담당", "입력", "처리", "산출물", "소요 시간", "통제 지점"]
    widths = [8, 16, 20, 30, 20, 12, 26]
    ws, _, _ = new_sheet(
        wb, "AS-IS", "07. 업무 프로세스 — 현재 업무 (AS-IS)",
        "현재 내부 직원이 수행하는 업무 흐름을 단계별로 적습니다.",
        headers, widths,
        example=["1", "영업점 직원", "고객 서류", "서류 확인 후 시스템 입력", "신청서", "20분", "이중 확인 없음"],
    )
    ws, _, _ = new_sheet(
        wb, "TO-BE", "07. 업무 프로세스 — 목표 업무 (TO-BE)",
        "변경 후 업무 흐름을 같은 형식으로 적습니다. 통제 지점이 빠지지 않았는지 확인하세요.",
        headers, widths,
        example=["1", "영업점 직원", "고객 서류(전자 제출)", "자동 검증 후 승인 요청", "승인 요청서", "5분", "승인자 이중 확인 (G6)"],
    )
    ws, first, last = new_sheet(
        wb, "RACI", "07. 업무 프로세스 — 역할과 권한 (RACI)",
        "R=실무(수행), A=승인, C=검토·협의, I=통보. 한 업무에 A는 한 명만 둡니다.",
        ["업무", "실무 (R)", "검토 (C)", "승인 (A)", "감사 (I)", "비고"],
        [28, 16, 16, 16, 16, 30],
        example=["고객 한도 상향", "영업점", "리스크 관리팀", "팀장", "내부감사", "G6 예외 기록"],
    )
    save(wb, XLSX_DIR / "07_업무프로세스.xlsx")


def build_test_plan():
    wb = Workbook()
    ws, first, last = new_sheet(
        wb, "검증계획", "08. 검증 계획 — 검증 범위와 방법",
        "검증 유형별로 대상, 방법, 통과 기준을 정합니다. 보안·규제 점검은 해당 부서 기준을 따릅니다.",
        ["검증 유형", "대상", "방법", "담당", "통과 기준", "결과"],
        [16, 28, 30, 14, 26, 14],
        example=["보안 점검", "인증·암호화 (G5)", "취약점 점검, 모의해킹", "보안팀", "[확인 필요]", "미실행"],
    )
    add_dropdown(ws, "F", first, last, ["미실행", "진행중", "통과", "조건부 통과", "실패"])

    ws, first, last = new_sheet(
        wb, "시나리오", "08. 검증 계획 — 테스트 시나리오",
        "수용 기준(04 기능 명세서)을 시나리오로 옮깁니다. 결과는 실행 후 갱신합니다.",
        ["ID", "시나리오", "기대 결과", "연결 명세", "결과"],
        [10, 40, 40, 16, 14],
        example=["TC-001", "잔액 조회 화면 진입", "계좌별 잔액이 마스킹 규칙대로 표시됨", "FN-01", "미실행"],
    )
    add_dropdown(ws, "E", first, last, TEST_RESULT)

    ws, first, last = new_sheet(
        wb, "결함", "08. 검증 계획 — 결함 관리",
        "치명 결함은 출시 차단으로 처리합니다.",
        ["결함 ID", "내용", "심각도", "관련 시나리오", "상태", "담당"],
        [12, 40, 12, 16, 14, 14],
        example=["BUG-001", "로그에 계좌 전체 번호 노출", "치명", "TC-001", "수정중", "개발 A"],
    )
    add_dropdown(ws, "C", first, last, SEVERITY)
    add_dropdown(ws, "E", first, last, DEFECT_STATUS)
    save(wb, XLSX_DIR / "08_검증계획.xlsx")


def build_decisions():
    wb = Workbook()
    ws, first, last = new_sheet(
        wb, "결정기록", "09. 의사결정 기록",
        "결정마다 예측과 리뷰일을 함께 적습니다. 리뷰일에 예측과 실제 결과를 비교합니다.",
        ["ID", "날짜", "결정 내용", "배경", "검토한 대안", "예측 (성과·리스크)", "리뷰일", "실제 결과", "담당"],
        [9, 12, 30, 26, 26, 28, 12, 26, 12],
        example=["D-001", "2026-11-16", "조회 전용으로 1차 출시", "규제 검토 기간 단축", "거래 기능 포함", "이용률 낮을 수 있음", "2027-02-16", "", "기획 A팀"],
    )
    ws, first, last = new_sheet(
        wb, "예외승인", "09. 의사결정 기록 — 예외 승인 (게이트 우회 시에만)",
        "게이트를 우회하면 사유, 승인자, 보완 조치를 반드시 남깁니다.",
        ["게이트", "사유", "승인자", "보완 조치", "날짜"],
        [18, 36, 14, 36, 12],
    )
    save(wb, XLSX_DIR / "09_의사결정기록.xlsx")


def build_release_checklist():
    wb = Workbook()
    docs = [
        "00 과제 정의서", "01 고객 리서치 (또는 생략 사유)", "02 PRD",
        "03 요구사항 정의서", "04 기능 명세서", "05 화면 설계서",
        "06 AI 기능 설계서 (해당 시)", "07 업무 프로세스 (해당 시)", "08 검증 계획과 결과",
    ]
    ws, first, last = new_sheet(
        wb, "문서완료", "10. 출시 점검표 — 문서 완료",
        "각 문서의 완료 여부를 표시합니다. 해당 없음이면 사유를 비고에 적습니다.",
        ["문서", "완료 여부", "비고"],
        [40, 14, 40],
    )
    for i, name in enumerate(docs):
        ws.cell(row=first + i, column=1, value=name)
    add_dropdown(ws, "B", first, first + len(docs) - 1, YES_NO_NA)

    ws, first, last = new_sheet(
        wb, "게이트판정", "10. 출시 점검표 — 게이트 판정",
        "모든 게이트가 '통과' 또는 '조건부 통과'여야 출시 가능합니다. 보류가 있으면 출시를 완료로 표시하지 않습니다.",
        ["게이트", "판정", "조건부 통과 시 보완 기한", "승인자"],
        [30, 16, 26, 16],
    )
    gates = [
        "G1 전자금융거래법", "G2 개인신용정보·개인정보", "G3 금융소비자보호",
        "G4 특금법(AML)", "G5 전자금융 보안", "G6 내부통제", "G7 AI 거버넌스",
    ]
    for i, name in enumerate(gates):
        ws.cell(row=first + i, column=1, value=name)
        ws.cell(row=first + i, column=2, value="미검토")
    gate_last = first + len(gates) - 1
    add_dropdown(ws, "B", first, gate_last, JUDGMENT)

    summary_row = gate_last + 2
    ws.cell(row=summary_row, column=1, value="통과 또는 조건부 통과 개수").font = Font(name=FONT, bold=True)
    ws.cell(
        row=summary_row,
        column=2,
        value=f'=COUNTIF(B{first}:B{gate_last},"통과")+COUNTIF(B{first}:B{gate_last},"조건부 통과")',
    )
    ws.cell(row=summary_row + 1, column=1, value="전체 게이트 수").font = Font(name=FONT, bold=True)
    ws.cell(row=summary_row + 1, column=2, value=f"=ROWS(B{first}:B{gate_last})")
    ws.cell(row=summary_row + 2, column=1, value="출시 조건 충족 여부").font = Font(name=FONT, bold=True)
    ws.cell(
        row=summary_row + 2,
        column=2,
        value=f'=IF(B{summary_row}=B{summary_row + 1},"충족","미충족")',
    )

    ws, first, last = new_sheet(
        wb, "운영준비", "10. 출시 점검표 — 운영 준비",
        "운영 준비 항목은 담당자와 완료 여부를 함께 적습니다.",
        ["항목", "담당", "완료 여부", "비고"],
        [36, 16, 14, 30],
    )
    items = ["장애 대응 절차와 담당자", "고객 공지 문구 (금소법 확인)", "직원 교육 완료 (해당 시)", "롤백 계획"]
    for i, name in enumerate(items):
        ws.cell(row=first + i, column=1, value=name)
    add_dropdown(ws, "C", first, first + len(items) - 1, YES_NO_NA)

    ws, first, last = new_sheet(
        wb, "최종판단", "10. 출시 점검표 — 최종 판단",
        "최종 판단은 사람이 합니다. 판단 근거에는 결정 ID(09 의사결정 기록)를 적습니다.",
        ["구분", "내용"],
        [24, 60],
    )
    ws.cell(row=first, column=1, value="출시 가능 여부")
    add_dropdown(ws, "B", first, first, ["가능", "불가", "단계 출시"])
    ws.cell(row=first + 1, column=1, value="판단 근거와 결정 ID")
    save(wb, XLSX_DIR / "10_출시점검표.xlsx")


def build_as_built():
    wb = Workbook()
    ws, first, last = new_sheet(
        wb, "차이기록", "11. As-built — 설계와 실제 구현의 차이",
        "출시 후 실제 구현과 문서가 다른 부분을 기록하고, 문서에 반영했는지 표시합니다.",
        ["문서", "항목", "설계 시 내용", "실제 구현", "차이 사유", "문서 반영 여부"],
        [18, 22, 30, 30, 26, 14],
        example=["04 기능 명세서", "잔액 표시 주기", "조회 시점 1회", "5분 캐시", "성능 개선", "예"],
    )
    add_dropdown(ws, "F", first, last, YES_NO_NA)
    ws, first, last = new_sheet(
        wb, "리뷰기록", "11. As-built — 리뷰 기록",
        "리뷰 일정과 확인자, 후속 조치를 적습니다.",
        ["리뷰일", "확인자", "확인 내용", "후속 조치", "기한"],
        [12, 14, 40, 36, 12],
    )
    save(wb, XLSX_DIR / "11_As-built.xlsx")


GATE_CHECKLISTS = {
    "G1 전자금융거래법": [
        "전자금융거래 해당 여부와 전자금융업 등록·위탁 요건 확인 [확인 필요]",
        "접근매체(인증수단·카드·계좌 정보) 발급·관리 방식",
        "거래 내역 확인, 오류 정정 절차 제공",
        "사고 발생 시 이용자 통지와 손해 배상 절차",
        "본인 확인 방식과 거래 한도 설정 근거",
        "전자문서 교부·보관 방식",
    ],
    "G2 개인신용정보·개인정보": [
        "수집 항목 목록과 필수·선택 구분",
        "수집·이용 목적과 동의 방식 (필수/선택 동의 분리)",
        "신용정보 조회·제공 여부와 신용정보법상 근거 [확인 필요]",
        "제3자 제공·위탁 여부와 계약·고지 절차",
        "보유 기간과 파기 방법",
        "고유식별정보(주민등록번호 등) 처리 근거",
        "마케팅 활용 시 별도 동의 여부",
        "정보주체 권리(열람·정정·삭제·처리정지) 행사 절차",
        "로그와 테스트 데이터에 실제 개인정보가 없는지 확인",
    ],
    "G3 금융소비자보호": [
        "판매 대상 상품의 분류와 적용 규제 [확인 필요]",
        "적합성 원칙: 고객 정보 파악 절차와 부적합 시 안내",
        "적정성 원칙: 판단 기준과 안내 문구",
        "설명의무: 핵심 설명서, 중요 사항(위험·비용·해지 조건) 설명",
        "불공정영업행위 금지: 상품 강요, 부당 결합 판매 여부",
        "부당권유 금지: 오인 유발 표현, 수익 단정 표현 여부",
        "광고 규제: 필수 기재 사항과 심의 절차",
        "위법계약해지권(청약철회 등) 안내 방식",
        "민원·분쟁 처리 경로 안내",
        "고령·취약 고객을 위한 추가 안내",
    ],
    "G4 특금법(AML)": [
        "고객확인(CDD) 절차 적용 여부 (신규 가입, 거래 시점)",
        "실제소유자 확인 절차 (법인 고객)",
        "고위험 고객 판정 기준과 강화된 확인(EDD)",
        "제재 대상 스크리닝 연계",
        "의심거래 탐지 시나리오와 보고(STR) 연계 절차",
        "고액현금거래 보고 대상 해당 여부",
        "가상자산 관련 기능이 있다면 별도 의무 확인 [확인 필요]",
        "기록 보관 기간과 방식",
    ],
    "G5 전자금융 보안": [
        "인증 방식(다중 인증, 생체, 간편 인증)과 위험 기반 인증",
        "접근 권한 최소화와 권한 승인 이력",
        "민감 정보 암호화(전송·저장)와 키 관리",
        "세션 관리(만료, 재인증, 중복 로그인 차단)",
        "외부 연계 API 보안(인증, 전송 보호, 호출 제한)",
        "앱 보안(위변조 탐지, 루팅·후킹 탐지) [확인 필요]",
        "망분리 정책에 따른 개발·운영 환경 구분",
        "취약점 점검과 모의해킹 일정",
        "장애·침해 사고 대응 절차와 보고 체계",
        "로그 수집과 이상 탐지 연계",
    ],
    "G6 내부통제": [
        "준법감시인 또는 준법 부서 사전 검토 완료",
        "이해상충 가능성 점검 (상품 수익과 고객 이익)",
        "직무 분리: 기획, 개발, 운영, 승인 담당 분리",
        "변경 승인 절차와 승인 이력 보관",
        "예외 처리와 게이트 우회 기록 (09 예외승인)",
        "감사 로그 항목과 보관 기간",
        "내부 직원의 고객 정보 접근 통제와 기록",
        "리스크 평가 결과와 잔여 리스크 수용자",
        "업무 매뉴얼과 교육 계획 반영 (07 업무 프로세스)",
    ],
    "G7 AI 거버넌스": [
        "AI 사용 목적과 자동화 범위 (06 AI 기능 설계서)",
        "사람의 최종 판단 지점 (Human-in-the-loop)",
        "외부 AI 서비스 사용 시 데이터 반출 범위와 차단 장치",
        "개인신용정보·고객 식별 정보가 학습·프롬프트에 들어가지 않는가 (G2 연계)",
        "편향 점검(집단별 결과 차이)과 기준값",
        "환각·오답 시 처리 방식과 고객 보호 장치 (G3 연계)",
        "고객에게 AI 사용을 알리는 문구",
        "로그 보관, 이상 탐지, 중단 기준과 담당자",
        "모델 변경과 재검증 주기",
        "개발자 대상 AI 사용 가이드 (사내 자료 준수)",
    ],
}


def build_gate_review():
    wb = Workbook()
    summary = wb.active
    summary.title = "요약"
    summary["A1"] = "게이트 검토표 — 요약"
    summary["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
    summary["A2"] = "각 게이트 시트에서 체크리스트를 확인한 뒤, 여기서 판정을 고릅니다. AI 기능이 없으면 G7을 '해당 없음'으로 판정하고 근거를 적습니다."
    summary["A2"].font = Font(name=FONT, size=9, italic=True, color="7F7F7F")
    headers = ["게이트", "판정", "체크 완료 수", "조건과 보완 사항", "판정 근거"]
    widths = [28, 16, 14, 40, 40]
    for col, (header, width) in enumerate(zip(headers, widths), start=1):
        cell = summary.cell(row=3, column=col, value=header)
        cell.font = Font(name=FONT, bold=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        summary.column_dimensions[get_column_letter(col)].width = width

    summary_first = 4
    for i, gate_name in enumerate(GATE_CHECKLISTS):
        row = summary_first + i
        sheet_name = gate_name.split(" ")[0]  # 예: G1
        summary.cell(row=row, column=1, value=gate_name)
        summary.cell(row=row, column=2, value="미검토")
        checklist_count = len(GATE_CHECKLISTS[gate_name])
        summary.cell(
            row=row,
            column=3,
            value=f'=COUNTIF(\'{sheet_name}\'!C4:C{3 + checklist_count},"예")&" / {checklist_count}"',
        )
        for col in range(1, 6):
            summary.cell(row=row, column=col).border = BORDER
            summary.cell(row=row, column=col).font = Font(name=FONT)
            summary.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    summary_last = summary_first + len(GATE_CHECKLISTS) - 1
    add_dropdown(summary, "B", summary_first, summary_last, JUDGMENT)
    summary.freeze_panes = "A4"

    for gate_name, items in GATE_CHECKLISTS.items():
        sheet_name = gate_name.split(" ")[0]
        ws = wb.create_sheet(sheet_name)
        ws["A1"] = f"{gate_name} 검토"
        ws["A1"].font = Font(name=FONT, size=14, bold=True, color="1F3864")
        ws["A2"] = "확인 여부는 '예 / 아니오 / 해당 없음' 중 고르고, 근거나 비고를 적습니다. 법령 조항은 최신 원문으로 확인하세요."
        ws["A2"].font = Font(name=FONT, size=9, italic=True, color="7F7F7F")
        for col, (header, width) in enumerate(zip(["No", "체크 항목", "확인 여부", "근거·비고"], [6, 60, 14, 40]), start=1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(name=FONT, bold=True)
            cell.fill = HEADER_FILL
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col)].width = width
        for i, item in enumerate(items, start=1):
            row = 3 + i
            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=item)
            for col in range(1, 5):
                cell = ws.cell(row=row, column=col)
                cell.border = BORDER
                cell.font = Font(name=FONT)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        add_dropdown(ws, "C", 4, 3 + len(items), YES_NO_NA)
        ws.freeze_panes = "A4"

    save(wb, GATE_DIR / "게이트_검토표.xlsx")


def main():
    build_requirements()
    build_process()
    build_test_plan()
    build_decisions()
    build_release_checklist()
    build_as_built()
    build_gate_review()


if __name__ == "__main__":
    main()
