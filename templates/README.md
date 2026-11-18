# 산출물 실물 템플릿

`docs/`의 메모 형식(md) 산출물을 강의와 실무에서 쓸 수 있는 Word, Excel, PowerPoint 파일로 옮긴 것입니다. 내용 구조는 `docs/`와 같습니다.

## 파일 연결표

| 산출물 (docs) | 실물 템플릿 | 형식 | 이유 |
|---|---|---|---|
| 00 과제 정의서 | `docx/00_과제정의서.docx` | Word | 문서형 서술과 표 |
| 01 고객 리서치 | `docx/01_고객_리서치.docx` | Word | 서술과 표 |
| 02 제품 요구사항 문서 | `docx/02_PRD.docx` | Word | 긴 서술 문서 |
| 03 요구사항 정의서 | `xlsx/03_요구사항정의서.xlsx` | Excel | ID 기준 목록 관리 |
| 04 기능 명세서 | `docx/04_기능명세서.docx` | Word | 기능별 서술과 예외 흐름 |
| 05 화면 설계서 | `pptx/05_화면설계서.pptx` | PowerPoint | 화면 이미지와 설명을 함께 보여주기 |
| 06 AI 기능 설계서 | `docx/06_AI기능설계서.docx` | Word | 서술과 평가 기준 |
| 07 업무 프로세스 | `xlsx/07_업무프로세스.xlsx` | Excel | AS-IS, TO-BE, 역할 분담표 표 |
| 08 검증 계획 | `xlsx/08_검증계획.xlsx` | Excel | 시나리오와 결함 추적 |
| 09 의사결정 기록 | `xlsx/09_의사결정기록.xlsx` | Excel | 예측과 리뷰일 비교 |
| 10 출시 점검표 | `xlsx/10_출시점검표.xlsx` | Excel | 판정 드롭다운과 자동 집계 |
| 11 실제 구현 기록 | `xlsx/11_As-built.xlsx` | Excel | 차이 기록 |
| 검토 관문 G1~G7 | `gates/게이트_검토표.xlsx` | Excel | 요약 시트와 검토 관문별 체크리스트 |

## 사용 원칙

- 실물 파일은 사람이 쓰기 위한 형식입니다. AI 작업 도구가 작업하는 기준은 여전히 `docs/`의 메모 형식(md)입니다.
- 실물 파일을 고쳤다면 내용을 `docs/`에 옮기거나, 반대로 `docs/`를 고친 뒤 실물 파일에 반영하세요. 두 곳이 어긋나면 `docs/`를 우선합니다.
- 예시 행(회색 기울임)은 참고용입니다. 실제 작성 시 지우세요.
- 판정 칸은 드롭다운으로 고릅니다. 집계 수식은 Excel이 열릴 때 계산됩니다.

## 다시 만들기

템플릿은 만들기 프로그램으로 만듭니다. 수정이 필요하면 만들기 프로그램을 고친 뒤 다시 실행하세요.

```bash
# Node.js: docx, pptxgenjs 필요
node tools/build-templates/build_docx.js
node tools/build-templates/build_pptx.js

# Python: openpyxl 필요
python tools/build-templates/build_xlsx.py
```

## 확인되지 않은 항목

- 생성 후 Excel 수식 재계산(LibreOffice)은 이 환경에서 실행하지 못했습니다. 공식은 단순 `COUNTIF`, `ROWS`, `IF`이므로 Excel에서 열어 값이 나오는지 한 번 확인하세요.
- PowerPoint와 Word에서 레이아웃을 실제로 열어 확인하지 않았습니다. 수업 전에 한 번 열어 보세요.
- 글꼴은 Windows 기본인 맑은 고딕(Malgun Gothic)을 사용했습니다.
