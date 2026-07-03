---
name: marknco-profile-doc
description: |
  마크앤컴퍼니 전용 기업 프로필 문서(Word + PDF) 자동 생성 스킬. 1페이지 기본 프로필은 물론
  고용·재무·매출 등 "성장 데이터"를 차트+표로 곁들인 2페이지 확장형까지 지원합니다.
  엑셀 접수 시트나 스크리닝 롱리스트를 받아 기업별 프로필 Word/PDF를 만들 때 반드시 이 스킬을 사용하세요.
  다음과 같은 말이 나오면 즉시 이 스킬을 참고하세요:
  "기업 프로필 만들어줘", "접수 시트로 PDF 만들어줘", "기업별 문서 생성",
  "1페이지 프로필", "K-PATH 문서", "스타트업 프로필 Word", "기업 소개 자료 만들어",
  "엑셀 → PDF", "Word 파일로 정리해줘", "회사별 파일로", "롱리스트 PDF로",
  "PDF 파일 하나로 합쳐줘", "고용 현황 차트", "재무 현황 그래프", "매출 손익 추이",
  "성장 데이터 넣어줘", "기업별 차트 자동으로".
  프로그램명, 목차, 컬럼 구성이 달라져도 이 스킬로 동일한 디자인이 유지됩니다.
---

# 마크앤컴퍼니 기업 프로필 생성 스킬

엑셀 접수 시트 → 기업별 1페이지(또는 성장 데이터 포함 2페이지) Word + PDF 자동 생성.
Pretendard 폰트, 딥 네이비 디자인 시스템 고정 적용.

디자인 스펙 전체: `references/design-system.md` 참고.

---

## 사전 확인 (스킬 시작 전)

1. **Pretendard OTF 파일** — 사용자가 업로드하거나 `~/.fonts/` 에 이미 설치되어 있는지 확인
   ```bash
   fc-list | grep -i pretendard
   ```
   없으면 사용자에게 OTF 파일 업로드 요청 → 설치:
   ```bash
   mkdir -p ~/.fonts && cp /mnt/user-data/uploads/PRETENDARD*.OTF ~/.fonts/ && fc-cache -f ~/.fonts
   ```
   npm 패키지로도 구할 수 있음(스타일별 OTF 포함): `npm install pretendard` → `node_modules/pretendard/dist/public/static/`
2. **docx npm 패키지** 확인
   ```bash
   npm list -g docx 2>/dev/null | grep docx || npm install -g docx
   ```
3. **LibreOffice** 확인
   ```bash
   which soffice || echo "없음"
   ```
4. **pdfplumber** 확인 (`overviewGroups` 사용 시 Step 4-1 보정에 필요)
   ```bash
   python3 -c "import pdfplumber" 2>/dev/null || pip install pdfplumber --break-system-packages -q
   ```
5. **qpdf** 확인 (여러 파일을 PDF 1개로 합칠 때 필요)
   ```bash
   which qpdf || echo "없음"
   ```
6. **matplotlib** 확인 (고용/재무/매출 등 "성장 데이터" 차트를 넣을 때만 필요)
   ```bash
   python3 -c "import matplotlib" 2>/dev/null || pip install matplotlib --break-system-packages -q
   ```

---

## Step 1: 엑셀 구조 파악

업로드된 엑셀을 읽어 시트 목록과 헤더를 파악한다.

```bash
python3 << 'PYEOF'
import openpyxl, json
wb = openpyxl.load_workbook('/mnt/user-data/uploads/[파일명].xlsx', data_only=True)
print("시트 목록:", wb.sheetnames)
for sname in wb.sheetnames:
    ws = wb[sname]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column+1)]
    count = sum(1 for r in range(2, ws.max_row+1) if ws.cell(row=r, column=1).value or ws.cell(row=r, column=2).value or ws.cell(row=r, column=3).value or ws.cell(row=r, column=4).value)
    print(f"  [{sname}] 헤더: {headers[:15]} ... 데이터행: {count}개")
PYEOF
```

사용자에게 다음을 확인:
- 메인 시트명 (기업 목록)
- 기업명 컬럼
- 제목/부제목 컬럼
- 오버뷰 표에 들어갈 컬럼들
- 본문 섹션 컬럼들
- 외부 데이터 시트 여부

**참고 — 소스 파일은 두 가지 패턴이 있을 수 있다:**
1. *지원서 접수 시트* (K-PATH 지원기업 프로필 등): 기업이 직접 작성한 응답 컬럼(제품 소개, 관심 트랙 등) + 필요시 별도 시트의 외부 데이터(혁신의숲 등)를 `--ext-sheet`로 연결.
2. *스크리닝 롱리스트* (혁신의숲 데이터 기반 등): 시트 하나에 이미 모든 정보(AI 요약, 투자 정보, 링크 등)가 컬럼으로 다 들어있는 경우. 이땐 `--ext-sheet` 없이 `--data-cols`만으로 전체를 한 번에 추출하면 된다 — 패널(외부 데이터 상자) 대신 `sections`로 통일해서 표시하는 것을 권장 (Step 3 참고).

---

## Step 2: 데이터 추출 → JSON

`scripts/extract_data.py` 실행. 반드시 실행 전에 아래 파라미터를 채운다:

```bash
python3 /path/to/skill/scripts/extract_data.py \
  --xlsx "/mnt/user-data/uploads/파일명.xlsx" \
  --sheet "정리 시트" \
  --name-col 4 \
  --data-cols "4-22" \
  --output /home/claude/work/companies.json
```

외부 데이터 시트가 있으면 추가:
```bash
python3 /path/to/skill/scripts/extract_data.py \
  --xlsx "/mnt/user-data/uploads/파일명.xlsx" \
  --sheet "정리 시트" \
  --name-col 4 \
  --data-cols "4-22" \
  --ext-sheet "innoclaw 데이터" \
  --ext-name-col 3 \
  --ext-data-cols "3-17" \
  --output /home/claude/work/companies.json \
  --ext-output /home/claude/work/ext_data.json
```

---

## Step 3: config.json 작성

사용자의 요구사항에 맞게 `/home/claude/work/config.json` 을 작성한다.

```json
{
  "programName": "2026 KT K-PATH",
  "header": { "enabled": true },
  "footer": { "enabled": true },
  "headerLeft": "Mark & Company · K-PATH",
  "footerText": "본 문서는 KT K-PATH 접수 자료를 기반으로 작성되었습니다.",
  "outputPrefix": "K-PATH_지원기업프로필",
  "fileNameField": "기업명",
  "nameField": "기업명",
  "subtitleField": "제품, 서비스명",

  "overviewRows": [
    { "label": "지원 분야",       "field": "지원 분야",           "type": "plain" },
    { "label": "기업 연혁",       "field": "기업 연혁",           "type": "plain" },
    { "label": "최근 총 고용인원", "field": "최근 총 고용인원",    "type": "plain" },
    { "label": "2025 매출액",     "field": "2025 매출액(원)",     "type": "money" },
    { "label": "최종 투자유치 단계","field": "최종 투자유치 단계", "type": "plain" },
    { "label": "홈페이지",        "field": "홈페이지 주소",       "type": "link"  }
  ],

  "sections": [
    { "heading": "제품·서비스 소개", "field": "제품 서비스 소개", "emptyText": "-" },
    { "heading": "주요 성과",       "field": "주요 성과", "emptyText": "-" },
    { "heading": "투자유치 이력",   "field": "투자유치 이력", "emptyText": "-" }
  ],

  "twoColumnSection": {
    "heading": "관심 트랙 / 기대사항",
    "leftLabel": "관심 트랙",
    "leftFields": ["관심 트랙 1", "관심 트랙 2", "관심 트랙 3"],
    "rightLabel": "기대사항",
    "rightFields": ["기대사항 1", "기대사항 2", "기대사항 3"]
  },

  "optionalTextSection": {
    "heading": "기타; 제안 아이디어",
    "field": "기타; 제안 아이디어"
  },

  "linkSection": {
    "heading": "사업계획서",
    "field": "사업계획서",
    "linkLabel": "사업계획서 파일 열기 ↗"
  },

  "panel": { "enabled": false }
}
```

**panel이 없으면**: `"panel": { "enabled": false }` 로 설정. (참고: 패널은 이탤릭체·베이지 배경의 구버전 표시 방식. 다른 섹션과 톤을 통일하고 싶다면 패널 대신 `sections`/`linkSection`으로 같은 정보를 표시하는 것을 권장 — 아래 3-1 참고)
**twoColumnSection이 없으면**: 키 자체를 제거.
**overviewRow type 종류**: `plain` / `link` / `money`
**빈 값 기본 표시**: `"-"` (모든 fallback 공통. 문구를 바꾸고 싶으면 각 섹션에 `emptyText` 지정)
**header/footer 끄기**: `"header": {"enabled": false}`, `"footer": {"enabled": false}`

### 3-1. 오버뷰 표가 길 때 → 2단 나란히 배치 (`overviewGroups`)

항목이 10개를 넘어가면 표 하나로는 세로로 너무 길어진다. `overviewRows` 대신 `overviewGroups`를 쓰면 좌/우 2개(이상) 표로 나란히 배치되어 공간을 절약한다.

```json
"overviewGroups": [
  { "rows": [ { "label": "...", "field": "...", "type": "plain" }, ... ] },
  { "rows": [ { "label": "...", "field": "...", "type": "money" }, ... ] }
]
```

**주의할 점 — 좌우 표 높이가 자동으로는 안 맞는다.**
- 필드 값 길이가 회사마다 다르면(예: 성과키워드가 콤마로 여러 개 나열된 경우, 긴 URL 등) 특정 회사만 그 셀이 줄바꿈되면서 한쪽 표가 더 길어진다.
- 글자 수로 줄바꿈 여부를 미리 예측하는 방법(한글/영문 폭 계산 등)은 근사치일 뿐이라 **절대 신뢰하면 안 된다** — 실제로 렌더링해서 확인해야 한다.
- **반드시 `scripts/balance_overview_groups.py` 로 전수 실측·보정할 것** (아래 Step 4-1).
- 행 하나를 의도적으로 더 크게 표시하고 싶으면(예: 표 전체가 항목 개수 차이로 구조적으로 1줄 차이나는 경우) 해당 row에 `"span": 2` 를 추가하면 그 행만 2배 높이가 된다.
- 값 길이 편차가 아주 큰 필드(성과키워드처럼 콤마로 여러 개 나열되는 태그성 필드 등)는 애초에 좌우 표 안에 넣지 말고 `wideOverviewRows` 로 표 아래 전체 폭 한 줄로 따로 빼는 것도 방법이다:
  ```json
  "wideOverviewRows": [
    { "label": "성과키워드", "field": "성과키워드", "type": "plain" }
  ]
  ```
  단, **어떤 회사는 표 안에 있고 어떤 회사는 표 밖에 있는 식으로 회사마다 다른 구성을 쓰지 말 것.** 표 구성(어느 필드가 어디 있는지)은 전체 기업이 항상 동일해야 한다 — 이건 통일성 요구사항이니 예외를 두지 않는다. `wideOverviewRows`를 쓰기로 했으면 전체 기업에 일괄 적용한다.

---

## Step 4: 문서 생성

```bash
node /path/to/skill/scripts/make_profile_docs.js \
  --companies /home/claude/work/companies.json \
  --config /home/claude/work/config.json \
  --out /home/claude/work/out \
  [--ext /home/claude/work/ext_data.json]   # 외부 데이터 있을 때만
```

---

## Step 4-1: 좌우 표 높이 보정 (`overviewGroups` 사용 시 필수)

config에 `overviewGroups`(2단 나란히 배치)를 썼다면 **반드시** 아래 스크립트로 전수 검증·보정한다. 사람 눈으로 몇 개만 훑어보고 넘어가면 안 된다 — 값 길이는 회사마다 다르므로 일부만 확인해서는 안전하지 않다는 것이 실제로 확인된 사실이다.

```bash
python3 /path/to/skill/scripts/balance_overview_groups.py \
  --companies /home/claude/work/companies.json \
  --config /home/claude/work/config.json \
  --docs-dir /home/claude/work/out \
  --script /path/to/skill/scripts/make_profile_docs.js
```

- 내부적으로 생성 → PDF 변환 → 표 테두리 실측(pdfplumber) → 회사별 보정값 주입 → 재생성을 2회(pass) 반복한다.
- 보정값은 companies.json에 `__overviewFillerTwips` 필드로 저장되며, 스크립트(make_profile_docs.js)가 이 값을 보고 **부족한 쪽 그룹의 모든 행에 균등하게** 높이를 나눠 늘린다 (특정 행 하나만 커지거나, 빈 행이 추가되는 방식이 아님 — 이 두 방식 모두 실제로 시도했다가 "보기 안 좋다"는 피드백을 받고 폐기된 방식이니 되돌리지 말 것).
- `overviewRows`(단일 표)만 쓰는 경우는 이 단계가 필요 없다.

---

## Step 4-2: 성장 데이터 섹션 (고용·재무·매출 현황 — 차트 + 표, 선택적)

"고용 현황", "재무 현황", "매출·손익 현황" 처럼 연도별/월별 추이를 차트+표로 보여달라는 요청이면 이 패턴을 쓴다. 1페이지 기본 프로필 **뒤에 이어서 2페이지째**에 자동으로 배치된다(`cfg.dataTables`가 있으면 스크립트가 페이지브레이크를 자동으로 넣음).

### 4-2-1. 차트 생성 (`generate_growth_charts.py`)

```bash
python3 /path/to/skill/scripts/generate_growth_charts.py \
  --companies /home/claude/work/companies.json \
  --out-dir /home/claude/work/charts \
  --years 2023,2024,2025 \
  --months 2025-01:2026-05
```

- 기업마다 PNG 3종(고용 시계열, 재무 그룹막대, 매출·손익 그룹막대)을 그려서 companies.json에 `__chart_employment` / `__chart_financial` / `__chart_revenue` 필드(파일 경로)로 저장한다.
- **반드시 매 회사마다 다른 이미지가 나오도록 `chartField` 방식을 쓸 것.** 특정 회사 하나로 그린 정적 PNG를 config에 `chart.path`로 고정해서 여러 회사에 그대로 쓰면 안 된다 — 실제로 이 실수를 했다가 전체 배치에서 엉뚱한 회사 차트가 뜨는 사고가 있었다.
- 연도/월 컬럼명이 다르면 스크립트 상단 docstring의 "전제조건" 부분을 참고해서 복사본을 만들어 컬럼명을 맞출 것.
- 데이터가 2개년/2개월 미만인 기업은 해당 차트가 자동으로 스킵되고(`None`), 표만 표시된다 — 정상 동작이며 에러 아님.
- 고용 시계열은 개월 수가 많으면(17개월 등) 표로 그대로 넣기엔 컬럼이 너무 많아진다. 그래서 표 대신 **요약 통계**(`__emp_latest`, `__emp_first`, `__emp_hires_sum`, `__emp_leaves_sum`, `__emp_net_change`)를 함께 계산해 저장한다 — 필요하면 이 필드들로 작은 요약 표를 만들어 차트 아래에 붙인다(config 예시는 4-2-2 참고).

### 4-2-2. config.json — `dataTables`

```json
"dataTables": [
  {
    "heading": "고용 현황",
    "chartField": "__chart_employment",
    "chartAspect": 4.75,
    "chartWidthPx": 640,
    "columns": ["최근월 근무인원", "누적 입사자", "누적 퇴사자", "순증감"],
    "rows": [
      { "label": "17개월 누적(25.01~26.05)",
        "fields": ["__emp_latest", "__emp_hires_sum", "__emp_leaves_sum", "__emp_net_change"],
        "types": ["number","number","number","number"], "unit": "명" }
    ]
  },
  {
    "heading": "재무 현황 (자산·부채·자본)",
    "chartField": "__chart_financial",
    "chartAspect": 4.34,
    "chartWidthPx": 620,
    "columns": ["2023", "2024", "2025", "3개년 성장률"],
    "rows": [
      { "label": "자산", "fields": ["2023년 자산","2024년 자산","2025년 자산","최근 3개년 재무 성장률"], "types": ["money","money","money","percent"] },
      { "label": "부채", "fields": ["2023년 부채","2024년 부채","2025년 부채","부채성장률_3개년"], "types": ["money","money","money","percent"] },
      { "label": "자본", "fields": ["2023년 자본","2024년 자본","2025년 자본","자본성장률_3개년"], "types": ["money","money","money","percent"] }
    ]
  },
  {
    "heading": "매출·손익 현황",
    "chartField": "__chart_revenue",
    "chartAspect": 4.34,
    "chartWidthPx": 620,
    "columns": ["2023", "2024", "2025", "3개년 성장률"],
    "rows": [
      { "label": "매출액", "fields": ["2023년 매출(원)","2024년 매출(원)","2025년 매출(원)","3개년 매출(원) 성장률"], "types": ["money","money","money","percent"] },
      { "label": "영업이익", "fields": ["2023 영업이익","2024 영업이익","2025 영업이익","3개년 영업이익 성장률"], "types": ["money","money","money","percent"] },
      { "label": "순이익", "fields": ["2023 순이익","2024 순이익","2025 순이익","3개년 순이익 성장률"], "types": ["money","money","money","percent"] }
    ]
  }
]
```

- `chartAspect`는 실제 생성된 PNG의 `width/height` 비율. 차트를 재생성했으면 반드시 다시 재서 갱신할 것(높이를 바꿨는데 비율을 안 갱신하면 이미지가 찌그러지거나 여백이 어긋난다).
- `heading` 대신 `chart`만 있고 `rows`가 없으면(=`rows` 생략) **표 없이 차트만** 렌더링된다 (고용현황처럼 raw 데이터가 너무 많아 표로 못 넣을 때).
- `type: "money"` 필드는 원(₩) 단위 raw 숫자를 넣으면 자동으로 억/백만/만원 단위를 골라서 표시한다(4-2-3 참고). `type: "percent"`는 소수(0.138 → "+13.8%")를 화살표+색상(양수=초록▲, 음수=빨강▼)으로 표시한다.
- 페이지 분리 기본값은 "2페이지째에 표시"(`cfg.dataTablesInline` 미설정 시). 남는 공간이 있으면 1페이지 흐름에 바로 이어붙이고 싶을 때만 `"dataTablesInline": true` + 각 테이블에 `"compact": true`를 추가 — 단, 압축 모드에서도 표가 안 들어가면 억지로 우기지 말고 2페이지로 되돌릴 것(실제로 한 페이지에 우겨넣으려다 성장률 정보가 손실되는 문제가 있었음).

### 4-2-3. 금액 단위 자동 처리 (`type: "money"`)

`fmtEok()`가 원 단위 raw 값을 받아 규모에 맞게 자동으로 단위를 고른다: **억원 → 백만원 → 만원 → 원**. 억 단위로만 강제 표시하면 소액 기업은 전부 "0.0억"으로 뭉개져서 실제 규모가 안 보이는 문제가 있었기 때문. 정확히 0원인 값만 "-"로 표시하고, 0.05억처럼 반올림하면 0으로 보이는 **근사값은 그대로(예: "1.4백만원") 표시**한다 — "0억대는 전부 -로" 처리하지 말 것. 이 구분을 헷갈려서 실제로 유효한 성장률 데이터를 지워버린 적이 있다.

### 4-2-4. 성장률(퍼센트) 계산 시 부호 규칙 — 기준값이 음수일 때

3개년 성장률처럼 `(v1-v0)/v0` 형태로 직접 계산해야 하는 필드(예: 부채·자본의 3개년 성장률처럼 원본 엑셀에 없어서 스킬이 계산해야 하는 경우)는 **분모에 `v0`가 아니라 `abs(v0)`를 쓸 것**:
```python
growth = (v1 - v0) / abs(v0) if v0 != 0 else None
```
기준값이 음수(예: 적자, 자본잠식)일 때 그냥 `v0`로 나누면 부호가 뒤집혀서 "적자가 더 커졌는데 성장률이 +270%"처럼 직관과 반대로 나온다. 실제로 KT 프로젝트에서 사용자가 엑셀에 미리 계산해 준 공식 값과 대조해서 이 버그를 발견했다 — 엑셀에 3개년 성장률이 이미 계산되어 있으면 **그 값을 그대로 쓰고 재계산하지 말 것**(같은 이유로 부호가 다를 수 있음).

### 4-2-5. 차트 자체 만들 때 주의점 (matplotlib, `generate_growth_charts.py` 참고)

- **범례는 그래프 바깥 위쪽에.** `loc="upper right"` 등으로 plot 안에 두면 값이 큰 기업에서 막대와 겹쳐 깨진다. `bbox_to_anchor=(0.5, 1.02), loc="lower center"` 로 바깥에 둘 것.
- **음수 막대의 라벨은 y축 여유(`set_ylim`)를 명시적으로 잡아줄 것.** 자동 스케일만 믿으면 라벨이 잘리거나 축 경계에서 깨진다.
- **`ax.annotate()` 호출 시 `va=` 파라미터를 절대 빠뜨리지 말 것.** 빠지면 라벨이 막대 중앙에 겹쳐서 렌더링된다(실제로 이 버그로 "라벨이 깨져 보인다"는 피드백을 받은 적 있음).
- 이중축(secondary axis) 차트에서 우측 축 눈금·라벨 색상은 데이터 계열과 같은 색(파란색 등)으로 맞추고 싶은 유혹이 있는데, 실제로는 **중립 회색(SUBTLE)이 더 깔끔하다** — 선/막대 색상만 데이터 구분용으로 쓰고 축 텍스트는 통일할 것.
- 차트 세로 크기를 키우면 2페이지 안에 안 들어갈 수 있다. `figsize`의 높이(height)를 조정한 뒤엔 반드시 실제 페이지 수를 재검증할 것 (Step 5).

---

## Step 5: 검증 + PDF 변환

```bash
cd /home/claude/work/out

# 1. 각 docx 검증
for f in *.docx; do
  python3 /mnt/skills/public/docx/scripts/office/validate.py "$f" 2>&1 | tail -2
done

# 2. PDF 변환
for f in *.docx; do
  python3 /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf "$f" > /dev/null 2>&1
done

# 3. 페이지 수 확인 (전부 1페이지여야 함)
for f in *.pdf; do
  pages=$(pdfinfo "$f" 2>/dev/null | grep "^Pages:" | awk '{print $2}')
  echo "$f: $pages pages"
done
```

2페이지가 나오는 파일이 있으면 → config.json 에서 해당 섹션 텍스트가 너무 길거나 섹션이 너무 많은 것. 섹션을 줄이거나 spacing 값을 줄여서 재생성.

**여러 회사를 한 PDF로 합쳐야 할 때** (예: "PDF 파일 하나에 다 넣어줘"):
```bash
python3 - << 'PYEOF'
import json, subprocess
data = json.load(open("/home/claude/work/companies.json"))
prefix = "K-PATH_지원기업프로필"  # config의 outputPrefix
files = [f"/home/claude/work/out/{prefix}_{c['기업명']}.pdf" for c in data]
subprocess.run(["qpdf", "--empty", "--pages", *files, "--", "/home/claude/work/합본.pdf"])
PYEOF
```
- 엑셀에 있는 순서(연번) 그대로 병합됨 — 별도 정렬 불필요.
- 파일명에 공백이 들어갈 수 있으므로 셸 문자열 결합(`tr`, `join`)이 아니라 **Python list를 subprocess에 직접 전달**할 것 (공백 포함 파일명이 깨지는 원인이 됨).

---

## Step 6: 파일 전달

기업별 개별 파일이 필요하면:
```bash
mkdir -p /mnt/user-data/outputs
cp /home/claude/work/out/*.docx /home/claude/work/out/*.pdf /mnt/user-data/outputs/
```

"PDF 하나로 합쳐줘" 같은 요청이면 Step 5의 병합 결과 파일 하나만 전달한다 (개별 파일 63개를 다 올릴 필요 없음):
```bash
cp /home/claude/work/합본.pdf /mnt/user-data/outputs/
```

`present_files` 로 전달 (여러 개면 PDF → docx 순서).

---

## 자주 쓰는 조정 포인트

| 증상 | 조정 |
|------|------|
| 2페이지로 넘침 | `sectionHeading` 의 `before` 값 줄이기 (300→240), `bodyPara` 마지막 줄 `after` 줄이기 (130→100) |
| 표 셀이 너무 좁음 | `overviewRows` 라벨 컬럼 `2400` → `2800` 으로 변경 |
| 오버뷰 항목이 10개+ | `overviewRows` → `overviewGroups` 로 전환해서 좌우 2단 배치 (Step 3-1) |
| 좌우 표 높이가 안 맞음 | **절대 눈대중/글자수로 판단하지 말고** `scripts/balance_overview_groups.py` 로 실측 보정 (Step 4-1). 빈 행 추가나 마지막 행만 늘리는 방식은 채택하지 말 것 — 전체 행에 균등 분배가 확정된 방식 |
| 외부 패널 항목 추가 | `panel.fields` 에 키-값 추가 (JS 스크립트도 같이 수정) — 단, 새 프로젝트는 패널보다 `sections`로 통일하는 것을 권장 |
| 섹션 순서 바꾸기 | config.json `sections` 배열 순서 변경 |
| 프로그램명 변경 | `headerLeft`, `footerText`, `outputPrefix` 수정 |
| 헤더/푸터 아예 없애기 | `"header": {"enabled": false}`, `"footer": {"enabled": false}` |
| 빈 값 표시 문구 변경 | 기본은 `"-"`. 섹션별로 다르게 하려면 해당 섹션에 `emptyText` 지정 |
| 여러 회사를 PDF 1개로 합치기 | Step 5 "여러 회사를 한 PDF로" 참고 (`qpdf`, Python list로 인자 전달) |
| 고용/재무/매출 차트+표 추가 | Step 4-2 (`generate_growth_charts.py` + `dataTables`) |
| 차트가 특정 회사에서만 이상하게 나옴 | `chart.path`(정적 경로)를 쓰고 있는지 확인 — 반드시 `chartField`(회사별 동적 경로)로 바꿀 것 |
| 차트 범례가 막대와 겹침 | 범례를 `bbox_to_anchor`로 그래프 바깥 위로 이동 (4-2-5) |
| 음수 막대 라벨이 깨지거나 잘림 | `set_ylim` 여유 확보 + `ax.annotate()`에 `va=` 파라미터 누락 여부 확인 (4-2-5) |
| 성장률 부호가 이상함(기준값이 음수일 때) | 분모를 `v0`가 아니라 `abs(v0)`로 계산했는지 확인, 엑셀에 공식값 있으면 그걸 우선 사용 (4-2-4) |
| 금액이 전부 "0.0억"으로 뭉개짐 | `type: "money"`가 자동으로 억/백만/만원 단위를 고르는지 확인 — 강제로 억 단위 포맷 쓰지 말 것 (4-2-3) |
