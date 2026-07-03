---
name: marknco-applicant-profile
description: |
  마크앤컴퍼니 K-PATH "접수시트"(지원기업 신청서) 전용 프로필 PDF 자동 생성 스킬.
  "정리 시트"(지원서 자체 응답)와 "innoclaw 데이터"(혁신의숲 보강 데이터) 두 시트가
  함께 있는 워크북을 받으면 이 스킬을 쓴다. 형제 스킬 marknco-profile-doc(롱리스트용,
  "전체 후보" 단일 시트)과 헷갈리지 말 것 — 시트 구성으로 구분한다 (아래 "이 스킬을 언제
  쓰는가" 참고).
  다음과 같은 말이 나오면 이 스킬을 참고하세요: "접수시트로 프로필 만들어줘",
  "지원기업 PDF", "신청서 기반 프로필", "정리시트+innoclaw 합쳐서".
---

# K-PATH 지원기업 프로필 생성 스킬

## 이 스킬을 언제 쓰는가 (형제 스킬과 구분)

입력 엑셀/시트를 열어서 시트 이름을 먼저 확인한다.

| 시트 구성 | 사용할 스킬 |
|-----------|-------------|
| "정리 시트" + "innoclaw 데이터" (또는 유사한 이름의 2개 시트: 지원서 응답 + 혁신의숲 보강) | **이 스킬 (marknco-applicant-profile)** |
| "전체 후보" 단일 시트 (혁신의숲 보강 데이터만, 지원서 응답 없음) | `marknco-profile-doc` (롱리스트용) |

두 스킬은 스크립트 대부분(`make_profile_docs.js`, `balance_overview_groups.py`,
`generate_growth_charts.py`)을 공유하지만 **config.json 구성과 데이터 결합 방식이 다르다.**
헷갈려서 반대로 쓰면 필드명이 안 맞아 "모집분야"에 "지원 분야"를 억지로 끼워 맞추는 식의
어설픈 처리가 나온다 — 실제로 이런 사고가 있었으니 시트 이름을 먼저 꼭 확인할 것.

---

## Step 1: 데이터 병합 (`merge_applicant_data.py`)

이 스킬 전용 스크립트로, 두 시트를 기업명 기준으로 합친다.

```bash
python3 scripts/merge_applicant_data.py \
  --xlsx 접수시트.xlsx \
  --form-sheet "정리 시트" --form-name-col 4 --form-data-cols "3-31" \
  --enrich-sheet "innoclaw 데이터" --enrich-name-col 3 --enrich-data-cols "2-106" \
  --output companies.json
```

- **정리 시트**: A~C열(참고 메모/스크리닝/순번)에 스크리닝 상태가 있다. 기본 skip-keywords
  `"취소,불성실"`로 "신청 취소"/"불성실 지원" 건을 자동 제외한다. **"중복 지원"처럼 다른
  문구로 스크리닝된 건은 이 기본값으로 안 걸러지니, 실행 전에 스크리닝 컬럼 값을 실제로
  훑어보고 필요하면 `--form-skip-keywords`에 추가할 것.**
- **innoclaw 데이터**: 스크리닝 컬럼이 없다 — 매칭 안 되는 기업은 `merged`에 정리시트 정보만
  남고 `__has_innoclaw: false`로 표시된다 (Step 3 참고).
- 컬럼 번호(`--form-data-cols` 등)는 시트 구조가 바뀌면 같이 바뀐다. **매번 먼저 엑셀을 열어서
  헤더 행을 직접 확인**하고, 스크립트 상단 docstring의 컬럼 매핑과 실제가 일치하는지 볼 것.
- **중복 기업명**은 경고로만 뜨고 자동으로 처리되지 않는다 — 파일명이 기업명 기준이라 중복이면
  나중 파일이 앞 파일을 덮어쓴다. 배치 생성 전에 사람이 원본 시트에서 확인/정리해야 한다.

---

## Step 2: config.json — `config/config_applicant.json` 그대로 재사용

이미 확정된 구성이 있으니 처음부터 새로 짜지 말고 재사용한다. 핵심 구조:

```json
{
  "nameField": "기업명",
  "subtitleField": "제품, 서비스명",
  "overviewGroups": [
    { "rows": [ 지원분야, 기업연혁, 최근총고용인원, 2025매출액, 최종투자유치단계, 홈페이지, "관심 트랙"(span:2) ] },
    { "rows": [ 사업자번호, 총괄책임자, 총괄직급, 총괄연락처, 총괄이메일, 실무자명, 실무자연락처, 실무자이메일 ] }
  ],
  "sections": [ 제품서비스소개, 주요성과, 기대사항, 투자유치이력 ],
  "optionalTextSection": { 기타;제안아이디어 },
  "linkSection": { 사업계획서 },
  "dataTablesCondField": "__has_innoclaw",
  "dataTables": [ 고용현황(차트만), 재무현황(차트+표), 매출·손익현황(차트+표) ]
}
```

- 오버뷰 표가 8:8로 균형 잡혀 있다 (좌: 기업정보, 우: 담당자정보). 항목을 추가/제거하면
  균형이 깨지니 `span`으로 다시 맞추거나 Step 3의 보정 스크립트에 맡길 것.
- **`dataTablesCondField: "__has_innoclaw"`가 핵심이다.** 이게 있어야 innoclaw 매칭이
  안 된 지원기업(한창 접수 중이라 혁신의숲 보강이 아직 안 붙은 신규 지원자 등)이
  **2페이지 없이 1페이지로만** 생성된다. 이 필드를 빼먹으면 매칭 안 된 기업도 전부
  빈 성장데이터 2페이지를 달고 나온다 — 반드시 확인할 것.
- 관심 트랙(`관심 트랙 1~3`)과 기대사항(`기대사항 1~3`)은 원본이 3개 분리 필드라서, Step 1의
  병합 스크립트가 `__관심트랙_통합`(콤마 나열, 표용), `__기대사항_통합`(불릿 리스트, 섹션용)
  으로 미리 합쳐준다. config에서 원본 3필드가 아니라 이 통합 필드를 참조한다.
- 담당자 연락처(총괄/실무자 이름·전화·이메일)는 내부 검토용 문서에만 노출되는 정보다.
  이 프로필을 외부 공유할 가능성이 있으면 그 전에 반드시 확인할 것.

---

## Step 3: 좌우 표 정렬 보정 — `overviewGroups` 쓰므로 필수

marknco-profile-doc SKILL.md Step 4-1과 완전히 동일한 스크립트를 그대로 쓴다.

```bash
python3 scripts/balance_overview_groups.py \
  --companies companies.json --config config/config_applicant.json \
  --docs-dir out --script scripts/make_profile_docs.js
```

---

## Step 4: 성장 데이터 차트 (innoclaw 매칭된 기업만 해당)

```bash
python3 scripts/generate_growth_charts.py \
  --companies companies.json --out-dir charts \
  --years 2023,2024,2025 --months 2025-01:2026-05
```

`__has_innoclaw: false`인 기업도 이 스크립트 자체는 그냥 통과시켜도 무방하다 (차트 만들
데이터가 없으니 자동으로 스킵됨 — `generate_growth_charts.py` 자체 로직). 문제는 Step 2의
`dataTablesCondField` 설정이 없으면 표/헤딩까지 억지로 나온다는 것이니, config를 다시 확인.

---

## Step 5: 생성 + 검증

marknco-profile-doc과 동일한 방식(`make_profile_docs.js` → validate.py → soffice PDF 변환).
단, **페이지 수가 기업마다 다를 수 있다** (innoclaw 매칭 여부에 따라 1페이지 또는 2페이지).
"전부 N페이지여야 한다"는 식의 일괄 검증 대신, "매칭된 기업=2페이지, 안 된 기업=1페이지"로
나눠서 확인할 것.

```bash
python3 -c "
import json
data = json.load(open('companies.json'))
matched = sum(1 for c in data if c.get('__has_innoclaw'))
print(f'innoclaw 매칭: {matched} / {len(data)}')
"
```

---

## 참고: 원본 스킬(marknco-profile-doc)과 공유하는 부분

- 폰트(Pretendard), 색상, 타이포그래피, 표/섹션 컴포넌트 디자인 전부 동일
  (`../marknco-profile-doc/references/design-system.md` 참고)
- `make_profile_docs.js`/`balance_overview_groups.py`/`generate_growth_charts.py`는
  두 스킬이 같은 코드를 복사해서 쓴다. **한쪽을 고치면 다른 쪽도 동기화할 것**
  (예: `isEmpty()`에 숫자 0을 빈 값으로 처리하는 수정, `dataTablesCondField` 기능 추가 등은
  두 스킬 모두에 반영되어 있어야 한다 — 잊어버리기 쉬우니 수정할 때마다 두 폴더 다 확인).
