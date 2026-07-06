# K-PATH 롱리스트 자동 PDF 생성 — 설정 가이드

구글 시트 업데이트 → 기업 프로필 PDF 자동 생성까지의 전체 파이프라인입니다.
직접 처음부터 끝까지 만들 필요 없이, 아래 순서대로 계정 설정만 하면 됩니다.
(코드/스킬은 이 폴더 안에 이미 다 준비되어 있습니다.)

```
구글시트 수정
   → [Apps Script] 메뉴에서 "PDF 생성 요청" 클릭 (또는 매일 자동 실행)
   → 시트를 .xlsx로 내보내서 GitHub의 input/latest.xlsx 로 커밋
   → Claude Code Routine의 웹훅을 호출해서 깨움
   → Routine이 이 저장소를 열어서 skills/marknco-profile-doc/SKILL.md 순서대로
     추출 → 표 정렬 보정 → 차트 생성 → 문서 생성 → 검증 → PDF 병합
   → 결과를 output/ 폴더에 커밋 + 푸시
```

---

## ⚠️ 지금 막혀있는 문제 → 이메일 확정 배포로 전환함

GitHub push가 몇 차례 시도 끝에 "성공"으로 표시된 적이 있었지만, 확인해보니 **존재하지도
않는 브랜치에 push했다고 스스로 잘못 보고한 사고**가 있었습니다. 즉 push 성공 여부를
신뢰할 수 없는 상태입니다. 그래서 이제 **이메일을 확정 배포 경로로 추가**했습니다
(`routine/ROUTINE_PROMPT.md` 참고) — GitHub push는 계속 시도하되, 성공/실패와 무관하게
완성된 PDF는 항상 `lsy@markncompany.co.kr` 로 이메일 발송됩니다.

**이걸 쓰려면 Routine에 Gmail 커넥터를 추가해야 해요**
1. `claude.ai/code/routines` → 이 루틴 편집
2. "커넥터" 섹션에서 **Gmail** 추가
3. 권한 요청 화면이 뜨면 **이메일 발송(Send)** 권한 승인
4. 저장 후 "Run Now"로 한 번 테스트 — `lsy@markncompany.co.kr` 로 메일이 오는지 확인

GitHub 저장소 쓰기 권한 문제(아래)는 시간 날 때 별도로 마저 해결하면 되지만, 당장 급하진
않습니다 — 이메일이 이미 확정 경로로 동작하니까요.



### (선택) GitHub push도 마저 고치고 싶다면

이메일이 확정 경로로 동작하니 급하진 않지만, 나중에 GitHub도 정상화하고 싶으면:

Routine이 처음 실행됐을 때 PDF는 정상적으로 다 만들었는데, **GitHub에 결과를 못 올리고
있어요** (`403 Resource not accessible by integration`). Routine에 연결된 GitHub App이
이 저장소에 **읽기 권한만** 가지고 있어서 그래요.

**고치는 법**
1. GitHub → 저장소 페이지 → **Settings**
2. 왼쪽 메뉴에서 **Integrations → GitHub Apps** (또는 조직 계정이면 조직 Settings →
   Third-party Access → GitHub Apps)
3. **Claude** (또는 Claude Code) 앱 찾아서 **Configure**
4. **Repository permissions → Contents** 를 **Read and write** 로 변경
5. 저장

또는:
1. `claude.ai/code/routines` 에서 이 루틴 열기
2. 저장소 연결 부분에서 **재연결(reconnect)** 하면서 권한 승인 화면이 다시 뜨면
   그때 "Contents: Read and write" 승인

권한 고친 다음엔 **push 후 원격 확인까지 되는지** 실제로 GitHub 저장소에 들어가서
파일이 진짜 올라왔는지 확인할 것 — "성공했다"는 로그만으로는 못 믿는다는 게 이미
확인된 상태입니다 (존재하지 않는 브랜치에 push 성공했다고 잘못 보고한 사례 있음).

---

## 스킬이 2개예요 — 어떤 데이터인지에 따라 자동으로 골라 씁니다

- `skills/marknco-profile-doc/` — **롱리스트**용 (혁신의숲 스크리닝 데이터, "전체 후보" 시트)
- `skills/marknco-applicant-profile/` — **K-PATH 접수시트**용 (지원기업 신청서, "정리 시트"
  + "innoclaw 데이터" 두 시트)

Routine 프롬프트가 `input/latest.xlsx`의 시트 이름을 보고 알아서 맞는 스킬을 골라 씁니다.
접수시트 쪽은 혁신의숲 데이터가 아직 안 붙은 신규 지원기업은 **1페이지만**, 붙은 기업은
**2페이지**(성장 데이터 포함)로 자동으로 갈립니다 — 정상 동작이에요.

---



- GitHub 계정 (조직 계정이면 더 좋음 — 담당자 개인 계정에 묶이지 않게)
- Claude Pro/Max/Team/Enterprise 플랜 + Claude Code 사용 가능 (Routines는 "Claude Code on the web" 활성화 필요)
- 이 자동화를 만질 구글 시트에 대한 편집 권한

---

## 1단계 — GitHub 저장소 만들기

1. GitHub에서 새 저장소 생성 (private 권장 — 기업 데이터가 들어가므로 공개 저장소 금지)
2. 이 폴더(`marknco-automation/`) 전체를 그 저장소에 push
   ```bash
   cd marknco-automation
   git init
   git add .
   git commit -m "init: K-PATH 자동화 초기 세팅"
   git branch -M main
   git remote add origin https://github.com/{owner}/{repo}.git
   git push -u origin main
   ```
3. **GitHub Personal Access Token 발급** (fine-grained 권장)
   - GitHub → Settings → Developer settings → Fine-grained tokens → Generate new token
   - Repository access: 방금 만든 저장소만 선택
   - Permissions: **Contents: Read and write** (이거 하나면 충분)
   - 발급된 토큰 문자열을 복사해둘 것 (다시 못 봄)

---

## 2단계 — Claude Code Routine 만들기

1. `claude.ai/code/routines` 접속 → **New routine**
2. **프롬프트**: `routine/ROUTINE_PROMPT.md` 파일 내용을 그대로 복사해서 붙여넣기
3. **저장소(Repositories)**: 1단계에서 만든 GitHub 저장소를 연결
   (처음 연결 시 GitHub App 설치 권한 요청이 뜨면 승인)
4. **트리거(Triggers)**: **API** 트리거 추가
   - 생성하면 웹훅 URL과 Bearer token이 발급됨 → **바로 복사해둘 것** (한 번만 보여줌)
5. 저장 후 **"Run Now"로 한 번 테스트 실행** 해볼 것 (input/latest.xlsx가 아직 없으면
   "입력 파일 없음"이라고 정상적으로 보고하고 끝나야 함 — 이게 정상 동작 확인)

---

## 3단계 — 구글 시트에 Apps Script 붙이기

1. 대상 구글 시트 열기 → **확장 프로그램 → Apps Script**
2. 기본 생성된 `Code.gs` 내용을 전부 지우고 `apps-script/Code.gs` 내용으로 교체
3. 저장 후 시트로 돌아와 새로고침 → 상단에 **"K-PATH 자동화"** 메뉴가 생김
4. 메뉴에서 **"최초 설정(1회만)"** 클릭 → 아래 4가지를 순서대로 입력
   - GitHub Personal Access Token (1단계에서 발급한 것)
   - GitHub 저장소 (`owner/repo-name` 형식)
   - Routine 웹훅 URL (2단계에서 발급된 것)
   - Routine Bearer Token (2단계에서 발급된 것)
5. 최초 실행 시 Apps Script 권한 승인 창이 뜨면 승인 (외부 시트 내보내기 + 네트워크 호출 권한 필요)

---

## 4단계 — 사용

- **수동**: 시트 업데이트 후 메뉴에서 **"K-PATH 자동화 → PDF 생성 요청"** 클릭
- **자동(매일)**: Apps Script 편집기에서 `installDailyTrigger` 함수를 한 번 실행 →
  매일 오전 9시 자동으로 최신 시트 기준 PDF 재생성
- 결과물은 GitHub 저장소의 `output/` 폴더에 `KPATH_롱리스트_날짜시각.pdf` 로 쌓임
- 진행 상황은 `claude.ai/code/routines` 에서 실행 로그로 실시간 확인 가능

---

## 참고 / 주의사항

- **셀 하나 바뀔 때마다 자동 실행(onEdit)은 넣지 않았습니다.** Routine 실행 횟수에
  일일 한도가 있어서(계정당 하루 실행 횟수 제한), 편집할 때마다 매번 돌리면 금방
  한도를 소진합니다. "완료 후 버튼 클릭" 또는 "하루 1회" 방식을 권장합니다.
- **디자인을 바꾸고 싶을 때는 이 자동화를 거치지 말고** 평소처럼 Claude 채팅에서
  `/marknco-profile-doc` 로 직접 요청하세요. Routine은 이미 확정된 디자인을
  그대로 반복 실행하는 용도이지, 새로운 디자인 결정을 내리는 용도가 아닙니다.
  (Routine 프롬프트에도 이 내용이 명시되어 있어 임의 변경은 시도하지 않습니다.)
- `config/config_confirmed.json` (기본 1페이지), `config/config_final_batch.json`
  (성장 데이터 포함 2페이지)는 지금까지 확정한 그대로 들어있습니다. 이후 채팅에서
  config를 다시 수정하면 이 폴더의 파일도 같이 갱신해서 GitHub에 반영해야
  Routine이 최신 디자인으로 돌아갑니다 — 채팅에서 "이것도 자동화 저장소에 반영해줘"
  라고 요청하면 그때그때 동기화해 드릴게요.
- 저장소는 반드시 **private**로 유지하세요 (기업 재무 데이터 포함).
