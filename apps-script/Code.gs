/**
 * K-PATH 롱리스트 자동화 — Apps Script
 * ------------------------------------------------------------
 * 이 파일을 구글 시트의 [확장 프로그램 > Apps Script] 에 붙여넣는다.
 *
 * 하는 일:
 *   1) 현재 시트를 .xlsx로 내보내기
 *   2) GitHub 저장소의 input/latest.xlsx 로 커밋(push)
 *   3) Claude Code Routine의 API 트리거(webhook)를 호출해서 PDF 생성 시작
 *
 * 최초 1회 설정 (직접 실행하지 말고 아래 setup() 함수를 한 번 실행):
 *   - GITHUB_TOKEN: GitHub Personal Access Token (repo 쓰기 권한, fine-grained 권장)
 *   - GITHUB_REPO:  "owner/repo-name" 형식
 *   - ROUTINE_URL:  Claude Code Routine 생성 시 발급되는 웹훅 URL
 *                   (예: https://api.anthropic.com/v1/claude_code/routines/trig_xxx/fire)
 *   - ROUTINE_TOKEN: 위 웹훅과 함께 발급되는 bearer token (한 번만 보여주므로 미리 복사해둘 것)
 *
 * 이 값들은 코드에 직접 쓰지 않고 [파일 > 프로젝트 속성 > 스크립트 속성] 에 저장한다
 * (아래 setup() 함수가 대화상자로 물어보고 자동 저장해줌).
 */

const SETTINGS_KEYS = ["GITHUB_TOKEN", "GITHUB_REPO", "ROUTINE_URL", "ROUTINE_TOKEN"];

/** 스프레드시트를 열면 커스텀 메뉴가 뜨도록 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("K-PATH 자동화")
    .addItem("PDF 생성 요청", "requestPdfGeneration")
    .addItem("최초 설정(1회만)", "setup")
    .addToUi();
}

/** 최초 1회: 필요한 값들을 물어보고 스크립트 속성에 저장 */
function setup() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const labels = {
    GITHUB_TOKEN: "GitHub Personal Access Token (repo 쓰기 권한)",
    GITHUB_REPO: "GitHub 저장소 (예: markncompany/kpath-automation)",
    ROUTINE_URL: "Claude Code Routine 웹훅 URL",
    ROUTINE_TOKEN: "Routine 웹훅 Bearer Token",
  };
  for (const key of SETTINGS_KEYS) {
    const resp = ui.prompt(labels[key], ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) {
      ui.alert("설정이 중단되었습니다.");
      return;
    }
    props.setProperty(key, resp.getResponseText().trim());
  }
  ui.alert("설정 완료. 이제 'PDF 생성 요청' 메뉴를 쓸 수 있습니다.");
}

/** 메인: 현재 시트를 내보내서 GitHub에 올리고 Routine을 깨운다 */
function requestPdfGeneration() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const missing = SETTINGS_KEYS.filter((k) => !props.getProperty(k));
  if (missing.length) {
    ui.alert("먼저 'K-PATH 자동화 > 최초 설정'을 실행해주세요. 누락: " + missing.join(", "));
    return;
  }

  ui.alert("PDF 생성을 요청합니다. 완료까지 몇 분 걸릴 수 있어요 (Claude Code 콘솔에서 진행 상황 확인 가능).");

  try {
    const xlsxBlob = exportSheetAsXlsx_();
    pushToGitHub_(xlsxBlob, props);
    triggerRoutine_(props);
    ui.alert("요청 완료! 결과 PDF는 GitHub 저장소의 output/ 폴더에 올라옵니다.");
  } catch (e) {
    ui.alert("오류 발생: " + e.message);
    throw e;
  }
}

/** 현재 스프레드시트 전체를 .xlsx Blob으로 내보내기 */
function exportSheetAsXlsx_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url =
    "https://docs.google.com/spreadsheets/export?id=" + ss.getId() + "&exportFormat=xlsx";
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });
  return resp.getBlob().setName("latest.xlsx");
}

/** GitHub Contents API로 input/latest.xlsx 를 생성/갱신 (커밋) */
function pushToGitHub_(blob, props) {
  const repo = props.getProperty("GITHUB_REPO");
  const token = props.getProperty("GITHUB_TOKEN");
  const path = "input/latest.xlsx";
  const apiUrl = "https://api.github.com/repos/" + repo + "/contents/" + path;

  // 기존 파일의 sha를 먼저 조회 (없으면 새로 생성, 있으면 갱신에 필요)
  let sha = null;
  const getResp = UrlFetchApp.fetch(apiUrl, {
    headers: { Authorization: "Bearer " + token, Accept: "application/vnd.github+json" },
    muteHttpExceptions: true,
  });
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  }

  const base64Content = Utilities.base64Encode(blob.getBytes());
  const payload = {
    message: "chore: 시트 데이터 갱신 (" + new Date().toISOString() + ")",
    content: base64Content,
  };
  if (sha) payload.sha = sha;

  const putResp = UrlFetchApp.fetch(apiUrl, {
    method: "put",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token, Accept: "application/vnd.github+json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  if (putResp.getResponseCode() >= 300) {
    throw new Error("GitHub 업로드 실패: " + putResp.getContentText());
  }
}

/** Claude Code Routine의 API 트리거(웹훅) 호출 */
function triggerRoutine_(props) {
  const url = props.getProperty("ROUTINE_URL");
  const token = props.getProperty("ROUTINE_TOKEN");
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      "anthropic-version": "2023-06-01",
      // Routine API가 아직 research preview라 베타 헤더가 필요할 수 있음.
      // 이 헤더 때문에 다른 에러(예: "beta header" 관련)가 나면 이 줄을 지우고 재시도할 것.
      "anthropic-beta": "experimental-cc-routine-2026-04-01",
    },
    payload: JSON.stringify({
      text: "새 데이터가 반영됐습니다. input/latest.xlsx 기준으로 PDF를 다시 생성해주세요.",
    }),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error("Routine 트리거 실패: " + resp.getContentText());
  }
}

/**
 * (선택) 매일 정해진 시각에 자동으로 실행하고 싶으면 아래 함수를 한 번만 실행해서
 * 시간 기반 트리거를 등록한다. 셀 하나 바뀔 때마다 매번 돌리는 onEdit 트리거는
 * Routine 실행 횟수 제한에 금방 걸리므로 권장하지 않는다 — 대신 "하루 한 번" 또는
 * 수동 버튼 클릭 방식을 권장.
 */
function installDailyTrigger() {
  ScriptApp.newTrigger("requestPdfGeneration")
    .timeBased()
    .everyDays(1)
    .atHour(9) // 매일 오전 9시
    .create();
  SpreadsheetApp.getUi().alert("매일 오전 9시 자동 실행이 등록되었습니다.");
}
