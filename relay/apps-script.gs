/**
 * 나라장터 AI 사업 모니터 — '갱신' 버튼 중계 (Google Apps Script 웹 앱)
 *
 * 화면(GitHub Pages) → 이 웹 앱 → GitHub Actions 수집 워크플로 실행(workflow_dispatch)
 *
 * 설치 (자세한 순서는 README "갱신 버튼 설정")
 *  1. script.google.com › 새 프로젝트 › 이 파일 내용을 그대로 붙여 넣기
 *  2. 프로젝트 설정(⚙) › 스크립트 속성 › GITHUB_TOKEN = fine-grained 토큰
 *     (이 저장소 하나에 Actions: Read and write 권한만)  ← 토큰은 여기에만. 코드에 쓰지 마세요
 *  3. 편집기에서 checkStatus 실행 → 권한 승인 (처음 한 번)
 *  4. 배포 › 새 배포 › 웹 앱 › 실행: 나 / 액세스: 모든 사용자 › 나온 /exec 주소를 config.json 의 refresh_relay_url 에
 *
 * - 쿨다운: 마지막 실행(예약 실행 포함) 후 COOLDOWN_MIN 분 이내면 실행하지 않음 — GitHub 실행 기록으로 판단, DB 없음
 * - 이미 대기·실행 중인 수집이 있으면 새로 실행하지 않음
 * - 응답: { status: "started" | "running" | "cooldown" | "ready" | "error", message, minutesLeft? }
 * - 참고: Apps Script 웹 앱은 응답 헤더를 정할 수 없어 CORS 출처 제한을 걸 수 없습니다.
 *   대신 30분 쿨다운과 토큰 권한(이 저장소 Actions만)이 남용을 막습니다.
 */
var REPO = "bono2610/g2b-monitor";
var WORKFLOW = "collect.yml";
var REF = "main";
var COOLDOWN_MIN = 30;
var ACTIVE = ["queued", "in_progress", "waiting", "requested", "pending"];

/** 화면의 갱신 버튼이 부름: 조건이 맞으면 수집 실행 */
function doPost(e) {
  return json_(run_(true));
}

/** 브라우저 주소창으로 열면 실행하지 않고 상태만 보여 줌 (확인용) */
function doGet(e) {
  return json_(run_(false));
}

/** 편집기에서 실행해 권한 승인·동작 확인용 (수집은 실행하지 않음) */
function checkStatus() {
  Logger.log(JSON.stringify(run_(false)));
}

function run_(dispatch) {
  var token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) return { status: "error", message: "스크립트 속성에 GITHUB_TOKEN 이 없어요" };
  // 두 사람이 동시에 눌러도 한 번만 실행되도록 잠금
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { status: "running", message: "수집 중" };
  try {
    var st = check_(token);
    if (!dispatch || st.status !== "ready") return st;
    var r = gh_(token, "/actions/workflows/" + encodeURIComponent(WORKFLOW) + "/dispatches", "post", { ref: REF });
    if (r.getResponseCode() === 204) {
      Utilities.sleep(3000);   // 다음 요청이 바로 '실행 중'을 보도록 GitHub에 실행이 잡힐 때까지 잠시 대기
      return { status: "started", message: "수집을 시작했어요" };
    }
    return { status: "error", message: "수집 실행 요청이 거절됐어요 (" + r.getResponseCode() + ")" };
  } catch (err) {
    return { status: "error", message: "GitHub에 연결하지 못했어요" };
  } finally {
    lock.releaseLock();
  }
}

// 최근 실행 기록으로 지금 상태 판단
function check_(token) {
  var r = gh_(token, "/actions/workflows/" + encodeURIComponent(WORKFLOW) + "/runs?per_page=10", "get");
  if (r.getResponseCode() !== 200) return { status: "error", message: "GitHub 실행 기록을 읽지 못했어요 (" + r.getResponseCode() + ")" };
  var runs = JSON.parse(r.getContentText()).workflow_runs || [];
  for (var i = 0; i < runs.length; i++) {
    if (ACTIVE.indexOf(runs[i].status) >= 0) return { status: "running", message: "수집 중" };
  }
  var last = runs[0];   // 최신순
  if (last) {
    var lastAt = new Date(last.run_started_at || last.created_at).getTime();
    var left = COOLDOWN_MIN * 60000 - (Date.now() - lastAt);
    if (left > 0) {
      var minutesLeft = Math.ceil(left / 60000);
      return { status: "cooldown", minutesLeft: minutesLeft, lastRunAt: new Date(lastAt).toISOString(), message: minutesLeft + "분 후 다시 가능" };
    }
  }
  return { status: "ready", lastRunAt: last ? (last.run_started_at || last.created_at) : null };
}

function gh_(token, path, method, body) {
  var opt = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: "Bearer " + token,
      "X-GitHub-Api-Version": "2022-11-28"
    }
  };
  if (body) { opt.contentType = "application/json"; opt.payload = JSON.stringify(body); }
  return UrlFetchApp.fetch("https://api.github.com/repos/" + REPO + path, opt);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
