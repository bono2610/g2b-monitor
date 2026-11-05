/*
 * 나라장터 AI 사업 모니터 — '갱신' 버튼 중계 함수 (Vercel 서버리스)
 *
 * 화면(GitHub Pages) → 이 함수 → GitHub Actions 수집 워크플로 실행(workflow_dispatch)
 *
 * - GitHub 토큰은 Vercel 환경변수 GITHUB_TOKEN 에만 둡니다. 응답·로그에 절대 싣지 않습니다.
 *   (이 저장소 하나에 Actions: Read and write 권한만 있는 fine-grained 토큰)
 * - 쿨다운: 마지막 실행(예약 실행 포함) 후 COOLDOWN_MIN 분 이내면 실행하지 않음 — GitHub API 실행 기록으로 판단, DB 없음
 * - 이미 대기·실행 중인 수집이 있으면 새로 실행하지 않음
 * - CORS: ALLOWED_ORIGIN(기본 https://bono2610.github.io) 출처만 허용
 *
 * POST /api/refresh → { status: "started" | "running" | "cooldown" | "error", message, minutesLeft? }
 * GET  /api/refresh → 실행하지 않고 상태만 알려 줌 (확인용)
 */
const REPO = process.env.GITHUB_REPO || "bono2610/g2b-monitor";
const WORKFLOW = process.env.WORKFLOW_FILE || "collect.yml";
const REF = process.env.WORKFLOW_REF || "main";
const COOLDOWN_MIN = Number(process.env.COOLDOWN_MIN || 30);
const ALLOWED_ORIGIN = process.env.ALLOWED_ORIGIN || "https://bono2610.github.io";
const ACTIVE = new Set(["queued", "in_progress", "waiting", "requested", "pending"]);

async function gh(path, init = {}) {
  return fetch(`https://api.github.com/repos/${REPO}${path}`, {
    ...init,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "g2b-monitor-relay",
      ...(init.body ? { "Content-Type": "application/json" } : {}),
    },
  });
}

// 최근 실행 기록으로 지금 상태 판단
async function check() {
  const r = await gh(`/actions/workflows/${encodeURIComponent(WORKFLOW)}/runs?per_page=10`);
  if (!r.ok) return { status: "error", http: 502, message: `GitHub 실행 기록을 읽지 못했어요 (${r.status})` };
  const { workflow_runs: runs = [] } = await r.json();
  if (runs.some(x => ACTIVE.has(x.status))) return { status: "running", message: "수집 중" };
  const last = runs[0];   // 최신순
  if (last) {
    const lastAt = Date.parse(last.run_started_at || last.created_at);
    const left = COOLDOWN_MIN * 60e3 - (Date.now() - lastAt);
    if (left > 0) {
      const minutesLeft = Math.ceil(left / 60e3);
      return { status: "cooldown", minutesLeft, lastRunAt: new Date(lastAt).toISOString(), message: `${minutesLeft}분 후 다시 가능` };
    }
  }
  return { status: "ready", lastRunAt: last ? last.run_started_at || last.created_at : null };
}

module.exports = async (req, res) => {
  const origin = req.headers.origin || "";
  res.setHeader("Vary", "Origin");
  res.setHeader("Cache-Control", "no-store");
  if (origin === ALLOWED_ORIGIN) {
    res.setHeader("Access-Control-Allow-Origin", ALLOWED_ORIGIN);
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
    res.setHeader("Access-Control-Max-Age", "86400");
  }
  const send = (code, body) => res.status(code).json(body);

  if (req.method === "OPTIONS") return res.status(origin === ALLOWED_ORIGIN ? 204 : 403).end();
  // 브라우저에서 오는 요청은 허용한 출처만 (주소창·curl 확인용 GET 은 Origin 이 없어 허용)
  if (origin && origin !== ALLOWED_ORIGIN) return send(403, { status: "error", message: "허용되지 않은 출처입니다" });
  if (!["GET", "POST"].includes(req.method)) return send(405, { status: "error", message: "GET 또는 POST만 가능합니다" });
  if (!process.env.GITHUB_TOKEN) return send(500, { status: "error", message: "중계 함수에 GITHUB_TOKEN 환경변수가 없어요" });

  try {
    const st = await check();
    if (st.status === "error") return send(st.http, { status: "error", message: st.message });
    if (req.method === "GET" || st.status !== "ready") return send(200, st);

    const r = await gh(`/actions/workflows/${encodeURIComponent(WORKFLOW)}/dispatches`, {
      method: "POST", body: JSON.stringify({ ref: REF }),
    });
    if (r.status === 204) return send(202, { status: "started", message: "수집을 시작했어요" });
    return send(502, { status: "error", message: `수집 실행 요청이 거절됐어요 (${r.status})` });
  } catch (e) {
    return send(502, { status: "error", message: "GitHub에 연결하지 못했어요" });
  }
};
