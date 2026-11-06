/*
 * 나라장터 AI 사업 모니터 — AI 도우미 (선택 기능, Claude API)
 *
 * - API 키가 없으면 아무것도 바뀌지 않음: 기존 화면·엑셀·보고서는 그대로 동작하고, 이 패널만 "키를 넣으면 쓸 수 있어요" 안내
 * - 키는 사용자가 직접 넣고 이 브라우저에만 저장(localStorage 또는 sessionStorage). 저장소·서버에는 키가 없음
 * - 브라우저에서 api.anthropic.com 을 바로 부름 (anthropic-dangerous-direct-browser-access 헤더)
 * - 숫자는 Claude 가 계산하지 않음: Claude 는 도구(필터 바꾸기·집계·유사사업 찾기·보고서 만들기)를 고르고,
 *   계산은 이 페이지의 코드가 해서 결과를 돌려줌 → 답변 숫자 = 화면 숫자
 * - Claude 에 보내는 것은 공개된 나라장터 공고 정보(공고명·기관·금액·낙찰 정보)뿐
 */
(() => {
  "use strict";
  const API = "https://api.anthropic.com/v1/messages";
  const MODELS = [
    { id: "claude-haiku-5-5", name: "Haiku 5.5 — 빠르고 저렴 (대화 추천)" },
    { id: "claude-sonnet-5-5", name: "Sonnet 5.5 — 문장 품질·정확도 높음 (보고서 추천)" },
  ];
  const K_KEY = "g2b-claude-key", K_CHAT = "g2b-claude-chat-model", K_REP = "g2b-claude-report-model";
  const store = {
    get(k) { try { return localStorage.getItem(k) ?? sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v, persist = true) {
      try { localStorage.removeItem(k); sessionStorage.removeItem(k); (persist ? localStorage : sessionStorage).setItem(k, v); } catch (e) {}
    },
    del(k) { try { localStorage.removeItem(k); sessionStorage.removeItem(k); } catch (e) {} },
    persisted(k) { try { return localStorage.getItem(k) != null; } catch (e) { return false; } },
  };
  const getKey = () => (store.get(K_KEY) || "").trim();
  const hasKey = () => /^sk-ant-/.test(getKey());
  const chatModel = () => store.get(K_CHAT) || MODELS[0].id;
  const reportModel = () => store.get(K_REP) || MODELS[1].id;
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ───────── Claude 호출 ─────────
  async function call({ model, system, messages, tools, max_tokens = 1500 }) {
    if (!hasKey()) throw new Error("Claude API 키가 없습니다");
    let r;
    try {
      r = await fetch(API, {
        method: "POST",
        headers: { "content-type": "application/json", "x-api-key": getKey(), "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true" },
        body: JSON.stringify({ model: model || chatModel(), max_tokens, system, messages, ...(tools ? { tools } : {}) }),
      });
    } catch (e) {
      throw new Error("Claude 서버에 연결하지 못했습니다. 인터넷 연결을 확인하세요 (기관망에서는 api.anthropic.com 접속이 막혀 있을 수 있어요)");
    }
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      const msg = j.error && j.error.message || "";
      if (r.status === 401) throw new Error("API 키가 올바르지 않습니다. ⚙ 설정에서 키를 확인하세요");
      if (r.status === 403) throw new Error("이 키로는 사용할 수 없는 요청입니다 (" + msg + ")");
      if (r.status === 404) throw new Error("모델을 찾을 수 없습니다. ⚙ 설정에서 다른 모델을 골라 보세요 (" + msg + ")");
      if (r.status === 429) throw new Error("요청이 너무 많거나 사용 한도에 걸렸습니다. 잠시 후 다시 시도하세요");
      if (r.status === 529 || r.status >= 500) throw new Error("Claude 서버가 혼잡합니다. 잠시 후 다시 시도하세요");
      throw new Error(`요청 실패 (${r.status}) ${msg}`);
    }
    return j;
  }
  const textOf = res => (res.content || []).filter(b => b.type === "text").map(b => b.text).join("\n").trim();

  // ───────── 보고서 빈칸 채우기 ─────────
  // payload: report.js 가 계산한 숫자 묶음. 돌려주는 값: 문장 배열 (○ 한 줄씩)
  async function reportInsights(kind, payload, scopeText) {
    const what = kind === "cost"
      ? "'적정 범위(안)' 아래 '추가 근거' 칸에 들어갈 문장. 유사 사례의 분포(중앙값·보통 범위 하위25~상위25%)를 근거로 이 분야 사업비를 어느 범위에서 검토하면 좋을지, 사업 범위·기간·장비 포함 여부에 따라 달라질 수 있다는 점, 사례 비교 시 유의점"
      : "'4. 시사점' 칸에 들어갈 문장. 도내 현황의 특징(규모·순위·분야 쏠림이나 공백), 타 시도와 비교한 시사점, 향후 검토할 사항";
    const system = "당신은 지방자치단체 행정 보고서 초안 작성을 돕는 보조자입니다. 주어진 JSON 숫자만 근거로 쓰고, JSON 에 없는 숫자·사실·사업명은 절대 만들지 마세요. " +
      "공문서 개조식 문체('~임', '~필요', '~검토' 등 명사형 종결)로 씁니다. 판단이 들어가는 표현은 '검토 필요', '~로 보임'처럼 단정하지 않습니다.";
    const prompt = `아래는 나라장터 공고 자동 수집 자료를 코드로 집계한 결과입니다.\n자료 범위: ${scopeText}\n\n${JSON.stringify(payload)}\n\n` +
      `${what}을 4~6개 작성하세요. 각 문장은 80자 안팎, 앞에 기호(○, -, 숫자)를 붙이지 마세요.\n` +
      `반드시 JSON 문자열 배열 하나만 출력하세요. 예: ["…", "…"]`;
    const res = await call({ model: reportModel(), system, messages: [{ role: "user", content: prompt }], max_tokens: 1200 });
    const t = textOf(res);
    let arr = null;
    try { arr = JSON.parse(t.slice(t.indexOf("["), t.lastIndexOf("]") + 1)); } catch (e) { arr = null; }
    if (!Array.isArray(arr)) arr = t.split("\n").map(s => s.replace(/^\s*([○\-•*]|\d+[.)])\s*/, "").trim()).filter(Boolean);
    arr = arr.map(s => String(s).replace(/^\s*([○\-•*]|\d+[.)])\s*/, "").trim()).filter(Boolean).slice(0, 8);
    if (!arr.length) throw new Error("AI 응답이 비어 있습니다");
    return arr;
  }

  // ───────── 대화 도구 ─────────
  const TOOLS = [
    { name: "set_filters",
      description: "대시보드 필터를 바꿉니다. 화면(요약·지도·차트·목록)이 실제로 바뀌고, 엑셀·보고서도 이 필터를 따릅니다. 넘긴 항목만 바뀌고 나머지는 그대로입니다. 항목을 풀려면 빈 문자열(\"\") 또는 0, 빈 배열을 넘기세요.",
      input_schema: { type: "object", properties: {
        query: { type: "string", description: "공고명·기관·낙찰업체에 들어간 글자 검색 (한 단어 또는 구절)" },
        type: { type: "string", description: "유형 (용역/물품 등, 목록 참고)" },
        region: { type: "string", description: "지역(시도) — 목록의 이름 그대로" },
        status: { type: "string", enum: ["", "open", "closed", "award"], description: "open=진행중, closed=마감(결과대기), award=낙찰완료" },
        period_days: { type: "integer", description: "최근 N일 공고만. 0=전체 기간" },
        min_amount: { type: "integer", description: "배정예산 하한(원). 0=제한 없음. 예: 1억=100000000" },
        keywords: { type: "array", items: { type: "string" }, description: "수집 키워드 칩 (목록 참고, 여러 개면 OR)" },
        local_only: { type: "boolean", description: "true=지자체 사업만" },
        field: { type: "string", description: "분야(자동분류) — 목록의 이름 그대로" },
      } } },
    { name: "reset_filters", description: "모든 필터를 기본값(최근 90일, 나머지 전체)으로 되돌립니다.", input_schema: { type: "object", properties: {} } },
    { name: "summarize",
      description: "현재 필터 결과를 집계합니다 (건수·배정예산 합계/중앙값/보통범위·낙찰 건수·평균 낙찰률·예산 집행률 중앙값). group_by 로 묶어서 볼 수 있습니다.",
      input_schema: { type: "object", properties: {
        group_by: { type: "string", enum: ["none", "region", "field", "kind", "status", "month", "org", "winner"], description: "묶음 기준. org=수요기관, winner=낙찰업체, kind=기관구분" },
        sort_by: { type: "string", enum: ["count", "budget", "award_amount"], description: "묶음 정렬 기준 (기본 count)" },
        top_n: { type: "integer", description: "묶음 최대 개수 (기본 10, 최대 30)" },
      } } },
    { name: "list_bids",
      description: "현재 필터 결과에서 공고를 골라 보여줍니다. 화면 목록의 정렬도 같이 바뀝니다.",
      input_schema: { type: "object", properties: {
        sort_by: { type: "string", enum: ["notice_desc", "budget_desc", "budget_asc", "rate_asc", "rate_desc", "close_asc"], description: "notice_desc=최근 공고순, rate=낙찰률(예정가격 대비), close_asc=마감 임박순" },
        limit: { type: "integer", description: "최대 20" },
      } } },
    { name: "find_similar",
      description: "사업비 적정성·유사사업 비교용. 현재 필터와 상관없이 전체 데이터에서 공고명에 키워드가 들어간 공고를 찾아 배정예산 분포(중앙값·하위25%·상위25%)와 사례를 돌려줍니다. 단가계약·예산 미기재 건은 통계에서 뺍니다.",
      input_schema: { type: "object", required: ["keywords"], properties: {
        keywords: { type: "array", items: { type: "string" }, description: "공고명에 들어갈 핵심어 (OR). 예: [\"챗봇\",\"민원상담\"]. 동의어를 2~4개 넣으면 좋음" },
        must_include: { type: "array", items: { type: "string" }, description: "모두 들어가야 하는 단어 (AND, 선택)" },
        field: { type: "string", description: "분야로 한정 (선택)" },
        region: { type: "string", description: "지역으로 한정 (선택)" },
        type: { type: "string", description: "유형으로 한정 (선택, 예: 용역)" },
        limit: { type: "integer", description: "사례 최대 개수 (기본 10, 최대 20)" },
      } } },
    { name: "download_report",
      description: "현재 필터 기준 보고서 초안 파일을 만들어 내려받게 합니다. kind=home: 도내 AI 사업 추진현황, kind=cost: ○○사업 사업비 산정 근거(field 필요).",
      input_schema: { type: "object", required: ["kind"], properties: {
        kind: { type: "string", enum: ["home", "cost"] },
        field: { type: "string", description: "kind=cost 일 때 분야" },
        ai_fill: { type: "boolean", description: "시사점/추가 근거 칸을 AI 초안으로 채울지 (기본 true)" },
        format: { type: "string", enum: ["hwpx", "docx"], description: "기본 hwpx(한글)" },
      } } },
    { name: "download_excel",
      description: "분석 엑셀(시도별·도내·사업비·업체·규모구간 시트)을 내려받게 합니다.",
      input_schema: { type: "object", properties: { scope: { type: "string", enum: ["filter", "all"], description: "filter=현재 필터(기본), all=전체 데이터" } } } },
  ];

  async function runTool(name, input) {
    const A = window.G2BApp;
    if (!A) return { error: "페이지가 아직 준비되지 않았습니다" };
    try {
      switch (name) {
        case "set_filters": return A.applyFilters(input || {});
        case "reset_filters": return A.resetFilters();
        case "summarize": return A.summarize(input || {});
        case "list_bids": return A.listBids(input || {});
        case "find_similar": return A.findSimilar(input || {});
        case "download_report": {
          if (!window.G2BReport) return { error: "보고서 모듈을 불러오지 못했습니다" };
          const r = await window.G2BReport.buildReport(input.kind, input.field, { scope: "filter", format: input.format || "hwpx", ai: input.ai_fill !== false });
          return { ok: true, file: r.name, rows: r.n, scope: r.scope, field: r.field, ai_filled: r.ai, ai_error: r.aiError || undefined };
        }
        case "download_excel": {
          if (!window.G2BReport) return { error: "분석 모듈을 불러오지 못했습니다" };
          const r = await window.G2BReport.downloadAnalysis(input.scope || "filter");
          return { ok: true, rows: r.n, scope: r.scope };
        }
      }
      return { error: "알 수 없는 도구: " + name };
    } catch (e) { return { error: e && e.message || String(e) }; }
  }

  function systemPrompt() {
    const A = window.G2BApp, v = A.vocab(), meta = A.meta();
    return [
      "당신은 '나라장터 AI 사업 모니터' 대시보드의 AI 도우미입니다. 사용자는 지방자치단체(제주) 공무원입니다.",
      `오늘: ${new Date().toISOString().slice(0, 10)} · 데이터: 조달청 나라장터 OpenAPI에서 공고명에 AI 관련 키워드가 든 입찰공고·낙찰정보를 하루 1회 수집 (마지막 수집 ${meta.updated || "-"}, 전체 ${v.total}건, 공고일 ${v.from}~${v.to})`,
      "",
      "## 규칙",
      "- 숫자(건수·금액·비율·순위)는 반드시 도구 결과에 있는 값만 씁니다. 직접 계산·추정하거나 기억으로 말하지 않습니다. 필요하면 도구를 먼저 부르세요.",
      "- 필터를 바꿔 달라는 말(예: '제주만', '최근 3개월', '1억 이상')은 set_filters 로 화면을 실제로 바꾼 뒤 결과 건수를 알려줍니다. 질문에 답하려고 필터가 필요할 때도 set_filters 를 쓰고, 바꿨다고 알려줍니다.",
      "- '최근 N개월'은 period_days=N×30. 금액은 원 단위 정수(1억=100000000).",
      "- 사업비 적정성·유사사업 질문은 find_similar 로 비교하고 중앙값과 보통 범위(하위25~상위25%)를 제시합니다. 사업 범위·기간이 달라 단순 비교는 어렵다는 점을 덧붙이고, 판단은 참고용이라고 밝힙니다.",
      "- 업체 관련 답은 공개 낙찰 정보를 정리한 참고용이며 업체 평가가 아니라고 짧게 밝힙니다.",
      "- 공고 세부 내용(과업지시서·첨부파일)은 이 데이터에 없습니다. 모르면 모른다고 하고 나라장터 원문 확인을 권합니다.",
      "- 지역·분야·기관구분은 기관명·공고명 규칙으로 추정한 값입니다.",
      "- 답은 한국어로 짧게(5~8줄 안팎), 필요하면 '- ' 목록. 금액은 억/백만원 단위로 읽기 쉽게. 마지막 줄에 '근거: …' 로 어떤 조건의 몇 건 기준인지 적습니다.",
      "- 보고서·엑셀을 요청하면 download_report / download_excel 을 부릅니다. 파일은 현재 필터 기준입니다.",
      "",
      "## 필터에 쓸 수 있는 값",
      `- 유형: ${v.types.join(", ")}`,
      `- 지역: ${v.regions.join(", ")} (전남광주=광주·전남 통합, 중앙·기타=중앙부처 등 지명 없는 기관)`,
      `- 분야: ${v.fields.join(", ")}`,
      `- 키워드 칩: ${v.keywords.join(", ")}`,
      "",
      `## 지금 화면 필터: ${A.filterSummary() || "기본(최근 90일)"} → ${A.getRows().length}건`,
    ].join("\n");
  }

  // ───────── 패널 UI ─────────
  const CSS = `
  .aip{position:fixed;top:0;right:0;bottom:0;width:min(420px,100vw);background:var(--surface);border-left:1px solid var(--border);
    box-shadow:-12px 0 32px rgba(0,0,0,.14);z-index:40;display:flex;flex-direction:column;transform:translateX(105%);transition:transform .2s ease}
  .aip.on{transform:none}
  .aip .hd{display:flex;align-items:center;gap:8px;padding:14px 14px 12px;border-bottom:1px solid var(--border);
    background:linear-gradient(135deg,var(--brand,#14315c),var(--brand-2,#1f4b8a));color:#fff}
  .aip .hd b{font-size:15px;flex:1}.aip .hd small{display:block;font-weight:400;font-size:11.5px;color:#cfe0f7}
  .aip .hd button{border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.1);color:#fff;border-radius:8px;padding:5px 9px;font:inherit;font-size:12px;cursor:pointer}
  .aip .hd button:hover{background:rgba(255,255,255,.2)}
  .aip .log{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}
  .aip .m{max-width:92%;padding:9px 12px;border-radius:12px;font-size:13.5px;line-height:1.6;word-break:keep-all;overflow-wrap:anywhere}
  .aip .m.u{align-self:flex-end;background:var(--series-1);color:#fff;border-bottom-right-radius:4px}
  .aip .m.a{align-self:flex-start;background:var(--surface-2);color:var(--text);border-bottom-left-radius:4px}
  .aip .m.a ul{margin:4px 0;padding-left:18px}.aip .m.a li{margin:1px 0}
  .aip .m.err{background:var(--warn-bg);color:var(--warn)}
  .aip .act{align-self:flex-start;font-size:12px;color:var(--text-2);background:var(--bg);border:1px dashed var(--border);border-radius:8px;padding:4px 9px}
  .aip .act.bad{color:var(--warn)}
  .aip .sugg{display:flex;flex-direction:column;gap:6px;margin-top:4px}
  .aip .sugg button{text-align:left;border:1px solid var(--border);background:var(--bg);color:var(--text);border-radius:10px;padding:8px 11px;font:inherit;font-size:13px;cursor:pointer}
  .aip .sugg button:hover{border-color:var(--series-1)}
  .aip .intro{color:var(--text-2);font-size:13px;line-height:1.65}
  .aip .intro b{color:var(--text)}
  .aip form.in{display:flex;gap:8px;padding:12px;border-top:1px solid var(--border);background:var(--surface)}
  .aip textarea{flex:1;resize:none;font:inherit;font-size:13.5px;color:var(--text);background:var(--bg);border:1px solid var(--border);border-radius:10px;padding:8px 10px;height:42px;max-height:120px}
  .aip textarea:disabled{opacity:.55}
  .aip form.in button{border:0;background:var(--series-1);color:#fff;border-radius:10px;padding:0 14px;font:inherit;font-weight:600;cursor:pointer}
  .aip form.in button:disabled{opacity:.5;cursor:default}
  .aip .set{padding:14px;border-bottom:1px solid var(--border);background:var(--bg);display:none;font-size:13px}
  .aip .set.on{display:block}
  .aip .set label{display:block;font-weight:600;font-size:12.5px;margin:10px 0 4px}
  .aip .set input[type=password],.aip .set input[type=text],.aip .set select{width:100%;font:inherit;font-size:13px;color:var(--text);background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:7px 10px}
  .aip .set .chk{display:flex;gap:6px;align-items:center;font-weight:400;margin-top:8px}
  .aip .set .row{display:flex;gap:6px;margin-top:12px;flex-wrap:wrap}
  .aip .set .row button{border:1px solid var(--border);background:var(--surface);color:var(--text);border-radius:8px;padding:6px 12px;font:inherit;font-size:12.5px;cursor:pointer}
  .aip .set .row button.pri{background:var(--series-1);border-color:var(--series-1);color:#fff}
  .aip .set .note{color:var(--muted);font-size:11.5px;line-height:1.55;margin-top:8px}
  .aip .set .st{font-size:12px;margin-top:8px}
  .aip .dots::after{content:"···";animation:aidots 1s steps(3) infinite;letter-spacing:2px}
  @keyframes aidots{0%{opacity:.3}50%{opacity:1}}
  #aiBtn .k{display:inline-block;width:7px;height:7px;border-radius:50%;background:#9ca3af;margin-left:2px}
  #aiBtn.has .k{background:#4ade80}
  `;

  let panel, log, ta, sendBtn, setBox, busy = false;
  let history = [];   // Claude messages (tool_use/tool_result 포함)

  const md = t => {   // 아주 단순한 마크다운: **굵게**, '- ' 목록, 줄바꿈
    const lines = esc(t).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").split("\n");
    let out = "", inList = false;
    lines.forEach(l => {
      const m = /^\s*[-•*]\s+(.*)$/.exec(l);
      if (m) { if (!inList) { out += "<ul>"; inList = true; } out += `<li>${m[1]}</li>`; }
      else { if (inList) { out += "</ul>"; inList = false; } out += l.trim() ? l.replace(/^#+\s*/, "") + "<br>" : (out.endsWith("<br>") ? "" : "<br>"); }
    });
    if (inList) out += "</ul>";
    return out.replace(/(<br>)+$/, "");
  };
  const add = (cls, html) => { const d = document.createElement("div"); d.className = cls; d.innerHTML = html; log.append(d); log.scrollTop = log.scrollHeight; return d; };

  const ACT_LABEL = { set_filters: "필터 적용", reset_filters: "필터 초기화", summarize: "집계", list_bids: "공고 고르기",
    find_similar: "유사사업 찾기", download_report: "보고서 만들기", download_excel: "분석 엑셀 만들기" };
  const actText = (name, out) => {
    if (out && out.error) return `⚠ ${ACT_LABEL[name] || name} 실패: ${out.error}`;
    if (name === "set_filters" || name === "reset_filters") return `🔧 ${ACT_LABEL[name]} → ${out.filter || "기본"} · ${out.count}건`;
    if (name === "find_similar") return `🔎 유사사업 ${out.matched}건 (통계 대상 ${out.stats ? out.stats.n : 0}건)`;
    if (name === "download_report") return `📄 ${out.file} 내려받음${out.ai_filled ? " · AI 초안 포함" : ""}`;
    if (name === "download_excel") return `📊 분석 엑셀 내려받음 (${out.rows}건)`;
    return `📐 ${ACT_LABEL[name] || name}`;
  };

  const SUGG = ["제주 최근 3개월 1억 이상 사업만 보여줘", "민원 챗봇 구축 3억이면 적정한 수준이야?", "지금 조건에서 낙찰 많이 받은 업체 TOP 5",
    "시도별로 사업 규모 비교해줘", "지금 조건으로 도내 현황 보고서 만들어줘"];

  function intro() {
    log.innerHTML = "";
    if (!hasKey()) {
      add("intro", `<b>AI 도우미는 선택 기능입니다.</b><br>Claude API 키를 넣으면 말로 필터를 바꾸고, 집계·유사사업 비교·보고서 시사점 초안을 받을 수 있어요.<br><br>
        키가 없어도 대시보드·엑셀·보고서는 지금처럼 모두 동작합니다.<br><br>
        <button type="button" class="btn primary" id="aiOpenSet">⚙ API 키 넣기</button>`);
      log.querySelector("#aiOpenSet").addEventListener("click", () => toggleSet(true));
      return;
    }
    const d = add("intro", `무엇이든 물어보세요. 필터를 바꾸면 화면도 같이 바뀝니다.<br><span style="color:var(--muted);font-size:12px">숫자는 화면 데이터를 코드가 계산한 값만 씁니다 · 공개된 공고 정보만 Claude로 보냅니다</span>`);
    const s = document.createElement("div"); s.className = "sugg";
    SUGG.forEach(q => { const b = document.createElement("button"); b.type = "button"; b.textContent = q; b.addEventListener("click", () => send(q)); s.append(b); });
    d.append(s);
  }

  async function send(text) {
    text = String(text || "").trim();
    if (!text || busy) return;
    if (!hasKey()) { toggleSet(true); return; }
    if (log.querySelector(".intro")) log.innerHTML = "";
    busy = true; sendBtn.disabled = true; ta.value = "";
    add("m u", esc(text));
    const mark = history.length;
    history.push({ role: "user", content: text });
    const DOTS = '<span class="dots">생각 중</span>';
    let bubble = add("m a", DOTS), done = false;
    try {
      for (let turn = 0; turn < 8 && !done; turn++) {
        const res = await call({ system: systemPrompt(), messages: history, tools: TOOLS, max_tokens: 1500 });
        history.push({ role: "assistant", content: res.content });
        const uses = (res.content || []).filter(b => b.type === "tool_use"), t = textOf(res);
        if (!uses.length) { bubble.innerHTML = md(t || "(응답 없음)"); done = true; break; }
        if (t) { bubble.innerHTML = md(t); bubble = null; }   // 중간 설명은 남겨 둠
        const results = [];
        for (const u of uses) {
          const out = (await runTool(u.name, u.input)) || {};
          add("act" + (out.error ? " bad" : ""), esc(actText(u.name, out)));
          results.push({ type: "tool_result", tool_use_id: u.id, content: JSON.stringify(out).slice(0, 12000), ...(out.error ? { is_error: true } : {}) });
        }
        history.push({ role: "user", content: results });
        if (bubble) log.append(bubble); else bubble = add("m a", DOTS);   // 대기 말풍선을 맨 아래로
      }
      if (!done) bubble.innerHTML = "도구 호출이 너무 길어져 멈췄습니다. 질문을 나눠서 다시 물어봐 주세요.";
    } catch (e) {
      if (!bubble) bubble = add("m a", "");
      bubble.className = "m a err"; bubble.textContent = e.message || String(e);
      history.length = mark;   // 실패한 질문은 기록에서 빼서 다음 질문이 꼬이지 않게
    } finally {
      if (!done && history.length > mark && history[history.length - 1].role === "user" && typeof history[history.length - 1].content !== "string") {
        // 도구 결과로 끝났는데 멈춘 경우: 기록을 질문 전으로 되돌림 (짝이 안 맞는 기록 방지)
        history.length = mark;
      }
      busy = false; sendBtn.disabled = !hasKey(); ta.focus();
      trim();
    }
  }
  // 대화가 길어지면 앞쪽을 잘라 비용을 줄임 (사용자 질문 경계에서만 자름)
  function trim() {
    while (history.length > 40) {
      history.shift();
      while (history.length && !(history[0].role === "user" && typeof history[0].content === "string")) history.shift();
    }
  }

  function toggleSet(on) {
    setBox.classList.toggle("on", on == null ? !setBox.classList.contains("on") : on);
    if (setBox.classList.contains("on")) {
      setBox.querySelector("#aiKey").value = getKey();
      setBox.querySelector("#aiRemember").checked = !getKey() || store.persisted(K_KEY);
      setBox.querySelector("#aiChatModel").value = chatModel();
      setBox.querySelector("#aiRepModel").value = reportModel();
      setBox.querySelector(".st").textContent = "";
    }
  }
  function refreshState() {
    const on = hasKey();
    ta.disabled = !on; sendBtn.disabled = !on || busy;
    ta.placeholder = on ? "예: 제주 빼고 최근 3개월 챗봇 사업 보여줘" : "API 키를 넣으면 쓸 수 있어요";
    const b = document.getElementById("aiBtn"); if (b) b.classList.toggle("has", on);
    document.dispatchEvent(new CustomEvent("g2b-ai-key"));
  }

  function build() {
    const st = document.createElement("style"); st.textContent = CSS; document.head.append(st);
    panel = document.createElement("aside");
    panel.className = "aip"; panel.setAttribute("aria-label", "AI 도우미");
    const opts = sel => MODELS.map(m => `<option value="${m.id}"${m.id === sel ? " selected" : ""}>${esc(m.name)}</option>`).join("");
    panel.innerHTML = `
      <div class="hd"><b>AI 도우미<small>Claude · 선택 기능</small></b>
        <button type="button" id="aiNew" title="대화를 비우고 새로 시작">새 대화</button>
        <button type="button" id="aiSetBtn" title="API 키·모델 설정">⚙ 설정</button>
        <button type="button" id="aiClose" aria-label="닫기">✕</button></div>
      <div class="set" id="aiSet">
        <label for="aiKey">Claude API 키</label>
        <input type="password" id="aiKey" placeholder="sk-ant-…" autocomplete="off" spellcheck="false">
        <label class="chk"><input type="checkbox" id="aiRemember"> 이 브라우저에 기억 (끄면 창을 닫을 때 지워짐)</label>
        <label for="aiChatModel">대화 모델</label><select id="aiChatModel">${opts(chatModel())}</select>
        <label for="aiRepModel">보고서 초안 모델</label><select id="aiRepModel">${opts(reportModel())}</select>
        <div class="row"><button type="button" class="pri" id="aiSave">저장</button><button type="button" id="aiTest">연결 확인</button><button type="button" id="aiDel">키 지우기</button></div>
        <div class="st" role="status"></div>
        <div class="note">· 키는 이 브라우저에만 저장되고 Claude(api.anthropic.com)로만 보냅니다. 저장소·서버에는 저장되지 않습니다.<br>
          · 함께 쓰는 PC에서는 '기억'을 끄세요. 사용량만큼 키 소유자에게 요금이 부과됩니다.<br>
          · 키는 <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noopener" style="color:inherit">Claude Console › API Keys</a>에서 발급합니다.</div>
      </div>
      <div class="log" id="aiLog" aria-live="polite"></div>
      <form class="in" id="aiForm"><textarea id="aiIn" rows="1" aria-label="AI 도우미에게 질문"></textarea><button type="submit" id="aiSend">보내기</button></form>`;
    document.body.append(panel);
    log = panel.querySelector("#aiLog"); ta = panel.querySelector("#aiIn"); sendBtn = panel.querySelector("#aiSend"); setBox = panel.querySelector("#aiSet");
    const stEl = setBox.querySelector(".st");
    panel.querySelector("#aiClose").addEventListener("click", () => panel.classList.remove("on"));
    panel.querySelector("#aiSetBtn").addEventListener("click", () => toggleSet());
    panel.querySelector("#aiNew").addEventListener("click", () => { if (!busy) { history = []; intro(); } });
    panel.querySelector("#aiForm").addEventListener("submit", e => { e.preventDefault(); send(ta.value); });
    ta.addEventListener("keydown", e => { if (e.key === "Enter" && !e.shiftKey && !e.isComposing) { e.preventDefault(); send(ta.value); } });
    const save = () => {
      const k = setBox.querySelector("#aiKey").value.trim(), keep = setBox.querySelector("#aiRemember").checked;
      if (k && !/^sk-ant-/.test(k)) { stEl.textContent = "⚠ Claude API 키는 sk-ant- 로 시작합니다"; return false; }
      if (k) store.set(K_KEY, k, keep); else store.del(K_KEY);
      store.set(K_CHAT, setBox.querySelector("#aiChatModel").value);
      store.set(K_REP, setBox.querySelector("#aiRepModel").value);
      refreshState(); return true;
    };
    setBox.querySelector("#aiSave").addEventListener("click", () => { if (save()) { stEl.textContent = hasKey() ? "✓ 저장했습니다" : "키 없이 저장했습니다"; history = []; intro(); setTimeout(() => toggleSet(false), 700); } });
    setBox.querySelector("#aiDel").addEventListener("click", () => { store.del(K_KEY); setBox.querySelector("#aiKey").value = ""; refreshState(); history = []; intro(); stEl.textContent = "키를 지웠습니다"; });
    setBox.querySelector("#aiTest").addEventListener("click", async () => {
      if (!save() || !hasKey()) { if (!hasKey()) stEl.textContent = "먼저 키를 입력하세요"; return; }
      stEl.textContent = "확인 중…";
      try { await call({ model: chatModel(), messages: [{ role: "user", content: "ping" }], max_tokens: 5 }); stEl.textContent = "✓ 연결 정상 (" + chatModel() + ")"; }
      catch (e) { stEl.textContent = "⚠ " + e.message; }
    });
    const btn = document.getElementById("aiBtn");
    if (btn) btn.addEventListener("click", () => { panel.classList.toggle("on"); if (panel.classList.contains("on")) (hasKey() ? ta : panel.querySelector("#aiSetBtn")).focus(); });
    document.addEventListener("keydown", e => { if (e.key === "Escape" && panel.classList.contains("on") && !document.querySelector("dialog[open]")) panel.classList.remove("on"); });
    refreshState(); intro();
  }

  window.G2BAI = { hasKey, call, reportInsights, open: () => panel && panel.classList.add("on") };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", build); else build();
})();
