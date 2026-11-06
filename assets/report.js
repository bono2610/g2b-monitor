/*
 * 나라장터 AI 사업 모니터 — 분석 엑셀 · 보고서 초안 (브라우저에서 생성, 서버 불필요)
 *
 * - 범위: 기본은 '현재 화면 필터' 결과, 선택하면 '전체 데이터'
 * - 숫자는 모두 코드가 집계, 문장은 고정 템플릿에 숫자만 채움
 * - Claude API 키가 있고 사용자가 고르면 '시사점'·'적정 범위 근거' 빈칸만 AI 초안으로 채움 (assets/ai.js) — 키가 없으면 빈칸 그대로
 * - 분류(지역·기관구분·분야)는 수집기가 config.json 규칙으로 계산해 둔 값을 그대로 씀
 * - 보고서는 HWPX(한글, assets/hwpx.js) 기본, .docx 도 선택 가능. 라이브러리는 버튼을 누를 때만 CDN에서 불러옴 (ExcelJS, docx)
 */
(() => {
  "use strict";

  // ───────── 라이브러리 지연 로딩 ─────────
  const LIBS = {
    ExcelJS: ["https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js",
              "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"],
    docx: ["https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js",
           "https://unpkg.com/docx@9.5.1/dist/index.iife.js"],
  };
  const loading = {};
  function loadLib(name) {
    if (window[name]) return Promise.resolve(window[name]);
    if (!loading[name]) {
      loading[name] = (async () => {
        for (const url of LIBS[name]) {
          try {
            await new Promise((res, rej) => {
              const s = document.createElement("script");
              s.src = url; s.onload = res; s.onerror = rej; document.head.append(s);
            });
            if (window[name]) return window[name];
          } catch (e) { /* 다음 CDN 시도 */ }
        }
        loading[name] = null;
        throw new Error(`${name} 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하세요.`);
      })();
    }
    return loading[name];
  }

  // ───────── 공통 도우미 ─────────
  const DEFAULT_RULES = {
    homeRegion: "제주",
    regions: ["제주", "서울", "부산", "대구", "인천", "대전", "울산", "세종", "경기", "강원", "충북", "충남", "전북", "전남광주", "경북", "경남", "중앙·기타"],
    fields: ["단가계약(통계제외)", "유지보수·운영", "교육·연수·행사", "감리·컨설팅·연구", "챗봇·민원상담", "생성형AI·LLM", "영상·CCTV 분석", "RPA·업무자동화", "디지털트윈", "데이터·빅데이터", "플랫폼·시스템 구축", "장비·물품 구매", "기타"],
    fieldRules: [],
    kinds: ["교육청·학교", "공공기관·출자출연", "지자체", "대학·연구기관", "기타"],
    excludePatterns: [],
    statsExcludeFields: ["단가계약(통계제외)"],
    similarExcludeFields: ["단가계약(통계제외)", "유지보수·운영", "교육·연수·행사", "감리·컨설팅·연구", "장비·물품 구매", "기타"],
    dedup: "공고명+수요기관이 같으면 1건",
  };
  const CASE_FIELDS = ["챗봇·민원상담", "생성형AI·LLM", "RPA·업무자동화"];
  const KIND_ORDER = ["지자체", "공공기관·출자출연", "교육청·학교", "대학·연구기관", "기타"];
  const OTHER = "중앙·기타";
  // 규모 구간 (배정예산 기준) — 분석 엑셀 '규모구간별통계' 시트용
  const BANDS = [["1억 미만", 0, 1e8], ["1~5억", 1e8, 5e8], ["5~10억", 5e8, 1e9], ["10억 이상", 1e9, null]];

  const sum = a => a.reduce((s, v) => s + (v || 0), 0);
  const avg = a => a.length ? sum(a) / a.length : null;
  const median = a => {
    if (!a.length) return null;
    const s = [...a].sort((x, y) => x - y), m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  };
  const pos = v => typeof v === "number" && v > 0;
  const pad = n => String(n).padStart(2, "0");
  const ymd = d => d ? `${d.getFullYear()}.${pad(d.getMonth() + 1)}.${pad(d.getDate())}` : "-";
  const stamp = (d = new Date()) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  const eok = n => (n / 1e8).toLocaleString("ko-KR", { maximumFractionDigits: 1 });
  const manwon = n => n == null ? "-" : Math.round(n / 1e4).toLocaleString("ko-KR");
  const comma = n => n == null ? "-" : Math.round(n).toLocaleString("ko-KR");
  const topOf = (arr) => {
    const m = new Map();
    arr.forEach(v => { if (v) m.set(v, (m.get(v) || 0) + 1); });
    let best = null;
    m.forEach((n, k) => { if (!best || n > best[1]) best = [k, n]; });
    return best;
  };
  function download(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  // ───────── 집계 (엑셀 수식과 같은 기준으로 계산 → 수식 결과값·보고서 숫자로 사용) ─────────
  function analyze(items, meta, statusLabel) {
    const R = { ...DEFAULT_RULES, ...(meta && meta.rules || {}) };
    const home = R.homeRegion;
    const rows = [...items].sort((a, b) => (+(b._notice || 0)) - (+(a._notice || 0)) || String(a.no).localeCompare(String(b.no)));
    rows.forEach(b => { b._st = statusLabel[b.status]; b._field = b.field || "기타"; b._kind = b.kind || "기타"; });

    const sidos = R.regions.filter(r => r !== OTHER);
    const grp = (list, name) => {
      const budgets = list.map(b => b.budget).filter(pos);
      return {
        name, n: list.length, budget: sum(list.map(b => b.budget)), budgetN: budgets.length,
        avgBudget: avg(budgets), median: median(budgets),
        min: budgets.length ? Math.min(...budgets) : null, max: budgets.length ? Math.max(...budgets) : null,
        awards: list.filter(b => b.status === "award").length,
        awardAmt: sum(list.map(b => b.awardAmt)), avgRate: avg(list.map(b => b.rate).filter(pos)),
        rateN: list.map(b => b.rate).filter(pos).length,
        open: list.filter(b => b.status === "open").length,
        closed: list.filter(b => b.status === "closed").length,
      };
    };
    const regionStats = [...sidos, OTHER].map(r => ({ ...grp(rows.filter(b => b.region === r), r), isSido: r !== OTHER }));
    const sidoStats = regionStats.filter(s => s.isSido);
    sidoStats.forEach(s => {
      s.budgetRank = 1 + sidoStats.filter(o => o.budget > s.budget).length;
      s.countRank = 1 + sidoStats.filter(o => o.n > s.n).length;
    });
    const total = grp(rows, "합계");

    const homeRows = rows.filter(b => b.region === home);
    const kinds = [...KIND_ORDER, ...R.kinds.filter(k => !KIND_ORDER.includes(k))];
    const homeKinds = kinds.map(k => grp(homeRows.filter(b => b._kind === k), k));
    const homeTotal = grp(homeRows, "합계");
    homeRows.forEach(b => {
      b._similar = !R.similarExcludeFields.includes(b._field) &&
        homeRows.some(o => o._field === b._field && o.demand !== b.demand) ? "유사사업 있음" : "";
    });

    const statFields = R.fields.filter(f => !R.statsExcludeFields.includes(f));
    const statRows = rows.filter(b => !R.statsExcludeFields.includes(b._field));
    const fieldStats = statFields.map(f => grp(statRows.filter(b => b._field === f), f));
    const fieldTotal = grp(statRows, "전체(통계제외 분야 제외)");
    const cases = f => rows.filter(b => b._field === f && pos(b.budget)).sort((a, b) => b.budget - a.budget);

    const wmap = new Map();
    rows.forEach(b => {
      const w = (b.winner || "").trim();
      if (!b.award || !w) return;
      if (!wmap.has(w)) wmap.set(w, []);
      wmap.get(w).push(b);
    });
    const winners = [...wmap].map(([name, list]) => ({
      name, n: list.length, amount: sum(list.map(b => b.awardAmt)), avgRate: avg(list.map(b => b.rate).filter(pos)),
      topField: topOf(list.map(b => b._field)), topOrg: topOf(list.map(b => b.demand)),
      local: list.filter(b => b._kind === "지자체").length, home: list.filter(b => b.region === home).length,
    })).sort((a, b) => b.n - a.n || b.amount - a.amount || a.name.localeCompare(b.name, "ko")).slice(0, 40);

    const dates = rows.map(b => b._notice).filter(Boolean).sort((a, b) => a - b);
    const from = dates[0] || null, to = dates[dates.length - 1] || null;
    const months = from ? Math.max(1, Math.round((to - from) / (30.44 * 864e5))) : 0;
    // 규모 구간별 통계: index.html 이 계산해 둔 b._exec(예산 집행률)·b._excl(통계 제외 사유) 사용
    const EX = ["단가계약", "예산 미기재", "비정상 비율"];
    const exCount = list => Object.fromEntries(EX.map(r => [r, list.filter(b => b._excl === r).length]));
    const awarded = rows.filter(b => b.status === "award");
    const bandOf = (list, k) => {
      const vals = list.map(b => b[k]).filter(pos);
      return { n: vals.length, avg: avg(vals), med: median(vals) };
    };
    const bandStats = BANDS.map(([name, lo, hi]) => {
      const inBand = b => pos(b.budget) && b.budget >= lo && (hi == null || b.budget < hi);
      const list = awarded.filter(b => !b._excl && inBand(b));
      const execList = list.filter(b => pos(b._exec));
      const ex = awarded.filter(b => b._excl && inBand(b));
      return { name, lo, hi, exec: bandOf(list, "_exec"), rate: bandOf(list, "rate"),
        budget: sum(execList.map(b => b.budget)), awardAmt: sum(execList.map(b => b.awardAmt)), exN: ex.length, ex: exCount(ex) };
    });
    const exAll = awarded.filter(b => b._excl), inc = awarded.filter(b => !b._excl);
    const bandTotal = { exec: bandOf(inc, "_exec"), rate: bandOf(inc, "rate"), exN: exAll.length, ex: exCount(exAll),
      noBudget: exCount(exAll.filter(b => !pos(b.budget))), noBudgetN: exAll.filter(b => !pos(b.budget)).length,
      budget: sum(inc.filter(b => pos(b._exec)).map(b => b.budget)), awardAmt: sum(inc.filter(b => pos(b._exec)).map(b => b.awardAmt)) };
    return { R, home, rows, sidos, regionStats, bandStats, bandTotal, EX, sidoStats, total, homeRows, kinds, homeKinds, homeTotal,
      statFields, fieldStats, fieldTotal, cases, winners, from, to, months,
      originalCount: sum(rows.map(b => b.dups || 1)) };
  }

  // ───────── 분석 엑셀 ─────────
  const FONT = "맑은 고딕";
  const NAVY = "FF1F3864", YELLOW = "FFFFF2CC", INPUT = "FFFFFF99", GRAY = "FFF2F2F2";
  const THIN = { style: "thin", color: { argb: "FFBFBFBF" } };
  const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  const NUM = "#,##0", MIL = "#,##0.0", PCT = "0.00", DATE = "yyyy-mm-dd";
  const q = s => `"${String(s).replace(/"/g, '""')}"`;
  // COUNTIFS 조건에 *, ?, ~ 가 들어 있으면 와일드카드로 해석되므로 이스케이프
  const crit = (ref, value) => /[*?~]/.test(String(value || "")) ?
    `SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${ref},"~","~~"),"*","~*"),"?","~?")` : ref;
  const xlDate = d => d ? new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())) : null;
  const F = (formula, result) => ({ formula, result: result == null ? "" : result });
  const AF = (formula, result, ref) => ({ formula, result: result == null ? "" : result, shareType: "array", ref });

  function styleHeader(row, from, to) {
    row.height = 30;
    for (let c = from; c <= to; c++) {
      const cell = row.getCell(c);
      cell.font = { name: FONT, size: 10, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = BORDER;
    }
  }
  function table(ws, top, headers, rows, opt = {}) {
    // rows: 셀 값 배열의 배열. 반환: 다음 빈 행 번호
    const hr = ws.getRow(top);
    headers.forEach((h, i) => { hr.getCell(i + 1).value = h; });
    styleHeader(hr, 1, headers.length);
    rows.forEach((vals, i) => {
      const r = ws.getRow(top + 1 + i);
      vals.forEach((v, j) => {
        const c = r.getCell(j + 1);
        c.value = v; c.border = BORDER;
        const fmt = opt.fmt && opt.fmt[j];
        if (fmt) c.numFmt = fmt;
        if (v && v.hyperlink) c.font = { name: FONT, size: 10, color: { argb: "FF0563C1" }, underline: true };
      });
      if (opt.rowStyle) opt.rowStyle(r, i);
    });
    return top + 1 + rows.length;
  }
  function title(ws, text, note) {
    ws.getCell("A1").value = text;
    ws.getCell("A1").font = { name: FONT, size: 14, bold: true };
    if (note) { ws.getCell("A2").value = note; ws.getCell("A2").font = { name: FONT, size: 9, color: { argb: "FF595959" } }; }
  }
  function subtitle(ws, row, text) {
    const c = ws.getCell(`A${row}`);
    c.value = text; c.font = { name: FONT, size: 11, bold: true, color: { argb: NAVY } };
  }
  function fill(row, from, to, argb) {
    for (let c = from; c <= to; c++) row.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
  }
  function applyFont(ws) {
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, c => {
      c.font = { size: 10, ...(c.font || {}), name: FONT };
    }));
  }
  const link = url => url ? { text: "원문", hyperlink: url } : "";

  async function buildAnalysisWorkbook(A, meta, scope = {}) {
    const ExcelJS = await loadLib("ExcelJS");
    const wb = new ExcelJS.Workbook();
    wb.creator = "나라장터 AI 사업 모니터";
    wb.created = new Date();
    wb.calcProperties.fullCalcOnLoad = true;   // 엑셀이 열 때 수식을 다시 계산
    const { R, home, rows } = A;
    const last = Math.max(2, rows.length + 1);
    const col = c => `'공고목록'!$${c}$2:$${c}$${last}`;
    const C = { st: col("A"), demand: col("F"), region: col("G"), kind: col("H"), field: col("I"),
      budget: col("N"), winner: col("P"), amt: col("Q"), rate: col("R"), exec: col("U"), excl: col("V") };
    const updated = meta && meta.updated ? String(meta.updated).replace("T", " ").slice(0, 16) : "-";
    const rr = (ctx && ctx.getRateRule && ctx.getRateRule()) || { lo: 50, hi: 150 };

    // 1. 안내
    const s1 = wb.addWorksheet("안내", { properties: { tabColor: { argb: NAVY } } });
    title(s1, "나라장터 AI 사업 분석자료", `다운로드 ${ymd(new Date())} · 내부 검토용 · 자동 생성`);
    const fieldRuleText = (R.fieldRules || []).map(([l, p]) => `${l}: ${p}`).join("\n") ||
      R.fields.join(", ");
    const guide = [
      ["자료 범위", scope.text || "전체 데이터 (필터 없음)"],
      ["자료 기준일", `마지막 수집 ${updated} (공고일 ${ymd(A.from)} ~ ${ymd(A.to)})`],
      ["출처", "조달청 나라장터 OpenAPI(공공데이터포털) — 입찰공고정보서비스, 낙찰정보서비스. 매일 1회 자동 수집"],
      ["수집 키워드", `공고명에 다음 단어가 들어간 용역·물품 공고: ${(meta && meta.keywords || []).join(", ")}`],
      ["제외 기준", `취소공고, AI가 인공지능이 아닌 공고(정규식): ${(R.excludePatterns || []).join(" / ") || "-"}`],
      ["대상 건수", `${rows.length.toLocaleString("ko-KR")}건 (중복 정리 전 원 공고 ${A.originalCount.toLocaleString("ko-KR")}건)`],
      ["중복 정리 기준", R.dedup + ". 묶인 건수는 '공고목록'의 '묶인 공고 수' 열"],
      ["배정예산", "발주기관이 공고에 적은 사업 예산(원, 부가세 포함). 미기재·0원은 평균·중앙값 등 통계에서 제외"],
      ["추정가격", "공고에 적힌 추정가격(원, 부가세 제외)"],
      ["낙찰률(예정가격 대비)", "낙찰금액 ÷ 예정가격 × 100 (나라장터 낙찰정보 제공값, %). 낙찰 정보가 연결된 공고만 표시. 다른 시트의 '낙찰률'도 이 값"],
      ["예산 집행률", "낙찰금액 ÷ 배정예산 × 100 (%) — 배정한 예산 중 실제로 낙찰된 비율(예산 편성 근거용). '공고목록' U열"],
      ["통계 제외", `규모구간별통계에서 평균·중앙값을 낼 때 빼는 낙찰 건('공고목록' V열): 단가계약 / 배정예산 미기재·0원 / 예산 집행률 ${rr.lo}% 미만 또는 ${rr.hi}% 초과(데이터 오류 가능성). 기준 숫자는 config.json의 rate_stats`],
      ["상태", "진행중: 마감 전 / 결과대기: 마감 후 낙찰 정보 없음 / 낙찰: 낙찰 정보 연결됨"],
      ["지역(추정)", "수요기관명으로 추정한 시도. 시도 정식명칭으로 시작하면 그 시도, 아니면 기관명에서 처음 나오는 지명. 중앙부처·전국 단위 기관은 '중앙·기타'"],
      ["기관구분", "수요기관명 규칙으로 추정: " + A.kinds.join(" / ") + " ('기타'는 중앙행정기관·군·민간업체 등)"],
      ["분야(자동분류)", "공고명 키워드 규칙으로 자동 분류(위에서부터 먼저 걸리는 것). 해당 없으면 물품은 '장비·물품 구매', 나머지는 '기타'\n" + fieldRuleText],
      ["주의사항", "· 공고명에 키워드가 없으면 수집되지 않아 실제 사업보다 적을 수 있습니다.\n· 지역·기관구분·분야는 규칙 기반 추정값이라 오분류가 있을 수 있습니다. 보고 전 원문을 확인하세요.\n· '단가계약(통계제외)' 분야는 사업비 통계에서 뺍니다.\n· 시트 3~6의 집계는 '공고목록' 시트를 참조하는 수식입니다. 공고목록을 고치면 집계도 바뀝니다.\n· 업체현황은 공개된 낙찰 정보를 정리한 내부 참고용이며 업체 평가 자료가 아닙니다."],
      ["시트 구성", "공고목록(자료 범위의 공고) / 시도별비교 / 도내현황(" + home + ") / 사업비근거(분야별 예산 분포·사례) / 업체현황(낙찰 상위 40개사) / 규모구간별통계(1억 미만·1~5억·5~10억·10억 이상, 예산 집행률·낙찰률(예정가격 대비))"],
    ];
    guide.forEach(([k, v], i) => {
      const r = s1.getRow(4 + i);
      r.getCell(1).value = k; r.getCell(2).value = v;
      r.getCell(1).font = { name: FONT, size: 10, bold: true };
      r.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY } };
      [1, 2].forEach(c => { r.getCell(c).border = BORDER; r.getCell(c).alignment = { vertical: "top", wrapText: true }; });
    });
    s1.columns = [{ width: 18 }, { width: 110 }];

    // 2. 공고목록
    const s2 = wb.addWorksheet("공고목록", { views: [{ state: "frozen", xSplit: 0, ySplit: 1 }] });
    const H2 = ["상태", "공고번호", "유형", "공고명", "공고기관", "수요기관", "지역(추정)", "기관구분", "분야", "계약방법",
      "공고일", "마감일", "공고월", "배정예산(원)", "추정가격(원)", "낙찰업체", "낙찰금액(원)", "낙찰률(예정가격 대비, %)", "묶인 공고 수", "원문링크",
      "예산 집행률(%)", "통계 제외 사유"];
    table(s2, 1, H2, rows.map(b => [
      b._st, b.no, b.type, b.title, b.org, b.demand, b.region, b._kind, b._field, b.method || "",
      xlDate(b._notice), xlDate(b._close), b._notice ? `${b._notice.getFullYear()}-${pad(b._notice.getMonth() + 1)}` : "",
      b.budget ?? null, b.estPrice ?? null, b.winner || "", b.awardAmt ?? null, b.rate ?? null, b.dups || 1, link(b.url),
      pos(b._exec) ? b._exec : null, b._excl || "",
    ]), { fmt: { 10: DATE, 11: DATE, 13: NUM, 14: NUM, 16: NUM, 17: PCT, 20: PCT } });
    s2.autoFilter = { from: "A1", to: `V${last}` };
    s2.columns.forEach((c, i) => { c.width = [8, 16, 6, 50, 24, 26, 10, 14, 16, 14, 11, 11, 9, 15, 15, 22, 15, 11, 8, 8, 11, 13][i]; });

    // 3. 시도별비교
    const s3 = wb.addWorksheet("시도별비교", { views: [{ state: "frozen", ySplit: 4 }] });
    title(s3, "시도별 AI 사업 비교", `공고목록 기준 수식. 금액 단위 백만원 · 건당 평균은 배정예산 기재 공고 기준 · 순위는 ${A.sidos.length}개 시도 대상`);
    const H3 = ["시도", "공고 건수", "배정예산 합계(백만원)", "건당 평균(백만원)", "낙찰 건수", "낙찰금액 합계(백만원)", "평균 낙찰률(%)", "예산 순위", "건수 순위"];
    const top3 = 4, first3 = top3 + 1, lastSido = first3 + A.sidos.length - 1;
    const r3 = A.regionStats.map((s, i) => {
      const r = first3 + i, a = `$A${r}`;
      return [s.name,
        F(`COUNTIFS(${C.region},${a})`, s.n),
        F(`SUMIFS(${C.budget},${C.region},${a})/1000000`, s.budget / 1e6),
        F(`IFERROR(AVERAGEIFS(${C.budget},${C.region},${a},${C.budget},">0")/1000000,0)`, (s.avgBudget || 0) / 1e6),
        F(`COUNTIFS(${C.region},${a},${C.st},"낙찰")`, s.awards),
        F(`SUMIFS(${C.amt},${C.region},${a})/1000000`, s.awardAmt / 1e6),
        F(`IFERROR(AVERAGEIFS(${C.rate},${C.region},${a},${C.rate},">0"),"")`, s.avgRate),
        s.isSido ? F(`RANK(C${r},$C$${first3}:$C$${lastSido})`, s.budgetRank) : "-",
        s.isSido ? F(`RANK(B${r},$B$${first3}:$B$${lastSido})`, s.countRank) : "-"];
    });
    const totR = first3 + r3.length, T = A.total;
    r3.push(["합계", F(`SUM(B${first3}:B${totR - 1})`, T.n), F(`SUM(C${first3}:C${totR - 1})`, T.budget / 1e6),
      F(`IFERROR(AVERAGEIFS(${C.budget},${C.budget},">0")/1000000,0)`, (T.avgBudget || 0) / 1e6),
      F(`SUM(E${first3}:E${totR - 1})`, T.awards), F(`SUM(F${first3}:F${totR - 1})`, T.awardAmt / 1e6),
      F(`IFERROR(AVERAGEIFS(${C.rate},${C.rate},">0"),"")`, T.avgRate), "", ""]);
    table(s3, top3, H3, r3, {
      fmt: { 1: NUM, 2: MIL, 3: MIL, 4: NUM, 5: MIL, 6: PCT, 7: NUM, 8: NUM },
      rowStyle: (row, i) => {
        const name = r3[i][0];
        if (name === home) { fill(row, 1, 9, YELLOW); row.getCell(1).font = { name: FONT, size: 10, bold: true }; }
        if (name === "합계") { fill(row, 1, 9, GRAY); row.eachCell(c => { c.font = { name: FONT, size: 10, bold: true }; }); }
      },
    });
    s3.columns.forEach((c, i) => { c.width = [12, 11, 16, 14, 11, 16, 13, 10, 10][i]; });

    // 4. 도내현황
    const s4 = wb.addWorksheet("도내현황", { properties: { tabColor: { argb: "FFFFC000" } } });
    title(s4, `${home} 도내 AI 사업 현황`, `지역(추정)이 '${home}'인 공고 · 기관구분별 소계는 공고목록 기준 수식`);
    subtitle(s4, 4, "1. 기관구분별 소계");
    const top4 = 5, first4 = top4 + 1;
    const H4a = ["기관구분", "공고 건수", "배정예산 합계(원)", "낙찰 건수", "진행중", "결과대기"];
    const hk = A.homeKinds.map((s, i) => {
      const a = `$A${first4 + i}`, reg = `${C.region},${q(home)},${C.kind},${a}`;
      return [s.name, F(`COUNTIFS(${reg})`, s.n), F(`SUMIFS(${C.budget},${reg})`, s.budget),
        F(`COUNTIFS(${reg},${C.st},"낙찰")`, s.awards), F(`COUNTIFS(${reg},${C.st},"진행중")`, s.open),
        F(`COUNTIFS(${reg},${C.st},"결과대기")`, s.closed)];
    });
    const hkTot = first4 + hk.length, HT = A.homeTotal;
    hk.push(["합계", ...["B", "C", "D", "E", "F"].map((L, j) =>
      F(`SUM(${L}${first4}:${L}${hkTot - 1})`, [HT.n, HT.budget, HT.awards, HT.open, HT.closed][j]))]);
    let next = table(s4, top4, H4a, hk, {
      fmt: { 1: NUM, 2: NUM, 3: NUM, 4: NUM, 5: NUM },
      rowStyle: (row, i) => { if (i === hk.length - 1) { fill(row, 1, 6, GRAY); row.eachCell(c => { c.font = { name: FONT, size: 10, bold: true }; }); } },
    });
    subtitle(s4, next + 1, "2. 사업 목록");
    s4.getCell(`A${next + 2}`).value = "중복 검토: 같은 분야를 도내 다른 수요기관도 발주했으면 '유사사업 있음' (유지보수·교육·물품·연구·기타 분야는 제외, 수식) · 노란 칸은 담당자 입력용";
    s4.getCell(`A${next + 2}`).font = { name: FONT, size: 9, color: { argb: "FF595959" } };
    const top4b = next + 3, first4b = top4b + 1;
    const H4b = ["No", "상태", "공고명", "수요기관", "기관구분", "분야", "공고일", "마감일", "배정예산(원)", "낙찰업체", "낙찰금액(원)", "낙찰률(%)", "묶인 공고 수", "중복 검토", "담당 부서 확인·비고", "원문링크"];
    const exc = R.similarExcludeFields.map(f => `$F{r}=${q(f)}`).join(",");
    const hl = A.homeRows.map((b, i) => {
      const r = first4b + i;
      const sim = `IF(OR(${exc.replace(/\{r\}/g, r)}),"",IF(COUNTIFS(${C.region},${q(home)},${C.field},$F${r})-COUNTIFS(${C.region},${q(home)},${C.field},$F${r},${C.demand},${crit(`$D${r}`, b.demand)})>0,"유사사업 있음",""))`;
      return [i + 1, b._st, b.title, b.demand, b._kind, b._field, xlDate(b._notice), xlDate(b._close), b.budget ?? null,
        b.winner || "", b.awardAmt ?? null, b.rate ?? null, b.dups || 1, F(sim, b._similar), "", link(b.url)];
    });
    if (!hl.length) hl.push(["", "", `${home} 지역 공고가 없습니다.`, "", "", "", "", "", "", "", "", "", "", "", "", ""]);
    table(s4, top4b, H4b, hl, {
      fmt: { 6: DATE, 7: DATE, 8: NUM, 10: NUM, 11: PCT },
      rowStyle: row => { row.getCell(15).fill = { type: "pattern", pattern: "solid", fgColor: { argb: INPUT } }; row.getCell(3).alignment = { wrapText: true, vertical: "top" }; },
    });
    s4.columns.forEach((c, i) => { c.width = [16, 9, 46, 24, 14, 16, 11, 11, 15, 20, 15, 9, 8, 13, 26, 8][i]; });

    // 5. 사업비근거
    const s5 = wb.addWorksheet("사업비근거");
    title(s5, "분야별 사업비 분포 (사업비 산정 근거)", `배정예산 0원·미기재 공고 제외 · '${R.statsExcludeFields.join(", ")}' 분야는 통계에서 제외 · 중앙값·최소·최대는 배열수식`);
    const H5 = ["분야", "공고 건수", "예산 기재 건수", "평균(원)", "중앙값(원)", "최소(원)", "최대(원)", "낙찰 건수", "평균 낙찰률(%)"];
    const top5 = 4, first5 = top5 + 1;
    const r5 = A.fieldStats.map((s, i) => {
      const r = first5 + i, a = `$A${r}`, cond = `(${C.field}=${a})*(${C.budget}>0)`;
      return [s.name, F(`COUNTIFS(${C.field},${a})`, s.n), F(`COUNTIFS(${C.field},${a},${C.budget},">0")`, s.budgetN),
        F(`IFERROR(AVERAGEIFS(${C.budget},${C.field},${a},${C.budget},">0"),"")`, s.avgBudget),
        AF(`IF($C${r}=0,"",MEDIAN(IF(${cond},${C.budget})))`, s.median, `E${r}`),
        AF(`IF($C${r}=0,"",MIN(IF(${cond},${C.budget})))`, s.min, `F${r}`),
        AF(`IF($C${r}=0,"",MAX(IF(${cond},${C.budget})))`, s.max, `G${r}`),
        F(`COUNTIFS(${C.field},${a},${C.st},"낙찰")`, s.awards),
        F(`IFERROR(AVERAGEIFS(${C.rate},${C.field},${a},${C.rate},">0"),"")`, s.avgRate)];
    });
    {
      const r = first5 + r5.length, FT = A.fieldTotal;
      const notEx = R.statsExcludeFields.map(f => `(${C.field}<>${q(f)})`).join("*") || "1";
      const exCount = R.statsExcludeFields.map(f => `-COUNTIFS(${C.field},${q(f)})`).join("");
      const exCountB = R.statsExcludeFields.map(f => `-COUNTIFS(${C.field},${q(f)},${C.budget},">0")`).join("");
      const exAw = R.statsExcludeFields.map(f => `-COUNTIFS(${C.field},${q(f)},${C.st},"낙찰")`).join("");
      const cond = `${notEx}*(${C.budget}>0)`;
      r5.push([FT.name,
        F(`COUNTA(${C.field})${exCount}`, FT.n),
        F(`COUNTIFS(${C.budget},">0")${exCountB}`, FT.budgetN),
        AF(`IF($C${r}=0,"",SUM(IF(${cond},${C.budget}))/$C${r})`, FT.avgBudget, `D${r}`),
        AF(`IF($C${r}=0,"",MEDIAN(IF(${cond},${C.budget})))`, FT.median, `E${r}`),
        AF(`IF($C${r}=0,"",MIN(IF(${cond},${C.budget})))`, FT.min, `F${r}`),
        AF(`IF($C${r}=0,"",MAX(IF(${cond},${C.budget})))`, FT.max, `G${r}`),
        F(`COUNTIFS(${C.st},"낙찰")${exAw}`, FT.awards),
        AF(`IFERROR(AVERAGE(IF(${notEx}*(${C.rate}>0),${C.rate})),"")`, FT.avgRate, `I${r}`)]);
    }
    next = table(s5, top5, H5, r5, {
      fmt: { 1: NUM, 2: NUM, 3: NUM, 4: NUM, 5: NUM, 6: NUM, 7: NUM, 8: PCT },
      rowStyle: (row, i) => { if (i === r5.length - 1) { fill(row, 1, 9, GRAY); row.eachCell(c => { c.font = { name: FONT, size: 10, bold: true }; }); } },
    });
    const H5c = ["No", "공고명", "수요기관", "지역(추정)", "기관구분", "공고일", "배정예산(원)", "낙찰금액(원)", "낙찰률(%)", "원문링크"];
    CASE_FIELDS.forEach(f => {
      const list = A.cases(f);
      subtitle(s5, next + 1, `사례: ${f} (배정예산순, ${list.length}건)`);
      const rowsC = list.map((b, i) => [i + 1, b.title, b.demand, b.region, b._kind, xlDate(b._notice), b.budget, b.awardAmt ?? null, b.rate ?? null, link(b.url)]);
      if (!rowsC.length) rowsC.push(["", "해당 분야 사례가 없습니다.", "", "", "", "", "", "", "", ""]);
      next = table(s5, next + 2, H5c, rowsC, { fmt: { 5: DATE, 6: NUM, 7: NUM, 8: PCT } });
    });
    s5.columns.forEach((c, i) => { c.width = [24, 46, 24, 14, 14, 11, 15, 15, 13, 8][i]; });

    // 6. 업체현황
    const s6 = wb.addWorksheet("업체현황");
    title(s6, "낙찰 업체 현황 (상위 40개사)", "※ 내부 참고용 — 나라장터에 공개된 낙찰 정보를 정리한 것이며 업체 평가·추천 자료가 아닙니다. 건수·금액·낙찰률·지자체·도내 건수는 공고목록 기준 수식");
    s6.getCell("A2").font = { name: FONT, size: 10, bold: true, color: { argb: "FFC00000" } };
    const H6 = ["순위", "업체명", "낙찰 건수", "낙찰금액 합계(원)", "평균 낙찰률(%)", "주요 분야", "주요 발주기관", "지자체 수행 건수", `${home} 수행 건수`];
    const top6 = 4, first6 = top6 + 1;
    const r6 = A.winners.map((w, i) => {
      const r = first6 + i, b = crit(`$B${r}`, w.name);
      return [i + 1, w.name, F(`COUNTIFS(${C.winner},${b})`, w.n), F(`SUMIFS(${C.amt},${C.winner},${b})`, w.amount),
        F(`IFERROR(AVERAGEIFS(${C.rate},${C.winner},${b},${C.rate},">0"),"")`, w.avgRate),
        w.topField ? `${w.topField[0]} (${w.topField[1]}건)` : "", w.topOrg ? `${w.topOrg[0]} (${w.topOrg[1]}건)` : "",
        F(`COUNTIFS(${C.winner},${b},${C.kind},"지자체")`, w.local), F(`COUNTIFS(${C.winner},${b},${C.region},${q(home)})`, w.home)];
    });
    if (!r6.length) r6.push(["", "낙찰 정보가 없습니다.", "", "", "", "", "", "", ""]);
    table(s6, top6, H6, r6, { fmt: { 2: NUM, 3: NUM, 4: PCT, 7: NUM, 8: NUM } });
    s6.columns.forEach((c, i) => { c.width = [6, 30, 10, 17, 13, 26, 34, 12, 12][i]; });

    // 7. 규모구간별 통계 (예산 집행률 · 낙찰률(예정가격 대비))
    const s7 = wb.addWorksheet("규모구간별통계");
    title(s7, "규모 구간별 평균 예산 집행률 · 낙찰률(예정가격 대비)", "공고목록 기준 수식 · 상태가 '낙찰'이고 통계 제외 사유(V열)가 없는 공고 · 구간은 배정예산 기준(하한 이상, 상한 미만) · 중앙값은 배열수식");
    const defs = [
      ["예산 집행률", "낙찰금액 ÷ 배정예산 × 100 — 배정한 예산 중 실제로 낙찰된 비율(예산 편성 근거용, 화면 기본 기준)"],
      ["낙찰률(예정가격 대비)", "낙찰금액 ÷ 예정가격 × 100 — 나라장터가 낙찰정보로 제공하는 공식 낙찰률(미제공 건은 계산에서 빠짐)"],
      ["통계 제외 기준", `단가계약 / 배정예산 미기재·0원 / 예산 집행률 ${rr.lo}% 미만 또는 ${rr.hi}% 초과(데이터 오류 가능성) — config.json의 rate_stats에서 변경`],
    ];
    defs.forEach(([k, v], i) => {
      const r = s7.getRow(4 + i);
      r.getCell(1).value = k; r.getCell(2).value = v;
      r.getCell(1).font = { name: FONT, size: 10, bold: true };
      r.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY } };
      s7.mergeCells(4 + i, 2, 4 + i, 13);
      [1, 2].forEach(c => { r.getCell(c).border = BORDER; r.getCell(c).alignment = { vertical: "top", wrapText: true }; });
    });
    const H7 = ["규모 구간(배정예산)", "하한(원)", "상한(원)", "통계 대상 건수(예산 집행률)", "평균 예산 집행률(%)", "중앙값 예산 집행률(%)",
      "통계 대상 건수(낙찰률)", "평균 낙찰률(예정가격 대비, %)", "중앙값 낙찰률(예정가격 대비, %)", "배정예산 합계(원)", "낙찰금액 합계(원)", "통계 제외 건수", "제외 사유"];
    const top7 = 8, first7 = top7 + 1;
    const exText = (extra) => A.EX.map((r, j) => `"${j ? " · " : ""}${r} "&COUNTIFS(${C.st},"낙찰",${C.excl},${q(r)}${extra})`).join("&");
    const exVal = c => A.EX.map(r => `${r} ${c[r]}`).join(" · ");
    const r7 = A.bandStats.map((s, i) => {
      const r = first7 + i;
      const band = `,${C.budget},">0",${C.budget},">="&$B${r}` + (s.hi == null ? "" : `,${C.budget},"<"&$C${r}`);
      const cond = `${C.st},"낙찰",${C.excl},""${band}`;
      const arr = `(${C.st}="낙찰")*(${C.excl}="")*(${C.budget}>0)*(${C.budget}>=$B${r})` + (s.hi == null ? "" : `*(${C.budget}<$C${r})`);
      return [s.name, s.lo, s.hi == null ? "" : s.hi,
        F(`COUNTIFS(${cond},${C.exec},">0")`, s.exec.n),
        F(`IFERROR(AVERAGEIFS(${C.exec},${cond},${C.exec},">0"),"")`, s.exec.avg),
        AF(`IF($D${r}=0,"",MEDIAN(IF(${arr}*(${C.exec}>0),${C.exec})))`, s.exec.med, `F${r}`),
        F(`COUNTIFS(${cond},${C.rate},">0")`, s.rate.n),
        F(`IFERROR(AVERAGEIFS(${C.rate},${cond},${C.rate},">0"),"")`, s.rate.avg),
        AF(`IF($G${r}=0,"",MEDIAN(IF(${arr}*(${C.rate}>0),${C.rate})))`, s.rate.med, `I${r}`),
        F(`SUMIFS(${C.budget},${cond},${C.exec},">0")`, s.budget),
        F(`SUMIFS(${C.amt},${cond},${C.exec},">0")`, s.awardAmt),
        F(`COUNTIFS(${C.st},"낙찰",${C.excl},"?*"${band})`, s.exN),
        F(exText(band), exVal(s.ex))];
    });
    {
      const r = first7 + r7.length, T7 = A.bandTotal;
      const nb = `,${C.budget},"<=0"`;
      r7.push(["구간 없음(배정예산 미기재·0원)", "", "", 0, "", "", 0, "", "", "", "",
        F(`COUNTIFS(${C.st},"낙찰",${C.excl},"?*"${nb})+COUNTIFS(${C.st},"낙찰",${C.excl},"?*",${C.budget},"")`, T7.noBudgetN),
        exVal(T7.noBudget)]);
      const tc = `${C.st},"낙찰",${C.excl},""`, arr = `(${C.st}="낙찰")*(${C.excl}="")`;
      r7.push(["전체", "", "",
        F(`COUNTIFS(${tc},${C.exec},">0")`, T7.exec.n),
        F(`IFERROR(AVERAGEIFS(${C.exec},${tc},${C.exec},">0"),"")`, T7.exec.avg),
        AF(`IF($D${r + 1}=0,"",MEDIAN(IF(${arr}*(${C.exec}>0),${C.exec})))`, T7.exec.med, `F${r + 1}`),
        F(`COUNTIFS(${tc},${C.rate},">0")`, T7.rate.n),
        F(`IFERROR(AVERAGEIFS(${C.rate},${tc},${C.rate},">0"),"")`, T7.rate.avg),
        AF(`IF($G${r + 1}=0,"",MEDIAN(IF(${arr}*(${C.rate}>0),${C.rate})))`, T7.rate.med, `I${r + 1}`),
        F(`SUMIFS(${C.budget},${tc},${C.exec},">0")`, T7.budget),
        F(`SUMIFS(${C.amt},${tc},${C.exec},">0")`, T7.awardAmt),
        F(`COUNTIFS(${C.st},"낙찰",${C.excl},"?*")`, T7.exN),
        F(exText(""), exVal(T7.ex))]);
    }
    table(s7, top7, H7, r7, {
      fmt: { 1: NUM, 2: NUM, 3: NUM, 4: "0.0", 5: "0.0", 6: NUM, 7: "0.0", 8: "0.0", 9: NUM, 10: NUM, 11: NUM },
      rowStyle: (row, i) => { if (i === r7.length - 1) { fill(row, 1, 13, GRAY); row.eachCell(c => { c.font = { name: FONT, size: 10, bold: true }; }); } },
    });
    s7.columns.forEach((c, i) => { c.width = [26, 14, 14, 12, 12, 12, 12, 14, 14, 17, 17, 10, 40][i]; });

    [s1, s2, s3, s4, s5, s6, s7].forEach(applyFont);
    return wb;
  }

  // ───────── 보고서 초안 (.docx) ─────────
  // 한글(HWP)에서도 열리도록 단순 구조만 사용: 굵은 글씨 문단, 고정폭(DXA) 표, 맑은 고딕
  const mil = n => n == null ? "-" : (n / 1e6).toLocaleString("ko-KR", { maximumFractionDigits: 1 });
  const pct1 = n => n == null ? "-" : n.toFixed(1);
  const periodText = A => A.from ? `최근 약 ${A.months}개월(${ymd(A.from)}~${ymd(A.to)})간` : "해당 기간";
  const fileSafe = s => String(s).replace(/[\\/:*?"<>|\s]+/g, "_");
  // 받침 유무로 조사 선택: josa("서울", "이", "가") → "서울이"
  const josa = (word, withB, without) => {
    const ch = String(word).replace(/[^가-힣0-9]/g, "").slice(-1);
    const code = ch ? ch.charCodeAt(0) - 0xAC00 : -1;
    const has = code >= 0 && code <= 11171 ? code % 28 !== 0 : /[013678]$/.test(ch);
    return word + (has ? withB : without);
  };
  const pp = (word, withB, without) => josa(word, withB, without).slice(String(word).length);   // 조사만

  function docKit(D) {
    const FONTS = { ascii: FONT, eastAsia: FONT, hAnsi: FONT, cs: FONT };
    const run = (text, o = {}) => new D.TextRun({ text: String(text ?? ""), font: FONTS, size: o.size || 21, bold: !!o.bold, color: o.color, italics: !!o.italics });
    const P = (text, o = {}) => new D.Paragraph({
      children: Array.isArray(text) ? text : [run(text, o)],
      alignment: o.align === "center" ? D.AlignmentType.CENTER : o.align === "right" ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT,
      spacing: { before: o.before ?? 0, after: o.after ?? 120, line: 300 }, indent: o.indent ? { left: o.indent } : undefined,
    });
    const H = (text, lvl = 1) => P(text, { bold: true, size: lvl === 1 ? 26 : 22, before: lvl === 1 ? 300 : 160, after: 120, color: lvl === 1 ? "1F3864" : undefined });
    const B = (text, o = {}) => P([run("○ ", o), run(text, o)], { indent: 200, ...o });
    const cell = (text, w, o = {}) => new D.TableCell({
      width: { size: w, type: D.WidthType.DXA },
      shading: o.fill ? { type: D.ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
      margins: { top: 50, bottom: 50, left: 80, right: 80 },
      verticalAlign: D.VerticalAlign.CENTER,
      children: String(text ?? "").split("\n").map(line => new D.Paragraph({
        alignment: o.align === "right" ? D.AlignmentType.RIGHT : o.align === "center" ? D.AlignmentType.CENTER : D.AlignmentType.LEFT,
        children: [run(line, { size: o.size || 18, bold: o.bold, color: o.color })],
      })),
    });
    const line = { style: D.BorderStyle.SINGLE, size: 4, color: "808080" };
    const borders = { top: line, bottom: line, left: line, right: line, insideHorizontal: line, insideVertical: line };
    // cols: [{ h: 머리글, w: 폭(twip), align }]
    const table = (cols, rows, o = {}) => new D.Table({
      width: { size: cols.reduce((s, c) => s + c.w, 0), type: D.WidthType.DXA },
      columnWidths: cols.map(c => c.w), layout: D.TableLayoutType.FIXED, borders,
      rows: [
        new D.TableRow({ tableHeader: true, children: cols.map(c => cell(c.h, c.w, { bold: true, fill: "1F3864", color: "FFFFFF", align: "center" })) }),
        ...rows.map((r, i) => new D.TableRow({
          children: r.map((v, j) => cell(v, cols[j].w, { align: cols[j].align, fill: o.highlight && o.highlight(i) ? "FFF2CC" : (o.bold && o.bold(i) ? "F2F2F2" : undefined), bold: o.bold && o.bold(i) })),
        })),
      ],
    });
    const box = (lines = 6, hint = "", filled = null) => new D.Table({
      width: { size: 9600, type: D.WidthType.DXA }, columnWidths: [9600], layout: D.TableLayoutType.FIXED, borders,
      rows: [new D.TableRow({ children: [new D.TableCell({
        width: { size: 9600, type: D.WidthType.DXA }, shading: { type: D.ShadingType.CLEAR, color: "auto", fill: "FFFFF0" },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [P(hint, { size: 18, color: "808080", italics: true }),
          ...(filled && filled.length ? filled : Array.from({ length: lines }, () => "")).map(t => P("○ " + t, { size: 21 }))],
      })] })],
    });
    const doc = children => new D.Document({
      creator: "나라장터 AI 사업 모니터", title: "보고서 초안",
      styles: { default: { document: { run: { font: FONTS, size: 21 } } } },
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children }],
    });
    return { run, P, H, B, table, box, doc, toBlob: d => D.Packer.toBlob(d), ext: "docx" };
  }

  // 범위·AI 안내 줄 (제목 아래)
  const scopeLine = (K, x) => x.scopeText ? K.P(`적용 조건: ${x.scopeText}`, { size: 18, align: "center", color: "1F3864", after: 200 }) : null;
  const AI_HINT = "※ AI(Claude) 초안 — 숫자는 위 표에서 코드가 계산한 값만 썼습니다. 내용은 담당자가 반드시 검토·수정하세요.";

  function homeReport(K, A, meta, x = {}) {
    const { home, total } = A;
    const hs = A.regionStats.find(s => s.name === home) || { n: 0, budget: 0, countRank: "-", budgetRank: "-" };
    const ht = A.homeTotal;
    const byCount = [...A.sidoStats].sort((a, b) => a.countRank - b.countRank)[0];
    const byBudget = [...A.sidoStats].sort((a, b) => a.budgetRank - b.budgetRank)[0];
    const share = (a, b) => b ? (a / b * 100).toFixed(1) : "0.0";
    const updated = meta && meta.updated ? String(meta.updated).slice(0, 10).replace(/-/g, ".") : ymd(new Date());
    const c = [];
    c.push(K.P(`${home} 도내 AI 사업 추진현황`, { bold: true, size: 32, align: "center", after: 80 }));
    c.push(K.P(`(자료: 조달청 나라장터 공고 자동 수집 · 기준일 ${updated} · 초안)`, { size: 18, align: "center", color: "595959", after: x.scopeText ? 60 : 240 }));
    const sl = scopeLine(K, x); if (sl) c.push(sl);

    c.push(K.H("1. 개요"));
    c.push(K.B(`${periodText(A)} 도내 기관이 나라장터에 낸 AI 관련 공고는 ${ht.n.toLocaleString("ko-KR")}건, 배정예산 ${eok(ht.budget)}억원으로 ${A.sidos.length}개 시도 중 건수 ${hs.countRank}위, 예산 ${hs.budgetRank}위입니다.`));
    c.push(K.B(`같은 기간 전체 공고(중앙·기타 포함)는 ${total.n.toLocaleString("ko-KR")}건, 배정예산 ${eok(total.budget)}억원이며, 도내 공고는 건수의 ${share(ht.n, total.n)}%, 예산의 ${share(ht.budget, total.budget)}%입니다.`));
    c.push(K.B(`도내 공고 ${ht.n}건 중 낙찰 ${ht.awards}건, 진행중 ${ht.open}건, 결과대기 ${ht.closed}건입니다.`));

    c.push(K.H("2. 도내 발주 현황"));
    c.push(K.H("가. 기관구분별", 2));
    c.push(K.B("기관구분별로는 " + A.homeKinds.filter(k => k.n).map(k => `${k.name} ${k.n}건(${eok(k.budget)}억원)`).join(", ") + (A.homeKinds.some(k => k.n) ? "입니다." : "공고가 없습니다.")));
    c.push(K.table([
      { h: "기관구분", w: 2600 }, { h: "공고 건수", w: 1300, align: "right" }, { h: "배정예산(백만원)", w: 1900, align: "right" },
      { h: "낙찰", w: 1200, align: "right" }, { h: "진행중", w: 1300, align: "right" }, { h: "결과대기", w: 1300, align: "right" },
    ], [...A.homeKinds, { ...ht, name: "합계" }].map(k => [k.name, k.n, mil(k.budget), k.awards, k.open, k.closed]),
    { bold: i => i === A.homeKinds.length }));
    c.push(K.H("나. 사업 목록", 2));
    c.push(K.P(`공고일 최근순 · 금액 단위 백만원 · 재공고 등 중복 공고는 1건으로 정리`, { size: 18, color: "595959" }));
    c.push(K.table([
      { h: "No", w: 450, align: "center" }, { h: "공고명", w: 2700 }, { h: "수요기관", w: 1650 }, { h: "분야", w: 1300 },
      { h: "공고일", w: 1300, align: "center" }, { h: "배정예산", w: 900, align: "right" }, { h: "상태(낙찰업체)", w: 1300 },
    ], A.homeRows.length ? A.homeRows.map((b, i) => [i + 1, b.title, b.demand, b._field, ymd(b._notice), mil(b.budget), b._st + (b.winner ? `\n(${b.winner})` : "")])
      : [["", "도내 공고가 없습니다.", "", "", "", "", ""]]));

    c.push(K.H("3. 타 시도 비교"));
    c.push(K.B(`${A.sidos.length}개 시도 중 공고 건수는 ${byCount.name}(${byCount.n}건)${pp(byCount.name, "이", "가")} 가장 많고, 배정예산은 ${byBudget.name}(${eok(byBudget.budget)}억원)${pp(byBudget.name, "이", "가")} 가장 큽니다. ${josa(home, "은", "는")} 건수 ${hs.countRank}위, 예산 ${hs.budgetRank}위입니다.`));
    c.push(K.P("배정예산 순 · 지역은 수요기관명 기준 추정 · 중앙부처·전국 단위 기관(중앙·기타)은 순위에서 제외", { size: 18, color: "595959" }));
    const ranked = [...A.sidoStats].sort((a, b) => a.budgetRank - b.budgetRank || a.countRank - b.countRank);
    c.push(K.table([
      { h: "예산 순위", w: 1000, align: "center" }, { h: "시도", w: 1300, align: "center" }, { h: "배정예산(억원)", w: 1600, align: "right" },
      { h: "공고 건수", w: 1300, align: "right" }, { h: "건수 순위", w: 1100, align: "center" }, { h: "건당 평균(백만원)", w: 1800, align: "right" },
      { h: "낙찰 건수", w: 1500, align: "right" },
    ], ranked.map(s => [s.budgetRank, s.name, eok(s.budget), s.n, s.countRank, mil(s.avgBudget), s.awards]),
    { highlight: i => ranked[i].name === home }));

    c.push(K.H("4. 시사점"));
    c.push(x.ai && x.ai.length ? K.box(0, AI_HINT, x.ai)
      : K.box(6, "※ 담당자 작성 — 위 현황과 타 시도 비교를 바탕으로 직접 작성하세요. (자동 생성하지 않음)"));

    c.push(K.H("5. 자료 한계"));
    [
      `공고명에 수집 키워드(${(meta && meta.keywords || []).join(", ")})가 들어간 용역·물품 공고만 집계하여 실제 사업 수보다 적을 수 있습니다.`,
      "지역·기관구분·분야는 기관명·공고명 규칙으로 추정한 값이라 일부 오분류가 있을 수 있습니다.",
      "재공고·정정공고는 공고명과 수요기관이 같으면 1건으로 묶었습니다.",
      "배정예산은 공고 기재 금액이며 실제 계약금액과 다를 수 있습니다. 낙찰 정보는 나라장터 낙찰 API에 연결된 건만 반영됩니다.",
      "조달청 나라장터를 거치지 않은 자체 계약·수의계약 등은 포함되지 않습니다.",
    ].forEach(t => c.push(K.B(t, { size: 19 })));
    return K.doc(c);
  }

  function costReport(K, A, meta, field, x = {}) {
    const fs = A.fieldStats.find(s => s.name === field) || A.fieldTotal;
    const list = A.cases(field);
    const LIMIT = 30;
    const updated = meta && meta.updated ? String(meta.updated).slice(0, 10).replace(/-/g, ".") : ymd(new Date());
    const c = [];
    c.push(K.P("○○사업 사업비 산정 근거", { bold: true, size: 32, align: "center", after: 80 }));
    c.push(K.P(`(참고 분야: ${field} · 자료: 조달청 나라장터 공고 자동 수집 · 기준일 ${updated} · 초안)`, { size: 18, align: "center", color: "595959", after: x.scopeText ? 60 : 240 }));
    const sl = scopeLine(K, x); if (sl) c.push(sl);

    c.push(K.H("1. 개요"));
    c.push(K.B(`${periodText(A)} 나라장터에 공고된 AI 관련 사업 중 '${field}' 분야는 ${fs.n}건이며, 배정예산이 기재된 ${fs.budgetN}건의 평균은 ${mil(fs.avgBudget)}백만원, 중앙값은 ${mil(fs.median)}백만원(최소 ${mil(fs.min)}백만원 ~ 최대 ${mil(fs.max)}백만원)입니다.`));
    c.push(K.B(fs.awards ? `이 중 낙찰은 ${fs.awards}건이며, 낙찰률이 공개된 ${fs.rateN}건의 평균 낙찰률은 ${pct1(fs.avgRate)}%입니다.` : "이 분야에 연결된 낙찰 정보는 아직 없습니다."));
    c.push(K.B(`전체 AI 관련 공고(통계제외 분야 제외) ${A.fieldTotal.n}건의 배정예산 중앙값은 ${mil(A.fieldTotal.median)}백만원입니다.`));

    c.push(K.H("2. 분야별 사업비 분포"));
    c.push(K.P(`금액 단위 백만원 · 배정예산 0원·미기재 제외 · '${A.R.statsExcludeFields.join(", ")}' 분야 제외 · 분야는 공고명 키워드 자동 분류`, { size: 18, color: "595959" }));
    const all = [...A.fieldStats, A.fieldTotal];
    c.push(K.table([
      { h: "분야", w: 2000 }, { h: "공고 건수", w: 900, align: "right" }, { h: "예산 기재", w: 900, align: "right" },
      { h: "평균", w: 1100, align: "right" }, { h: "중앙값", w: 1100, align: "right" }, { h: "최소", w: 1000, align: "right" },
      { h: "최대", w: 1200, align: "right" }, { h: "평균 낙찰률(%)", w: 1300, align: "right" },
    ], all.map(s => [s.name === A.fieldTotal.name ? "전체" : s.name, s.n, s.budgetN, mil(s.avgBudget), mil(s.median), mil(s.min), mil(s.max), pct1(s.avgRate)]),
    { highlight: i => all[i].name === field, bold: i => i === all.length - 1 }));

    c.push(K.H(`3. '${field}' 분야 사례 (배정예산순)`));
    c.push(K.P(list.length > LIMIT ? `상위 ${LIMIT}건 (전체 ${list.length}건은 분석 엑셀의 공고목록 참고) · 금액 단위 백만원` : `${list.length}건 · 금액 단위 백만원`, { size: 18, color: "595959" }));
    c.push(K.table([
      { h: "No", w: 450, align: "center" }, { h: "공고명", w: 2600 }, { h: "수요기관", w: 1600 }, { h: "지역", w: 950, align: "center" },
      { h: "공고일", w: 1300, align: "center" }, { h: "배정예산", w: 900, align: "right" }, { h: "낙찰금액", w: 900, align: "right" },
      { h: "낙찰률(%)", w: 900, align: "right" },
    ], list.length ? list.slice(0, LIMIT).map((b, i) => [i + 1, b.title, b.demand, b.region, ymd(b._notice), mil(b.budget), b.awardAmt ? mil(b.awardAmt) : "-", b.rate ? b.rate.toFixed(1) : "-"])
      : [["", "해당 분야 사례가 없습니다.", "", "", "", "", "", ""]]));

    c.push(K.H("4. 적정 범위(안)"));
    c.push(K.P("※ 담당자 작성 — 위 분포와 사례를 참고해 사업 범위·규모에 맞게 직접 작성하세요. (자동 생성하지 않음)", { size: 18, color: "808080", italics: true }));
    c.push(K.table([{ h: "구분", w: 1600, align: "center" }, { h: "금액(백만원)", w: 2000, align: "right" }, { h: "산정 근거·비고", w: 6000 }],
      [["하한", "", ""], ["적정", "", ""], ["상한", "", ""]]));
    c.push(x.ai && x.ai.length ? K.box(0, AI_HINT, x.ai) : K.box(4, "※ 담당자 작성 — 사업 범위 비교, 단가·인력 산정 등 추가 근거"));

    c.push(K.H("5. 유의사항"));
    [
      "분야는 공고명 키워드 규칙으로 자동 분류한 결과로, 사업 내용이 다른 공고가 섞여 있을 수 있습니다. 사례는 원문을 확인한 뒤 인용하세요.",
      "배정예산은 공고 기재 금액(부가세 포함)이며 실제 계약금액과 다를 수 있습니다. 낙찰률은 나라장터 제공값입니다.",
      "사업 기간·범위·포함 장비 등이 달라 금액을 그대로 비교하기 어렵습니다. 비슷한 범위의 사례를 골라 비교하세요.",
      "단가계약은 통계에서 제외했습니다. 재공고·정정공고는 1건으로 묶었습니다.",
    ].forEach(t => c.push(K.B(t, { size: 19 })));
    return K.doc(c);
  }

  // ───────── 범위 (현재 필터 / 전체 데이터) ─────────
  // 도내현황 보고서는 '도내 vs 타 시도' 비교가 목적이라 지역 필터만 빼고 거름(noRegion)
  function scopeRows(scope, kind) {
    if (scope === "all" || !ctx.getRows) return { rows: ctx.getItems(), text: "전체 데이터 (필터 없음)" };
    const noRegion = kind === "home";
    const f = (ctx.filterSummary && ctx.filterSummary({ noRegion })) || "";
    return { rows: ctx.getRows({ noRegion }),
      text: (f ? `현재 화면 필터 — ${f}` : "현재 화면 필터 (조건 없음 = 전체)") +
        (noRegion && ctx.regionFilter && ctx.regionFilter() ? ` · 지역 필터(${ctx.regionFilter()})는 시도 비교를 위해 적용하지 않음` : "") };
  }

  // AI 에 넘길 숫자 묶음: 보고서 표에 들어간 값과 같은 값만 (AI 는 문장만 씀)
  function aiPayload(kind, A, field) {
    const m = v => v == null ? null : Math.round(v / 1e6 * 10) / 10;   // 백만원
    if (kind === "home") {
      const hs = A.regionStats.find(s => s.name === A.home) || {};
      const fieldCnt = {}; A.homeRows.forEach(b => { fieldCnt[b._field] = (fieldCnt[b._field] || 0) + 1; });
      return {
        보고서: `${A.home} 도내 AI 사업 추진현황`, 기간: periodText(A), 도: A.home,
        도내: { 공고: A.homeTotal.n, 배정예산_백만원: m(A.homeTotal.budget), 낙찰: A.homeTotal.awards, 진행중: A.homeTotal.open, 결과대기: A.homeTotal.closed,
          건수순위: hs.countRank, 예산순위: hs.budgetRank, 시도수: A.sidos.length },
        전체: { 공고: A.total.n, 배정예산_백만원: m(A.total.budget) },
        도내_기관구분별: A.homeKinds.filter(k => k.n).map(k => ({ 구분: k.name, 공고: k.n, 배정예산_백만원: m(k.budget) })),
        도내_분야별_건수: fieldCnt,
        도내_사업_예시: A.homeRows.slice(0, 15).map(b => ({ 공고명: b.title, 기관: b.demand, 분야: b._field, 배정예산_백만원: m(b.budget), 상태: b._st })),
        시도별_예산상위: [...A.sidoStats].sort((a, b) => a.budgetRank - b.budgetRank).slice(0, 6)
          .map(s => ({ 시도: s.name, 공고: s.n, 배정예산_백만원: m(s.budget), 건당평균_백만원: m(s.avgBudget) })),
      };
    }
    const fs = A.fieldStats.find(s => s.name === field) || A.fieldTotal, list = A.cases(field);
    const sorted = list.map(b => b.budget).sort((a, b) => a - b);
    const qt = p => { if (!sorted.length) return null; const i = (sorted.length - 1) * p, lo = Math.floor(i);
      return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (i - lo); };
    return {
      보고서: "○○사업 사업비 산정 근거", 분야: field, 기간: periodText(A),
      분야통계: { 공고: fs.n, 예산기재: fs.budgetN, 평균_백만원: m(fs.avgBudget), 중앙값_백만원: m(fs.median), 하위25_백만원: m(qt(.25)), 상위25_백만원: m(qt(.75)),
        최소_백만원: m(fs.min), 최대_백만원: m(fs.max), 낙찰: fs.awards, 평균낙찰률: fs.avgRate == null ? null : +fs.avgRate.toFixed(1) },
      전체_중앙값_백만원: m(A.fieldTotal.median),
      규모구간별_예산집행률_중앙값: A.bandStats.map(s => ({ 구간: s.name, 건수: s.exec.n, 중앙값: s.exec.med == null ? null : +s.exec.med.toFixed(1) })),
      사례_상위: list.slice(0, 12).map(b => ({ 공고명: b.title, 기관: b.demand, 지역: b.region, 배정예산_백만원: m(b.budget),
        낙찰금액_백만원: m(b.awardAmt), 낙찰률: b.rate ? +b.rate.toFixed(1) : null })),
    };
  }

  // 보고서 만들기: opts = { scope: "filter"|"all", format: "hwpx"|"docx", ai: bool, onStatus }
  async function buildReport(kind, field, opts = {}) {
    const meta = ctx.getMeta();
    const sc = scopeRows(opts.scope || "filter", kind);
    const A = analyze(sc.rows, meta, ctx.statusLabel);
    if (kind === "cost" && !field) field = (A.fieldStats.filter(s => s.n).sort((a, b) => b.n - a.n)[0] || A.fieldStats[0] || {}).name || "기타";
    let ai = null, aiError = null;
    if (opts.ai && window.G2BAI && window.G2BAI.hasKey()) {
      try {
        opts.onStatus && opts.onStatus("AI가 시사점 작성 중…");
        ai = await window.G2BAI.reportInsights(kind, aiPayload(kind, A, field), sc.text);
      } catch (e) { aiError = e && e.message || String(e); ai = null; }
    }
    opts.onStatus && opts.onStatus("파일 만드는 중…");
    const K = (opts.format || "hwpx") === "docx" ? docKit(await loadLib("docx")) : window.G2BHwpx.kit();
    const x = { scopeText: sc.text, ai };
    const doc = kind === "cost" ? costReport(K, A, meta, field, x) : homeReport(K, A, meta, x);
    if (K.ext === "hwpx") doc.title = kind === "cost" ? `사업비 산정 근거 (${field})` : `${A.home} AI 사업 추진현황`;
    const blob = await K.toBlob(doc);
    const tag = sc.text.startsWith("전체") ? "" : "_필터";
    const name = (kind === "cost" ? `사업비_산정근거_${fileSafe(field)}` : `${A.home}_AI사업_추진현황`) + `${tag}_초안_${stamp()}.${K.ext}`;
    download(blob, name);
    return { name, n: sc.rows.length, ai: !!ai, aiError, scope: sc.text, field };
  }

  // ───────── 진입점 ─────────
  let ctx = null;   // { getItems, getMeta, statusLabel, getRows, filterSummary, regionFilter }
  async function withBusy(btn, label, fn) {
    const old = btn.innerHTML;
    btn.disabled = true; btn.textContent = label;
    try { return await fn(t => { btn.textContent = t; }); }
    catch (e) { console.error(e); alert("파일을 만들지 못했습니다: " + (e && e.message || e)); }
    finally { btn.disabled = false; btn.innerHTML = old; }
  }

  async function downloadAnalysis(scope = "filter") {
    const sc = scopeRows(scope, "analysis");
    const A = analyze(sc.rows, ctx.getMeta(), ctx.statusLabel);
    const wb = await buildAnalysisWorkbook(A, ctx.getMeta(), sc);
    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `나라장터_AI사업_분석자료${scope === "all" ? "_전체" : "_필터"}_${stamp()}.xlsx`);
    return { n: sc.rows.length, scope: sc.text };
  }

  function init(c) {
    ctx = c;
    const ab = document.getElementById("anaBtn"), abAll = document.getElementById("anaAllBtn");
    if (ab) ab.addEventListener("click", () => withBusy(ab, "생성 중…", () => downloadAnalysis("filter")));
    if (abAll) abAll.addEventListener("click", () => withBusy(abAll, "생성 중…", () => downloadAnalysis("all")));

    // 보고서 초안 대화상자
    const rb = document.getElementById("repBtn"), dlg = document.getElementById("repDlg");
    if (!rb || !dlg) return;
    const sel = dlg.querySelector("#repField");
    const R = { ...DEFAULT_RULES, ...(c.getMeta() && c.getMeta().rules || {}) };
    const val = name => (dlg.querySelector(`input[name=${name}]:checked`) || {}).value;
    const fillFields = () => {   // 분야 선택지: 고른 범위 기준 건수
      const keep = sel.value, counts = new Map();
      scopeRows(val("repScope"), "cost").rows.forEach(b => counts.set(b.field || "기타", (counts.get(b.field || "기타") || 0) + 1));
      const fields = R.fields.filter(f => !R.statsExcludeFields.includes(f));
      sel.innerHTML = fields.map(f => `<option value="${f.replace(/"/g, "&quot;")}">${f} (${counts.get(f) || 0}건)</option>`).join("");
      if (fields.includes(keep)) sel.value = keep;
      else { const best = fields.slice().sort((a, b) => (counts.get(b) || 0) - (counts.get(a) || 0))[0]; if (best) sel.value = best; }
    };
    const sync = () => {
      const cost = val("repKind") === "cost";
      sel.disabled = !cost; dlg.querySelector("#repFieldRow").classList.toggle("off", !cost);
      const sc = scopeRows(val("repScope"), val("repKind"));
      dlg.querySelector("#repScopeNote").textContent = `${sc.rows.length.toLocaleString("ko-KR")}건 · ${sc.text}`;
      const has = !!(window.G2BAI && window.G2BAI.hasKey()), aiBox = dlg.querySelector("#repAi");
      aiBox.disabled = !has; if (!has) aiBox.checked = false;
      dlg.querySelector("#repAiNote").textContent = has ? (cost ? "'적정 범위 근거' 칸을 AI 초안으로 채움" : "'시사점' 칸을 AI 초안으로 채움") + " · 숫자는 코드 계산값만 사용"
        : "Claude API 키를 넣으면 쓸 수 있어요 (AI 도우미 › ⚙ 설정). 키가 없으면 빈칸으로 둡니다";
    };
    dlg.querySelectorAll("input[name=repKind],input[name=repScope]").forEach(r => r.addEventListener("change", () => { fillFields(); sync(); }));
    rb.addEventListener("click", () => { fillFields(); sync(); if (dlg.showModal) dlg.showModal(); else dlg.setAttribute("open", ""); });
    dlg.querySelector("#repClose").addEventListener("click", () => dlg.close ? dlg.close() : dlg.removeAttribute("open"));
    const go = dlg.querySelector("#repGo");
    go.addEventListener("click", () => withBusy(go, "생성 중…", async status => {
      const r = await buildReport(val("repKind"), sel.value, { scope: val("repScope"), format: val("repFmt"), ai: dlg.querySelector("#repAi").checked, onStatus: status });
      if (r && r.aiError) alert("AI 초안은 만들지 못해 해당 칸을 빈칸으로 두었습니다.\n사유: " + r.aiError);
      if (dlg.close) dlg.close(); else dlg.removeAttribute("open");
    }));
  }

  let inited = false;
  const initOnce = c => { if (!inited) { inited = true; init(c); } };
  window.G2BReport = { init: initOnce, analyze, loadLib, buildReport, downloadAnalysis };
  if (window.__g2bReportCfg) initOnce(window.__g2bReportCfg);   // 데이터가 먼저 준비된 경우
})();
