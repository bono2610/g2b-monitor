/*
 * 나라장터 AI 사업 모니터 — 분석 엑셀 · 보고서 초안 (브라우저에서 생성, 서버 불필요)
 *
 * - 외부 LLM 없음: 모든 숫자는 bids.json 집계, 문장은 고정 템플릿에 숫자만 채움
 * - 분류(지역·기관구분·분야)는 수집기가 config.json 규칙으로 계산해 둔 값을 그대로 씀
 * - 라이브러리는 버튼을 누를 때만 CDN에서 불러옴 (ExcelJS, docx)
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
    regions: ["제주", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종", "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "중앙·기타"],
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
    return { R, home, rows, sidos, regionStats, sidoStats, total, homeRows, kinds, homeKinds, homeTotal,
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

  async function buildAnalysisWorkbook(A, meta) {
    const ExcelJS = await loadLib("ExcelJS");
    const wb = new ExcelJS.Workbook();
    wb.creator = "나라장터 AI 사업 모니터";
    wb.created = new Date();
    wb.calcProperties.fullCalcOnLoad = true;   // 엑셀이 열 때 수식을 다시 계산
    const { R, home, rows } = A;
    const last = Math.max(2, rows.length + 1);
    const col = c => `'공고목록'!$${c}$2:$${c}$${last}`;
    const C = { st: col("A"), demand: col("F"), region: col("G"), kind: col("H"), field: col("I"),
      budget: col("N"), winner: col("P"), amt: col("Q"), rate: col("R") };
    const updated = meta && meta.updated ? String(meta.updated).replace("T", " ").slice(0, 16) : "-";

    // 1. 안내
    const s1 = wb.addWorksheet("안내", { properties: { tabColor: { argb: NAVY } } });
    title(s1, "나라장터 AI 사업 분석자료", `다운로드 ${ymd(new Date())} · 내부 검토용 · 자동 생성`);
    const fieldRuleText = (R.fieldRules || []).map(([l, p]) => `${l}: ${p}`).join("\n") ||
      R.fields.join(", ");
    const guide = [
      ["자료 기준일", `마지막 수집 ${updated} (공고일 ${ymd(A.from)} ~ ${ymd(A.to)})`],
      ["출처", "조달청 나라장터 OpenAPI(공공데이터포털) — 입찰공고정보서비스, 낙찰정보서비스. 매일 1회 자동 수집"],
      ["수집 키워드", `공고명에 다음 단어가 들어간 용역·물품 공고: ${(meta && meta.keywords || []).join(", ")}`],
      ["제외 기준", `취소공고, AI가 인공지능이 아닌 공고(정규식): ${(R.excludePatterns || []).join(" / ") || "-"}`],
      ["대상 건수", `${rows.length.toLocaleString("ko-KR")}건 (중복 정리 전 원 공고 ${A.originalCount.toLocaleString("ko-KR")}건)`],
      ["중복 정리 기준", R.dedup + ". 묶인 건수는 '공고목록'의 '묶인 공고 수' 열"],
      ["배정예산", "발주기관이 공고에 적은 사업 예산(원, 부가세 포함). 미기재·0원은 평균·중앙값 등 통계에서 제외"],
      ["추정가격", "공고에 적힌 추정가격(원, 부가세 제외)"],
      ["낙찰률", "낙찰금액 ÷ 예정가격 × 100 (나라장터 낙찰정보 제공값, %). 낙찰 정보가 연결된 공고만 표시"],
      ["상태", "진행중: 마감 전 / 결과대기: 마감 후 낙찰 정보 없음 / 낙찰: 낙찰 정보 연결됨"],
      ["지역(추정)", "수요기관명으로 추정한 시도. 시도 정식명칭으로 시작하면 그 시도, 아니면 기관명에서 처음 나오는 지명. 중앙부처·전국 단위 기관은 '중앙·기타'"],
      ["기관구분", "수요기관명 규칙으로 추정: " + A.kinds.join(" / ") + " ('기타'는 중앙행정기관·군·민간업체 등)"],
      ["분야(자동분류)", "공고명 키워드 규칙으로 자동 분류(위에서부터 먼저 걸리는 것). 해당 없으면 물품은 '장비·물품 구매', 나머지는 '기타'\n" + fieldRuleText],
      ["주의사항", "· 공고명에 키워드가 없으면 수집되지 않아 실제 사업보다 적을 수 있습니다.\n· 지역·기관구분·분야는 규칙 기반 추정값이라 오분류가 있을 수 있습니다. 보고 전 원문을 확인하세요.\n· '단가계약(통계제외)' 분야는 사업비 통계에서 뺍니다.\n· 시트 3~6의 집계는 '공고목록' 시트를 참조하는 수식입니다. 공고목록을 고치면 집계도 바뀝니다.\n· 업체현황은 공개된 낙찰 정보를 정리한 내부 참고용이며 업체 평가 자료가 아닙니다."],
      ["시트 구성", "공고목록(전체 공고) / 시도별비교 / 도내현황(" + home + ") / 사업비근거(분야별 예산 분포·사례) / 업체현황(낙찰 상위 40개사)"],
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
      "공고일", "마감일", "공고월", "배정예산(원)", "추정가격(원)", "낙찰업체", "낙찰금액(원)", "낙찰률(%)", "묶인 공고 수", "원문링크"];
    table(s2, 1, H2, rows.map(b => [
      b._st, b.no, b.type, b.title, b.org, b.demand, b.region, b._kind, b._field, b.method || "",
      xlDate(b._notice), xlDate(b._close), b._notice ? `${b._notice.getFullYear()}-${pad(b._notice.getMonth() + 1)}` : "",
      b.budget ?? null, b.estPrice ?? null, b.winner || "", b.awardAmt ?? null, b.rate ?? null, b.dups || 1, link(b.url),
    ]), { fmt: { 10: DATE, 11: DATE, 13: NUM, 14: NUM, 16: NUM, 17: PCT } });
    s2.autoFilter = { from: "A1", to: `T${last}` };
    s2.columns.forEach((c, i) => { c.width = [8, 16, 6, 50, 24, 26, 10, 14, 16, 14, 11, 11, 9, 15, 15, 22, 15, 9, 8, 8][i]; });

    // 3. 시도별비교
    const s3 = wb.addWorksheet("시도별비교", { views: [{ state: "frozen", ySplit: 4 }] });
    title(s3, "시도별 AI 사업 비교", "공고목록 기준 수식. 금액 단위 백만원 · 건당 평균은 배정예산 기재 공고 기준 · 순위는 17개 시도 대상");
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

    [s1, s2, s3, s4, s5, s6].forEach(applyFont);
    return wb;
  }

  // ───────── 진입점 ─────────
  let ctx = null;   // { getItems, getMeta, statusLabel }
  async function withBusy(btn, label, fn) {
    const old = btn.textContent;
    btn.disabled = true; btn.textContent = label;
    try { await fn(); }
    catch (e) { console.error(e); alert("파일을 만들지 못했습니다: " + (e && e.message || e)); }
    finally { btn.disabled = false; btn.textContent = old; }
  }

  async function downloadAnalysis() {
    const A = analyze(ctx.getItems(), ctx.getMeta(), ctx.statusLabel);
    const wb = await buildAnalysisWorkbook(A, ctx.getMeta());
    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `나라장터_AI사업_분석자료_${stamp()}.xlsx`);
  }

  function init(c) {
    ctx = c;
    const ab = document.getElementById("anaBtn");
    if (ab) ab.addEventListener("click", () => withBusy(ab, "생성 중…", downloadAnalysis));
  }

  window.G2BReport = { init, analyze, loadLib };
})();
