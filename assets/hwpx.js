/*
 * 나라장터 AI 사업 모니터 — HWPX(한글 문서) 작성기 (브라우저에서 생성, 외부 라이브러리 없음)
 *
 * - report.js 의 보고서 함수가 쓰는 문서 도구(P·H·B·table·box·doc)와 같은 모양으로 만들어
 *   같은 보고서 내용을 .docx 와 .hwpx 로 모두 뽑을 수 있게 함
 * - HWPX = OWPML(KS X 6101) XML 을 ZIP 으로 묶은 파일. mimetype 은 압축하지 않고 맨 앞에 둠
 * - 단위: 보고서 함수는 워드 단위(twip = 1/1440인치, 글자 크기 = 반 포인트)로 값을 넘기고, 여기서 HWPUNIT(1/7200인치)로 바꿈
 */
(() => {
  "use strict";

  // ───────── ZIP (무압축 STORE) ─────────
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  const crc32 = u8 => { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const enc = new TextEncoder();

  function zip(files) {   // files: [[이름, 문자열]] → Uint8Array
    const d = new Date();
    const time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    const parts = [], central = [];
    let offset = 0;
    files.forEach(([name, text]) => {
      const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, time, true); lh.setUint16(12, date, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nb, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true); ch.setUint16(12, time, true); ch.setUint16(14, date, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, nb.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nb);
      offset += 30 + nb.length + data.length;
    });
    const cdSize = central.reduce((s, p) => s + p.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((s, p) => s + p.length, 0));
    let p = 0; all.forEach(a => { out.set(a, p); p += a.length; });
    return out;
  }

  // ───────── XML 도우미 ─────────
  const NS = {
    ha: "http://www.hancom.co.kr/hwpml/2011/app", hp: "http://www.hancom.co.kr/hwpml/2011/paragraph",
    hs: "http://www.hancom.co.kr/hwpml/2011/section", hc: "http://www.hancom.co.kr/hwpml/2011/core",
    hh: "http://www.hancom.co.kr/hwpml/2011/head", hpf: "http://www.hancom.co.kr/schema/2011/hpf",
    opf: "http://www.idpf.org/2007/opf/",
  };
  const xmlns = keys => keys.map(k => ` xmlns:${k}="${NS[k]}"`).join("");
  const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>';
  // XML 에 못 넣는 제어문자 제거 + 특수문자 이스케이프
  const x = s => String(s ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const hex = c => "#" + String(c || "000000").replace("#", "").toUpperCase();
  const FONT = "맑은 고딕";
  const LANGS = ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"];
  const all7 = v => `hangul="${v}" latin="${v}" hanja="${v}" japanese="${v}" other="${v}" symbol="${v}" user="${v}"`;
  const TW = 5;   // twip → HWPUNIT

  // ───────── 문서 도구 (report.js docKit 과 같은 모양) ─────────
  function kit() {
    // 글자·문단·테두리 모양 등록부: 같은 조합은 같은 번호를 씀
    const charPrs = [], paraPrs = [], fills = [];
    const charId = (o = {}) => {
      const k = [o.size || 21, !!o.bold, !!o.italics, hex(o.color)].join("|");
      let i = charPrs.findIndex(c => c.k === k);
      if (i < 0) { charPrs.push({ k, size: o.size || 21, bold: !!o.bold, italic: !!o.italics, color: hex(o.color) }); i = charPrs.length - 1; }
      return i;
    };
    const paraId = (o = {}) => {
      const align = o.align === "center" ? "CENTER" : o.align === "right" ? "RIGHT" : o.align === "justify" ? "JUSTIFY" : "LEFT";
      const k = [align, o.indent || 0, o.hang || 0, o.before || 0, o.after || 0, o.line || 160].join("|");
      let i = paraPrs.findIndex(p => p.k === k);
      if (i < 0) { paraPrs.push({ k, align, left: ((o.indent || 0) + (o.hang || 0)) * TW, intent: -(o.hang || 0) * TW, prev: (o.before || 0) * TW, next: (o.after || 0) * TW, line: o.line || 160 }); i = paraPrs.length - 1; }
      return i;
    };
    // 테두리/배경: 1 = 쪽 테두리(없음), 2 = 글자 테두리(없음), 3~ = 표 칸(실선 + 배경색)
    const fillId = color => {
      const k = color ? hex(color) : "";
      let i = fills.indexOf(k);
      if (i < 0) { fills.push(k); i = fills.length - 1; }
      return 3 + i;
    };
    charId(); paraId(); fillId("");   // 0번 기본값

    let pid = 0, tid = 0;
    const run = (text, o = {}) => ({ text: String(text ?? ""), o });
    const pXml = (runs, o = {}) => {
      const body = runs.map(r => `<hp:run charPrIDRef="${charId(r.o)}"><hp:t>${x(r.text)}</hp:t></hp:run>`).join("");
      return `<hp:p id="${++pid}" paraPrIDRef="${paraId(o)}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${body || '<hp:run charPrIDRef="0"/>'}</hp:p>`;
    };
    const P = (text, o = {}) => pXml(Array.isArray(text) ? text : [run(text, o)], {
      align: o.align, indent: o.indent, hang: o.hang, before: o.before ?? 0, after: o.after ?? 120,
    });
    const H = (text, lvl = 1) => P(text, { bold: true, size: lvl === 1 ? 26 : 22, before: lvl === 1 ? 300 : 160, after: 120, color: lvl === 1 ? "1F3864" : undefined });
    const B = (text, o = {}) => P([run("○ ", o), run(text, o)], { indent: 200, hang: 290, ...o });

    const CM = { l: 280, r: 280, t: 120, b: 120 };
    const cellXml = (text, w, row, col, o = {}) => {
      const paras = String(text ?? "").split("\n").map(line =>
        pXml([run(line, { size: o.size || 17, bold: o.bold, color: o.color })], { align: o.align, after: 0, line: 130 })).join("");
      return `<hp:tc name="" header="${o.header ? 1 : 0}" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="${fillId(o.fill)}">` +
        `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${paras}</hp:subList>` +
        `<hp:cellAddr colAddr="${col}" rowAddr="${row}"/><hp:cellSpan colSpan="1" rowSpan="1"/>` +
        `<hp:cellSz width="${w}" height="${o.h || 1100}"/><hp:cellMargin left="${CM.l}" right="${CM.r}" top="${CM.t}" bottom="${CM.b}"/></hp:tc>`;
    };
    const tblXml = (widths, rowsCells, o = {}) => {
      const W = widths.reduce((s, w) => s + w, 0), rowH = o.rowH || 1100;
      const trs = rowsCells.map((cells, r) => `<hp:tr>${cells.map((c, j) => cellXml(c.text, widths[j], r, j, { ...c, h: c.h || rowH })).join("")}</hp:tr>`).join("");
      const H = rowsCells.reduce((s, cells) => s + Math.max(...cells.map(c => c.h || rowH)), 0);
      return `<hp:p id="${++pid}" paraPrIDRef="${paraId({ after: 160 })}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">` +
        `<hp:tbl id="${1000 + (++tid)}" zOrder="${tid}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rowsCells.length}" colCnt="${widths.length}" cellSpacing="0" borderFillIDRef="${fillId("")}" noAdjust="0">` +
        `<hp:sz width="${W}" widthRelTo="ABSOLUTE" height="${H}" heightRelTo="ABSOLUTE" protect="0"/>` +
        `<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
        `<hp:outMargin left="0" right="0" top="0" bottom="400"/><hp:inMargin left="${CM.l}" right="${CM.r}" top="${CM.t}" bottom="${CM.b}"/>` +
        `${trs}</hp:tbl></hp:run></hp:p>`;
    };
    // cols: [{ h: 머리글, w: 폭(twip), align }] — 워드와 같은 값
    const table = (cols, rows, o = {}) => {
      const widths = cols.map(c => Math.round(c.w * TW));
      const head = cols.map(c => ({ text: c.h, bold: true, fill: "1F3864", color: "FFFFFF", align: "center", header: true }));
      const body = rows.map((r, i) => r.map((v, j) => ({ text: v, align: cols[j].align,
        fill: o.highlight && o.highlight(i) ? "FFF2CC" : (o.bold && o.bold(i) ? "F2F2F2" : ""), bold: o.bold && o.bold(i) })));
      return tblXml(widths, [head, ...body]);
    };
    // 담당자 작성 칸: 회색 안내문 + ○ 줄. filled 가 있으면 그 문장들로 채움
    const box = (lines = 6, hint = "", filled = null) => {
      const items = filled && filled.length ? filled : Array.from({ length: lines }, () => "");
      const paras = [hint ? pXml([run(hint, { size: 18, color: "808080", italics: true })], { after: 60, line: 150 }) : "",
        ...items.map(t => pXml([run("○ " + t, { size: 21 })], { after: 60, line: 160, hang: 290 }))].join("");
      const w = 9600 * TW;
      return `<hp:p id="${++pid}" paraPrIDRef="${paraId({ after: 160 })}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">` +
        `<hp:tbl id="${1000 + (++tid)}" zOrder="${tid}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="TABLE" repeatHeader="0" rowCnt="1" colCnt="1" cellSpacing="0" borderFillIDRef="${fillId("")}" noAdjust="0">` +
        `<hp:sz width="${w}" widthRelTo="ABSOLUTE" height="${(items.length + 1) * 1800}" heightRelTo="ABSOLUTE" protect="0"/>` +
        `<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
        `<hp:outMargin left="0" right="0" top="0" bottom="400"/><hp:inMargin left="850" right="850" top="400" bottom="400"/>` +
        `<hp:tr><hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${fillId("FFFFF0")}">` +
        `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${paras}</hp:subList>` +
        `<hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/><hp:cellSz width="${w}" height="${(items.length + 1) * 1800}"/>` +
        `<hp:cellMargin left="850" right="850" top="400" bottom="400"/></hp:tc></hp:tr></hp:tbl></hp:run></hp:p>`;
    };

    // ───────── 머리(header.xml) ─────────
    function headerXml() {
      const fontfaces = LANGS.map(l => `<hh:fontface lang="${l}" fontCnt="1"><hh:font id="0" face="${FONT}" type="TTF" isEmbedded="0">` +
        `<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4" contrast="0" strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/></hh:font></hh:fontface>`).join("");
      const border = (type, color = "#000000") => ["left", "right", "top", "bottom"].map(s => `<hh:${s}Border type="${type}" width="0.12 mm" color="${color}"/>`).join("");
      const bf = (id, type, fill) => `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">` +
        `<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>${border(type, type === "NONE" ? "#000000" : "#808080")}` +
        `<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>` +
        (fill ? `<hc:fillBrush><hc:winBrush faceColor="${fill}" hatchColor="#999999" alpha="0"/></hc:fillBrush>` : "") + `</hh:borderFill>`;
      const borderFills = [bf(1, "NONE"), bf(2, "NONE"), ...fills.map((f, i) => bf(3 + i, "SOLID", f))];
      const chars = charPrs.map((c, i) => `<hh:charPr id="${i}" height="${c.size * 50}" textColor="${c.color}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">` +
        `<hh:fontRef ${all7(0)}/><hh:ratio ${all7(100)}/><hh:spacing ${all7(0)}/><hh:relSz ${all7(100)}/><hh:offset ${all7(0)}/>` +
        (c.italic ? "<hh:italic/>" : "") + (c.bold ? "<hh:bold/>" : "") +
        `<hh:underline type="NONE" shape="SOLID" color="#000000"/><hh:strikeout shape="NONE" color="#000000"/><hh:outline type="NONE"/>` +
        `<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>`);
      const paras = paraPrs.map((p, i) => `<hh:paraPr id="${i}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">` +
        `<hh:align horizontal="${p.align}" vertical="BASELINE"/><hh:heading type="NONE" idRef="0" level="0"/>` +
        `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>` +
        `<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>` +
        `<hh:margin><hc:intent value="${p.intent}" unit="HWPUNIT"/><hc:left value="${p.left}" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/>` +
        `<hc:prev value="${p.prev}" unit="HWPUNIT"/><hc:next value="${p.next}" unit="HWPUNIT"/></hh:margin>` +
        `<hh:lineSpacing type="PERCENT" value="${p.line}" unit="HWPUNIT"/>` +
        `<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>`);
      return DECL + `<hh:head${xmlns(["ha", "hp", "hs", "hc", "hh", "hpf", "opf"])} version="1.4" secCnt="1">` +
        `<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/><hh:refList>` +
        `<hh:fontfaces itemCnt="${LANGS.length}">${fontfaces}</hh:fontfaces>` +
        `<hh:borderFills itemCnt="${borderFills.length}">${borderFills.join("")}</hh:borderFills>` +
        `<hh:charProperties itemCnt="${chars.length}">${chars.join("")}</hh:charProperties>` +
        `<hh:tabProperties itemCnt="1"><hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/></hh:tabProperties>` +
        `<hh:paraProperties itemCnt="${paras.length}">${paras.join("")}</hh:paraProperties>` +
        `<hh:styles itemCnt="1"><hh:style id="0" type="PARA" name="바탕글" engName="Normal" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/></hh:styles>` +
        `</hh:refList><hh:compatibleDocument targetProgram="HWP201X"><hh:layoutCompatibility/></hh:compatibleDocument>` +
        `<hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption><hh:trackchangeConfig flags="56"/></hh:head>`;
    }

    // 첫 문단에 쪽 설정(A4, 여백 20mm)
    const SEC = `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">` +
      `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="0" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0">` +
      `<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>` +
      `<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>` +
      `<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>` +
      `<hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY"><hp:margin header="2835" footer="2835" gutter="0" left="5669" right="5669" top="4252" bottom="4252"/></hp:pagePr>` +
      `<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>` +
      `<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr>` +
      `<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>` +
      `<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>` +
      ["BOTH", "EVEN", "ODD"].map(t => `<hp:pageBorderFill type="${t}" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>`).join("") +
      `</hp:secPr><hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl></hp:run></hp:p>`;

    const doc = (children, meta = {}) => ({ children, title: meta.title || "보고서 초안" });
    const files = d => {
      const section = DECL + `<hs:sec${xmlns(["ha", "hp", "hs", "hc", "hh", "hpf", "opf"])}>${SEC}${d.children.join("")}</hs:sec>`;
      const header = headerXml();   // 본문을 다 만든 뒤에 만들어야 쓰인 모양이 모두 등록됨
      const plain = d.children.join("\n").replace(/<[^>]+>/g, " ").replace(/\s+\n/g, "\n").slice(0, 1000);
      return [
        ["mimetype", "application/hwp+zip"],
        ["version.xml", DECL + `<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="1" xmlVersion="1.4" application="Hancom Office Hangul" appVersion="11, 0, 0, 0"/>`],
        ["META-INF/container.xml", DECL + `<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="${NS.hpf}"><ocf:rootfiles>` +
          `<ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/></ocf:rootfiles></ocf:container>`],
        ["META-INF/manifest.xml", DECL + `<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`],
        ["Contents/content.hpf", DECL + `<opf:package${xmlns(["ha", "hp", "hs", "hc", "hh", "hpf", "opf"])} version="" unique-identifier="" id="">` +
          `<opf:metadata><opf:title>${x(d.title)}</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">나라장터 AI 사업 모니터</opf:meta></opf:metadata>` +
          `<opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>` +
          `<opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest>` +
          `<opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`],
        ["Contents/header.xml", header],
        ["Contents/section0.xml", section],
        ["settings.xml", DECL + `<ha:HWPApplicationSetting xmlns:ha="${NS.ha}" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>`],
        ["Preview/PrvText.txt", plain],
      ];
    };
    const toBytes = d => zip(files(d));
    const toBlob = async d => new Blob([toBytes(d)], { type: "application/hwp+zip" });
    return { run, P, H, B, table, box, doc, toBlob, toBytes, files, ext: "hwpx" };
  }

  const api = { kit, zip };
  if (typeof window !== "undefined") window.G2BHwpx = api;
  if (typeof module !== "undefined") module.exports = api;
})();
