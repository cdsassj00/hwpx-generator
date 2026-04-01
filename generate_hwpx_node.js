/**
 * HWPX 공공행정문서 생성기 v2 (JavaScript / Node.js)
 * ====================================================
 * Python 버전과 동일한 접근: python-hwpx base template을 내장하고
 * xmldom으로 XML을 패치하는 방식.
 *
 * 설치: npm install jszip @xmldom/xmldom
 */

const JSZip = require('jszip');
const fs = require('fs');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const TEMPLATE_FILES = require('./hwpx_template_node');

// ── 네임스페이스 ──────────────────────────────────────────────
const HH = 'http://www.hancom.co.kr/hwpml/2011/head';
const HC = 'http://www.hancom.co.kr/hwpml/2011/core';
const HP = 'http://www.hancom.co.kr/hwpml/2011/paragraph';

// ── 표준 번호체계 ─────────────────────────────────────────────
const NUM_L1 = ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ'];
const NUM_L3 = ['가', '나', '다', '라', '마', '바', '사', '아'];
const BULLET_H2 = '□';
const BULLET_L1 = '○';
const BULLET_L2 = '-';
const NUM_L1_SET = new Set(NUM_L1);

const DOC_TEXT_WIDTH = 42520;
const NAVY = '#003366';
const NAVY_LINE = '#315F97';
const TITLE_BG = '#DFEAF5';

// ── 폰트 정의 ─────────────────────────────────────────────────
const REQUIRED_FONTS = [
  { face: '맑은 고딕',    familyType: 'FCAT_GOTHIC',  weight: '6', proportion: '4' },
  { face: 'HY헤드라인M',  familyType: 'FCAT_GOTHIC',  weight: '6', proportion: '0' },
  { face: '휴먼명조',     familyType: 'FCAT_MYUNGJO', weight: '6', proportion: '0' },
];

// ── charPr 정의 [name, height, fontFace, bold, color] ─────────
const CHAR_STYLES = [
  ['meta',       1300, '휴먼명조',    false, '#000000'],
  ['title',      2000, 'HY헤드라인M', false, '#000000'],
  ['sec_num',    1500, '맑은 고딕',   true,  '#FFFFFF'],
  ['sec_title',  1600, 'HY헤드라인M', false, '#000000'],
  ['h2',         1600, 'HY헤드라인M', false, '#000000'],
  ['body',       1500, '휴먼명조',    false, '#000000'],
  ['body_bold',  1500, '휴먼명조',    true,  '#000000'],
  ['note',       1200, '맑은 고딕',   false, '#000000'],
  ['tbl_hdr',    1200, '맑은 고딕',   true,  '#000000'],
  ['tbl_body',   1200, '맑은 고딕',   false, '#000000'],
  ['conclusion', 1500, '휴먼명조',    false, '#000000'],
];

// ── paraPr 정의 [name, halign, left, indent, lspct, prev, next] ─
const PARA_STYLES = [
  ['meta',       'CENTER',  0,    0, 130,    0,  700],
  ['title',      'CENTER',  0,    0, 130,    0, 2000],
  ['sec_num',    'CENTER',  0,    0, 160,    0,    0],
  ['sec_title',  'JUSTIFY', 0,    0, 160,    0,    0],
  ['h2',         'JUSTIFY', 1400, 0, 160, 2400,  400],
  ['body',       'JUSTIFY', 2800, 0, 160,  400,    0],
  ['body_sub',   'JUSTIFY', 4200, 0, 160,  200,    0],
  ['note',       'JUSTIFY', 2800, 0, 150, 1000,    0],
  ['tbl',        'CENTER',  0,    0, 160,    0,    0],
  ['conclusion', 'JUSTIFY', 2800, 0, 160, 1200,    0],
  ['spacer',     'JUSTIFY', 0,    0, 130,    0,    0],
  ['l3_heading', 'JUSTIFY', 2800, 0, 160, 1400,  400],
  ['attach',     'JUSTIFY', 0,    0, 160, 2800,    0],
];


// ── 오늘 날짜 ─────────────────────────────────────────────────
function today() {
  const d = new Date();
  const wd = ['일', '월', '화', '수', '목', '금', '토'][d.getDay()];
  const Y = d.getFullYear();
  const M = String(d.getMonth() + 1).padStart(2, '0');
  const D = String(d.getDate()).padStart(2, '0');
  return `${Y}. ${M}. ${D}.(${wd})`;
}


// ── DOM 헬퍼 ──────────────────────────────────────────────────
function getElementsByTagNS(parent, ns, localName) {
  return Array.from(parent.getElementsByTagNameNS(ns, localName));
}

function deepClone(node) {
  return node.cloneNode(true);
}


// ═══════════════════════════════════════════════════════════════
// STEP 1: base 문서 빌드 (python-hwpx 역할)
// ═══════════════════════════════════════════════════════════════

function buildBaseSectionXml(doc) {
  const parser = new DOMParser();

  // Load base section template
  const baseSection = Buffer.from(TEMPLATE_FILES['Contents/section0.xml'], 'base64').toString('utf-8');
  const sectionDoc = parser.parseFromString(baseSection, 'text/xml');
  const sec = sectionDoc.documentElement;

  // Get the first paragraph (secPr) - keep it
  const allPs = getElementsByTagNS(sec, HP, 'p');
  const secPrPara = allPs[0];

  // Remove placeholder paragraphs (only direct children of sec, except secPr)
  for (let i = allPs.length - 1; i >= 1; i--) {
    const p = allPs[i];
    if (p.parentNode === sec) {
      sec.removeChild(p);
    }
  }

  // Build content
  const dt = doc.doc_type || '서면보고';
  const date = doc.date || today();
  const dept = doc.dept || '';
  const author = doc.author || '';
  const metaParts = [dt, date, `${dept} ${author}`.trim()].filter(Boolean);
  addPara(sectionDoc, sec, metaParts.join(' | '));
  addPara(sectionDoc, sec, '');

  // 제목 → 1×1 표
  addTable(sectionDoc, sec, 1, 1, [[doc.title || '보고서']]);
  addPara(sectionDoc, sec, '');

  // 섹션 렌더링
  const cnt = [0, 0, 0];

  // 불릿 중복 방지
  function strip(text) {
    return (text || '').replace(/^[\s]*(○|□|⇒|※|[-])\s*/, '').replace(/^[\s]*(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ)[.\s]*/, '').trim();
  }

  function render(s, level) {
    cnt[level - 1]++;
    for (let i = level; i < 3; i++) cnt[i] = 0;
    const heading = strip(s.heading || '');

    if (level === 1) {
      const num = NUM_L1[Math.min(cnt[0] - 1, NUM_L1.length - 1)];
      if (cnt[0] > 1) addPara(sectionDoc, sec, '');
      // 대제목 → 1×3 표
      addTable(sectionDoc, sec, 1, 3, [[num, '', heading]]);

      for (const sub of (s.subsections || [])) render(sub, level + 1);
      for (const para of (s.paragraphs || [])) addPara(sectionDoc, sec, `${BULLET_L1} ${strip(para)}`);
    } else if (level === 2) {
      addPara(sectionDoc, sec, `${BULLET_H2} ${heading}`);
      for (const para of (s.paragraphs || [])) addPara(sectionDoc, sec, `${BULLET_L1} ${strip(para)}`);
      for (const sub of (s.subsections || [])) render(sub, level + 1);
    } else {
      const num = NUM_L3[Math.min(cnt[2] - 1, NUM_L3.length - 1)];
      addPara(sectionDoc, sec, `${num}. ${heading}`);
      for (const para of (s.paragraphs || [])) addPara(sectionDoc, sec, `${BULLET_L2} ${strip(para)}`);
      for (const sub of (s.subsections || [])) render(sub, level + 1);
    }

    for (const note of (s.notes || [])) addPara(sectionDoc, sec, `※ ${strip(note)}`);
    for (const conc of (s.conclusions || [])) addPara(sectionDoc, sec, `⇒ ${strip(conc)}`);

    if (s.table && s.table.rows && s.table.rows.length > 0) {
      addPara(sectionDoc, sec, '');
      const colCount = Math.max(...s.table.rows.map(r => r.length));
      addTable(sectionDoc, sec, s.table.rows.length, colCount, s.table.rows);
      addPara(sectionDoc, sec, '');
    }
  }

  for (const s of (doc.sections || [])) render(s, 1);

  // 붙임
  (doc.attachments || []).forEach((att, i) => {
    addPara(sectionDoc, sec, `붙임${i + 1}  ${att}`);
  });

  // 담당자 표
  if (doc.contacts && doc.contacts.length > 0) {
    addPara(sectionDoc, sec, '');
    const rows = [['담당 부서', '담당자', '연락처']];
    for (const c of doc.contacts) rows.push([c.dept || '', c.name || '', c.tel || '']);
    addTable(sectionDoc, sec, rows.length, 3, rows);
  }

  return new XMLSerializer().serializeToString(sectionDoc);
}

function addPara(xmlDoc, parent, text) {
  const p = xmlDoc.createElementNS(HP, 'hp:p');
  p.setAttribute('id', String(Math.floor(Math.random() * 4294967295)));
  p.setAttribute('paraPrIDRef', '0');
  p.setAttribute('styleIDRef', '0');
  p.setAttribute('pageBreak', '0');
  p.setAttribute('columnBreak', '0');
  p.setAttribute('merged', '0');

  const run = xmlDoc.createElementNS(HP, 'hp:run');
  run.setAttribute('charPrIDRef', '0');
  const t = xmlDoc.createElementNS(HP, 'hp:t');
  t.appendChild(xmlDoc.createTextNode(text));
  run.appendChild(t);
  p.appendChild(run);
  parent.appendChild(p);
  return p;
}

function addTable(xmlDoc, parent, rowCnt, colCnt, data) {
  const p = xmlDoc.createElementNS(HP, 'hp:p');
  p.setAttribute('id', String(Math.floor(Math.random() * 4294967295)));
  p.setAttribute('paraPrIDRef', '0');
  p.setAttribute('styleIDRef', '0');
  p.setAttribute('pageBreak', '0');
  p.setAttribute('columnBreak', '0');
  p.setAttribute('merged', '0');

  const run = xmlDoc.createElementNS(HP, 'hp:run');
  run.setAttribute('charPrIDRef', '0');

  const tbl = xmlDoc.createElementNS(HP, 'hp:tbl');
  tbl.setAttribute('id', String(Math.floor(Math.random() * 4294967295)));
  tbl.setAttribute('zOrder', '0');
  tbl.setAttribute('numberingType', 'TABLE');
  tbl.setAttribute('textWrap', 'TOP_AND_BOTTOM');
  tbl.setAttribute('textFlow', 'BOTH_SIDES');
  tbl.setAttribute('lock', '0');
  tbl.setAttribute('dropcapstyle', 'None');
  tbl.setAttribute('pageBreak', 'CELL');
  tbl.setAttribute('repeatHeader', '0');
  tbl.setAttribute('rowCnt', String(rowCnt));
  tbl.setAttribute('colCnt', String(colCnt));
  tbl.setAttribute('cellSpacing', '0');
  tbl.setAttribute('borderFillIDRef', '3');
  tbl.setAttribute('noAdjust', '0');

  const colW = Math.floor(DOC_TEXT_WIDTH / colCnt);
  const rowH = 3600;

  const sz = xmlDoc.createElementNS(HP, 'hp:sz');
  sz.setAttribute('width', String(DOC_TEXT_WIDTH));
  sz.setAttribute('widthRelTo', 'ABSOLUTE');
  sz.setAttribute('height', String(rowH * rowCnt));
  sz.setAttribute('heightRelTo', 'ABSOLUTE');
  sz.setAttribute('protect', '0');
  tbl.appendChild(sz);

  const pos = xmlDoc.createElementNS(HP, 'hp:pos');
  for (const [k, v] of Object.entries({treatAsChar:'1',affectLSpacing:'0',flowWithText:'1',allowOverlap:'0',holdAnchorAndSO:'0',vertRelTo:'PARA',horzRelTo:'COLUMN',vertAlign:'TOP',horzAlign:'LEFT',vertOffset:'0',horzOffset:'0'})) {
    pos.setAttribute(k, v);
  }
  tbl.appendChild(pos);

  for (const marginName of ['outMargin', 'inMargin']) {
    const m = xmlDoc.createElementNS(HP, `hp:${marginName}`);
    m.setAttribute('left', '0'); m.setAttribute('right', '0');
    m.setAttribute('top', '0'); m.setAttribute('bottom', '0');
    tbl.appendChild(m);
  }

  for (let ri = 0; ri < rowCnt; ri++) {
    const tr = xmlDoc.createElementNS(HP, 'hp:tr');
    for (let ci = 0; ci < colCnt; ci++) {
      const text = (data[ri] && ci < data[ri].length) ? String(data[ri][ci]) : '';
      const tc = xmlDoc.createElementNS(HP, 'hp:tc');
      tc.setAttribute('name', '');
      tc.setAttribute('header', '0');
      tc.setAttribute('hasMargin', '0');
      tc.setAttribute('protect', '0');
      tc.setAttribute('editable', '0');
      tc.setAttribute('dirty', '1');
      tc.setAttribute('borderFillIDRef', '3');

      const subList = xmlDoc.createElementNS(HP, 'hp:subList');
      subList.setAttribute('id', '');
      subList.setAttribute('textDirection', 'HORIZONTAL');
      subList.setAttribute('lineWrap', 'BREAK');
      subList.setAttribute('vertAlign', 'CENTER');
      subList.setAttribute('linkListIDRef', '0');
      subList.setAttribute('linkListNextIDRef', '0');
      subList.setAttribute('textWidth', '0');
      subList.setAttribute('textHeight', '0');
      subList.setAttribute('hasTextRef', '0');
      subList.setAttribute('hasNumRef', '0');

      const cellP = xmlDoc.createElementNS(HP, 'hp:p');
      cellP.setAttribute('paraPrIDRef', '0');
      cellP.setAttribute('styleIDRef', '0');
      cellP.setAttribute('pageBreak', '0');
      cellP.setAttribute('columnBreak', '0');
      cellP.setAttribute('merged', '0');
      cellP.setAttribute('id', String(Math.floor(Math.random() * 4294967295)));

      const cellRun = xmlDoc.createElementNS(HP, 'hp:run');
      cellRun.setAttribute('charPrIDRef', '0');
      const cellT = xmlDoc.createElementNS(HP, 'hp:t');
      cellT.appendChild(xmlDoc.createTextNode(text));
      cellRun.appendChild(cellT);
      cellP.appendChild(cellRun);
      subList.appendChild(cellP);
      tc.appendChild(subList);

      const cellAddr = xmlDoc.createElementNS(HP, 'hp:cellAddr');
      cellAddr.setAttribute('colAddr', String(ci));
      cellAddr.setAttribute('rowAddr', String(ri));
      tc.appendChild(cellAddr);

      const cellSpan = xmlDoc.createElementNS(HP, 'hp:cellSpan');
      cellSpan.setAttribute('colSpan', '1');
      cellSpan.setAttribute('rowSpan', '1');
      tc.appendChild(cellSpan);

      const cellSz = xmlDoc.createElementNS(HP, 'hp:cellSz');
      cellSz.setAttribute('width', String(colW));
      cellSz.setAttribute('height', String(rowH));
      tc.appendChild(cellSz);

      const cellMargin = xmlDoc.createElementNS(HP, 'hp:cellMargin');
      cellMargin.setAttribute('left', '0'); cellMargin.setAttribute('right', '0');
      cellMargin.setAttribute('top', '0'); cellMargin.setAttribute('bottom', '0');
      tc.appendChild(cellMargin);

      tr.appendChild(tc);
    }
    tbl.appendChild(tr);
  }

  run.appendChild(tbl);
  p.appendChild(run);
  parent.appendChild(p);
}


// ═══════════════════════════════════════════════════════════════
// STEP 2: header.xml 패치
// ═══════════════════════════════════════════════════════════════

function patchHeader(headerXml) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(headerXml, 'text/xml');
  const styleMap = {};
  const bfMap = {};

  // ── 폰트 등록 ──
  const fontIdMap = {};
  const fontfaces = getElementsByTagNS(doc, HH, 'fontface');
  for (const ff of fontfaces) {
    const lang = ff.getAttribute('lang');
    let cnt = parseInt(ff.getAttribute('fontCnt') || '2');
    for (const { face, familyType, weight, proportion } of REQUIRED_FONTS) {
      let exists = false;
      const fonts = getElementsByTagNS(ff, HH, 'font');
      for (const f of fonts) {
        if (f.getAttribute('face') === face) {
          if (!fontIdMap[face]) fontIdMap[face] = {};
          fontIdMap[face][lang] = f.getAttribute('id');
          exists = true;
          break;
        }
      }
      if (!exists) {
        const font = doc.createElementNS(HH, 'hh:font');
        font.setAttribute('id', String(cnt));
        font.setAttribute('face', face);
        font.setAttribute('type', 'TTF');
        font.setAttribute('isEmbedded', '0');
        const ti = doc.createElementNS(HH, 'hh:typeInfo');
        ti.setAttribute('familyType', familyType);
        ti.setAttribute('weight', weight);
        ti.setAttribute('proportion', proportion);
        for (const a of ['contrast','strokeVariation','armStyle','letterform','midline','xHeight']) {
          ti.setAttribute(a, '0');
        }
        font.appendChild(ti);
        ff.appendChild(font);
        if (!fontIdMap[face]) fontIdMap[face] = {};
        fontIdMap[face][lang] = String(cnt);
        cnt++;
      }
    }
    ff.setAttribute('fontCnt', String(cnt));
  }

  function getFontId(fontName) {
    const ids = fontIdMap[fontName] || {};
    return ids['HANGUL'] || ids[Object.keys(ids)[0]] || '0';
  }

  // ── charPr 추가 ──
  const cps = getElementsByTagNS(doc, HH, 'charProperties')[0];
  const baseCp = getElementsByTagNS(cps, HH, 'charPr').find(el => el.getAttribute('id') === '0');
  let cpCnt = parseInt(cps.getAttribute('itemCnt') || '7');

  for (let i = 0; i < CHAR_STYLES.length; i++) {
    const [name, height, fontName, bold, color] = CHAR_STYLES[i];
    const newId = cpCnt + i;
    styleMap[`cp_${name}`] = String(newId);

    const cp = deepClone(baseCp);
    cp.setAttribute('id', String(newId));
    cp.setAttribute('height', String(height));
    cp.setAttribute('textColor', color);

    const fref = getElementsByTagNS(cp, HH, 'fontRef')[0];
    if (fref) {
      const fid = getFontId(fontName);
      for (const la of ['hangul','latin','hanja','japanese','other','symbol','user']) {
        const specific = (fontIdMap[fontName] || {})[la.toUpperCase()] || fid;
        fref.setAttribute(la, specific);
      }
    }

    const existingBold = getElementsByTagNS(cp, HH, 'bold')[0];
    if (bold && !existingBold) {
      const boldEl = doc.createElementNS(HH, 'hh:bold');
      // Insert after fontRef or at beginning
      const frefEl = getElementsByTagNS(cp, HH, 'fontRef')[0];
      if (frefEl && frefEl.nextSibling) {
        // Just append before underline
        const underline = getElementsByTagNS(cp, HH, 'underline')[0];
        if (underline) cp.insertBefore(boldEl, underline);
        else cp.appendChild(boldEl);
      } else {
        cp.appendChild(boldEl);
      }
    } else if (!bold && existingBold) {
      cp.removeChild(existingBold);
    }

    cps.appendChild(cp);
  }
  cps.setAttribute('itemCnt', String(cpCnt + CHAR_STYLES.length));

  // ── paraPr 추가 ──
  const pps = getElementsByTagNS(doc, HH, 'paraProperties')[0];
  const basePp = getElementsByTagNS(pps, HH, 'paraPr').find(el => el.getAttribute('id') === '0');
  let ppCnt = parseInt(pps.getAttribute('itemCnt') || '0');

  for (const [name, halign, leftMargin, indent, linePct, prev, nxt] of PARA_STYLES) {
    const newPpId = ppCnt;
    ppCnt++;
    const pp = deepClone(basePp);
    pp.setAttribute('id', String(newPpId));

    const alignEl = getElementsByTagNS(pp, HH, 'align')[0];
    if (alignEl) alignEl.setAttribute('horizontal', halign);

    // patch margins in hp:switch/hp:case and hp:switch/hp:default
    const margins = getElementsByTagNS(pp, HH, 'margin');
    for (const margin of margins) {
      for (let c = margin.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const localName = c.localName;
        if (localName === 'left') c.setAttribute('value', String(leftMargin));
        else if (localName === 'intent') c.setAttribute('value', String(indent));
        else if (localName === 'prev') c.setAttribute('value', String(prev));
        else if (localName === 'next') c.setAttribute('value', String(nxt));
      }
    }

    const lineSpacings = getElementsByTagNS(pp, HH, 'lineSpacing');
    for (const ls of lineSpacings) {
      ls.setAttribute('type', 'PERCENT');
      ls.setAttribute('value', String(linePct));
    }

    pps.appendChild(pp);
    styleMap[`pp_${name}`] = String(newPpId);
  }
  pps.setAttribute('itemCnt', String(ppCnt));

  // ── borderFill 추가 ──
  const bfs = getElementsByTagNS(doc, HH, 'borderFills')[0];
  let bfCnt = parseInt(bfs.getAttribute('itemCnt') || '3');

  function addBorderFill(key, borders, fillColor) {
    bfCnt++;
    bfMap[key] = String(bfCnt);

    const bf = doc.createElementNS(HH, 'hh:borderFill');
    bf.setAttribute('id', String(bfCnt));
    bf.setAttribute('threeD', '0');
    bf.setAttribute('shadow', '0');
    bf.setAttribute('centerLine', 'NONE');
    bf.setAttribute('breakCellSeparateLine', '0');

    for (const t of ['slash', 'backSlash']) {
      const el = doc.createElementNS(HH, `hh:${t}`);
      el.setAttribute('type', 'NONE');
      el.setAttribute('Crooked', '0');
      el.setAttribute('isCounter', '0');
      bf.appendChild(el);
    }

    for (const [side, type, width, color] of borders) {
      const el = doc.createElementNS(HH, `hh:${side}`);
      el.setAttribute('type', type);
      el.setAttribute('width', width);
      el.setAttribute('color', color);
      bf.appendChild(el);
    }

    if (fillColor) {
      const fb = doc.createElementNS(HC, 'hc:fillBrush');
      const wb = doc.createElementNS(HC, 'hc:winBrush');
      wb.setAttribute('faceColor', fillColor);
      wb.setAttribute('hatchColor', '#000000');
      wb.setAttribute('alpha', '0');
      fb.appendChild(wb);
      bf.appendChild(fb);
    }

    bfs.appendChild(bf);
  }

  // title_box: 연하늘 배경, 하단/우측 굵은선
  addBorderFill('title_box', [
    ['leftBorder', 'SOLID', '0.12 mm', '#000000'], ['rightBorder', 'SOLID', '0.5 mm', '#000000'],
    ['topBorder', 'SOLID', '0.12 mm', '#000000'], ['bottomBorder', 'SOLID', '0.5 mm', '#000000'],
  ], TITLE_BG);

  // sec_num_box: 남색 배경
  addBorderFill('sec_num_box', [
    ['leftBorder', 'SOLID', '0.1 mm', NAVY], ['rightBorder', 'SOLID', '0.1 mm', NAVY],
    ['topBorder', 'SOLID', '0.1 mm', NAVY], ['bottomBorder', 'SOLID', '0.1 mm', NAVY],
  ], NAVY);

  // sec_gap: 투명
  addBorderFill('sec_gap', [
    ['leftBorder', 'NONE', '0.12 mm', '#000000'], ['rightBorder', 'NONE', '0.12 mm', '#000000'],
    ['topBorder', 'NONE', '0.12 mm', '#000000'], ['bottomBorder', 'NONE', '0.12 mm', '#000000'],
  ], null);

  // sec_title_box: 하단 남색 밑줄
  addBorderFill('sec_title_box', [
    ['leftBorder', 'NONE', '0.25 mm', NAVY_LINE], ['rightBorder', 'NONE', '0.25 mm', NAVY_LINE],
    ['topBorder', 'NONE', '0.25 mm', NAVY_LINE], ['bottomBorder', 'SOLID', '0.5 mm', NAVY_LINE],
  ], null);

  // sec_hdr_tbl: 테두리 없음
  addBorderFill('sec_hdr_tbl', [
    ['leftBorder', 'NONE', '0.12 mm', '#000000'], ['rightBorder', 'NONE', '0.12 mm', '#000000'],
    ['topBorder', 'NONE', '0.12 mm', '#000000'], ['bottomBorder', 'NONE', '0.12 mm', '#000000'],
  ], null);

  // tbl_hdr: 표 헤더
  addBorderFill('tbl_hdr', [
    ['leftBorder', 'SOLID', '0.12 mm', '#000000'], ['rightBorder', 'SOLID', '0.12 mm', '#000000'],
    ['topBorder', 'SOLID', '0.4 mm', '#000000'], ['bottomBorder', 'SOLID', '0.4 mm', '#000000'],
  ], null);

  // tbl_body: 표 본문
  addBorderFill('tbl_body', [
    ['leftBorder', 'SOLID', '0.12 mm', '#000000'], ['rightBorder', 'SOLID', '0.12 mm', '#000000'],
    ['topBorder', 'SOLID', '0.12 mm', '#000000'], ['bottomBorder', 'SOLID', '0.12 mm', '#000000'],
  ], null);

  bfs.setAttribute('itemCnt', String(bfCnt));

  const headerOut = new XMLSerializer().serializeToString(doc);
  return { headerXml: headerOut, styleMap, bfMap };
}


// ═══════════════════════════════════════════════════════════════
// STEP 3: section0.xml 패치
// ═══════════════════════════════════════════════════════════════

function patchSection(sectionXml, styleMap, bfMap, docTitle, doc) {
  const parser = new DOMParser();
  const sectionDoc = parser.parseFromString(sectionXml, 'text/xml');

  const secHeadings = (doc.sections || []).map(s => s.heading || '');

  // ── 표 서식 적용 ──
  const tables = getElementsByTagNS(sectionDoc, HP, 'tbl');
  for (const tbl of tables) {
    const rows = parseInt(tbl.getAttribute('rowCnt') || '0');
    const cols = parseInt(tbl.getAttribute('colCnt') || '0');

    // 첫 번째 셀 텍스트
    let firstText = '';
    const tElems = getElementsByTagNS(tbl, HP, 't');
    for (const te of tElems) {
      const txt = (te.textContent || '').trim();
      if (txt) { firstText = txt; break; }
    }

    if (rows === 1 && cols === 1 && firstText === (docTitle || '').trim()) {
      applyTitleTable(tbl, styleMap, bfMap, sectionDoc);
    } else if (rows === 1 && cols === 3 && NUM_L1_SET.has(firstText.replace('.', ''))) {
      applySecHeaderTable(tbl, styleMap, bfMap, sectionDoc);
    } else {
      applyDataTable(tbl, styleMap, bfMap, sectionDoc);
    }
  }

  // ── 문단 스타일 적용 ──
  const paragraphs = getElementsByTagNS(sectionDoc, HP, 'p');
  for (const p of paragraphs) {
    // Skip paragraphs inside tables (subList children)
    if (p.parentNode && p.parentNode.localName === 'subList') continue;

    // Skip paragraphs that contain tables (these are wrapper paragraphs for tables)
    if (getElementsByTagNS(p, HP, 'tbl').length > 0) continue;

    // Only get direct child runs (not runs nested in tables within this paragraph)
    const runs = [];
    for (let c = p.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && c.localName === 'run' && c.namespaceURI === HP) {
        runs.push(c);
      }
    }

    let text = '';
    for (const r of runs) {
      const t = getElementsByTagNS(r, HP, 't')[0];
      if (t && t.textContent) { text = t.textContent.trim(); break; }
    }

    // Skip secPr paragraph
    if (getElementsByTagNS(p, HP, 'secPr').length > 0) continue;

    let cpKey = null, ppKey = null;

    if (!text) {
      ppKey = 'pp_spacer';
    } else if (isMeta(text)) {
      cpKey = 'cp_meta'; ppKey = 'pp_meta';
    } else if (text.startsWith(BULLET_H2)) {
      cpKey = 'cp_h2'; ppKey = 'pp_h2';
    } else if (text.startsWith('⇒')) {
      cpKey = 'cp_conclusion'; ppKey = 'pp_conclusion';
    } else if (text.startsWith('※')) {
      cpKey = 'cp_note'; ppKey = 'pp_note';
    } else if (text.startsWith('붙임')) {
      cpKey = 'cp_sec_title'; ppKey = 'pp_attach';
    } else if (text.startsWith(BULLET_L1)) {
      cpKey = 'cp_body'; ppKey = 'pp_body';
    } else if (text.startsWith(BULLET_L2 + ' ')) {
      cpKey = 'cp_body'; ppKey = 'pp_body_sub';
    } else if (NUM_L3.some(n => text.startsWith(n + '.'))) {
      cpKey = 'cp_h2'; ppKey = 'pp_l3_heading';
    } else {
      cpKey = 'cp_body'; ppKey = 'pp_body';
    }

    if (cpKey && styleMap[cpKey]) {
      for (const r of runs) {
        if (getElementsByTagNS(r, HP, 'secPr').length > 0) continue;
        r.setAttribute('charPrIDRef', styleMap[cpKey]);
      }
    }
    if (ppKey && styleMap[ppKey]) {
      p.setAttribute('paraPrIDRef', styleMap[ppKey]);
    }
  }

  return new XMLSerializer().serializeToString(sectionDoc);
}

function isMeta(text) {
  return text.includes('|') && (text.includes('서면보고') || text.includes('기안') ||
    text.includes('대면보고') || text.includes('보고'));
}

function applyTitleTable(tbl, styleMap, bfMap, doc) {
  tbl.setAttribute('borderFillIDRef', bfMap.title_box || '1');
  const sz = getElementsByTagNS(tbl, HP, 'sz')[0];
  if (sz) sz.setAttribute('width', String(DOC_TEXT_WIDTH));

  for (const tc of getElementsByTagNS(tbl, HP, 'tc')) {
    tc.setAttribute('borderFillIDRef', bfMap.title_box || '1');
    const csz = getElementsByTagNS(tc, HP, 'cellSz')[0];
    if (csz) csz.setAttribute('width', String(DOC_TEXT_WIDTH));

    for (const run of getElementsByTagNS(tc, HP, 'run')) {
      run.setAttribute('charPrIDRef', styleMap.cp_title || '0');
    }
    for (const p of getElementsByTagNS(tc, HP, 'p')) {
      if (styleMap.pp_title) p.setAttribute('paraPrIDRef', styleMap.pp_title);
    }
  }
}

function applySecHeaderTable(tbl, styleMap, bfMap, doc) {
  tbl.setAttribute('borderFillIDRef', bfMap.sec_hdr_tbl || '1');
  const sz = getElementsByTagNS(tbl, HP, 'sz')[0];
  if (sz) sz.setAttribute('width', String(DOC_TEXT_WIDTH));

  const trs = getElementsByTagNS(tbl, HP, 'tr');
  if (!trs.length) return;
  const tcs = getElementsByTagNS(trs[0], HP, 'tc');
  if (tcs.length < 3) return;

  const w0 = 2573, w1 = 566, w2 = DOC_TEXT_WIDTH - w0 - w1;

  // Cell 0: 남색 번호 박스
  tcs[0].setAttribute('borderFillIDRef', bfMap.sec_num_box || '1');
  const csz0 = getElementsByTagNS(tcs[0], HP, 'cellSz')[0];
  if (csz0) { csz0.setAttribute('width', String(w0)); csz0.setAttribute('height', '2466'); }
  for (const run of getElementsByTagNS(tcs[0], HP, 'run')) run.setAttribute('charPrIDRef', styleMap.cp_sec_num || '0');
  for (const p of getElementsByTagNS(tcs[0], HP, 'p')) p.setAttribute('paraPrIDRef', styleMap.pp_sec_num || '0');

  // Cell 1: 간격
  tcs[1].setAttribute('borderFillIDRef', bfMap.sec_gap || '1');
  const csz1 = getElementsByTagNS(tcs[1], HP, 'cellSz')[0];
  if (csz1) { csz1.setAttribute('width', String(w1)); csz1.setAttribute('height', '2466'); }

  // Cell 2: 제목 (하단 밑줄)
  tcs[2].setAttribute('borderFillIDRef', bfMap.sec_title_box || '1');
  const csz2 = getElementsByTagNS(tcs[2], HP, 'cellSz')[0];
  if (csz2) { csz2.setAttribute('width', String(w2)); csz2.setAttribute('height', '2466'); }
  for (const run of getElementsByTagNS(tcs[2], HP, 'run')) run.setAttribute('charPrIDRef', styleMap.cp_sec_title || '0');
  for (const p of getElementsByTagNS(tcs[2], HP, 'p')) p.setAttribute('paraPrIDRef', styleMap.pp_sec_title || '0');

  // subList vertAlign → CENTER
  for (const sl of getElementsByTagNS(tbl, HP, 'subList')) sl.setAttribute('vertAlign', 'CENTER');
}

function applyDataTable(tbl, styleMap, bfMap, doc) {
  const colCnt = parseInt(tbl.getAttribute('colCnt') || '3');
  const colW = Math.floor(DOC_TEXT_WIDTH / colCnt);

  const sz = getElementsByTagNS(tbl, HP, 'sz')[0];
  if (sz) sz.setAttribute('width', String(DOC_TEXT_WIDTH));

  const trs = getElementsByTagNS(tbl, HP, 'tr');
  trs.forEach((tr, ri) => {
    const tcs = getElementsByTagNS(tr, HP, 'tc');
    tcs.forEach((tc, ci) => {
      const csz = getElementsByTagNS(tc, HP, 'cellSz')[0];
      if (csz) csz.setAttribute('width', String(colW));

      if (ri === 0) {
        tc.setAttribute('borderFillIDRef', bfMap.tbl_hdr || '3');
        tc.setAttribute('header', '1');
        for (const run of getElementsByTagNS(tc, HP, 'run')) run.setAttribute('charPrIDRef', styleMap.cp_tbl_hdr || '0');
      } else {
        tc.setAttribute('borderFillIDRef', bfMap.tbl_body || '3');
        for (const run of getElementsByTagNS(tc, HP, 'run')) run.setAttribute('charPrIDRef', styleMap.cp_tbl_body || '0');
      }

      for (const p of getElementsByTagNS(tc, HP, 'p')) {
        if (styleMap.pp_tbl) p.setAttribute('paraPrIDRef', styleMap.pp_tbl);
      }
    });
  });
}


// ═══════════════════════════════════════════════════════════════
// STEP 4: ZIP 조립
// ═══════════════════════════════════════════════════════════════

async function createGovHwpx(doc, outputPath) {
  outputPath = outputPath || doc.output || '보고서.hwpx';

  // 1. Build base section
  const baseSectionXml = buildBaseSectionXml(doc);

  // 2. Load and patch header
  const baseHeaderXml = Buffer.from(TEMPLATE_FILES['Contents/header.xml'], 'base64').toString('utf-8');
  const { headerXml, styleMap, bfMap } = patchHeader(baseHeaderXml);

  // 3. Patch section
  const sectionXml = patchSection(baseSectionXml, styleMap, bfMap, doc.title || '', doc);

  // 4. Assemble ZIP (exact same structure as python-hwpx)
  const zip = new JSZip();

  // Add all template files first (preserving order)
  for (const [filename, b64data] of Object.entries(TEMPLATE_FILES)) {
    if (filename === 'Contents/header.xml' || filename === 'Contents/section0.xml') continue;

    const data = Buffer.from(b64data, 'base64');
    if (filename === 'mimetype') {
      zip.file(filename, data, { compression: 'STORE', createFolders: false });
    } else {
      zip.file(filename, data, { createFolders: false });
    }
  }

  // Add patched header and section
  zip.file('Contents/header.xml', headerXml, { createFolders: false });
  zip.file('Contents/section0.xml', sectionXml, { createFolders: false });

  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
    // Ensure mimetype is first and uncompressed
  });

  fs.writeFileSync(outputPath, buffer);
  console.log(`✅ HWPX 생성 완료: ${outputPath}`);
  return outputPath;
}


// ═══════════════════════════════════════════════════════════════
// CLI / 데모
// ═══════════════════════════════════════════════════════════════
if (require.main === module) {
  const demo = {
    title: 'AI시대 행정문서 작성 가이드라인(안) 보고',
    doc_type: '서면보고',
    date: today(),
    dept: '혁신행정담당관',
    author: '박은희 사무관',
    sections: [
      {
        heading: '추진 배경',
        subsections: [
          {
            heading: '추진 필요성',
            paragraphs: ['정부는 사람과 AI 모두 쉽게 읽고 작성할 수 있는 표준화된 문서 작성 가이드라인을 수립할 필요가 있음.'],
          },
          {
            heading: '현황 분석',
            paragraphs: ['현행 문서 형식의 한계를 분석함.'],
          },
        ],
      },
      {
        heading: '주요 추진 내용',
        subsections: [
          {
            heading: '가이드라인 배포 및 시범 실시',
            paragraphs: ['행정안전부는 AI 친화적 보고서 작성을 위해 가이드라인을 전 부서에 배포하고 시범 실시함.'],
            table: { rows: [['구분','기존 방식','개선 방식'],['문서 형식','개조식','서술식'],['표 작성','셀 병합 허용','셀 병합 금지']] },
          },
        ],
        conclusions: ['보안 체계와 AI 혁신은 동시 추구 가치이며, 안전한 혁신이 가능하도록 제도 마련 필요'],
      },
    ],
    contacts: [{ dept: '혁신행정담당관', name: '박은희 사무관', tel: '044-205-1473' }],
  };

  createGovHwpx(demo, '보고서_v2.hwpx')
    .then(p => console.log(`파일: ${p}`))
    .catch(e => console.error(e));
}

module.exports = { createGovHwpx, today };
