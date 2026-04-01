/**
 * HWPX 공공행정문서 생성기 — 브라우저 버전
 * ==========================================
 * Node.js 의존성 없이 브라우저에서 직접 실행 가능.
 * 필요: JSZip (CDN으로 로드)
 *
 * 사용법:
 *   // 기본 스타일
 *   const blob = await HwpxGenerator.createGovHwpx(doc);
 *
 *   // 커스텀 스타일
 *   const blob = await HwpxGenerator.createGovHwpx(doc, null, {
 *     colors: { navy: '#003366', title_bg: '#DFEAF5' },
 *     fonts: { title: 'HY헤드라인M', body: '휴먼명조' },
 *     sizes: { title: 2000, body: 1500 },
 *     indent: { h2: 1400, body: 2800 },
 *     spacing: { body_line: 160, h2_before: 2400 },
 *   });
 */

(function (root) {
  'use strict';

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

  // ═══════════════════════════════════════════════════════════════
  // 기본 스타일 설정 (사용자가 options로 오버라이드 가능)
  // ═══════════════════════════════════════════════════════════════
  const DEFAULTS = {
    colors: {
      navy:       '#003366',   // 대제목 박스 배경
      navy_line:  '#315F97',   // 대제목 밑줄
      title_bg:   '#DFEAF5',   // 보고서 제목 배경
      text:       '#000000',   // 기본 글자색
      sec_num:    '#FFFFFF',   // 대제목 번호 글자색 (흰색)
    },
    fonts: {
      title:      'HY헤드라인M',  // 제목, 대제목, 중제목
      body:       '휴먼명조',     // 본문
      ui:         '맑은 고딕',    // 메타, 번호, 표, 주석
    },
    sizes: {  // 글자 크기 (HWPUNIT, 100 = 1pt)
      meta:       1300,
      title:      2000,
      sec_num:    1500,
      sec_title:  1600,
      h2:         1600,
      body:       1500,
      note:       1200,
      tbl_hdr:    1200,
      tbl_body:   1200,
      conclusion: 1500,
    },
    indent: {  // 들여쓰기 (HWPUNIT, 7200 = 1인치 = 25.4mm)
      h2:         1400,   // □ 중제목
      body:       2800,   // ○ 본문
      body_sub:   4200,   // - 하위
      note:       2800,   // ※ 주석
      conclusion: 2800,   // ⇒ 결론
      l3_heading: 2800,   // 가. 소제목
    },
    spacing: {  // 줄간격/문단간격 (HWPUNIT)
      body_line:    160,  // 본문 줄간격 (%)
      meta_line:    130,  // 메타 줄간격 (%)
      h2_before:   2400,  // □ 중제목 위 간격
      h2_after:     400,  // □ 중제목 아래 간격
      body_before:  400,  // ○ 본문 위 간격
      title_after: 2000,  // 제목 아래 간격
      meta_after:   700,  // 메타 아래 간격
    },
  };

  /** 옵션 병합: 사용자가 넘긴 값만 오버라이드 */
  function mergeOptions(userOpts) {
    if (!userOpts) return JSON.parse(JSON.stringify(DEFAULTS));
    const merged = JSON.parse(JSON.stringify(DEFAULTS));
    for (const group of Object.keys(DEFAULTS)) {
      if (userOpts[group] && typeof userOpts[group] === 'object') {
        Object.assign(merged[group], userOpts[group]);
      }
    }
    return merged;
  }

  /** 옵션에서 내부 스타일 배열 생성 */
  function buildCharStyles(o) {
    return [
      ['meta',       o.sizes.meta,       o.fonts.ui,    false, o.colors.text],
      ['title',      o.sizes.title,      o.fonts.title, false, o.colors.text],
      ['sec_num',    o.sizes.sec_num,    o.fonts.ui,    true,  o.colors.sec_num],
      ['sec_title',  o.sizes.sec_title,  o.fonts.title, false, o.colors.text],
      ['h2',         o.sizes.h2,         o.fonts.title, false, o.colors.text],
      ['body',       o.sizes.body,       o.fonts.body,  false, o.colors.text],
      ['body_bold',  o.sizes.body,       o.fonts.body,  true,  o.colors.text],
      ['note',       o.sizes.note,       o.fonts.ui,    false, o.colors.text],
      ['tbl_hdr',    o.sizes.tbl_hdr,    o.fonts.ui,    true,  o.colors.text],
      ['tbl_body',   o.sizes.tbl_body,   o.fonts.ui,    false, o.colors.text],
      ['conclusion', o.sizes.conclusion, o.fonts.ui,    true,  o.colors.text],
    ];
  }

  function buildParaStyles(o) {
    return [
      ['meta',       'CENTER',  0,              0, o.spacing.meta_line,  0,                    o.spacing.meta_after],
      ['title',      'CENTER',  0,              0, o.spacing.meta_line,  0,                    o.spacing.title_after],
      ['sec_num',    'CENTER',  0,              0, o.spacing.body_line,  0,                    0],
      ['sec_title',  'JUSTIFY', 0,              0, o.spacing.body_line,  0,                    0],
      ['h2',         'JUSTIFY', o.indent.h2,    0, o.spacing.body_line,  o.spacing.h2_before,  o.spacing.h2_after],
      ['body',       'JUSTIFY', o.indent.body,  0, o.spacing.body_line,  o.spacing.body_before, 0],
      ['body_sub',   'JUSTIFY', o.indent.body_sub, 0, o.spacing.body_line, 200, 0],
      ['note',       'JUSTIFY', o.indent.note,  0, 150,                  1000,                  0],
      ['tbl',        'CENTER',  0,              0, o.spacing.body_line,  0,                    0],
      ['conclusion', 'JUSTIFY', o.indent.conclusion, 0, o.spacing.body_line, 1200, 0],
      ['spacer',     'JUSTIFY', 0,              0, o.spacing.meta_line,  0,                    0],
      ['l3_heading', 'JUSTIFY', o.indent.l3_heading, 0, o.spacing.body_line, 1400, 400],
      ['attach',     'JUSTIFY', 0,              0, o.spacing.body_line,  2800,                  0],
    ];
  }

  /** 옵션에서 사용할 폰트 목록 추출 */
  function buildRequiredFonts(o) {
    const fontSet = new Set([o.fonts.title, o.fonts.body, o.fonts.ui]);
    const fontMap = {
      '맑은 고딕':    { familyType: 'FCAT_GOTHIC',  weight: '6', proportion: '4' },
      'HY헤드라인M':  { familyType: 'FCAT_GOTHIC',  weight: '6', proportion: '0' },
      '휴먼명조':     { familyType: 'FCAT_MYUNGJO', weight: '6', proportion: '0' },
    };
    // 기본 3폰트는 항상 포함, 추가 폰트는 GOTHIC 기본값
    return Array.from(fontSet).map(face => ({
      face,
      familyType: (fontMap[face] || {}).familyType || 'FCAT_GOTHIC',
      weight:     (fontMap[face] || {}).weight     || '6',
      proportion: (fontMap[face] || {}).proportion || '0',
    }));
  }

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

  // ── Base64 디코딩 (브라우저용) ─────────────────────────────────
  function b64ToUint8Array(b64) {
    const bin = atob(b64);
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }

  function b64ToUtf8(b64) {
    const bytes = b64ToUint8Array(b64);
    return new TextDecoder('utf-8').decode(bytes);
  }

  // ═══════════════════════════════════════════════════════════════
  // STEP 1: base 문서 빌드
  // ═══════════════════════════════════════════════════════════════

  function buildBaseSectionXml(doc, TEMPLATE_FILES) {
    const parser = new DOMParser();
    const baseSection = b64ToUtf8(TEMPLATE_FILES['Contents/section0.xml']);
    const sectionDoc = parser.parseFromString(baseSection, 'text/xml');
    const sec = sectionDoc.documentElement;

    const allPs = getElementsByTagNS(sec, HP, 'p');

    for (let i = allPs.length - 1; i >= 1; i--) {
      const p = allPs[i];
      if (p.parentNode === sec) sec.removeChild(p);
    }

    const dt = doc.doc_type || '서면보고';
    const date = doc.date || today();
    const dept = doc.dept || '';
    const author = doc.author || '';
    const metaParts = [dt, date, `${dept} ${author}`.trim()].filter(Boolean);
    addPara(sectionDoc, sec, metaParts.join(' | '));
    addPara(sectionDoc, sec, '');

    addTable(sectionDoc, sec, 1, 1, [[doc.title || '보고서']]);
    addPara(sectionDoc, sec, '');

    const cnt = [0, 0, 0];

    // 불릿 중복 방지: 텍스트에 이미 불릿이 있으면 제거 후 다시 붙임
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

    (doc.attachments || []).forEach((att, i) => {
      addPara(sectionDoc, sec, `붙임${i + 1}  ${att}`);
    });

    if (doc.contacts && doc.contacts.length > 0) {
      addPara(sectionDoc, sec, '');
      const rows = [['담당 부서', '담당자', '연락처']];
      for (const c of doc.contacts) rows.push([c.dept || '', c.name || '', c.tel || '']);
      addTable(sectionDoc, sec, rows.length, 3, rows);
    }

    return new XMLSerializer().serializeToString(sectionDoc);
  }

  // ═══════════════════════════════════════════════════════════════
  // STEP 1-B: 보도자료 문서 빌드
  // ═══════════════════════════════════════════════════════════════

  function buildPressSectionXml(doc, TEMPLATE_FILES) {
    const parser = new DOMParser();
    const baseSection = b64ToUtf8(TEMPLATE_FILES['Contents/section0.xml']);
    const sectionDoc = parser.parseFromString(baseSection, 'text/xml');
    const sec = sectionDoc.documentElement;

    // 기존 문단 모두 제거 (첫 번째 제외)
    const allPs = getElementsByTagNS(sec, HP, 'p');
    for (let i = allPs.length - 1; i >= 1; i--) {
      const p = allPs[i];
      if (p.parentNode === sec) sec.removeChild(p);
    }

    // ── 보도자료 스타일 ID 매핑 (원본 템플릿 header.xml 기반) ──
    const S = {
      // 헤더 테이블 영역
      hdr_label:    { pp: 37, cp: 42 },  // "보도자료" 라벨
      hdr_time_lbl: { pp: 38, cp: 44 },  // "보도시점" 라벨
      hdr_time_val: { pp: 1,  cp: 44 },  // 시점 값 (온라인/지면)
      // 제목 영역
      title:        { pp: 39, cp: 48 },  // 대제목 (h=2600, 함초롬돋움)
      subtitle:     { pp: 39, cp: 49 },  // 부제목 (h=2600, 함초롬돋움)
      // 리드문
      lead_first:   { pp: 40, cp: 50 },  // 첫 리드문 (- 으로 시작)
      lead_rest:    { pp: 41, cp: 50 },  // 나머지 리드문
      // 관련 국정과제
      policy:       { pp: 42, cp: 53 },
      // 본문 □/○ (h=1400, 함초롬바탕)
      body_h1:      { pp: 47, cp: 54 },  // □ 소제목
      body_h2:      { pp: 48, cp: 54 },  // ○ 중제목
      body_p:       { pp: 49, cp: 54 },  // 본문 문단
      body_h1_alt:  { pp: 50, cp: 54 },  // □ 소제목 (대안)
      body_sub1:    { pp: 58, cp: 60 },  // 세부항목 1
      body_sub2:    { pp: 59, cp: 60 },  // 세부항목 2
      body_sub3:    { pp: 60, cp: 60 },  // 세부항목 3
      // 담당자/붙임
      contact:      { pp: 55, cp: 64 },
      attach:       { pp: 56, cp: 66 },
      // 빈 줄
      blank:        { pp: 43, cp: 50 },
    };

    const date = doc.date || today();

    function strip(text) {
      return (text || '').replace(/^[\s]*(○|□|⇒|※|[-])\s*/, '').replace(/^[\s]*(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ)[.\s]*/, '').trim();
    }
    function ap(text, style) {
      return addPara(sectionDoc, sec, text, {
        paraPrIDRef: style.pp,
        charPrIDRef: style.cp,
      });
    }

    // ── 보도자료 헤더 표 (보도자료 + 보도시점) ──
    // 원본: 2x2 테이블 — row0: [보도자료 라벨, 빈칸], row1: [보도시점, 시간값]
    const onlineTime = doc.press_time_online || `${date} 12:00 이후`;
    const printTime = doc.press_time_print || `${date} 조간부터`;
    addTable(sectionDoc, sec, 2, 2, [
      ['보도자료', ''],
      ['보도시점', `(온라인) ${onlineTime}\n(지  면) ${printTime}`],
    ]);
    ap('', S.blank);

    // ── 대제목 + 부제목 ──
    ap(doc.title || '보도자료 제목', S.title);
    if (doc.subtitle) {
      ap(doc.subtitle, S.subtitle);
    }
    ap('', S.blank);

    // ── 리드문 (- 으로 시작하는 요약) ──
    if (doc.lead_lines && doc.lead_lines.length > 0) {
      doc.lead_lines.forEach((line, i) => {
        const style = i === 0 ? S.lead_first : S.lead_rest;
        ap(`- ${strip(line)}`, style);
      });
      ap('', S.blank);
    }

    // ── 관련 국정과제 ──
    if (doc.policy_ref) {
      ap(` 【관련 국정과제】 ${doc.policy_ref}`, S.policy);
      ap('', S.blank);
    }

    // ── 본문 (□/○ 패턴) ──
    const cnt = [0, 0, 0];
    let h1Count = 0;

    function render(s, level) {
      cnt[level - 1]++;
      for (let i = level; i < 3; i++) cnt[i] = 0;
      const heading = strip(s.heading || '');

      if (level === 1) {
        // □ 소제목 — 번갈아가며 pp=47, pp=50 사용 (원본 패턴)
        const h1Style = (h1Count++ % 2 === 0) ? S.body_h1 : S.body_h1_alt;
        ap(`□ ${heading}`, h1Style);
        for (const sub of (s.subsections || [])) render(sub, level + 1);
        for (const para of (s.paragraphs || [])) ap(` ○ ${strip(para)}`, S.body_p);
      } else if (level === 2) {
        ap(` ○ ${heading}`, S.body_h2);
        for (const para of (s.paragraphs || [])) ap(`  - ${strip(para)}`, S.body_p);
        for (const sub of (s.subsections || [])) render(sub, level + 1);
      } else {
        const num = NUM_L3[Math.min(cnt[2] - 1, NUM_L3.length - 1)];
        // 세부항목 — pp=58/59/60 순환
        const subStyles = [S.body_sub1, S.body_sub2, S.body_sub3];
        const subStyle = subStyles[Math.min(cnt[2] - 1, subStyles.length - 1)];
        ap(`  ${num}. ${heading}`, subStyle);
        for (const para of (s.paragraphs || [])) ap(`   - ${strip(para)}`, S.body_p);
      }

      for (const note of (s.notes || [])) ap(`※ ${strip(note)}`, S.body_p);
      for (const conc of (s.conclusions || [])) ap(`⇒ ${strip(conc)}`, S.body_p);

      if (s.table && s.table.rows && s.table.rows.length > 0) {
        ap('', S.blank);
        const colCount = Math.max(...s.table.rows.map(r => r.length));
        addTable(sectionDoc, sec, s.table.rows.length, colCount, s.table.rows);
        ap('', S.blank);
      }
    }

    for (const s of (doc.sections || [])) render(s, 1);

    // ── 담당자 표 ──
    if (doc.contacts && doc.contacts.length > 0) {
      ap('', S.blank);
      const rows = [['담당 부서', '담당자', '연락처']];
      for (const c of doc.contacts) rows.push([c.dept || '', c.name || '', c.tel || '']);
      addTable(sectionDoc, sec, rows.length, 3, rows);
    }

    // ── 붙임 ──
    if (doc.attachments && doc.attachments.length > 0) {
      ap('', S.blank);
      doc.attachments.forEach((att, i) => {
        ap(`붙임${i + 1}  ${att}`, S.attach);
      });
    }

    return new XMLSerializer().serializeToString(sectionDoc);
  }

  function addPara(xmlDoc, parent, text, opts) {
    const o = opts || {};
    const pp = o.paraPrIDRef != null ? String(o.paraPrIDRef) : '0';
    const cp = o.charPrIDRef != null ? String(o.charPrIDRef) : '0';

    const p = xmlDoc.createElementNS(HP, 'hp:p');
    p.setAttribute('id', String(Math.floor(Math.random() * 4294967295)));
    p.setAttribute('paraPrIDRef', pp);
    p.setAttribute('styleIDRef', '0');
    p.setAttribute('pageBreak', '0');
    p.setAttribute('columnBreak', '0');
    p.setAttribute('merged', '0');

    const run = xmlDoc.createElementNS(HP, 'hp:run');
    run.setAttribute('charPrIDRef', cp);
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

  function patchHeader(headerXml, opts) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(headerXml, 'text/xml');
    const styleMap = {};
    const bfMap = {};

    const CHAR_STYLES = buildCharStyles(opts);
    const PARA_STYLES = buildParaStyles(opts);
    const REQUIRED_FONTS = buildRequiredFonts(opts);

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
        const underline = getElementsByTagNS(cp, HH, 'underline')[0];
        if (underline) cp.insertBefore(boldEl, underline);
        else cp.appendChild(boldEl);
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

    addBorderFill('title_box', [
      ['leftBorder', 'SOLID', '0.12 mm', '#000000'], ['rightBorder', 'SOLID', '0.5 mm', '#000000'],
      ['topBorder', 'SOLID', '0.12 mm', '#000000'], ['bottomBorder', 'SOLID', '0.5 mm', '#000000'],
    ], opts.colors.title_bg);

    addBorderFill('sec_num_box', [
      ['leftBorder', 'SOLID', '0.1 mm', opts.colors.navy], ['rightBorder', 'SOLID', '0.1 mm', opts.colors.navy],
      ['topBorder', 'SOLID', '0.1 mm', opts.colors.navy], ['bottomBorder', 'SOLID', '0.1 mm', opts.colors.navy],
    ], opts.colors.navy);

    addBorderFill('sec_gap', [
      ['leftBorder', 'NONE', '0.12 mm', '#000000'], ['rightBorder', 'NONE', '0.12 mm', '#000000'],
      ['topBorder', 'NONE', '0.12 mm', '#000000'], ['bottomBorder', 'NONE', '0.12 mm', '#000000'],
    ], null);

    addBorderFill('sec_title_box', [
      ['leftBorder', 'NONE', '0.25 mm', opts.colors.navy_line], ['rightBorder', 'NONE', '0.25 mm', opts.colors.navy_line],
      ['topBorder', 'NONE', '0.25 mm', opts.colors.navy_line], ['bottomBorder', 'SOLID', '0.5 mm', opts.colors.navy_line],
    ], null);

    addBorderFill('sec_hdr_tbl', [
      ['leftBorder', 'NONE', '0.12 mm', '#000000'], ['rightBorder', 'NONE', '0.12 mm', '#000000'],
      ['topBorder', 'NONE', '0.12 mm', '#000000'], ['bottomBorder', 'NONE', '0.12 mm', '#000000'],
    ], null);

    addBorderFill('tbl_hdr', [
      ['leftBorder', 'SOLID', '0.12 mm', '#000000'], ['rightBorder', 'SOLID', '0.12 mm', '#000000'],
      ['topBorder', 'SOLID', '0.4 mm', '#000000'], ['bottomBorder', 'SOLID', '0.4 mm', '#000000'],
    ], null);

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

    // ── 표 서식 적용 ──
    const tables = getElementsByTagNS(sectionDoc, HP, 'tbl');
    for (const tbl of tables) {
      const rows = parseInt(tbl.getAttribute('rowCnt') || '0');
      const cols = parseInt(tbl.getAttribute('colCnt') || '0');

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
      if (p.parentNode && p.parentNode.localName === 'subList') continue;
      if (getElementsByTagNS(p, HP, 'tbl').length > 0) continue;

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

  function isPressTitle(text) {
    // 보도자료 대제목: 길고 □/○로 시작하지 않는 첫 번째 의미 있는 문단
    return text.length > 15 && !text.startsWith('□') && !text.startsWith('○') &&
      !text.startsWith('-') && !text.startsWith('※') && !text.startsWith('⇒') &&
      !text.startsWith('보도') && !text.startsWith('【') && !text.startsWith('붙임');
  }

  function applyTitleTable(tbl, styleMap, bfMap) {
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

  function applySecHeaderTable(tbl, styleMap, bfMap) {
    tbl.setAttribute('borderFillIDRef', bfMap.sec_hdr_tbl || '1');
    const sz = getElementsByTagNS(tbl, HP, 'sz')[0];
    if (sz) sz.setAttribute('width', String(DOC_TEXT_WIDTH));

    const trs = getElementsByTagNS(tbl, HP, 'tr');
    if (!trs.length) return;
    const tcs = getElementsByTagNS(trs[0], HP, 'tc');
    if (tcs.length < 3) return;

    const w0 = 2573, w1 = 566, w2 = DOC_TEXT_WIDTH - w0 - w1;

    tcs[0].setAttribute('borderFillIDRef', bfMap.sec_num_box || '1');
    const csz0 = getElementsByTagNS(tcs[0], HP, 'cellSz')[0];
    if (csz0) { csz0.setAttribute('width', String(w0)); csz0.setAttribute('height', '2466'); }
    for (const run of getElementsByTagNS(tcs[0], HP, 'run')) run.setAttribute('charPrIDRef', styleMap.cp_sec_num || '0');
    for (const p of getElementsByTagNS(tcs[0], HP, 'p')) p.setAttribute('paraPrIDRef', styleMap.pp_sec_num || '0');

    tcs[1].setAttribute('borderFillIDRef', bfMap.sec_gap || '1');
    const csz1 = getElementsByTagNS(tcs[1], HP, 'cellSz')[0];
    if (csz1) { csz1.setAttribute('width', String(w1)); csz1.setAttribute('height', '2466'); }

    tcs[2].setAttribute('borderFillIDRef', bfMap.sec_title_box || '1');
    const csz2 = getElementsByTagNS(tcs[2], HP, 'cellSz')[0];
    if (csz2) { csz2.setAttribute('width', String(w2)); csz2.setAttribute('height', '2466'); }
    for (const run of getElementsByTagNS(tcs[2], HP, 'run')) run.setAttribute('charPrIDRef', styleMap.cp_sec_title || '0');
    for (const p of getElementsByTagNS(tcs[2], HP, 'p')) p.setAttribute('paraPrIDRef', styleMap.pp_sec_title || '0');

    for (const sl of getElementsByTagNS(tbl, HP, 'subList')) sl.setAttribute('vertAlign', 'CENTER');
  }

  function applyDataTable(tbl, styleMap, bfMap) {
    const colCnt = parseInt(tbl.getAttribute('colCnt') || '3');
    const colW = Math.floor(DOC_TEXT_WIDTH / colCnt);

    const sz = getElementsByTagNS(tbl, HP, 'sz')[0];
    if (sz) sz.setAttribute('width', String(DOC_TEXT_WIDTH));

    const trs = getElementsByTagNS(tbl, HP, 'tr');
    trs.forEach((tr, ri) => {
      const tcs = getElementsByTagNS(tr, HP, 'tc');
      tcs.forEach((tc) => {
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
  // STEP 4: ZIP 조립 (브라우저용 — Blob 반환)
  // ═══════════════════════════════════════════════════════════════

  async function createGovHwpx(doc, templateFiles, styleOptions) {
    // 인자 유연 처리: createGovHwpx(doc), createGovHwpx(doc, opts), createGovHwpx(doc, tmpl, opts)
    if (templateFiles && !styleOptions && !templateFiles['mimetype']) {
      styleOptions = templateFiles;
      templateFiles = null;
    }

    // 템플릿 타입에 따라 기본 템플릿 선택
    const templateType = doc._templateType || 'report';  // 'report' | 'press'

    if (!templateFiles) {
      if (templateType === 'press' && typeof HWPX_PRESS_RELEASE_TEMPLATE !== 'undefined') {
        templateFiles = HWPX_PRESS_RELEASE_TEMPLATE;
      } else if (typeof HWPX_TEMPLATE_FILES !== 'undefined') {
        templateFiles = HWPX_TEMPLATE_FILES;
      }
    }
    if (!templateFiles) throw new Error('HWPX 템플릿 파일이 필요합니다.');

    const opts = mergeOptions(styleOptions);

    // 1. Build base section (템플릿 타입에 따라 분기)
    const baseSectionXml = templateType === 'press'
      ? buildPressSectionXml(doc, templateFiles)
      : buildBaseSectionXml(doc, templateFiles);

    // 2. Load and patch header
    const baseHeaderXml = b64ToUtf8(templateFiles['Contents/header.xml']);
    const { headerXml, styleMap, bfMap } = patchHeader(baseHeaderXml, opts);

    // 3. Patch section (보도자료는 buildPressSectionXml에서 이미 스타일 적용됨, patchSection 건너뜀)
    const sectionXml = templateType === 'press'
      ? baseSectionXml
      : patchSection(baseSectionXml, styleMap, bfMap, doc.title || '', doc);

    // 4. Assemble ZIP
    const zip = new JSZip();

    for (const [filename, b64data] of Object.entries(templateFiles)) {
      if (filename === 'Contents/header.xml' || filename === 'Contents/section0.xml') continue;

      const data = b64ToUint8Array(b64data);
      if (filename === 'mimetype') {
        zip.file(filename, data, { compression: 'STORE', createFolders: false });
      } else {
        zip.file(filename, data, { createFolders: false });
      }
    }

    zip.file('Contents/header.xml', headerXml, { createFolders: false });
    zip.file('Contents/section0.xml', sectionXml, { createFolders: false });

    const blob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      mimeType: 'application/hwp+zip',
    });

    return blob;
  }

  // ── 다운로드 헬퍼 ──
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || '보고서.hwpx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ── Public API ──
  root.HwpxGenerator = {
    createGovHwpx: createGovHwpx,
    downloadBlob: downloadBlob,
    today: today,
    DEFAULTS: DEFAULTS,
    getDefaults: function() { return JSON.parse(JSON.stringify(DEFAULTS)); },
    TEMPLATE_TYPES: {
      report: { label: '서면보고', description: '행정안전부 서면보고 양식' },
      press:  { label: '보도자료', description: '행정안전부 보도자료 양식' },
    },
  };

})(typeof window !== 'undefined' ? window : this);
