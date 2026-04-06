---
name: hwpx-gov-doc-js
description: >
  행정안전부 'AI 친화 행정문서 작성 가이드라인' 기반으로 공공 보고서를 HWPX 파일로 자동 생성하는 스킬 (JavaScript/Node.js 버전).
  검증된 base64 템플릿을 내장하여 한컴 오피스에서 정상 동작하는 서면보고·보도자료를 생성함.
  사용자가 "HWPX", "한글 문서", "공공 보고서", "행정 문서", "서면보고", "보도자료",
  "공문", "보고서 작성", "AI로 보고서"를 언급하면 이 스킬을 사용한다.
---

# HWPX 공공행정문서 자동생성 스킬 (JavaScript)

## 환경 요구사항

```bash
npm install jszip @xmldom/xmldom
```

## 지원 문서 유형

| 유형 | 설명 |
|------|------|
| **서면보고** | Ⅰ/Ⅱ 대제목 표 + □/○ 본문 + ⇒ 결론 |
| **보도자료** | 보도시점 표 + 제목/부제/리드문 + □/○ 본문 + 담당자 표 |

## 핵심 가이드라인 규칙 (반드시 적용)

1. **주어·서술어 명확** — 개조식 금지, 완전한 문장으로 서술 (~함, ~임, ~예정임)
2. **표준 번호체계** — Ⅰ. → 1. → 가. → 1) 순서 준수
3. **셀 병합 금지** — 단순 표 구조만 사용
4. **특수기호 대신 번호체계** — heading, paragraphs에 불릿 기호(□, ○, Ⅰ 등)를 넣지 마세요. 시스템이 자동 추가합니다.

## 라이브러리 파일 위치

이 스킬은 저장소 루트의 JS 라이브러리를 사용합니다:

```
(저장소 루트)/
├── generate_hwpx_node.js       ← Node.js 생성기
├── generate_hwpx_browser.js    ← 브라우저 생성기
├── hwpx_template_node.js       ← 서면보고 템플릿 (Node.js)
├── hwpx_template_browser.js    ← 서면보고 템플릿 (브라우저)
└── hwpx_press_template.js      ← 보도자료 템플릿
```

## 워크플로우

### Step 1 — 입력 수집

사용자로부터 받을 정보:
```
필수: 문서 제목, 추진 배경, 주요 내용
선택: 작성 부서, 담당자, 작성일, 문서 종류(서면보고/보도자료), 연락처, 붙임
```

### Step 2 — 내용 정제 (가이드라인 적용)

입력 내용을 변환:
- "AI 도입 필요" → "정부는 AI를 도입할 필요가 있음."
- 주어 없는 문장 → 주어 추가
- 모든 문장 끝 → ~함 / ~임 / ~예정임 / ~있음

### Step 3 — JSON 구성 후 생성

**CLI에서 바로 실행:**

```bash
cd /path/to/hwpx-generator     # 저장소 루트로 이동
npm install                     # 첫 실행 시 의존성 설치 (jszip, @xmldom/xmldom)
node test_example.js            # 데모 실행
```

**Node.js에서 사용:**

```javascript
// ⚠️ 저장소 루트의 JS 파일을 require (경로 주의)
const { createGovHwpx, today } = require('/path/to/hwpx-generator/generate_hwpx_node.js');

const doc = {
  title: 'AI시대 행정문서 작성 가이드라인(안) 보고',
  doc_type: '서면보고',    // 또는 '보도자료'
  dept: '혁신행정담당관',
  author: '박은희 사무관',
  date: today(),
  sections: [
    {
      heading: '추진 배경',
      paragraphs: [
        '정부는 표준화된 문서 작성 가이드라인을 수립할 필요가 있음.',
      ],
      table: {
        rows: [
          ['정보 종류', 'AI 인식 수준'],
          ['일반 문자', '정확하게 인식'],
        ]
      }
    },
    {
      heading: '주요 내용',
      paragraphs: ['전 부서에 가이드라인을 배포하고 시범 실시함.'],
      subsections: [
        {
          heading: '문서 작성 원칙',
          paragraphs: ['주어와 서술어를 명확히 기술함.']
        }
      ]
    }
  ],
  contacts: [
    { dept: '혁신행정담당관', name: '박은희 사무관', tel: '044-205-1473' }
  ]
};

createGovHwpx(doc, '보고서.hwpx')
  .then(path => console.log('생성 완료:', path));
```

**보도자료:**

```javascript
const doc = {
  title: '보도자료 헤드라인',
  subtitle: '부제목',
  doc_type: '보도자료',
  _templateType: 'press',
  press_time_online: '2026. 4. 1.(수) 12:00 이후',
  press_time_print: '2026. 4. 2.(목) 조간부터',
  lead_lines: ['리드문 1', '리드문 2'],
  policy_ref: '관련 정책명',
  sections: [
    {
      heading: '소제목',
      subsections: [
        { heading: '세부항목', paragraphs: ['본문 내용'] }
      ]
    }
  ],
  contacts: [{ dept: '담당부서', name: '담당자', tel: '전화번호' }],
  attachments: ['첨부자료명']
};
```

**서식 커스텀 (선택):**

```javascript
const opts = {
  colors: { navy: '#003366', title_bg: '#DFEAF5', navy_line: '#315F97' },
  fonts: { title: 'HY헤드라인M', body: '휴먼명조', ui: '맑은 고딕' },
  sizes: { title: 2000, body: 1500, h2: 1600 },
  indent: { h2: 1400, body: 2800 },
  spacing: { body_line: 160, h2_before: 2400 }
};

createGovHwpx(doc, '보고서.hwpx', opts);
```

### Step 4 — 파일 제공

생성된 .hwpx 파일 경로를 사용자에게 알려준다.

## JSON 스키마

### 서면보고
```json
{
  "title": "보고서 제목",
  "doc_type": "서면보고",
  "dept": "부서명",
  "author": "작성자",
  "sections": [
    {
      "heading": "대제목 (Ⅰ 자동)",
      "subsections": [
        {
          "heading": "중제목 (□ 자동)",
          "paragraphs": ["본문 (○ 자동)"],
          "subsections": [
            { "heading": "소제목 (가 자동)", "paragraphs": ["세부 (- 자동)"] }
          ]
        }
      ],
      "table": { "rows": [["헤더1", "헤더2"], ["값1", "값2"]] },
      "conclusions": ["결론 (⇒ 자동)"],
      "notes": ["주석 (※ 자동)"]
    }
  ],
  "contacts": [{ "dept": "부서", "name": "이름", "tel": "연락처" }],
  "attachments": ["첨부 파일명"]
}
```

### 보도자료
```json
{
  "title": "헤드라인",
  "subtitle": "부제목",
  "doc_type": "보도자료",
  "_templateType": "press",
  "press_time_online": "2026. 4. 1.(수) 12:00 이후",
  "press_time_print": "2026. 4. 2.(목) 조간부터",
  "lead_lines": ["리드문"],
  "policy_ref": "관련 국정과제",
  "sections": [{ "heading": "소제목", "paragraphs": ["본문"] }],
  "contacts": [{ "dept": "부서", "name": "이름", "tel": "연락처" }],
  "attachments": ["첨부자료명"]
}
```

## 파일 구조

```
skill-js/
├── SKILL.md              ← 이 파일
└── references/
    ├── guideline_rules.md
    └── hwpx_format.md
```

라이브러리 코드는 저장소 루트의 `generate_hwpx_node.js`, `hwpx_template_node.js` 등을 참조합니다.
