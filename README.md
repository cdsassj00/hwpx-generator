# hwpx-generator

**AI 친화 한글(.hwpx) 공공행정문서 생성기**

행정안전부 「AI 친화 행정문서 작성 가이드라인」 기반으로 한글(.hwpx) 파일을 자동 생성합니다.
JSON 구조만 전달하면 표준 번호체계(Ⅰ/□/○), 서식, 표가 적용된 완성 문서를 만들어줍니다.

## 지원 문서 유형

| 유형 | 설명 |
|------|------|
| **서면보고** | Ⅰ/Ⅱ 대제목 표 + □/○ 본문 + ⇒ 결론 |
| **보도자료** | 보도시점 표 + 제목/부제/리드문 + □/○ 본문 + 담당자 표 |

## 설치

```bash
npm install hwpx-generator
```

## 빠른 시작 (Node.js)

```javascript
const { createGovHwpx, today } = require('hwpx-generator');

// 서면보고 예시
const doc = {
  title: '디지털 전환 추진 현황 보고',
  doc_type: '서면보고',
  dept: '디지털정부혁신실',
  author: '홍길동',
  date: today(),
  sections: [
    {
      heading: '추진 배경',
      subsections: [
        {
          heading: '디지털 전환 필요성',
          paragraphs: [
            '코로나19 이후 비대면 행정서비스 수요가 급증하고 있다.',
            '주요국 대비 디지털 정부 경쟁력 강화가 시급한 상황이다.'
          ]
        }
      ],
      conclusions: ['디지털 전환 가속화를 위한 범정부 추진체계 필요']
    },
    {
      heading: '주요 추진 내용',
      subsections: [
        {
          heading: 'AI 행정서비스 도입',
          paragraphs: ['민원 자동 분류, 문서 요약 등 AI 기반 서비스를 도입한다.']
        }
      ]
    }
  ],
  contacts: [{ dept: '디지털정부혁신실', name: '홍길동', tel: '044-205-1234' }]
};

createGovHwpx(doc, '보고서.hwpx')
  .then(path => console.log('생성 완료:', path))
  .catch(err => console.error(err));
```

## 브라우저에서 사용

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="hwpx_template_browser.js"></script>
<script src="hwpx_press_template.js"></script>
<script src="generate_hwpx_browser.js"></script>

<script>
const doc = { title: '보고서 제목', sections: [...] };

HwpxGenerator.createGovHwpx(doc).then(blob => {
  HwpxGenerator.downloadBlob(blob, '보고서.hwpx');
});
</script>
```

## JSON 구조 — 서면보고

```json
{
  "title": "보고서 제목 (~보고)",
  "doc_type": "서면보고",
  "dept": "부서명",
  "author": "작성자",
  "date": "2026. 04. 01.",
  "sections": [
    {
      "heading": "대제목 텍스트 (Ⅰ, Ⅱ 자동)",
      "subsections": [
        {
          "heading": "중제목 텍스트 (□ 자동)",
          "paragraphs": ["본문 (○ 자동)", "두 번째 문단"]
        }
      ],
      "conclusions": ["결론 (⇒ 자동)"],
      "notes": ["주석 (※ 자동)"]
    }
  ],
  "contacts": [{ "dept": "부서", "name": "이름", "tel": "연락처" }]
}
```

## JSON 구조 — 보도자료

```json
{
  "_templateType": "press",
  "title": "보도자료 헤드라인",
  "subtitle": "부제목",
  "press_time_online": "2026. 4. 1.(수) 12:00 이후",
  "press_time_print": "2026. 4. 2.(목) 조간부터",
  "lead_lines": ["리드문 1", "리드문 2"],
  "policy_ref": "관련 정책명",
  "sections": [
    {
      "heading": "소제목 (□ 자동)",
      "subsections": [
        {
          "heading": "세부항목 (○ 자동)",
          "paragraphs": ["본문 내용"]
        }
      ]
    }
  ],
  "contacts": [{ "dept": "담당부서", "name": "담당자", "tel": "전화번호" }],
  "attachments": ["첨부자료명"]
}
```

## 주요 규칙

- heading, paragraphs에 **불릿 기호를 넣지 마세요** (□, ○, Ⅰ 등). 시스템이 자동 추가합니다.
- sections는 최소 2개 이상 권장
- paragraphs는 서술형 문장으로 작성

## 서식 커스텀

```javascript
const opts = {
  colors: { navy: '#003366', title_bg: '#DFEAF5', navy_line: '#315F97' },
  sizes: { title: 2000, body: 1500, h2: 1600 },
  spacing: { body_line: 160 }
};

createGovHwpx(doc, '보고서.hwpx', opts);
```

## 라이브 데모

[https://cdsahwpx.netlify.app](https://cdsahwpx.netlify.app)

## 개발 및 배포

**한국데이터사이언티스트협회 (Korea Data Scientist Association)**

- 홈페이지: [https://cdsa.kr](https://cdsa.kr)
- 담당: 신성진 (sjshin@cdsa.or.kr)
- 행정안전부 AI전문인재 양성과정 교육 자료로 개발

## 라이선스

MIT
