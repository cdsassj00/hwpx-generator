# HWPX & AI — 한글 문서 자동 생성기

**행정안전부 AI전문인재 교육용** · 한국데이터사이언티스트협회

> 자연어로 묘사하면, 한글(.hwpx) 문서가 됩니다.

라이브 데모: [https://cdsahwpx.netlify.app](https://cdsahwpx.netlify.app)

---

## 이 저장소의 구성

이 저장소는 **4가지**로 구성되어 있습니다.

```
hwpx-generator/
│
├── ❶ JS 라이브러리 ─────────── 핵심 엔진. HWPX 파일을 만드는 코드.
│   ├── generate_hwpx_browser.js     브라우저용 생성기
│   ├── generate_hwpx_node.js        Node.js용 생성기
│   ├── hwpx_template_browser.js     서면보고 템플릿 (base64)
│   ├── hwpx_template_node.js        Node.js용 동일 템플릿
│   ├── hwpx_press_template.js       보도자료 템플릿 (base64)
│   └── package.json
│
├── ❷ JS 스킬 ──────────────── Claude가 ❶을 쓸 수 있게 해주는 설명서.
│   └── skill-js/
│       ├── SKILL.md                 Claude에게 주는 사용법 + 규칙
│       └── references/              가이드라인 참고자료
│
├── ❸ Python 스킬 ──────────── 같은 기능의 Python 버전. 의존성 0.
│   └── skill-python/
│       ├── SKILL.md                 Claude에게 주는 사용법 + 규칙
│       ├── scripts/
│       │   ├── generate_hwpx.py     Python 생성기 (표준 라이브러리만)
│       │   ├── report_template.py   서면보고 템플릿 (base64)
│       │   └── press_template.py    보도자료 템플릿 (base64)
│       └── references/
│
├── ❹ 웹 데모 ──────────────── 비개발자가 브라우저에서 바로 쓰는 UI.
│   ├── index.html                   이 파일 하나가 전체 UI
│   ├── netlify.toml                 배포 설정
│   └── netlify/
│       ├── edge-functions/chat.ts   Claude API SSE 프록시
│       └── functions/chat.mjs       동기 fallback
│
└── README.md
```

### 각각이 뭔지 한 줄로

| # | 이름 | 한 줄 설명 | 누가 쓰나 |
|---|------|-----------|----------|
| ❶ | **JS 라이브러리** | JSON 넣으면 .hwpx 파일을 만들어주는 코드 | 개발자 |
| ❷ | **JS 스킬** | "❶을 이렇게 쓰면 돼"라고 Claude AI에게 알려주는 설명서 | Claude (Node.js 환경) |
| ❸ | **Python 스킬** | ❶과 같은 기능을 Python으로 다시 만든 것 + Claude 설명서 | Claude (Python 환경, Timely 등) |
| ❹ | **웹 데모** | ❶을 브라우저 UI로 감싼 것. 클릭만으로 문서 생성 | 공무원, 비개발자 |

### 관계도

```
❶ JS 라이브러리 (핵심 엔진)
    │
    ├──→ ❷ JS 스킬이 이걸 "이렇게 써" 하고 Claude에게 알려줌
    │
    └──→ ❹ 웹 데모가 이걸 <script>로 불러서 UI를 입힘

❸ Python 스킬
    └──→ ❶과 같은 base64 템플릿을 쓰되, Python으로 독립 구현
```

---

## 스킬(Skill)이 뭔가요?

**Skill**은 Claude AI가 특정 작업을 잘 하도록 모아놓은 "전문 지식 꾸러미"입니다.

SKILL.md(사용 설명서) + 코드 + 참고자료를 폴더로 묶으면 스킬이 됩니다. Claude에게 이 폴더를 주면, 사용자가 "보고서 만들어줘"라고 말했을 때 스킬 안의 코드와 규칙을 활용해서 전문가처럼 만들어줍니다.

비유하면: 요리사에게 "백종원 레시피북"을 건네주는 것. 요리사는 원래도 요리하지만, 그 책이 있으면 특정 요리를 훨씬 잘 만듭니다.

**스킬 구하는 곳:** Anthropic GitHub, Claude Code 플러그인 마켓플레이스, 직접 만들기

---

## 템플릿 원리

HWPX는 ZIP 파일입니다. 안에 XML 여러 개가 들어있어요.

```
보고서.hwpx (ZIP)
├── mimetype                 "application/hwp+zip"
├── Contents/
│   ├── header.xml           ← 폰트, 스타일 정의 (템플릿 그대로 사용)
│   ├── section0.xml         ← 실제 문서 내용 (이것만 새로 조립)
│   └── content.hpf          ← 메타데이터 (템플릿 그대로)
├── META-INF/                ← 컨테이너 (템플릿 그대로)
└── ...
```

한컴 오피스가 까다롭게 검사하는 `header.xml` 등은 실제 한컴에서 만든 원본을 base64로 인코딩해서 **그대로 사용**합니다. 내용이 담기는 `section0.xml`만 새로 조립하기 때문에 "손상된 파일" 오류가 없습니다.

❶ JS 라이브러리, ❸ Python 스킬 모두 **동일한 base64 템플릿**을 공유합니다.

---

## 빠른 시작

### ❹ 웹 데모 (가장 쉬움)

[https://cdsahwpx.netlify.app](https://cdsahwpx.netlify.app) 접속 → 주제 입력 → HWPX 다운로드

### ❶ JS 라이브러리 — 브라우저

```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="hwpx_template_browser.js"></script>
<script src="hwpx_press_template.js"></script>
<script src="generate_hwpx_browser.js"></script>

<script>
const doc = {
  title: '디지털 전환 현황 보고',
  doc_type: '서면보고',
  sections: [{ heading: '추진 배경', paragraphs: ['비대면 서비스 수요가 급증함.'] }],
  contacts: [{ dept: '디지털정부혁신실', name: '홍길동', tel: '044-205-1234' }]
};

HwpxGenerator.createGovHwpx(doc).then(blob => {
  HwpxGenerator.downloadBlob(blob, '보고서.hwpx');
});
</script>
```

### ❶ JS 라이브러리 — Node.js

```bash
cd /path/to/hwpx-generator   # 저장소 루트로 이동
npm install                   # jszip, @xmldom/xmldom 설치
node test_example.js          # 데모 실행 → 서면보고_예시.hwpx 생성
```

### ❸ Python 스킬

```bash
cd /path/to/hwpx-generator/skill-python/scripts   # 스킬 scripts 폴더로 이동
python generate_hwpx.py                  # 서면보고 데모
python generate_hwpx.py -t press         # 보도자료 데모
python generate_hwpx.py -j my_data.json  # 내 JSON으로 생성
```

---

## JSON 입력 형식

heading, paragraphs에 **불릿 기호(□, ○, Ⅰ 등)를 넣지 마세요** — 시스템이 자동 추가합니다.

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
          "paragraphs": ["본문 (○ 자동)"]
        }
      ],
      "conclusions": ["결론 (⇒ 자동)"]
    }
  ],
  "contacts": [{ "dept": "부서", "name": "이름", "tel": "연락처" }]
}
```

### 보도자료

```json
{
  "title": "헤드라인",
  "subtitle": "부제목",
  "doc_type": "보도자료",
  "press_time_online": "2026. 4. 1.(수) 12:00 이후",
  "press_time_print": "2026. 4. 2.(목) 조간부터",
  "lead_lines": ["리드문 1"],
  "sections": [{ "heading": "소제목", "paragraphs": ["본문"] }],
  "contacts": [{ "dept": "담당부서", "name": "담당자", "tel": "전화번호" }]
}
```

---

## 개발

**한국데이터사이언티스트협회** · 행정안전부 AI전문인재 양성과정 교육 자료 · MIT
