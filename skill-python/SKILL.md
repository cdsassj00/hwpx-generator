---
name: hwpx-gov-doc
description: >
  행정안전부 'AI 친화 행정문서 작성 가이드라인' 기반으로 공공 보고서를 HWPX 파일로 자동 생성하는 스킬.
  외부 라이브러리 의존성 없이(Zero-Dependency) Python 표준 라이브러리만으로 동작하므로
  Timely, Claude Code, 로컬 등 어떤 환경에서든 손상 없는 HWPX 파일을 생성함.
  사용자가 "HWPX", "한글 문서", "공공 보고서", "행정 문서", "서면보고", "보도자료",
  "공문", "보고서 작성", "AI로 보고서"를 언급하면 항상 이 스킬을 사용한다.
---

# HWPX 공공행정문서 자동생성 스킬 (Zero-Dependency)

## 환경 요구사항

```bash
# 설치할 것 없음! Python 3.7+ 표준 라이브러리만 사용
python scripts/generate_hwpx.py
```

- **의존성 없음**: zipfile, base64, xml.etree.ElementTree 등 표준 라이브러리만 사용
- python-hwpx 불필요, pip install 불필요
- Python 3.7 이상
- Windows/Mac/Linux/Timely/Claude Code 모든 환경 동작

## 핵심 원리

JS 브라우저 버전(https://github.com/cdsassj00/hwpx-generator)과 **동일한 검증된 base64 템플릿**을 Python에 내장.
header.xml, content.hpf 등 Hancom Office가 요구하는 모든 메타데이터가 이미 포함되어 있어
**어떤 환경에서든 "손상된 파일" 오류가 발생하지 않음**.

## 핵심 가이드라인 규칙 (반드시 적용)

1. **주어·서술어 명확** — 개조식 금지, 완전한 문장으로 서술 (~함, ~임, ~예정임)
2. **표준 번호체계** — Ⅰ. → 1. → 가. → 1) 순서 준수
3. **셀 병합 금지** — 단순 표 구조만 사용
4. **불필요한 서식 제거** — 박스·음영 최소화
5. **특수기호 대신 번호체계** — ㅁ, ㅇ, - 단독 사용 금지

## 워크플로우

### Step 1 — 입력 수집

사용자로부터 받을 정보:
```
필수: 문서 제목, 추진 배경, 주요 내용
선택: 작성 부서, 담당자, 작성일, 문서 종류(서면보고/보도자료), 담당자 연락처, 붙임 사항
```

### Step 2 — 내용 정제 (가이드라인 적용)

입력 내용을 변환:
- "AI 도입 필요" → "정부는 AI를 도입할 필요가 있음."
- 주어 없는 문장 → 주어 추가
- 모든 문장 끝 → ~함 / ~임 / ~예정임 / ~있음

### Step 3 — 스크립트 실행

**CLI에서 바로 실행:**

```bash
cd skill-python/scripts
python generate_hwpx.py                    # 서면보고 데모
python generate_hwpx.py -t press           # 보도자료 데모
python generate_hwpx.py -j my_data.json    # 내 JSON 파일로 생성
```

**Python 코드에서 호출:**

```python
import sys, os

# ⚠️ 스킬 scripts 폴더로 경로 이동 (import를 위해 필수)
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'skill-python', 'scripts'))
# 또는 절대경로: sys.path.insert(0, '/path/to/hwpx-generator/skill-python/scripts')

from generate_hwpx import create_gov_hwpx

# 서면보고 예시
doc = {
    "title": "AI시대 행정문서 작성 가이드라인(안) 보고",
    "doc_type": "서면보고",
    "date": "2026. 4. 6.(월)",
    "dept": "혁신행정담당관",
    "author": "박은희 사무관",
    "sections": [
        {
            "heading": "추진 배경",
            "paragraphs": [
                "정부는 사람과 AI 모두 쉽게 읽고 작성할 수 있는 표준화된 문서 작성 가이드라인을 수립할 필요가 있음.",
            ],
            "table": {
                "rows": [
                    ["정보 종류", "AI 인식 수준"],
                    ["일반 문자", "정확하게 인식"],
                ]
            }
        },
        {
            "heading": "주요 내용",
            "paragraphs": ["행정안전부는 전 부서에 가이드라인을 배포하고 시범 실시함."],
            "subsections": [
                {
                    "heading": "세부 추진 사항",
                    "paragraphs": ["주어와 서술어를 명확히 기술하여 모호함을 제거함."]
                }
            ]
        }
    ],
    "contacts": [
        {"dept": "혁신행정담당관", "name": "박은희 사무관", "tel": "044-205-1473"}
    ],
    "output": "보고서.hwpx"
}

create_gov_hwpx(doc)
```

```python
# 보도자료 예시
doc = {
    "title": "보도자료 헤드라인",
    "subtitle": "부제목",
    "doc_type": "보도자료",
    "press_time_online": "2026. 4. 1.(수) 12:00 이후",
    "press_time_print": "2026. 4. 2.(목) 조간부터",
    "lead_lines": ["리드문 1", "리드문 2"],
    "policy_ref": "관련 정책명",
    "sections": [
        {
            "heading": "소제목",
            "subsections": [
                {
                    "heading": "세부항목",
                    "paragraphs": ["본문 내용"]
                }
            ]
        }
    ],
    "contacts": [{"dept": "담당부서", "name": "담당자", "tel": "전화번호"}],
    "attachments": ["첨부자료명"],
    "output": "보도자료.hwpx"
}

create_gov_hwpx(doc)
```

### Step 4 — 파일 제공

생성된 .hwpx 파일 경로를 사용자에게 알려준다.

## Timely/래퍼 플랫폼에서 사용 시

파일을 바이트로 직접 반환해야 할 경우:

```python
from generate_hwpx import create_gov_hwpx_bytes
import base64

hwpx_bytes = create_gov_hwpx_bytes(doc)

# base64로 인코딩해서 전달 (바이너리 손상 방지)
b64_str = base64.b64encode(hwpx_bytes).decode('ascii')
```

## JSON 스키마

### 서면보고
```json
{
  "title": "보고서 제목",
  "doc_type": "서면보고",
  "date": "2026. 4. 6.(월)",
  "dept": "부서명",
  "author": "작성자",
  "sections": [
    {
      "heading": "대제목 (Ⅰ. 자동)",
      "paragraphs": ["본문"],
      "subsections": [
        {
          "heading": "중제목 (□ 자동)",
          "paragraphs": ["본문"],
          "subsections": [
            {
              "heading": "소제목 (가. 자동)",
              "paragraphs": ["본문"]
            }
          ]
        }
      ],
      "table": {"rows": [["헤더1", "헤더2"], ["데이터1", "데이터2"]]},
      "notes": ["주석"],
      "conclusions": ["결론"]
    }
  ],
  "contacts": [{"dept": "부서", "name": "이름", "tel": "전화번호"}],
  "attachments": ["첨부자료명"],
  "output": "파일명.hwpx"
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
  "lead_lines": ["리드문 1", "리드문 2"],
  "policy_ref": "관련 국정과제",
  "sections": [...],
  "contacts": [...],
  "attachments": [...],
  "output": "파일명.hwpx"
}
```

## 파일 구조

```
hwpx-skill/
├── SKILL.md          ← 이 파일
├── scripts/
│   ├── generate_hwpx.py       ← 메인 생성기 (의존성 없음)
│   ├── report_template.py     ← 서면보고 base64 템플릿
│   └── press_template.py      ← 보도자료 base64 템플릿
└── references/
    ├── guideline_rules.md     ← 가이드라인 규칙
    └── hwpx_format.md         ← HWPX 포맷 설명
```
