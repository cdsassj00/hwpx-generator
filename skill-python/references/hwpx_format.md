# HWPX 파일 포맷 참조

## 구조 개요

HWPX는 ZIP 아카이브 기반의 XML 문서 포맷 (한컴 한글 2014 이후).

```
document.hwpx (ZIP)
├── mimetype                    ← "application/hwp+zip" (압축 없이 첫 파일)
├── META-INF/
│   └── manifest.xml           ← 포함 파일 목록
├── Contents/
│   ├── content.hpf            ← OPF 패키지 선언 (spine 순서)
│   ├── header.xml             ← 스타일·폰트·단락 서식 정의
│   └── section0.xml           ← 본문 내용 (섹션 추가 시 section1.xml …)
└── settings.xml               ← 문서 보호 등 설정
```

## 핵심 네임스페이스

| 접두사 | URI |
|--------|-----|
| hp     | http://www.hancom.com/hwpml/2012/paragraph |
| hh     | http://www.hancom.com/hwpml/2012/core |
| opf    | http://www.idpf.org/2007/opf |

## 본문 구조 (section0.xml)

```xml
<hp:sec>
  <hp:p>          <!-- 단락 -->
    <hp:pPr>      <!-- 단락 속성 -->
      <hp:pStyle hp:val="Normal"/>
      <hp:ind hp:left="0" hp:right="0"/>
    </hp:pPr>
    <hp:r>        <!-- 텍스트 런 -->
      <hp:rPr>    <!-- 런 속성 (볼드 등) -->
        <hp:b/>
      </hp:rPr>
      <hp:t>내용</hp:t>
    </hp:r>
  </hp:p>

  <hp:tbl hp:rowCount="3" hp:colCount="2" hp:width="8000">
    <hp:tr>
      <hp:tc>
        <hp:tcPr hp:colSpan="1" hp:rowSpan="1"/>  <!-- 셀 병합 없음 -->
        <hp:p>...</hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hp:sec>
```

## 가이드라인 준수 규칙

### 금지 항목
- `hp:colSpan` > 1 또는 `hp:rowSpan` > 1 → 셀 병합 금지
- `<hp:tbl>` 안에 `<hp:tbl>` → 중첩 표 금지
- 특수기호(ㅁ, ㅇ, -, •) 단독 단락 구분 → 표준 번호체계로 대체

### 권장 항목
- `<hp:pStyle>` 반드시 명시
- 주어+서술어 완전 문장으로 `<hp:t>` 작성
- 들여쓰기는 `<hp:ind hp:left="N">` 으로 (시각적 공백 문자 금지)

## 폰트 크기 단위

HWP에서 폰트 크기 = pt × 100
- 10pt → 1000
- 12pt → 1200
- 14pt → 1400
