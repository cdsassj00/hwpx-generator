"""
HWPX 공공행정문서 자동생성기 (Zero-Dependency)
==============================================
외부 라이브러리 없이 Python 표준 라이브러리만으로 HWPX 생성.
JS 브라우저 버전과 동일한 검증된 base64 템플릿을 내장하여
어떤 환경(Timely, Claude Code, 로컬)에서든 손상 없는 파일을 보장.

설치: 없음 (Python 3.7+ 표준 라이브러리만 사용)
실행: python generate_hwpx.py
"""

import base64
import io
import os
import re
import sys
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

# ── 네임스페이스 ──────────────────────────────────────────────
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HS = "http://www.hancom.co.kr/hwpml/2011/section"
HH = "http://www.hancom.co.kr/hwpml/2011/head"
HC = "http://www.hancom.co.kr/hwpml/2011/core"

# ElementTree에 네임스페이스 등록
for prefix, uri in [("hp", HP), ("hs", HS), ("hh", HH), ("hc", HC)]:
    ET.register_namespace(prefix, uri)

# ── 번호 체계 ──────────────────────────────────────────────────
NUM_L1 = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ"]
NUM_L3 = ["가", "나", "다", "라", "마", "바", "사", "아"]
BULLET_H2 = "□"
BULLET_L1 = "○"
BULLET_L2 = "-"

DOC_TEXT_WIDTH = 42520


def _today():
    d = datetime.now()
    wd = ["월", "화", "수", "목", "금", "토", "일"][d.weekday()]
    return d.strftime(f"%Y. %m. %d.({wd})")


def _strip(text):
    """불릿/번호 접두사 제거"""
    text = re.sub(r"^[\s]*(○|□|⇒|※|[-])\s*", "", text or "")
    text = re.sub(r"^[\s]*(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ)[.\s]*", "", text)
    return text.strip()


def _rand_id():
    import random
    return str(random.randint(100000000, 4294967295))


# ═══════════════════════════════════════════════════════════════
# XML 빌더 헬퍼
# ═══════════════════════════════════════════════════════════════

def _add_para(parent, text, pp="0", cp="0"):
    """<hp:p> 단락 추가"""
    p = ET.SubElement(parent, f"{{{HP}}}p")
    p.set("id", _rand_id())
    p.set("paraPrIDRef", str(pp))
    p.set("styleIDRef", "0")
    p.set("pageBreak", "0")
    p.set("columnBreak", "0")
    p.set("merged", "0")

    run = ET.SubElement(p, f"{{{HP}}}run")
    run.set("charPrIDRef", str(cp))
    t = ET.SubElement(run, f"{{{HP}}}t")
    t.text = text
    return p


def _add_table(parent, rows_data, col_count=None):
    """<hp:tbl> 표 추가 (셀 병합 없음)"""
    if not rows_data:
        return
    row_cnt = len(rows_data)
    col_cnt = col_count or max(len(r) for r in rows_data)
    col_w = DOC_TEXT_WIDTH // col_cnt
    row_h = 3600

    # 표를 감싸는 단락
    p = ET.SubElement(parent, f"{{{HP}}}p")
    p.set("id", _rand_id())
    for attr in ["paraPrIDRef", "styleIDRef", "pageBreak", "columnBreak", "merged"]:
        p.set(attr, "0")

    run = ET.SubElement(p, f"{{{HP}}}run")
    run.set("charPrIDRef", "0")

    tbl = ET.SubElement(run, f"{{{HP}}}tbl")
    tbl.set("id", _rand_id())
    for k, v in {
        "zOrder": "0", "numberingType": "TABLE", "textWrap": "TOP_AND_BOTTOM",
        "textFlow": "BOTH_SIDES", "lock": "0", "dropcapstyle": "None",
        "pageBreak": "CELL", "repeatHeader": "0", "rowCnt": str(row_cnt),
        "colCnt": str(col_cnt), "cellSpacing": "0", "borderFillIDRef": "3",
        "noAdjust": "0",
    }.items():
        tbl.set(k, v)

    sz = ET.SubElement(tbl, f"{{{HP}}}sz")
    sz.set("width", str(DOC_TEXT_WIDTH))
    sz.set("widthRelTo", "ABSOLUTE")
    sz.set("height", str(row_h * row_cnt))
    sz.set("heightRelTo", "ABSOLUTE")
    sz.set("protect", "0")

    pos = ET.SubElement(tbl, f"{{{HP}}}pos")
    for k, v in {
        "treatAsChar": "1", "affectLSpacing": "0", "flowWithText": "1",
        "allowOverlap": "0", "holdAnchorAndSO": "0", "vertRelTo": "PARA",
        "horzRelTo": "COLUMN", "vertAlign": "TOP", "horzAlign": "LEFT",
        "vertOffset": "0", "horzOffset": "0",
    }.items():
        pos.set(k, v)

    for margin_name in ["outMargin", "inMargin"]:
        m = ET.SubElement(tbl, f"{{{HP}}}{margin_name}")
        for d in ["left", "right", "top", "bottom"]:
            m.set(d, "0")

    for ri in range(row_cnt):
        tr = ET.SubElement(tbl, f"{{{HP}}}tr")
        for ci in range(col_cnt):
            text = str(rows_data[ri][ci]) if ri < len(rows_data) and ci < len(rows_data[ri]) else ""

            tc = ET.SubElement(tr, f"{{{HP}}}tc")
            for k, v in {
                "name": "", "header": "0", "hasMargin": "0",
                "protect": "0", "editable": "0", "dirty": "1",
                "borderFillIDRef": "3",
            }.items():
                tc.set(k, v)

            sub_list = ET.SubElement(tc, f"{{{HP}}}subList")
            for k, v in {
                "id": "", "textDirection": "HORIZONTAL", "lineWrap": "BREAK",
                "vertAlign": "CENTER", "linkListIDRef": "0",
                "linkListNextIDRef": "0", "textWidth": "0",
                "textHeight": "0", "hasTextRef": "0", "hasNumRef": "0",
            }.items():
                sub_list.set(k, v)

            cell_p = ET.SubElement(sub_list, f"{{{HP}}}p")
            cell_p.set("paraPrIDRef", "0")
            cell_p.set("styleIDRef", "0")
            cell_p.set("pageBreak", "0")
            cell_p.set("columnBreak", "0")
            cell_p.set("merged", "0")
            cell_p.set("id", _rand_id())

            cell_run = ET.SubElement(cell_p, f"{{{HP}}}run")
            cell_run.set("charPrIDRef", "0")
            cell_t = ET.SubElement(cell_run, f"{{{HP}}}t")
            cell_t.text = text

            addr = ET.SubElement(tc, f"{{{HP}}}cellAddr")
            addr.set("colAddr", str(ci))
            addr.set("rowAddr", str(ri))

            span = ET.SubElement(tc, f"{{{HP}}}cellSpan")
            span.set("colSpan", "1")
            span.set("rowSpan", "1")

            cell_sz = ET.SubElement(tc, f"{{{HP}}}cellSz")
            cell_sz.set("width", str(col_w))
            cell_sz.set("height", str(row_h))

            cell_margin = ET.SubElement(tc, f"{{{HP}}}cellMargin")
            for d in ["left", "right", "top", "bottom"]:
                cell_margin.set(d, "0")


# ═══════════════════════════════════════════════════════════════
# 템플릿 로딩
# ═══════════════════════════════════════════════════════════════

def _load_template(doc_type="서면보고"):
    """검증된 base64 템플릿 로드"""
    script_dir = os.path.dirname(os.path.abspath(__file__))

    if doc_type == "보도자료":
        from press_template import PRESS_TEMPLATE
        return PRESS_TEMPLATE
    else:
        from report_template import REPORT_TEMPLATE
        return REPORT_TEMPLATE


# ═══════════════════════════════════════════════════════════════
# 서면보고 section0.xml 빌드
# ═══════════════════════════════════════════════════════════════

def _build_report_section(doc, template_files):
    """서면보고용 section0.xml 생성"""
    # 템플릿의 section0.xml 파싱
    base_xml = base64.b64decode(template_files["Contents/section0.xml"]).decode("utf-8")
    root = ET.fromstring(base_xml)

    # 기존 단락 제거 (첫 번째 제외 - secPr 포함)
    paras = root.findall(f"{{{HP}}}p")
    for p in paras[1:]:
        root.remove(p)

    # 메타 줄
    dt = doc.get("doc_type", "서면보고")
    date = doc.get("date", _today())
    dept = doc.get("dept", "")
    author = doc.get("author", "")
    meta_parts = [p for p in [dt, date, f"{dept} {author}".strip()] if p]
    _add_para(root, " | ".join(meta_parts))
    _add_para(root, "")

    # 제목 표
    _add_table(root, [[doc.get("title", "보고서")]], col_count=1)
    _add_para(root, "")

    # 섹션 렌더링
    cnt = [0, 0, 0]

    def render(section, level):
        cnt[level - 1] += 1
        for i in range(level, 3):
            cnt[i] = 0

        heading = _strip(section.get("heading", ""))

        if level == 1:
            num = NUM_L1[min(cnt[0] - 1, len(NUM_L1) - 1)]
            if cnt[0] > 1:
                _add_para(root, "")
            _add_table(root, [[num, "", heading]], col_count=3)
            for sub in section.get("subsections", []):
                render(sub, level + 1)
            for para in section.get("paragraphs", []):
                _add_para(root, f"{BULLET_L1} {_strip(para)}")
        elif level == 2:
            _add_para(root, f"{BULLET_H2} {heading}")
            for para in section.get("paragraphs", []):
                _add_para(root, f"{BULLET_L1} {_strip(para)}")
            for sub in section.get("subsections", []):
                render(sub, level + 1)
        else:
            num = NUM_L3[min(cnt[2] - 1, len(NUM_L3) - 1)]
            _add_para(root, f"{num}. {heading}")
            for para in section.get("paragraphs", []):
                _add_para(root, f"{BULLET_L2} {_strip(para)}")
            for sub in section.get("subsections", []):
                render(sub, level + 1)

        for note in section.get("notes", []):
            _add_para(root, f"※ {_strip(note)}")
        for conc in section.get("conclusions", []):
            _add_para(root, f"⇒ {_strip(conc)}")

        if "table" in section:
            rows = section["table"].get("rows", [])
            if rows:
                _add_para(root, "")
                _add_table(root, rows)
                _add_para(root, "")

    for s in doc.get("sections", []):
        render(s, 1)

    # 붙임
    for i, att in enumerate(doc.get("attachments", []), 1):
        _add_para(root, f"붙임{i}  {att}")

    # 담당자 표
    contacts = doc.get("contacts", [])
    if contacts:
        _add_para(root, "")
        rows = [["담당 부서", "담당자", "연락처"]]
        for c in contacts:
            rows.append([c.get("dept", ""), c.get("name", ""), c.get("tel", "")])
        _add_table(root, rows, col_count=3)

    return ET.tostring(root, encoding="unicode", xml_declaration=True)


# ═══════════════════════════════════════════════════════════════
# 보도자료 section0.xml 빌드
# ═══════════════════════════════════════════════════════════════

def _build_press_section(doc, template_files):
    """보도자료용 section0.xml 생성"""
    base_xml = base64.b64decode(template_files["Contents/section0.xml"]).decode("utf-8")
    root = ET.fromstring(base_xml)

    # 기존 단락 제거 (첫 번째 제외)
    paras = root.findall(f"{{{HP}}}p")
    for p in paras[1:]:
        root.remove(p)

    date = doc.get("date", _today())

    # 보도자료 스타일 ID (프레스 템플릿의 header.xml 기준)
    S = {
        "title":      {"pp": "39", "cp": "48"},
        "subtitle":   {"pp": "39", "cp": "49"},
        "lead_first": {"pp": "40", "cp": "50"},
        "lead_rest":  {"pp": "41", "cp": "50"},
        "policy":     {"pp": "42", "cp": "53"},
        "body_h1":    {"pp": "47", "cp": "54"},
        "body_h2":    {"pp": "48", "cp": "54"},
        "body_p":     {"pp": "49", "cp": "54"},
        "contact":    {"pp": "55", "cp": "64"},
        "attach":     {"pp": "56", "cp": "66"},
        "blank":      {"pp": "43", "cp": "50"},
    }

    def ap(text, style):
        _add_para(root, text, pp=style["pp"], cp=style["cp"])

    # 보도자료 헤더 표
    online = doc.get("press_time_online", f"{date} 12:00 이후")
    print_t = doc.get("press_time_print", f"{date} 조간부터")
    _add_table(root, [
        ["보도자료", ""],
        ["보도시점", f"(온라인) {online}\n(지  면) {print_t}"],
    ], col_count=2)
    ap("", S["blank"])

    # 대제목 + 부제목
    ap(doc.get("title", "보도자료 제목"), S["title"])
    if doc.get("subtitle"):
        ap(doc["subtitle"], S["subtitle"])
    ap("", S["blank"])

    # 리드문
    for i, line in enumerate(doc.get("lead_lines", [])):
        style = S["lead_first"] if i == 0 else S["lead_rest"]
        ap(f"- {_strip(line)}", style)
    if doc.get("lead_lines"):
        ap("", S["blank"])

    # 관련 국정과제
    if doc.get("policy_ref"):
        ap(f" 【관련 국정과제】 {doc['policy_ref']}", S["policy"])
        ap("", S["blank"])

    # 본문
    cnt = [0, 0, 0]

    def render(s, level):
        cnt[level - 1] += 1
        for i in range(level, 3):
            cnt[i] = 0
        heading = _strip(s.get("heading", ""))

        if level == 1:
            ap(f"□ {heading}", S["body_h1"])
            for sub in s.get("subsections", []):
                render(sub, level + 1)
            for para in s.get("paragraphs", []):
                ap(f" ○ {_strip(para)}", S["body_p"])
        elif level == 2:
            ap(f" ○ {heading}", S["body_h2"])
            for para in s.get("paragraphs", []):
                ap(f"  - {_strip(para)}", S["body_p"])
            for sub in s.get("subsections", []):
                render(sub, level + 1)
        else:
            num = NUM_L3[min(cnt[2] - 1, len(NUM_L3) - 1)]
            ap(f"  {num}. {heading}", S["body_p"])
            for para in s.get("paragraphs", []):
                ap(f"   - {_strip(para)}", S["body_p"])

        for note in s.get("notes", []):
            ap(f"※ {_strip(note)}", S["body_p"])
        for conc in s.get("conclusions", []):
            ap(f"⇒ {_strip(conc)}", S["body_p"])

        if "table" in s:
            rows = s["table"].get("rows", [])
            if rows:
                ap("", S["blank"])
                _add_table(root, rows)
                ap("", S["blank"])

    for s in doc.get("sections", []):
        render(s, 1)

    # 담당자
    contacts = doc.get("contacts", [])
    if contacts:
        ap("", S["blank"])
        rows = [["담당 부서", "담당자", "연락처"]]
        for c in contacts:
            rows.append([c.get("dept", ""), c.get("name", ""), c.get("tel", "")])
        _add_table(root, rows, col_count=3)

    # 붙임
    for i, att in enumerate(doc.get("attachments", []), 1):
        ap(f"붙임{i}  {att}", S["attach"])

    return ET.tostring(root, encoding="unicode", xml_declaration=True)


# ═══════════════════════════════════════════════════════════════
# 메인 생성 함수
# ═══════════════════════════════════════════════════════════════

def create_gov_hwpx(doc: dict) -> str:
    """
    JSON dict → HWPX 파일 생성 (의존성 없음)

    Parameters:
        doc: 문서 데이터 딕셔너리
            - title: 제목
            - doc_type: "서면보고" | "보도자료"
            - date, dept, author: 메타 정보
            - sections: [{heading, paragraphs, subsections, table, notes, conclusions}]
            - contacts: [{dept, name, tel}]
            - attachments: [str]
            - output: 출력 파일 경로
            # 보도자료 전용:
            - subtitle, press_time_online, press_time_print
            - lead_lines, policy_ref

    Returns:
        저장된 파일 경로
    """
    doc_type = doc.get("doc_type", doc.get("_templateType", "서면보고"))
    if doc_type in ("press", "보도자료"):
        doc_type = "보도자료"

    template = _load_template(doc_type)

    # section0.xml 생성
    if doc_type == "보도자료":
        new_section = _build_press_section(doc, template)
    else:
        new_section = _build_report_section(doc, template)

    # HWPX ZIP 조립
    output_path = doc.get("output", "보고서.hwpx")
    buf = io.BytesIO()

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, b64_data in template.items():
            raw = base64.b64decode(b64_data)

            if filename == "mimetype":
                # mimetype은 압축하지 않음 (HWPX 규격)
                zf.writestr(filename, raw, compress_type=zipfile.ZIP_STORED)
            elif filename == "Contents/section0.xml":
                # 새로 생성한 section 사용
                zf.writestr(filename, new_section.encode("utf-8"))
            else:
                # 나머지는 템플릿 그대로
                zf.writestr(filename, raw)

    # 파일로 저장
    with open(output_path, "wb") as f:
        f.write(buf.getvalue())

    print(f"✅ HWPX 생성 완료: {output_path} ({os.path.getsize(output_path):,} bytes)")
    return output_path


def create_gov_hwpx_bytes(doc: dict) -> bytes:
    """
    HWPX를 bytes로 반환 (파일 저장 없이)
    Timely 등 래퍼 플랫폼에서 직접 바이트를 다룰 때 사용.
    """
    doc_type = doc.get("doc_type", doc.get("_templateType", "서면보고"))
    if doc_type in ("press", "보도자료"):
        doc_type = "보도자료"

    template = _load_template(doc_type)

    if doc_type == "보도자료":
        new_section = _build_press_section(doc, template)
    else:
        new_section = _build_report_section(doc, template)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, b64_data in template.items():
            raw = base64.b64decode(b64_data)
            if filename == "mimetype":
                zf.writestr(filename, raw, compress_type=zipfile.ZIP_STORED)
            elif filename == "Contents/section0.xml":
                zf.writestr(filename, new_section.encode("utf-8"))
            else:
                zf.writestr(filename, raw)

    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════
# 데모
# ═══════════════════════════════════════════════════════════════

DEMO_REPORT = {
    "title": "AI시대 행정문서 작성 가이드라인(안) 보고",
    "doc_type": "서면보고",
    "date": _today(),
    "dept": "혁신행정담당관",
    "author": "박은희 사무관",
    "sections": [
        {
            "heading": "추진 배경",
            "paragraphs": [
                "정부는 사람과 AI 모두 쉽게 읽고 작성할 수 있는 표준화된 문서 작성 가이드라인을 수립할 필요가 있음.",
                "현행 개조식·셀병합 문서 형식은 AI 모델이 문단 간 관계를 오인식하는 오류를 유발함.",
            ],
            "table": {
                "rows": [
                    ["정보 종류", "AI 인식 수준"],
                    ["일반 문자", "정확하게 인식"],
                    ["단순 표", "표 내용의 의미를 명확하게 인식"],
                ]
            },
        },
        {
            "heading": "주요 추진 내용",
            "paragraphs": [
                "행정안전부는 AI 친화적 보고서 작성을 위해 가이드라인을 전 부서에 배포하고 시범 실시함.",
            ],
            "subsections": [
                {
                    "heading": "문서 작성 원칙",
                    "paragraphs": [
                        "주어와 서술어를 명확히 기술하여 모호함을 제거함.",
                        "공문서 표준 번호체계(Ⅰ., 1., 가., 1))를 준수함.",
                    ],
                },
            ],
        },
    ],
    "contacts": [
        {"dept": "혁신행정담당관", "name": "박은희 사무관", "tel": "044-205-1473"},
    ],
    "output": "서면보고_테스트.hwpx",
}

DEMO_PRESS = {
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
                    "paragraphs": ["본문 내용"],
                }
            ],
        }
    ],
    "contacts": [{"dept": "담당부서", "name": "담당자", "tel": "전화번호"}],
    "attachments": ["첨부자료명"],
    "output": "보도자료_테스트.hwpx",
}

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="HWPX 공공행정문서 생성기 (Zero-Dependency)")
    parser.add_argument("--type", "-t", choices=["report", "press"], default="report")
    parser.add_argument("--output", "-o", default="")
    parser.add_argument("--json", "-j", default="", help="JSON 파일 경로")
    args = parser.parse_args()

    if args.json:
        import json
        with open(args.json, "r", encoding="utf-8") as f:
            doc = json.load(f)
    else:
        doc = DEMO_PRESS if args.type == "press" else DEMO_REPORT

    if args.output:
        doc["output"] = args.output

    create_gov_hwpx(doc)
