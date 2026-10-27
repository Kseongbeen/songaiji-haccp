# -*- coding: utf-8 -*-
"""
40. 한우 농장 HACCP 발표 초안 v6 (PPT)

v5(39_한우_PPT_v5.py)에서 달라진 점
- 데이터 재검증: 인증 농장 수를 '업체명+시도+시군구' 기준으로 중복 제거
  (업체명만 쓰면 다른 지역의 같은 이름 농장이 하나로 합쳐져 과소 집계된다)
- 출처 재검증: 법령 원문(국가법령정보센터 API), 정부 문서 첨부(hwpx), 논문 원문으로 수치 확인
  · 60세 이상 63%는 '축산 농가 전체' 수치 → 표기 수정
  · Grimes 외(2025)는 60,843마리, 지방 무늬는 '늘어나는 경향' → 표기 수정
  · 2026년 스마트HACCP 실제 사업은 '가금류 도축장 시스템 개발 + 도축장·집유장 8곳 보급'
  · 축사로(국립축산과학원)가 이미 약품 구매·사용·처방 기록과 이력제 연동을 제공 → 해결 A를 축사로 연계로 재설계
- 법적 검토 슬라이드 추가(개인정보, 이력정보 비밀 보호, 판매기록 규정, 출하 '차단'이 아닌 '경고')
- 모든 본문 슬라이드에 발표 대본(노트) 추가
- 디자인: 둥근 모서리 카드, 알약 모양 태그

산출: output/한우HACCP_발표초안_v6.pptx
"""

import importlib.util
import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("ppt38", os.path.join(HERE, "38_축산_PPT.py"))
v4 = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(v4)

FONT = v4.FONT
INK, SUB, MUTE, LINE, BG, WHITE = v4.INK, v4.SUB, v4.MUTE, v4.LINE, v4.BG, v4.WHITE
ACC, ACC_SOFT, GREEN, GRAY = v4.ACC, v4.ACC_SOFT, v4.GREEN, v4.GRAY
GREEN_SOFT = RGBColor(0xDF, 0xF0, 0xE8)
W, H, M = v4.W, v4.H, v4.M
text, bars = v4.text, v4.bars
OUT = os.path.join(HERE, "output")
RADIUS = Inches(0.18)


# ─────────────────────────────────────────────────────────────
# 둥근 디자인 도구
# ─────────────────────────────────────────────────────────────
def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = WHITE
    return s


def _round(sh, w, h, r=RADIUS):
    sh.adjustments[0] = min(0.5, r / min(w, h))


def box(slide, x, y, w, h, fill=BG, line=None, lines=(), anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, r=RADIUS, pad=0.2):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    _round(sh, w, h, r)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    sh.shadow.inherit = False
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(pad)
    tf.margin_top = tf.margin_bottom = Inches(0.14)
    for i, (t, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(size * 0.4)
        rr = p.add_run()
        rr.text = t
        rr.font.name = FONT
        rr.font.size = Pt(size)
        rr.font.bold = bold
        rr.font.color.rgb = color
    return sh


def pill(slide, x, y, w, h, label, fill, color, size=12, line=None):
    return box(slide, x, y, w, h, fill=fill, line=line, lines=[(label, size, True, color)],
               anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, r=h / 2, pad=0.05)


def chevron(slide, x, y, w=Inches(0.3), h=Inches(0.36)):
    sh = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = GRAY
    sh.line.fill.background()
    sh.shadow.inherit = False


def header(slide, section, title, sub=None):
    pill(slide, M, Inches(0.32), Inches(0.4) + Inches(0.165) * len(section), Inches(0.34), section, ACC_SOFT, ACC, size=11)
    text(slide, M, Inches(0.72), W - 2 * M, Inches(0.7), [(title, 25, True, INK)])
    if sub:
        text(slide, M, Inches(1.38), W - 2 * M, Inches(0.5), [(sub, 14, False, SUB)])


def table(slide, x, y, w, rows, col_w, row_h=Inches(0.42), hot_rows=(), size=13, num_cols=()):
    """둥근 카드 위에 선 없는 표를 얹는다. 강조 행은 글자색으로만 표시한다."""
    total_h = row_h * len(rows)
    box(slide, x - Inches(0.08), y - Inches(0.08), w + Inches(0.16), total_h + Inches(0.16), fill=BG)
    hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, row_h)
    _round(hdr, w, row_h, Inches(0.12))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = LINE; hdr.line.fill.background(); hdr.shadow.inherit = False
    gt = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w, total_h)
    tbl = gt.table
    tblPr = gt._element.graphic.graphicData.tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    for j, cw in enumerate(col_w):
        tbl.columns[j].width = cw
    for i, row in enumerate(rows):
        tbl.rows[i].height = row_h
        for j, v in enumerate(row):
            c = tbl.cell(i, j)
            c.fill.background()
            c.margin_left = c.margin_right = Inches(0.12)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if (j in num_cols and i > 0) else PP_ALIGN.LEFT
            rr = p.add_run()
            rr.text = str(v)
            rr.font.name = FONT
            rr.font.size = Pt(size - 1 if i == 0 else size)
            rr.font.bold = (i == 0) or (i in hot_rows)
            rr.font.color.rgb = SUB if i == 0 else (ACC if i in hot_rows else INK)
    _strip_borders(tbl)
    return tbl


def _strip_borders(tbl):
    from pptx.oxml.ns import qn
    from lxml import etree
    for row in tbl.rows:
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(tag), w="0")
                etree.SubElement(ln, qn("a:noFill"))


# ─────────────────────────────────────────────────────────────
# 참고문헌: 키 → (저자/기관, 제목, 날짜, URL). 슬라이드 등장 순서대로 번호를 매긴다.
# ─────────────────────────────────────────────────────────────
LAW = "https://www.law.go.kr/법령/"
REFS = {
    "haccp": ("한국식품안전관리인증원", "HACCP 인증업체정보 (공공데이터포털 15033311, 이용허락범위 제한 없음)", "2026-09-08 수집", "https://www.data.go.kr/data/15033311/openapi.do"),
    "smart": ("한국식품안전관리인증원", "축산물 스마트HACCP 인증업체 정보 서비스 (공공데이터포털 15118113, 이용허락범위 제한 없음)", "2026-09-14 수집", "https://www.data.go.kr/data/15118113/openapi.do"),
    "tech": ("한국식품안전관리인증원", "스마트HACCP 기술지원 업체 정보 (공공데이터포털 15131084, 이용허락범위 제한 없음)", "2026-09-14 수집", "https://www.data.go.kr/data/15131084/openapi.do"),
    "census": ("국가데이터처", "2026년 2분기 가축동향조사 결과 (보도자료 첨부 통계표, 2026-06-01 기준, 소는 잠정치)", "2026-07-24", "https://www.korea.kr/briefing/pressReleaseView.do?newsId=156772070"),
    "consult": ("농림축산식품부", "2024년 축산물 HACCP 컨설팅 지원사업 시행지침 (한국한우협회 게시 첨부)", "2024", "https://www.ihanwoo.org/19324"),
    "plan24": ("농림축산식품부", "2024년 생산단계 축산물 스마트 HACCP 현장구축사업 추진계획 (공고 제2024-321호, 단계별 중장기 추진 로드맵)", "2024-07-26", "https://www.mafra.go.kr/bbs/home/791/570945/artclView.do"),
    "plan25": ("농림축산식품부", "2025년 생산단계 축산물 스마트 HACCP 현장구축사업 추진계획 (공고 제2024-509호)", "2024-12-02", "https://www.mafra.go.kr/bbs/home/791/572428/artclView.do"),
    "fsnews": ("대한급식신문", "축산물 안전 사각지대 해소된다", "2026-04-02", "https://www.fsnews.co.kr/news/articleView.html?idxno=60482"),
    "foodnews26": ("식품저널", "해썹인증원, 축산물 생산단계 디지털 안전관리 본격 지원", "2026-04-13", "https://www.foodnews.co.kr/news/articleView.html?idxno=118111"),
    "law_ahl": ("국가법령정보센터", "축산물 위생관리법 제9조·제9조의2·제9조의3 (안전관리인증기준, 유효기간, 조사·평가)", "시행 2026-12-31 기준 조문", LAW + "축산물위생관리법/제9조"),
    "age": ("한국농어민신문", "청년농 육성한다지만…높은 진입장벽 탓 후계인력 확보 '막막' (통계청 2020년 기준 축산 농가 경영주 연령 인용)", "2022-04-05", "http://www.agrinet.co.kr/news/articleView.html?idxno=308392"),
    "grade_rule": ("농림축산식품부", "축산물 등급판정 세부기준 (고시 제2026-80호)", "2026-07-31", "https://www.law.go.kr/행정규칙/축산물등급판정세부기준"),
    "mafra_pls": ("농림축산식품부·농림축산검역본부", "동물용의약(외)품 안전사용 안내 (휴약기간, 잔류위반 농가 제재)", "현행", "https://www.mafra.go.kr/bbs/PLS/830/571195/download.do"),
    "kdi_pls": ("KDI 경제교육·정보센터", "축·수산물 동물용의약품 PLS 제도 2024년 1월 1일 본격 시행", "2023-12-28", "https://eiec.kdi.re.kr/policy/materialView.do?num=246571"),
    "law_pharm": ("국가법령정보센터", "약사법 제85조(동물용 의약품 등에 대한 특례)·제98조(과태료)", "현행", LAW + "약사법/제85조"),
    "akaike23": ("Akaike H, Nagai M, Okatani AT, Morita Y", "Food safety, livestock health, and productivity of a farm fatting Japanese black cattle following implementation of a certificated HACCP system. J Vet Med Sci 85(9):985-989", "2023", "https://doi.org/10.1292/jvms.23-0034"),
    "schneider": ("Schneider MJ, Tait RG, Busby WD, Reecy JM", "An evaluation of bovine respiratory disease complex in feedlot cattle. J Anim Sci 87(5):1821-1827", "2009", "https://doi.org/10.2527/jas.2008-1283"),
    "wilson": ("Wilson BK, Step DL, Maxwell CL 외", "Effect of bovine respiratory disease during the receiving period on steer finishing performance, efficiency, carcass characteristics, and lung scores. Prof Anim Sci 33(1):24-36", "2017", "https://pmc.ncbi.nlm.nih.gov/articles/PMC7147665/"),
    "blake": ("Blakebrough-Hall C, McMeniman JP, González LA", "An evaluation of the economic effects of bovine respiratory disease on animal performance, carcass traits, and economic outcomes in feedlot cattle defined using four BRD diagnosis methods. J Anim Sci 98(2):skaa005", "2020", "https://doi.org/10.1093/jas/skaa005"),
    "grimes": ("Grimes BB, McEvers TJ, Tennant TC, Lawrence TE", "Association of bovine lung lesions at slaughter with carcass performance, value and liver-abscess severity. Transl Anim Sci 9:txaf112", "2025", "https://pmc.ncbi.nlm.nih.gov/articles/PMC12405687/"),
    "grade_stat": ("한국농어민신문", "한우 1++ 등급 비중이 가장 높은 사육월령은? (2024 축산물등급판정 통계연보 인용)", "2025-04-04", "https://www.agrinet.co.kr/news/articleView.html?idxno=335756"),
    "mafra_short": ("농림축산식품부", "한우 사육기간 단축으로 두 마리 토끼 잡는다 (보도자료, 강원대 박병기 교수팀 분석 인용)", "2024-04-03", "https://www.mafra.go.kr/bbs/home/792/569966/artclView.do"),
    "chuksaro": ("식품저널", "'축사로', 무항생제 한우 기록 관리 서비스 (동물약품 구매·사용, 수의사 처방관리 항목 추가, 이력제 정보 자동 연동)", "2023-10-18", "http://www.foodnews.co.kr/news/articleView.html?idxno=104991"),
    "chuksaro2": ("농민신문", "한우 농장 기록, '축사로'로 똑똑하게 관리하세요", "2023-03-02", "https://www.nongmin.com/article/20230302500093"),
    "law_vet": ("국가법령정보센터", "수의사법 제12조의2·제12조의3, 같은 법 시행규칙 제11조(처방전 기재사항)", "현행", LAW + "수의사법/제12조의2"),
    "evet_news": ("이엠디", "수의사 전자처방전 발급 의무화 2020년 2월 28일부터 시행", "2020-02-23", "https://www.mdon.co.kr/news/article.html?no=25680"),
    "law_rule": ("국가법령정보센터", "동물용 의약품등 취급규칙 제3조의2·제22조 (판매기록 작성·보존)", "현행", LAW + "동물용의약품등취급규칙/제22조"),
    "law_trace": ("국가법령정보센터", "가축 및 축산물 이력관리에 관한 법률 제5조(출생 등의 신고)·제25조(이력정보의 공개 등)·제32조(벌칙)", "현행", LAW + "가축및축산물이력관리에관한법률/제25조"),
    "trace_api": ("축산물품질평가원", "쇠고기이력정보 (공공데이터포털 15056898, 공공누리 제1유형 출처표시, 개체번호 1건씩 조회)", "명세 확인 2026-09-17", "https://www.data.go.kr/data/15056898/openapi.do"),
    "law_privacy": ("국가법령정보센터", "개인정보 보호법 제15조(수집·이용)·제17조(제공)", "현행", LAW + "개인정보보호법/제15조"),
    "onepass": ("축산물품질평가원", "축산물원패스 (등급판정결과 조회·통합증명서)", "현행", "https://play.google.com/store/apps/details?id=kr.or.ekape.kapemcp.android"),
    "ai_grade": ("축산신문", "한우 온라인 구매도 '마블링' 보고 고른다…AI 등급판정 이미지 민간 개방", "2026-08-27", "https://www.chuksannews.co.kr/news/article.html?no=273287"),
    "weather": ("부산일보", "한우 도축전 AI로 도체등급 예측…날씨 빅데이터 콘테스트 최우수상", "2026-08-06", "https://www.busan.com/view/busan/view.php?code=2026080617425417698"),
    "calc2019": ("이데일리", "축평원, 소고기 등급 예측 가능한 진단 프로그램 개발", "2019-11-04", "https://edaily.co.kr/News/Read?mediaCodeNo=257&newsId=03955686622682440"),
    "ultra": ("농기자재신문", "거세한우의 지방 축적 과다 방지를 위한 초음파 진단시스템 개발", "2021-12-23", "http://newsam.co.kr/mobile/article.html?no=33861"),
    "genome": ("축산경제신문", "한우 암소 유전체 정보 이용하면 농가 맞춤형 개량 가능", "2020-12-11", "http://www.chukkyung.co.kr/news/articleView.html?idxno=59178"),
    "pinna": ("Pinna D, Guarnido-Lopez P, Nagaoka S 외 (Research Square 프리프린트, 동료 검토 전)", "A machine learning approach to predict meat production factors in Japanese black cattle fattening", "2025-07-28", "https://doi.org/10.21203/rs.3.rs-6653315/v1"),
    "inoue": ("Inoue K, Honda T, Oyama K", "Genetic relationships between internal diseases diagnosed at slaughter and carcass traits in Japanese Black cattle. J Anim Sci 93(6):2714-2721", "2015", "https://doi.org/10.2527/jas.2014-8497"),
}
_order = []


def cite(*keys):
    for k in keys:
        assert k in REFS, k
        if k not in _order:
            _order.append(k)
    return "[참고 " + ", ".join(str(n) for n in sorted(_order.index(k) + 1 for k in keys)) + "]"


def source(slide, keys, t=""):
    text(slide, M, Inches(7.02), W - 2 * M, Inches(0.3), [((cite(*keys) + " " + t).strip(), 9, False, MUTE)])


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t


# ─────────────────────────────────────────────────────────────
# 본문 슬라이드
# ─────────────────────────────────────────────────────────────
def cover(prs):
    s = blank(prs)
    box(s, Inches(0.6), Inches(0.6), W - Inches(1.2), H - Inches(1.2), fill=BG, r=Inches(0.4))
    pill(s, Inches(1.2), Inches(1.6), Inches(6.6), Inches(0.46), "2026 한국식품안전관리인증원 공공데이터 분석·활용 경진대회", WHITE, SUB, size=13)
    text(s, Inches(1.2), Inches(2.45), Inches(11), Inches(2.2),
         [("기록 부담은 줄이고,", 42, True, INK), ("출하 수익은 높이고", 42, True, ACC)])
    text(s, Inches(1.2), Inches(4.5), Inches(11), Inches(1.0),
         [("스마트HACCP에서 비어 있는 한우 농장을 수의사 전자처방·약 판매기록·축사로·축산물이력제 연계로 채우는 방안", 18, False, SUB)])
    pill(s, Inches(1.2), Inches(5.75), Inches(2.9), Inches(0.42), "기획서 초안 v6 · 조원 검토용", WHITE, MUTE, size=12)
    notes(s, "【발표 대본】 안녕하십니까. 저희는 한우 농장이 HACCP과 스마트HACCP에서 왜 비어 있는지, "
             "그리고 농장주가 스스로 쓰고 싶어지는 방식으로 이 빈칸을 어떻게 채울지 제안드리겠습니다. "
             "핵심은 두 가지입니다. 기록은 이미 있는 데이터로 자동으로 채우고, 그 기록을 농가의 돈이 되는 정보로 돌려드리는 것입니다.")


def s_data(prs):
    s = blank(prs)
    header(s, "데이터", "분석에 사용한 데이터",
           "한국식품안전관리인증원 공개 데이터 3종을 가축 통계·정책 문서·법령과 연결했습니다.")
    rows = [
        ["데이터", "제공 기관", "규모·기준", "쓴 곳"],
        ["HACCP 인증업체정보", "한국식품안전관리인증원", "41,069건 (2026-09-08)", "축종별 인증 농장 수"],
        ["축산물 스마트HACCP 인증업체", "한국식품안전관리인증원", "298곳 (563행 → 인증번호 기준)", "스마트HACCP 등록 업종, 중요관리점"],
        ["스마트HACCP 기술지원 업체", "한국식품안전관리인증원", "2,719건", "등록 준비 단계 현황"],
        ["가축동향조사 (2026년 2분기)", "국가데이터처", "축종·규모별 농장 수", "인증률 분모"],
        ["컨설팅 시행지침 · 스마트HACCP 추진계획", "농림축산식품부", "2019~2026년", "예산, 계획 대비 실제 사업"],
    ]
    table(s, M, Inches(2.05), W - 2 * M, rows,
          [Inches(3.7), Inches(2.9), Inches(3.1), Inches(2.43)], row_h=Inches(0.56), size=13)
    box(s, M, Inches(5.65), W - 2 * M, Inches(1.1), fill=GREEN_SOFT, lines=[
        ("집계 원칙", 13, True, GREEN),
        ("인증 농장은 '업체명 + 시도 + 시군구'가 같으면 한 곳으로 셉니다. 이름만 쓰면 다른 지역의 같은 이름 농장이 합쳐지고, "
         "인증번호로 세면 한 농장의 여러 인증이 따로 잡히기 때문입니다.", 12, False, INK)])
    source(s, ["haccp", "smart", "tech", "census", "consult", "plan24", "plan25"],
           "공개 데이터 3종은 모두 '이용허락범위 제한 없음'입니다.")
    notes(s, "【발표 대본】 주최 기관이 공개한 데이터 세 가지를 중심으로, 분모가 되는 가축동향조사와 정부 계획 문서를 붙였습니다. "
             "인증 농장 수는 중복을 어떻게 걷어내느냐에 따라 달라지므로 기준을 먼저 밝힙니다.\n"
             "【검증 메모】 한우 인증 농장 수: 업체명 기준 1,676곳, 업체명+시군구 기준 1,835곳, 인증번호 기준 1,884건. "
             "v5까지 쓴 1,607곳은 품목이 '한우' 하나인 건만, 업체명 기준으로 센 값이라 과소 집계였습니다.")


def s_status(prs):
    s = blank(prs)
    header(s, "현황", "한우 농장은 HACCP 2.5%, 스마트HACCP은 0건입니다")
    text(s, M, Inches(1.5), Inches(7), Inches(0.35), [("축종별 농장 HACCP 인증률", 13, True, SUB)])
    bars(s, M, Inches(1.85), Inches(7.2), Inches(4.85),
         ["닭", "돼지", "오리", "젖소", "한우"], [61.4, 24.2, 20.9, 11.8, 2.5], hot={4}, vmax=75)
    box(s, Inches(8.2), Inches(1.55), Inches(4.5), Inches(1.65), fill=WHITE, line=LINE, lines=[
        ("한우 농장 HACCP 인증", 13, False, SUB),
        ("72,389곳 중 1,835곳", 26, True, ACC)], anchor=MSO_ANCHOR.MIDDLE)
    box(s, Inches(8.2), Inches(3.35), Inches(4.5), Inches(1.65), fill=WHITE, line=LINE, lines=[
        ("농장 HACCP 인증 중 스마트HACCP 등록", 13, False, SUB),
        ("6,294건 중 0건", 26, True, ACC)], anchor=MSO_ANCHOR.MIDDLE)
    box(s, Inches(8.2), Inches(5.15), Inches(4.5), Inches(1.6), fill=BG, lines=[
        ("큰 농장만 따져도 낮습니다", 13, True, INK),
        ("50마리 이상 한·육우 농장 20,763곳으로 분모를 좁혀도 최대 8.9%입니다.", 12, False, SUB)])
    source(s, ["haccp", "smart", "census"],
           "닭(3천 마리 이상)·오리(2천 마리 이상)는 큰 농가만 분모에 들어가 인증률이 높게 나옵니다. 한우는 세는 방식에 따라 2.3~2.6%.")
    notes(s, "【발표 대본】 한우 농장 7만 2천여 곳 중 HACCP 인증을 받은 곳은 1,835곳, 2.5%로 다섯 축종 중 가장 낮습니다. "
             "스마트HACCP으로 넘어가면 농장은 모든 축종을 합쳐도 0건입니다. "
             "한두 마리 키우는 농가가 분모에 많아서 낮은 것 아니냐는 질문이 나올 수 있어, 50마리 이상 농장만 분모로 잡아 봤고 그래도 최대 8.9%입니다.\n"
             "【검증 메모】 분모: 가축동향조사 2026년 2분기 통계표(한우 72,389 / 한·육우 75,647 / 50마리 미만 54,884 / 젖소 5,342 / 돼지 5,465 / 닭 3,123가구 / 오리 502가구). "
             "8.9% = 한·육우 인증 1,845곳 ÷ (75,647 − 54,884). 인증 농가가 모두 50마리 이상이라는 가정의 상한입니다. "
             "닭은 분기별 가구 수 변동(2,747~3,123)을 반영하면 61.4~69.8%, 오리는 18.8~33.4%입니다. "
             "6,294건 = 가축사육업 5,972 + 종축업 289 + 부화업 33 (인증번호 기준).")


def s_cause_structure(prs):
    s = blank(prs)
    header(s, "원인 1 · 구조", "인증을 받지 않아도 팔 수 있고, 받아도 값이 오르지 않습니다")
    steps = [("농장", "원하면 신청", False), ("도축장", "법으로 의무", True),
             ("가공·포장", "법으로 의무", True), ("정육점·마트", "원하면 신청", False)]
    bw, gap = Inches(2.72), Inches(0.42)
    for i, (head, tag, must) in enumerate(steps):
        x = M + (bw + gap) * i
        box(s, x, Inches(1.55), bw, Inches(1.35), fill=ACC_SOFT if i == 0 else BG, lines=[
            (head, 17, True, ACC if i == 0 else INK)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        pill(s, x + Inches(0.5), Inches(2.3), bw - Inches(1.0), Inches(0.42), tag,
             GREEN if must else WHITE, WHITE if must else MUTE, size=12, line=None if must else LINE)
        if i < len(steps) - 1:
            chevron(s, x + bw + Inches(0.06), Inches(2.05))
    cw = (W - 2 * M - Inches(0.3) * 2) / 3
    cards = [("50마리 미만 한·육우 농장", "72.6%", "75,647곳 중 54,884곳 (2026년 6월)"),
             ("60세 이상 축산 농가 경영주", "63%", "2020년 기준 · 한우만의 수치는 아님"),
             ("인증이 반영되는 판매 가격", "없음", "등급판정 세부기준에 인증 항목 없음")]
    for i, (head, big, small) in enumerate(cards):
        x = M + (cw + Inches(0.3)) * i
        box(s, x, Inches(3.15), cw, Inches(2.35), fill=BG, lines=[
            (head, 14, False, SUB), (big, 50, True, ACC), (small, 11, False, MUTE)],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, M, Inches(5.7), W - 2 * M, Inches(1.1), fill=ACC_SOFT, lines=[
        ("인증을 받으면 3년마다 연장 심사, 해마다 1회 이상 조사·평가를 받습니다. 법에 적힌 혜택은 인증 명칭 사용, 기술·교육 지원, "
         "평가 우수 시 조사 면제 등이고 소 값에 반영되는 장치는 없습니다. 안전만 내세워서는 농가가 움직이지 않습니다.", 14, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE)
    source(s, ["law_ahl", "census", "age", "grade_rule"], "도축업·집유업은 자체안전관리인증기준 운용, 식육포장처리업 등은 인증기준 준수가 의무입니다.")
    notes(s, "【발표 대본】 소고기가 소비자에게 가는 길에서 도축장과 가공·포장 단계는 HACCP이 법으로 의무입니다. 농장은 원하면 신청하는 곳입니다. "
             "그런데 한·육우 농장의 73%는 50마리 미만이고, 축산 농가 경영주의 63%는 60세 이상입니다. "
             "인증을 받으면 3년마다 연장하고 매년 조사·평가를 받아야 하는데, 소 값이 매겨지는 등급 판정 기준에는 인증 여부가 들어가지 않습니다. "
             "부담은 있고 돈은 안 되는 구조입니다.\n"
             "【검증 메모】 축산물 위생관리법 제9조②③④(의무·자율 구분), 제9조⑨(인증 명칭 사용)·⑩(기술·정보 제공, 교육훈련), 제9조의2(유효기간 3년), 제9조의3①③(연 1회 이상 조사·평가, 우수 시 면제 등 지원 '할 수 있다'). "
             "63%는 기사 원문상 '축산 농가' 경영주(60~69세 41.4%, 70세 이상 21.6%)로, 한우 농가만의 수치가 아닙니다. "
             "등급판정 세부기준(고시 제2026-80호)의 육질등급 항목은 근내지방도·육색·지방색·조직감·성숙도뿐입니다. "
             "정책사업 선정 가점 등 다른 우대가 있는지는 모든 사업지침을 확인하지 못했으므로 '판매 가격에 반영되지 않는다'고만 말합니다.")


def s_cause_design(prs):
    s = blank(prs)
    header(s, "원인 2 · 설계", "스마트HACCP은 센서로 재는 항목에 맞춰져 있습니다",
           "스마트HACCP 등록 업체의 중요관리점(반드시 지켜야 하는 핵심 관리 항목) 668개 중 90.3%가 기계로 재는 항목입니다.")
    bars(s, M, Inches(2.0), Inches(6.2), Inches(4.3),
         ["냉장·냉동 보관", "금속 검출", "가열·살균", "이물 검출", "그 외"], [379, 159, 55, 10, 65],
         hot=set(), fmt='#,##0"개"', vmax=460)
    box(s, Inches(7.2), Inches(2.05), Inches(2.65), Inches(3.3), fill=BG, lines=[
        ("가공 공장", 15, True, INK), ("", 6, False, INK),
        ("핵심 항목", 11, False, MUTE), ("냉장 온도, 금속 검출", 14, False, INK), ("", 6, False, INK),
        ("확인 방법", 11, False, MUTE), ("센서가 재서 자동 기록", 14, True, INK)])
    box(s, Inches(10.05), Inches(2.05), Inches(2.65), Inches(3.3), fill=ACC_SOFT, lines=[
        ("한우 농장", 15, True, ACC), ("", 6, False, INK),
        ("핵심 항목", 11, False, MUTE), ("휴약기간 (약을 쓴 뒤 출하하면 안 되는 기간)", 14, False, INK), ("", 6, False, INK),
        ("확인 방법", 11, False, MUTE), ("약 기록과 소 번호를 대조", 14, True, ACC)])
    box(s, Inches(7.2), Inches(5.55), Inches(5.5), Inches(0.85), fill=WHITE, line=LINE, lines=[
        ("농장에 필요한 것은 센서가 아니라 기록 연결입니다.", 15, True, INK)], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    source(s, ["smart", "akaike23"], "'완제품 냉장보관', '완제품냉장보관공정'처럼 표기가 다른 항목은 핵심 단어로 묶었습니다.")
    notes(s, "【발표 대본】 스마트HACCP은 온도계나 금속 검출기 같은 기계가 값을 재서 자동으로 남기는 방식입니다. "
             "실제로 등록 업체의 중요관리점 90%가 냉장·금속·가열처럼 기계로 재는 항목입니다. "
             "그런데 소 농장에서 가장 중요한 관리 항목은 '약을 쓴 소를 휴약기간이 지나기 전에 출하하지 않는 것'입니다. "
             "일본의 HACCP 인증 흑소 농장도 이것을 중요관리점으로 정하고, 출하할 소의 귀표 번호를 약 기록과 대조했습니다. 이건 센서로 잴 수 없고 기록을 연결해야 합니다.\n"
             "【검증 메모】 298곳의 중요관리점 표기 668개를 핵심어로 분류(냉장·냉동 379, 금속 159, 가열·살균 55, 이물 10, 기타 65). "
             "기계 측정형 603개 = 90.3%. 인증번호별 첫 행 기준입니다.")


def s_cause_policy(prs):
    s = blank(prs)
    header(s, "원인 3 · 지원", "농장은 계획에만 있었고, 컨설팅 예산은 5년 새 64% 줄었습니다")
    text(s, M, Inches(1.5), Inches(6.4), Inches(0.4), [("생산단계 스마트HACCP: 로드맵과 실제 사업", 13, True, SUB)])
    rows = [["연도", "2024년 로드맵", "실제 사업"],
            ["2024", "집유장", "집유장 2곳"],
            ["2025", "도축장(소·돼지), 농장(젖소)", "도축장(포유류), 집유장"],
            ["2026", "도축장(가금), 농장(소·돼지·가금)", "가금 도축장, 8곳 보급"]]
    table(s, M, Inches(1.95), Inches(6.4), rows, [Inches(0.8), Inches(3.0), Inches(2.6)],
          row_h=Inches(0.6), hot_rows=(2, 3), size=13)
    text(s, M, Inches(4.55), Inches(6.4), Inches(2.3), [
        ("스마트HACCP 농장 0건은 농가가 외면해서가 아니라, 사업 대상에 농장이 들어간 적이 없기 때문입니다.", 14, True, ACC),
        ("2025년 신청서에는 '□ 농장' 칸까지 있었고, 2024년 계획의 추진 과제에도 '사육농장 공정별' 센서가 들어 있었습니다.", 12, False, SUB),
        ("※ 집유장 = 농장에서 짠 우유를 모으는 곳 · 2026년 8곳은 도축장·집유장", 11, False, MUTE)])
    text(s, Inches(7.4), Inches(1.5), Inches(5.3), Inches(0.4), [("축산물 HACCP 컨설팅 지원 예산 (백만 원, 국비+지방비+자부담)", 12, True, SUB)])
    bars(s, Inches(7.3), Inches(1.9), Inches(5.4), Inches(3.5),
         ["2019", "2020", "2021", "2022", "2023", "2024"], [2800, 2800, 2593, 2593, 2033, 1015],
         hot={5}, horizontal=False, fmt='#,##0', vmax=3300)
    box(s, Inches(7.4), Inches(5.55), Inches(5.3), Inches(1.25), fill=ACC_SOFT, lines=[
        ("2024년 예산을 전부 한우 농장에 써도 개별 컨설팅 169곳, 단체 컨설팅 약 330곳", 13, True, ACC),
        ("50마리 이상 농장만 2만 곳이라, 농가마다 컨설턴트를 붙이는 방식으로는 닿기 어렵습니다.", 12, False, INK)])
    source(s, ["plan24", "plan25", "fsnews", "foodnews26", "consult"], "2026년 실제 사업은 보도 기준입니다.")
    notes(s, "【발표 대본】 농림축산식품부의 2024년 로드맵에는 2025년 젖소 농장, 2026년 소·돼지·가금 농장이 들어 있었습니다. "
             "하지만 실제 사업은 2024년 집유장, 2025년 포유류 도축장과 집유장, 2026년 가금 도축장 시스템 개발과 도축장·집유장 8곳 보급이었습니다. "
             "그래서 스마트HACCP 농장이 0건인 것은 농가 탓이 아닙니다. 한편 농장 인증을 돕는 컨설팅 예산은 2019년 28억 원에서 2024년 10억 원으로 64% 줄었습니다.\n"
             "【검증 메모】 로드맵: 2024 추진계획 hwpx의 '단계별 중장기 추진 로드맵' 그림(1차년도 집유장 / 2차년도 도축장(소,돼지)·농장(젖소) / 3차년도 도축장(가금)·농장(소,돼지,가금)). "
             "2025 추진계획: 사업대상 도축장(포유류에 한함)·집유장, 국비 551백만 원, 신청서 '□집유장 □도축장 □농장'. "
             "2026: 대한급식신문(2026-04-02)·식품저널 보도 '가금류 도축장 시스템 개발, 전국 8개 작업장(집유장·도축장), 구축비 1억 원 이내 최대 70%'. "
             "예산: 2024 시행지침 '연도별 재정투입 현황'(합계 기준, 2024년 국고 406·지방비 305·자부담 304). 이 예산은 도축장·집유장·사료공장 컨설팅도 포함합니다. "
             "단가: 한우 개별 600만 원(1,015÷6=169), 단체는 기본 600만 원+농가당 300만 원(한 단체로 몰면 약 337곳).")


def s_problem(prs):
    s = blank(prs)
    header(s, "문제 정의", "부담은 크고, 돌아오는 돈은 없습니다")
    box(s, M, Inches(1.5), W - 2 * M, Inches(1.2), fill=ACC_SOFT, lines=[
        ("한우 농장이 HACCP·스마트HACCP에서 비어 있는 이유는 안전 의식이 아니라 계산입니다. "
         "기록은 손으로 해야 하고, 인증으로 버는 돈은 없으며, 제도는 공장에 맞춰져 있습니다.", 17, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE, pad=0.3)
    items = [
        ("부담", "매번 손으로 기록", "작고 나이 든 농가에게 약 기록과 기준서 관리가 큰 일", "기록을 자동으로 채운다"),
        ("보상", "인증해도 값은 그대로", "등급은 도축 후에야 알고, 기록이 판매 판단에 쓰이지 않음", "기록을 돈이 되는 정보로 돌려준다"),
        ("제도", "센서형 스마트HACCP", "농장은 사업 대상에서 빠졌고, 컨설팅 예산은 줄어듦", "이미 있는 시스템을 연결한다"),
    ]
    bw = (W - 2 * M - Inches(0.3) * 2) / 3
    for i, (tag, head, body, to) in enumerate(items):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(2.95), bw, Inches(2.2), fill=WHITE, line=LINE, lines=[
            (tag, 13, True, ACC), (head, 19, True, INK), ("", 6, False, INK), (body, 14, False, SUB)])
        pill(s, x + Inches(0.2), Inches(5.35), bw - Inches(0.4), Inches(0.62), to, GREEN_SOFT, GREEN, size=14)
    notes(s, "【발표 대본】 정리하면, 한우 농장이 비어 있는 이유는 안전 의식이 부족해서가 아니라 계산이 맞지 않아서입니다. "
             "기록은 손으로 해야 하고, 인증해도 값이 오르지 않고, 제도는 공장에 맞춰져 있습니다. "
             "그래서 해결책도 세 가지 방향입니다. 기록은 자동으로 채우고, 기록을 돈이 되는 정보로 돌려주고, 새 시스템 대신 이미 있는 시스템을 연결합니다.")


def s_loss(prs):
    s = blank(prs)
    header(s, "돈과 연결되는 지점 1", "휴약기간을 한 번 놓치면 손실이 큽니다",
           "휴약기간 = 가축에게 약을 쓴 뒤, 고기에 약 성분이 남지 않도록 출하하지 않고 기다리는 기간입니다.")
    items = [("전량 폐기", "잔류허용기준을 넘은 축산물은 전량 폐기됩니다."),
             ("과태료", "사용 기준이 정해진 동물약의 기준을 어기면 100만 원 이하 과태료 (수의사 처방에 따른 사용은 예외)"),
             ("6개월 규제검사", "잔류위반 농가로 지정되면 6개월간 출하 가축마다 검사를 받고 출하가 제한됩니다.")]
    bw = (W - 2 * M - Inches(0.3) * 2) / 3
    for i, (head, body) in enumerate(items):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(2.1), bw, Inches(2.35), fill=BG, lines=[
            (head, 24, True, ACC), ("", 6, False, INK), (body, 14, False, INK)], anchor=MSO_ANCHOR.MIDDLE, pad=0.3)
    box(s, M, Inches(4.65), W - 2 * M, Inches(1.0), fill=WHITE, line=LINE, lines=[
        ("2024년 1월부터 잔류허용기준이 없는 약은 고기 1kg에 0.01mg 이하(일률기준)만 허용됩니다. 실수할 여지가 더 작아졌습니다.", 15, False, INK)],
        anchor=MSO_ANCHOR.MIDDLE, pad=0.3)
    box(s, M, Inches(5.85), W - 2 * M, Inches(0.95), fill=ACC_SOFT, lines=[
        ("지금은 사람이 약품 기록부를 보고 확인합니다. 자동 확인은 안전 장치이자 손실 보험입니다.", 16, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE, pad=0.3)
    source(s, ["mafra_pls", "kdi_pls", "law_pharm"], "과태료 근거: 약사법 제85조제3항·제98조제1항제10호")
    notes(s, "【발표 대본】 휴약기간을 한 번 놓치면 농가 손실이 큽니다. 기준을 넘은 고기는 전량 폐기되고, 사용 기준 위반에는 과태료가 나오며, "
             "잔류위반 농가로 지정되면 6개월 동안 출하하는 소마다 검사를 받습니다. 2024년부터는 기준이 없는 약에 0.01mg 일률기준이 적용돼 더 엄격해졌습니다. "
             "그래서 자동 확인은 안전만을 위한 것이 아니라 농가 입장에서 손실을 막는 보험입니다.\n"
             "【검증 메모】 폐기·6개월 집중 관리·규제검사·출하 제한: 농식품부·검역본부 '동물용의약(외)품 안전사용 안내' 원문. "
             "과태료: 약사법 제98조①10호 '제85조제3항을 위반하여 동물용 의약품등의 사용 기준을 지키지 아니한 자' 100만 원 이하. "
             "단, 제85조③ 단서에 따라 수의사 처방에 따라 사용한 경우에는 사용 기준 적용이 제외되므로 '과태료가 무조건 나온다'고 말하면 안 됩니다. "
             "PLS: 2024-01-01부터 소·돼지·닭·우유·달걀, 일률기준 0.01mg/kg.")


def s_health(prs):
    s = blank(prs)
    header(s, "돈과 연결되는 지점 2", "아팠던 소는 덜 받고, 관리하면 아픈 소가 줄었습니다",
           "약 기록은 곧 '어느 소가 아팠는지'의 기록입니다. 해외 연구에서 치료 이력은 출하 성적과 연결됐습니다.")
    text(s, M, Inches(2.0), Inches(6.0), Inches(0.35), [("해외 연구: 호흡기 병으로 치료받은 소의 출하 성적", 13, True, SUB)])
    facts = [("-39.6kg", "3번 이상 치료한 소의 도체중 · 1마리당 약 385호주달러 손해 (호주, 898마리)"),
             ("-8.2kg", "치료받은 소의 도체중 (미국, 5,976마리)"),
             ("70% → 36%", "상위 두 품질 등급 비율: 치료 0회 → 3~4회 (미국, 174마리)")]
    y = Inches(2.4)
    for big, small in facts:
        box(s, M, y, Inches(6.0), Inches(1.05), fill=BG, lines=[(big, 22, True, ACC), (small, 12, False, SUB)],
            anchor=MSO_ANCHOR.MIDDLE, pad=0.25)
        y += Inches(1.15)
    box(s, Inches(6.9), Inches(2.0), Inches(5.8), Inches(3.25), fill=GREEN_SOFT, lines=[
        ("일본 흑소 농장 (280마리, 가족 4명, 2020년 인증)", 15, True, GREEN), ("", 4, False, INK),
        ("출하할 소의 귀표 번호를 약·주사 바늘 기록과 손으로 대조", 13, False, INK),
        ("• 기준 이탈 0건 (2019~2021년 출하 때마다 점검)", 13, True, INK),
        ("• 들여온 뒤 6개월 안 호흡기 치료: 78마리 → 28~39마리", 13, False, INK),
        ("• 최고 등급(5등급) 비율: 36.5% → 66~78%", 13, False, INK),
        ("• 5등급 단가는 내렸지만 전체 매출은 증가", 13, False, INK)], anchor=MSO_ANCHOR.MIDDLE, pad=0.3)
    box(s, Inches(6.9), Inches(5.4), Inches(5.8), Inches(1.4), fill=WHITE, line=LINE, lines=[
        ("주의할 점", 12, True, INK),
        ("도체중 감소는 연구마다 일관되지만 품질 등급은 엇갈립니다(미국 6만여 마리 연구에서는 지방 무늬가 오히려 늘어나는 경향). "
         "일본 사례는 농장 1곳이고 사료 첨가제·감시 카메라 등도 함께 바꿨습니다.", 11, False, SUB)], pad=0.25)
    source(s, ["blake", "schneider", "wilson", "akaike23", "grimes"], "도체중 = 머리·가죽·내장·발을 뗀 몸통 무게. 한우 대상 연구는 찾지 못했습니다. 상세는 부록 1.")
    notes(s, "【발표 대본】 약 기록은 어느 소가 아팠는지의 기록이기도 합니다. 해외 연구에서 호흡기 병으로 여러 번 치료받은 소는 도체중이 수십 kg 가볍고, "
             "상위 등급 비율도 낮았습니다. 일본의 한 흑소 농장은 HACCP을 도입하며 출하 전 약 기록 대조를 중요관리점으로 삼았고, "
             "기준 이탈이 한 번도 없었으며, 아픈 소가 줄고 최고 등급 비율이 늘었습니다. 다만 한 농장 사례이고 다른 개선도 함께 했기 때문에 "
             "HACCP 때문이라고 단정하지는 않겠습니다.\n"
             "【검증 메모】 Akaike 2023 원문: 2017년 5등급 58/159(36.5%), 2018 66.0%, 2019 77.5%, 2020 77.1%, 2021 73.1%. "
             "BRD 치료 78→28·39·29·38. 사망 4→5·2·2·1(유의차 없음). 5등급 단가 2,798엔(2017)→2,666엔(2021)/kg, 총매출은 2021년이 더 높음. "
             "저자도 '주로 GMP(기본 위생 관리) 효과'로 봄. Blakebrough-Hall 2020: ≥3회 치료 39.6kg, AUD 384.97 감소. "
             "Schneider 2009: 8.16±1.38kg. Wilson 2017: Prime+Choice 0회 70.3%, 1회 56.5%, 2회 60.2%, 3~4회 36.2%(선형 감소 P=0.03). "
             "Grimes 2025: 60,843 도체, 폐 경화 심할수록 도체중 감소, 근내지방 점수는 증가 경향.")


def s_timing(prs):
    s = blank(prs)
    header(s, "돈과 연결되는 지점 3", "언제 팔지에 따라 등급과 사료비가 갈립니다",
           "오래 키우면 최고 등급 비율이 오르지만 사료비가 들고, 일찍 팔면 사료비는 줄지만 등급을 놓칠 수 있습니다.")
    box(s, M, Inches(2.1), Inches(5.95), Inches(2.9), fill=BG, lines=[
        ("더 키우면", 15, True, INK), ("", 4, False, INK),
        ("거세우 평균 출하 (2024년)", 12, False, MUTE), ("31.6개월", 30, True, INK),
        ("최고 등급(1++) 비중이 가장 높은 월령", 12, False, MUTE), ("35개월 · 43.9%", 30, True, ACC)], pad=0.3)
    box(s, Inches(6.75), Inches(2.1), Inches(5.95), Inches(2.9), fill=BG, lines=[
        ("일찍 팔면", 15, True, INK), ("", 4, False, INK),
        ("30개월 대신 24개월에 출하할 때", 12, False, MUTE), ("사료비 약 32% ↓", 30, True, INK),
        ("온실가스 약 25% ↓ · 강원대 박병기 교수팀", 12, False, MUTE)], pad=0.3)
    box(s, M, Inches(5.25), W - 2 * M, Inches(1.5), fill=ACC_SOFT, anchor=MSO_ANCHOR.MIDDLE, pad=0.3, lines=[
        ("정답은 소마다 다릅니다", 15, True, ACC),
        ("평균 출하는 최고 등급 비중이 가장 높은 월령보다 3~4개월 이릅니다. 소별 치료 이력과 성장 정보를 보고 "
         "'지금 팔지, 더 키울지' 판단할 수 있다면 그것이 농가가 기록을 남길 이유가 됩니다.", 14, False, INK)])
    source(s, ["grade_stat", "mafra_short"], "2024 축산물등급판정 통계연보(한국농어민신문 인용), 농림축산식품부 보도자료")
    notes(s, "【발표 대본】 농가의 실제 고민은 언제 팔지입니다. 2024년 거세우는 평균 31.6개월에 출하됐는데, 1++ 비중은 35개월에서 43.9%로 가장 높았습니다. "
             "반대로 정부는 사육 기간을 30개월에서 24개월로 줄이면 사료비가 32% 줄어든다는 분석을 내놨습니다. "
             "더 키우는 게 이득인지 일찍 파는 게 이득인지는 소마다 다릅니다. 이 판단을 돕는 정보를 주면 농가가 기록할 이유가 생깁니다.\n"
             "【검증 메모】 31.6개월·35개월 43.9%: 한국농어민신문 2025-04-04(2024 통계연보 인용). 암소는 평균 53.4개월, 1++ 최고 33개월 25.4%(번식우가 섞여 비교 제외). "
             "32%·25%: 농식품부 보도자료 2024-04-03 원문 '현행 30개월 사육 대비 24개월 사육 시 사료비용이 약 32% 절감, 온실가스 약 25% 감소'.")


def s_direction(prs):
    s = blank(prs)
    header(s, "해결 방향", "두 가지를 함께 합니다",
           "안전 기록만 요구하면 쓰지 않습니다. 기록 부담을 없애고, 그 기록으로 돈이 되는 정보를 돌려줍니다.")
    for i, (tag, head, body, gain) in enumerate([
        ("A", "기록 부담을 없앱니다",
         "수의사 전자처방과 약 판매기록을 농가 동의를 받아 축사로에 자동으로 넣고, 출하 신고 전에 휴약기간을 자동으로 확인합니다.",
         "농가가 얻는 것: 기록 시간 절약, 폐기·과태료 예방"),
        ("B", "돈이 되는 정보를 돌려줍니다",
         "쌓인 약·치료 기록과 축산물이력제 정보로 소별 출하 판단 정보(예상 등급 확률, 더 키울 때 사료비)를 보여줍니다.",
         "농가가 얻는 것: 출하 시기 판단, 아픈 소 조기 발견"),
    ]):
        x = M + (Inches(5.95) + Inches(0.2)) * i
        box(s, x, Inches(2.1), Inches(5.95), Inches(3.3), fill=BG, lines=[
            (tag, 40, True, GREEN), (head, 22, True, INK), ("", 6, False, INK), (body, 15, False, SUB)], pad=0.35)
        pill(s, x + Inches(0.35), Inches(5.6), Inches(5.25), Inches(0.62), gain, GREEN_SOFT, GREEN, size=14)
    notes(s, "【발표 대본】 해결책은 두 축입니다. A는 부담을 없애는 쪽, B는 돈을 돌려주는 쪽입니다. "
             "A만 하면 여전히 농가에게 남는 게 없고, B만 하면 기록이 쌓이지 않습니다. 둘을 함께 해야 돌아갑니다.")


def s_a(prs):
    s = blank(prs)
    header(s, "해결 A", "이미 있는 약 기록을 모아 휴약기간을 자동 확인합니다",
           "새 앱을 만들지 않습니다. 한우 농장용 HACCP 기록 프로그램 '축사로'에 이미 있는 기록을 흘려 넣습니다.")
    steps = [("① 기록 발생", "수의사 전자처방\n약 판매처 판매기록"),
             ("② 자동 전송", "농가가 동의하면\n축사로로 전송"),
             ("③ 소 선택", "\"몇 번 소에 쓰셨나요?\"\n번호만 고르기"),
             ("④ 출하 가능일", "제품별 휴약기간으로\n자동 계산"),
             ("⑤ 출하 전 경고", "이력제 출하 신고 전\n남은 기간 알림")]
    bw, gap = Inches(2.2), Inches(0.28)
    for i, (head, body) in enumerate(steps):
        x = M + (bw + gap) * i
        box(s, x, Inches(1.95), bw, Inches(1.75), fill=GREEN if i == 2 else BG, lines=[
            (head, 15, True, WHITE if i == 2 else INK), ("", 3, False, INK),
            (body, 12, False, WHITE if i == 2 else SUB)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, pad=0.1)
        if i < len(steps) - 1:
            chevron(s, x + bw + Inches(0.01), Inches(2.65), w=Inches(0.26), h=Inches(0.34))
    box(s, M, Inches(3.95), Inches(5.95), Inches(2.85), fill=WHITE, line=LINE, pad=0.25, lines=[
        ("이미 있는 것 (법령·운영 중)", 14, True, GREEN), ("", 2, False, INK),
        ("• 처방대상 약은 전자처방 의무, 처방전에 농장명·동물 종류·약·수량 기재 (2020년 2월~)", 12, False, INK),
        ("• 항생제·호르몬제·구충제 등 판매 시 판매일·제품명·수량·구매자 기록 보존 의무", 12, False, INK),
        ("• 축사로: 약품 구매·사용·수의사 처방 기록 항목, 이력제 정보 자동 연동 (2023년~)", 12, False, INK),
        ("• 소를 출하하면 농장경영자가 이력제에 신고하는 의무", 12, False, INK)])
    box(s, Inches(6.75), Inches(3.95), Inches(5.95), Inches(2.85), fill=ACC_SOFT, pad=0.25, lines=[
        ("빈틈과 새로 필요한 것", 14, True, ACC), ("", 2, False, INK),
        ("• 판매기록은 서식·전산 제출 규정이 없고 농장식별번호 항목도 없음 → 전산 제출 표준", 12, False, INK),
        ("• 약국은 주사용 항생제 등을 뺀 처방대상 약을 처방전 없이 팔 수 있어 전자처방만으로는 누락", 12, False, INK),
        ("• 처방전에 소 이력번호는 필수가 아님 → 농가의 소 선택 한 번", 12, False, INK),
        ("• 전자처방·판매기록 → 축사로 연계는 아직 확인되지 않음 → 시범 연계", 12, False, INK)])
    source(s, ["law_vet", "evet_news", "law_rule", "chuksaro", "law_trace", "law_pharm"])
    notes(s, "【발표 대본】 해결 A는 새 앱이 아닙니다. 국립축산과학원의 한우 농장 HACCP 프로그램 '축사로'에는 이미 약품 구매·사용과 수의사 처방을 적는 칸이 있고, "
             "이력제 정보도 자동으로 불러옵니다. 문제는 농가가 그 칸을 손으로 채워야 한다는 점입니다. "
             "그런데 처방대상 약은 2020년부터 수의사가 전자처방으로만 발급하고, 처방전에는 농장명과 약, 수량이 들어갑니다. "
             "항생제나 호르몬제를 팔 때는 판매처가 판매 기록을 남겨야 합니다. 이 두 기록을 농가 동의를 받아 축사로로 보내면, "
             "농가는 '몇 번 소에 썼는지'만 고르면 되고, 출하 가능일 계산과 출하 신고 전 경고는 자동으로 됩니다.\n"
             "【검증 메모】 수의사법 제12조의2②(처방전은 수의사처방관리시스템으로 발급), 시행규칙 제11조③(동물 종류·성별·연령·체중, 소유자 성명·생년월일·전화번호, 농장명, 성분·용량·용법·처방일수·판매수량). "
             "개체별 처방 원칙, 같은 축사는 군별 처방 가능 → 이력번호가 항상 있지는 않음. "
             "동물용 의약품등 취급규칙 제22조③(제3조의2① 각 호 약품: 호르몬제·항균제·생물학적제제·마약류·마취제·살충제·구충제 판매 시 판매일자·제품명·수량·용도·구매자 기록 1년 이상 보존), "
             "제22조①14호(처방대상 약: 처방전 2년·판매기록 3년 보존). 약사법 제85조⑦(약국개설자는 주사용 항생물질·주사용 생물학적 제제 외 처방대상 약을 처방전 없이 판매 가능). "
             "이력법 제5조①3호(도축 출하 포함 양도 시 신고). 축사로 기능: 식품저널 2023-10-18, 농민신문 2023-03-02. "
             "제품별 휴약기간을 기계가 읽을 수 있는 공개 데이터로 받을 수 있는지는 확인하지 못했습니다(검역본부 허가 현황 데이터에 휴약기간 항목 여부 확인 필요).")


def s_b(prs):
    s = blank(prs)
    header(s, "해결 B", "기록한 농가에는 소별 출하 판단 정보를 돌려드립니다",
           "등급은 도축 후에야 알 수 있습니다. 쌓인 기록을 팔기 전 판단에 쓰도록 돌려주는 것이 기록의 보상입니다.")
    box(s, M, Inches(2.05), Inches(5.0), Inches(4.75), fill=WHITE, line=LINE, pad=0.3, lines=[
        ("화면 예시 (설명용 가상 수치)", 12, True, MUTE), ("", 6, False, INK),
        ("352번 소 · 31개월", 20, True, INK),
        ("휴약기간 끝남 · 출하 가능", 15, True, GREEN), ("", 8, False, INK),
        ("지금 팔면", 12, False, MUTE), ("1++ 40% · 1+ 35%", 20, True, INK), ("", 4, False, INK),
        ("한 달 더 키우면", 12, False, MUTE), ("1++ 48% · 사료비 +32만 원", 20, True, INK), ("", 8, False, INK),
        ("최근 6개월 치료 1회 (자동 기록)", 13, False, SUB)])
    box(s, Inches(5.9), Inches(2.05), Inches(6.8), Inches(2.2), fill=BG, pad=0.3, lines=[
        ("무엇을 쓰나요", 14, True, INK), ("", 3, False, INK),
        ("• 축산물이력제: 출생일·성별·도축 이력 / 축사로: 혈통·초음파 입력값", 13, False, SUB),
        ("• A에서 쌓인 약·치료 기록 ← 기존 예측에서 찾지 못한 정보", 13, True, GREEN),
        ("• 등급은 하나로 단정하지 않고 확률로 보여줍니다", 13, False, SUB)])
    box(s, Inches(5.9), Inches(4.4), Inches(6.8), Inches(2.4), fill=ACC_SOFT, pad=0.3, lines=[
        ("근거 수준과 진행 방식", 14, True, ACC), ("", 3, False, INK),
        ("• 출하 전 등급 예측은 아직 연구·대회 단계이고 정확도도 높지 않습니다 (부록 2)", 13, False, INK),
        ("• 치료 이력과 도체중의 관계는 일관되지만, 품질 등급 영향은 엇갈립니다", 13, False, INK),
        ("• 그래서 시범 농장에서 '치료 기록을 넣으면 예측이 나아지는가'부터 측정합니다", 13, True, INK)])
    source(s, ["grade_rule", "trace_api", "chuksaro2", "onepass", "weather", "pinna"],
           "화면은 축사로 또는 축산물원패스에 기능을 더하는 방식을 검토합니다.")
    notes(s, "【발표 대본】 해결 B는 기록한 만큼 돌려드리는 부분입니다. 축사로에는 이미 혈통, 초음파 측정값, 출하 일정 같은 정보가 있고, 이력제에는 출생일과 성별이 있습니다. "
             "여기에 A에서 자동으로 쌓인 치료 기록을 더해, '지금 팔면 1++ 확률이 얼마, 한 달 더 키우면 얼마, 그때 사료비는 얼마'를 보여주는 것이 목표입니다. "
             "다만 출하 전 등급 예측은 아직 연구 단계이고 정확도도 높지 않습니다. 그래서 효과를 약속하지 않고, 시범 농장에서 치료 기록이 예측을 개선하는지부터 측정하겠습니다.\n"
             "【검증 메모】 화면 수치는 모두 가상입니다. 축사로 기능(혈통 유전능력, 비육 관리·출하 일정, 초음파 육질 진단 입력): 농민신문 2023-03-02. "
             "축산물원패스: 등급판정결과 조회 제공(축산물품질평가원 앱). 예측 정확도 사례: Pinna 외 2025 프리프린트, 흑소 44마리, 도체 형질 R²=0.34. "
             "【법적 주의】 예측값은 참고 정보임을 화면에 고지해야 합니다(실제 판정은 축산물품질평가사가 등급판정 세부기준에 따라 함).")


def s_legal(prs):
    s = blank(prs)
    header(s, "법적 검토", "제안이 지켜야 할 법적 조건",
           "제안 내용을 관련 법령 원문과 대조했습니다. 시범 사업은 현행법 안에서, 확산은 규정 정비와 함께 갑니다.")
    items = [
        ("개인정보", "처방전·판매기록에는 농장주 성명·생년월일·전화번호가 들어 있습니다.",
         "농가 동의를 받아 수집·제공 (개인정보 보호법 제15조·제17조)"),
        ("이력정보 비밀", "이력정보 중 경영·영업상 비밀은 보호되고 목적 외 사용이 금지됩니다.",
         "농장별 등급 데이터는 공식 제공 절차와 가명 처리로 (이력법 제25조, 위반 시 제32조 처벌)"),
        ("판매기록 전산화", "현행 규칙은 기록 항목과 보존 기간만 정하고 서식·전산 제출·농장번호는 정하지 않습니다.",
         "시범은 판매처 자율 참여, 확산 때 취급규칙(농식품부령) 개정 검토"),
        ("출하 경고", "휴약기간이 남았다고 시스템이 출하 신고를 막을 법적 근거는 없습니다.",
         "'차단'이 아니라 '경고'로 설계, 최종 판단은 농가·수의사"),
        ("예측 정보", "예상 등급은 공식 등급판정이 아닙니다.",
         "확률과 '참고용' 고지, 실제 판정은 축산물품질평가사"),
        ("공공데이터", "사용 데이터는 이용허락 제한 없음 또는 공공누리 제1유형(출처표시)입니다.",
         "출처 표기, 발표 자료에 개별 농장명·대표자명은 싣지 않음"),
    ]
    cw = (W - 2 * M - Inches(0.25) * 2) / 3
    for i, (tag, issue, how) in enumerate(items):
        x = M + (cw + Inches(0.25)) * (i % 3)
        y = Inches(2.0) + Inches(2.45) * (i // 3)
        box(s, x, y, cw, Inches(2.3), fill=BG, pad=0.22, lines=[
            (tag, 14, True, ACC), (issue, 12, False, SUB), ("", 3, False, INK), ("→ " + how, 12, True, GREEN)])
    source(s, ["law_privacy", "law_trace", "law_rule", "law_pharm", "trace_api", "haccp"])
    notes(s, "【발표 대본】 제안을 법령 원문과 대조했을 때 지켜야 할 조건은 여섯 가지입니다. "
             "처방전과 판매기록에는 개인정보가 들어 있으므로 농가 동의를 받아야 하고, 이력정보는 영업상 비밀이 보호되므로 공식 절차와 가명 처리를 거칩니다. "
             "판매기록의 전산 제출은 지금 규정에 없으므로 시범은 자율 참여로 하고, 확산할 때 규칙 개정을 검토합니다. "
             "출하를 막는 것이 아니라 경고만 하고, 예상 등급은 참고용임을 밝힙니다.\n"
             "【검증 메모】 개인정보 보호법 제15조①1호·제17조①1호(동의). 이력법 제25조③(경영·영업상 비밀 보호, 방역 효율·축산산업 발전 목적 외 사용 금지)·④(공공기관 직원 등의 누설 금지), 제32조①(3년 이하 징역 또는 3천만 원 이하 벌금). "
             "취급규칙 제22조③은 '판매일자, 제품명, 수량, 용도 및 구매자 등에 관한 기록을 1년 이상 보존'만 규정. "
             "HACCP 인증업체정보에는 대표자명·주소가 들어 있어 발표 자료에는 집계값만 씁니다. "
             "이 슬라이드는 법률 자문이 아니라 조문 대조 결과입니다. 실제 사업화 전에는 소관 부처 유권해석이 필요합니다.")


def s_effect(prs):
    s = blank(prs)
    header(s, "기대 효과", "누구에게 무엇이 달라지나요")
    items = [
        ("농가", "약 기록을 손으로 쓰지 않고, 폐기·과태료를 미리 막고, 소별로 출하 시기를 판단할 수 있습니다."),
        ("한국식품안전관리인증원", "농장 심사를 전산 기록으로 확인하고, 스마트HACCP 농장 0건을 풀 첫 모델을 얻습니다."),
        ("소비자", "농장 단계의 약 관리까지 확인된 한우를 받습니다."),
    ]
    bw = (W - 2 * M - Inches(0.3) * 2) / 3
    for i, (who, body) in enumerate(items):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(1.55), bw, Inches(2.15), fill=BG, pad=0.3, lines=[(who, 17, True, GREEN), ("", 4, False, INK), (body, 14, False, INK)])
    rows = [["성과 지표", "현재", "시범 사업 목표"],
            ["스마트HACCP 등록 한우 농장", "0건", "시범 농장 첫 등록"],
            ["출하 소 중 휴약기간 자동 확인 비율", "측정 안 됨", "시범 농장 출하 건 기준으로 측정 시작"],
            ["농가 약 기록 시간", "측정 안 됨", "도입 전후 비교"],
            ["예상 등급 정확도", "기준 없음", "치료 기록을 넣었을 때 개선 폭 측정"]]
    table(s, M + Inches(0.08), Inches(4.05), W - 2 * M - Inches(0.16), rows,
          [Inches(4.6), Inches(2.3), Inches(5.07)], row_h=Inches(0.54), size=14)
    notes(s, "【발표 대본】 농가는 기록 시간을 줄이고 손실을 막고 출하 판단 정보를 얻습니다. 인증원은 농장 심사를 전산 기록으로 할 수 있고, "
             "스마트HACCP 농장 0건을 풀 첫 모델을 얻습니다. 목표 숫자는 부풀리지 않고, 시범 사업에서 처음으로 측정하는 것을 목표로 잡았습니다.")


def s_limit(prs):
    s = blank(prs)
    header(s, "한계와 검증", "아직 확인이 필요한 부분과 검증 방법")
    box(s, M, Inches(1.55), Inches(5.95), Inches(5.25), fill=BG, pad=0.3, lines=[
        ("아직 확인이 필요한 것", 16, True, INK), ("", 6, False, INK),
        ("연계 가능성", 14, True, INK),
        ("전자처방·판매기록을 축사로로 보낼 수 있는지, 판매기록 전산화 비율은 얼마인지", 12, False, SUB), ("", 4, False, INK),
        ("휴약기간 데이터", 14, True, INK),
        ("제품별 휴약기간을 기계가 읽을 수 있는 형태로 받을 수 있는지", 12, False, SUB), ("", 4, False, INK),
        ("효과 근거", 14, True, INK),
        ("HACCP이 등급을 올리는지는 일본 농장 1곳 사례뿐이고, 치료 기록과 품질 등급의 관계는 연구마다 엇갈림", 12, False, SUB)])
    box(s, Inches(6.75), Inches(1.55), Inches(5.95), Inches(5.25), fill=WHITE, line=GREEN, pad=0.3, lines=[
        ("검증 방법", 16, True, GREEN), ("", 6, False, INK),
        ("비교 농장을 두고 전후를 봅니다", 14, True, INK),
        ("인증 농장의 인증 전후 등급 변화를, 같은 기간 인증받지 않은 비슷한 규모 농장의 변화와 비교합니다. "
         "같은 시기 전국 등급이 함께 오른 효과를 걸러냅니다.", 12, False, SUB), ("", 4, False, INK),
        ("필요한 데이터", 14, True, INK),
        ("• 인증일: HACCP 인증업체정보", 12, False, SUB),
        ("• 소별 등급·농장: 공개 API는 소 1마리씩 조회 → 축산물품질평가원에 가명 처리 자료 제공 요청", 12, False, SUB),
        ("• 농장 연결: 한우 인증 건의 17.8%가 같은 이름을 공유해 주소까지 맞춰 연결", 12, False, SUB)])
    source(s, ["haccp", "trace_api", "akaike23", "grimes"])
    notes(s, "【발표 대본】 한계도 분명히 말씀드립니다. 전자처방과 판매기록을 축사로로 실제로 보낼 수 있는지, 휴약기간 데이터를 기계가 읽을 수 있는지는 시범 사업에서 확인해야 합니다. "
             "HACCP이 등급을 올린다는 근거도 아직 약합니다. 그래서 효과는 인증 농장과 비슷한 비인증 농장을 같은 기간 비교하는 방식으로 검증하겠습니다.\n"
             "【검증 메모】 17.8% = 한우 품목 가축사육업 인증 1,884건 중 업체명이 2건 이상 등장하는 건의 비율(v5의 17.6%를 원자료로 다시 계산). "
             "업체명 기준으로는 7.6%, 서로 다른 시군구에 같은 이름이 있는 경우는 6.3%입니다.")


# ─────────────────────────────────────────────────────────────
# 부록
# ─────────────────────────────────────────────────────────────
def s_appendix_divider(prs):
    s = blank(prs)
    box(s, Inches(0.6), Inches(0.6), W - Inches(1.2), H - Inches(1.2), fill=BG, r=Inches(0.4))
    text(s, Inches(1.2), Inches(2.6), Inches(11), Inches(2.4),
         [("부록", 40, True, INK),
          ("1. 해외 연구 상세   2. 기존 등급 예측 방법   3. 데이터·출처 검증 결과   4. 참고문헌", 17, False, SUB)])


def s_app_research(prs):
    s = blank(prs)
    header(s, "부록 1", "치료 이력과 출하 성적에 관한 연구")
    rows = [["연구", "대상", "결과"],
            ["일본 · Akaike 외 (2023)", "흑소 농장 1곳, 2017~2021년", "호흡기 치료 78 → 28~39마리, 5등급 비율 36.5% → 66~78%. 비교 농장 없음, 사료 첨가제 등 동시 개선"],
            ["미국 · Schneider 외 (2009)", "비육우 5,976마리", "호흡기 병 치료 소는 도체중 8.2kg 가볍고 근내지방 점수도 낮음"],
            ["미국 · Wilson 외 (2017)", "거세우 174마리", "Prime·Choice 비율: 치료 0회 70.3%, 1회 56.5%, 2회 60.2%, 3~4회 36.2%"],
            ["호주 · Blakebrough-Hall 외 (2020)", "거세우 898마리", "3번 이상 치료 소는 도체중 39.6kg 가볍고 1마리당 약 385호주달러 손해"],
            ["미국 · Grimes 외 (2025)", "도체 60,843두", "폐 병변이 심할수록 도체중·가치 감소, 근내지방 점수는 증가 경향"],
            ["일본 · Inoue 외 (2015)", "흑소 5,788마리", "간·복부지방 괴사, 대장 염증과 근내지방도의 유전 상관이 양(+) → 마블링 개량과 일부 질병 위험이 함께 갈 수 있음"]]
    table(s, M + Inches(0.08), Inches(1.6), W - 2 * M - Inches(0.16), rows, [Inches(3.1), Inches(2.5), Inches(6.37)], row_h=Inches(0.74), size=12)
    source(s, ["akaike23", "schneider", "wilson", "blake", "grimes", "inoue"], "모든 수치는 논문 원문 또는 초록으로 확인했습니다.")


def s_app_methods(prs):
    s = blank(prs)
    header(s, "부록 2", "이미 있는 등급 판정·예측 방법")
    rows = [["방법", "누가", "언제", "한계"],
            ["AI 등급판정 장비 (51개 도축장)", "축산물품질평가원", "도축 후", "판정 보조, 최종 판정은 품질평가사"],
            ["등급 예측 계산기 (2019)", "축산물품질평가원", "도축 후", "도체중·등지방 등 도축 후 값 필요"],
            ["날씨·혈통·개체정보로 예측", "날씨 빅데이터 콘테스트 수상작 (2026)", "출하 전", "대회 수상작, 서비스 아님"],
            ["초음파 진단 시스템", "㈜한국육류연구소 주관", "출하 전", "육량등급 83.8% 일치, 장비 대당 2천만 원"],
            ["암소 유전체 분석", "농협·국립축산과학원", "번식 선택 때", "타고난 유전능력만 확인"],
            ["환경·행동·혈액으로 예측", "해외 연구 (프리프린트)", "출하 전", "44마리, 도체 형질 R²=0.34"]]
    table(s, M + Inches(0.08), Inches(1.6), W - 2 * M - Inches(0.16), rows, [Inches(3.3), Inches(3.5), Inches(1.5), Inches(3.67)], row_h=Inches(0.6), size=13)
    box(s, M, Inches(6.1), W - 2 * M, Inches(0.7), fill=GREEN_SOFT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, lines=[
        ("약·치료 기록을 예측에 넣은 방법은 찾지 못했습니다. 해결 B가 더하는 부분입니다.", 14, True, GREEN)])
    source(s, ["ai_grade", "calc2019", "weather", "ultra", "genome", "pinna"])


def s_app_check(prs):
    s = blank(prs)
    header(s, "부록 3", "데이터·출처 검증 결과 (v5 → v6에서 고친 것)")
    rows = [["항목", "v5 표기", "검증 결과 (v6)"],
            ["한우 인증 농장", "1,607곳 · 2.2%", "1,835곳 · 2.5% (업체명+시군구, '한우' 포함 품목)"],
            ["축종별 인증률", "닭 52.9 · 돼지 22.1 · 젖소 11.1", "닭 61.4 · 돼지 24.2 · 젖소 11.8 (같은 기준으로 재계산)"],
            ["50마리 이상 분모", "최대 7.7%", "최대 8.9%"],
            ["60세 이상 63%", "한우 농가처럼 표기", "축산 농가 경영주 전체(2020) 수치로 수정"],
            ["2026년 스마트HACCP", "닭·오리 도축장, 작업장 8곳", "가금 도축장 시스템 개발 + 도축장·집유장 8곳 보급"],
            ["미국 대규모 연구", "6만 마리, 지방 무늬 줄지 않음", "60,843두, 근내지방 점수는 증가 경향"],
            ["해결 A 설계", "판매처→앱 신규", "축사로(이미 약품·처방 기록 칸, 이력제 연동)에 연계"],
            ["이름 중복", "17.6%", "17.8% (원자료 재계산)"]]
    table(s, M + Inches(0.08), Inches(1.6), W - 2 * M - Inches(0.16), rows, [Inches(2.6), Inches(3.9), Inches(5.47)], row_h=Inches(0.54), size=12)
    text(s, M, Inches(6.55), W - 2 * M, Inches(0.4), [
        ("법령은 국가법령정보센터 원문, 정책 수치는 정부 공고 첨부파일, 논문 수치는 원문 PDF로 확인했습니다. 확인하지 못한 항목은 '한계와 검증' 슬라이드에 남겼습니다.", 11, False, SUB)])


def s_refs(prs):
    refs = [(i + 1, REFS[k]) for i, k in enumerate(_order)]
    per = 13
    pages = [refs[i:i + per] for i in range(0, len(refs), per)]
    for pi, chunk in enumerate(pages):
        s = blank(prs)
        header(s, "부록 4", "참고문헌 ({}/{})".format(pi + 1, len(pages)))
        tb = s.shapes.add_textbox(M, Inches(1.5), W - 2 * M, Inches(5.8))
        tf = tb.text_frame
        tf.word_wrap = True
        for k, (no, (org, title, date, url)) in enumerate(chunk):
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.space_after = Pt(1)
            rr = p.add_run()
            rr.text = "[{}] {}, 「{}」, {}".format(no, org, title, date)
            rr.font.name = FONT; rr.font.size = Pt(10); rr.font.bold = True; rr.font.color.rgb = INK
            p2 = tf.add_paragraph()
            p2.space_after = Pt(3)
            r2 = p2.add_run()
            r2.text = "      " + url
            r2.font.name = FONT; r2.font.size = Pt(8.5); r2.font.color.rgb = MUTE
            r2.hyperlink.address = url


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    cover(prs)
    s_data(prs)
    s_status(prs)
    s_cause_structure(prs); s_cause_design(prs); s_cause_policy(prs)
    s_problem(prs)
    s_loss(prs); s_health(prs); s_timing(prs)
    s_direction(prs); s_a(prs); s_b(prs); s_legal(prs)
    s_effect(prs); s_limit(prs)
    s_appendix_divider(prs); s_app_research(prs); s_app_methods(prs); s_app_check(prs)
    s_refs(prs)
    unused = [k for k in REFS if k not in _order]
    # python-pptx 기본 서식에 박힌 문서 정보(2013-01-27, 'Steve Canny')를 지금 값으로 바꾼다
    import datetime
    now = datetime.datetime.now()
    cp = prs.core_properties
    cp.title = "한우 HACCP 발표 초안 v6"
    cp.subject = "2026 한국식품안전관리인증원 공공데이터 분석·활용 경진대회"
    cp.author = ""
    cp.last_modified_by = ""
    cp.created = now
    cp.modified = now
    cp.revision = 1
    p = os.path.join(OUT, "한우HACCP_발표초안_v6.pptx")
    prs.save(p)
    print("[저장]", p, len(prs.slides), "장, 참고문헌", len(_order), "건, 미인용", unused)


if __name__ == "__main__":
    main()
