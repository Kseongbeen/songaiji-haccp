# -*- coding: utf-8 -*-
"""
38. 한우 농장 HACCP 발표 초안 (PPT)

순서: EDA → 문제 인식 → 해결 방안
차트는 이미지가 아니라 PowerPoint 기본 차트로 넣는다. 조원들이 숫자와 색을 바로 고칠 수 있게.

산출: output/한우HACCP_발표초안.pptx
"""

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

OUT = "output"
FONT = "맑은 고딕"

INK = RGBColor(0x15, 0x20, 0x1B)
SUB = RGBColor(0x4C, 0x57, 0x52)
MUTE = RGBColor(0x8A, 0x94, 0x8F)
LINE = RGBColor(0xDD, 0xE3, 0xDF)
BG = RGBColor(0xF4, 0xF6, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACC = RGBColor(0xB8, 0x41, 0x1A)        # 농장·문제
ACC_SOFT = RGBColor(0xF6, 0xE6, 0xDE)
GRAY = RGBColor(0x9A, 0xA3, 0x9F)       # 비교 대상
GREEN = RGBColor(0x0D, 0x7D, 0x57)      # 해결

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.6)


# ─────────────────────────────────────────────────────────────
# 공통 도구
# ─────────────────────────────────────────────────────────────
def blank(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = WHITE
    return s


def text(slide, x, y, w, h, lines, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    """lines: [(글자, 크기, 굵게, 색), ...]"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.02)
    for i, (t, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(size * 0.45)
        r = p.add_run()
        r.text = t
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def header(slide, section, title, sub=None):
    text(slide, M, Inches(0.35), Inches(6), Inches(0.35), [(section, 12, True, ACC)])
    text(slide, M, Inches(0.65), W - 2 * M, Inches(0.7), [(title, 26, True, INK)])
    if sub:
        text(slide, M, Inches(1.35), W - 2 * M, Inches(0.5), [(sub, 14, False, SUB)])


def source(slide, t):
    text(slide, M, Inches(7.0), W - 2 * M, Inches(0.3), [(t, 9, False, MUTE)])


def box(slide, x, y, w, h, fill=BG, line=None, lines=(), anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
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
    tf.margin_left = tf.margin_right = Inches(0.18)
    tf.margin_top = tf.margin_bottom = Inches(0.12)
    for i, (t, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(size * 0.4)
        r = p.add_run()
        r.text = t
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return sh


def arrow(slide, x, y, w=Inches(0.35), h=Inches(0.35)):
    sh = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = MUTE
    sh.line.fill.background()
    sh.shadow.inherit = False


def table(slide, x, y, w, rows, col_w, row_h=Inches(0.42), hot_rows=(), size=13, num_cols=()):
    t = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w, row_h * len(rows)).table
    for j, cw in enumerate(col_w):
        t.columns[j].width = cw
    for i, row in enumerate(rows):
        t.rows[i].height = row_h
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = BG if i == 0 else (ACC_SOFT if i in hot_rows else WHITE)
            c.margin_left = c.margin_right = Inches(0.1)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if (j in num_cols and i > 0) else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(v)
            r.font.name = FONT
            r.font.size = Pt(size - 1 if i == 0 else size)
            r.font.bold = (i == 0) or (i in hot_rows)
            r.font.color.rgb = SUB if i == 0 else (ACC if i in hot_rows else INK)
    return t


def style_axis(ax, size=11, visible=True):
    ax.tick_labels.font.size = Pt(size)
    ax.tick_labels.font.name = FONT
    ax.tick_labels.font.color.rgb = SUB
    ax.format.line.color.rgb = LINE
    ax.visible = visible


def bars(slide, x, y, w, h, cats, vals, hot=(), horizontal=True, fmt='0.0"%"', vmax=None, label_size=12):
    d = CategoryChartData()
    d.categories = cats
    d.add_series("값", vals)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    ch = slide.shapes.add_chart(kind, x, y, w, h, d).chart
    ch.has_title = False
    ch.has_legend = False
    ch.font.name = FONT
    plot = ch.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = fmt
    dl.number_format_is_linked = False
    dl.font.size = Pt(label_size)
    dl.font.name = FONT
    dl.font.color.rgb = INK
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    ser = plot.series[0]
    for i in range(len(vals)):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = ACC if i in hot else GRAY
    style_axis(ch.category_axis, 12)
    va = ch.value_axis
    va.has_major_gridlines = False
    va.visible = False
    if vmax is not None:
        va.maximum_scale = vmax
    va.minimum_scale = 0
    if horizontal:
        ch.category_axis.reverse_order = True
    return ch


# ─────────────────────────────────────────────────────────────
# 슬라이드
# ─────────────────────────────────────────────────────────────
def cover(prs):
    s = blank(prs)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.25), H)
    bar.fill.solid(); bar.fill.fore_color.rgb = ACC; bar.line.fill.background()
    text(s, Inches(1.0), Inches(2.0), Inches(11), Inches(0.5),
         [("2026 한국식품안전관리인증원 공공데이터 분석·활용 경진대회", 14, False, SUB)])
    text(s, Inches(1.0), Inches(2.6), Inches(11.3), Inches(1.8),
         [("기록은 자동으로, 결과는 등급으로", 40, True, INK),
          ("HACCP·스마트HACCP에서 비어 있는 한우 농장을 기존 데이터 연계로 채우는 방안", 20, False, SUB)])
    text(s, Inches(1.0), Inches(5.6), Inches(8), Inches(0.5), [("발표 초안 · 조원 검토용", 13, False, MUTE)])


def divider(prs, no, name, desc):
    s = blank(prs)
    s.background.fill.fore_color.rgb = BG
    text(s, Inches(1.0), Inches(2.5), Inches(3), Inches(1.2), [(no, 60, True, ACC)])
    text(s, Inches(1.0), Inches(3.7), Inches(11), Inches(1.4), [(name, 36, True, INK), (desc, 17, False, SUB)])


# ─────────────────────────────── 시작하기 전에
def s_terms(prs):
    s = blank(prs)
    header(s, "시작하기 전에", "이 자료에서 쓰는 용어", "처음 보시는 분도 이해하실 수 있도록 자주 나오는 용어를 먼저 정리했습니다.")
    rows = [
        ["용어", "뜻"],
        ["HACCP (해썹)", "식품·축산물을 만들고 기르는 과정에서 위험 요소를 미리 찾아 중점적으로 관리하는 위생 관리 기준입니다."],
        ["한국식품안전관리인증원", "HACCP 인증 심사와 사후 관리를 맡는 식품의약품안전처 산하 공공기관입니다. 이번 공모전의 주최 기관입니다."],
        ["스마트HACCP", "온도·금속 검사 같은 관리 기록을 기계와 전산으로 자동 기록하는 HACCP 운영 방식입니다."],
        ["중요관리점", "위험을 막기 위해 반드시 지켜야 하는 핵심 관리 항목입니다. 예) 냉장 온도, 출하 전 휴약기간"],
        ["휴약기간", "가축에게 약을 쓴 뒤 고기나 우유에 약 성분이 남지 않도록 팔지 않고 기다려야 하는 기간입니다."],
        ["거세우", "고기용으로 키우려고 어릴 때 거세한 수소입니다. 한우 고기로 팔리는 소의 약 절반입니다."],
        ["축산물이력제", "소 한 마리마다 태어난 날·옮겨 간 농장·도축한 날·등급을 기록하고 조회할 수 있게 한 제도입니다."],
        ["축사로", "국립축산과학원이 운영하는 농장용 HACCP 기록·가축 관리 전산 프로그램입니다."],
        ["전자처방 (eVET)", "수의사가 처방이 필요한 동물 약의 처방전을 전자로 발급하는 시스템입니다."],
    ]
    table(s, M, Inches(1.9), W - 2 * M, rows, [Inches(2.8), Inches(9.33)], row_h=Inches(0.49), size=12)
    source(s, "[참고 1, 2, 12, 13, 14, 20, 39] 용어 뜻은 각 기관 자료와 법령 설명을 바탕으로 쉽게 풀어 썼습니다.")


def s_withdrawal(prs):
    s = blank(prs)
    header(s, "배경 1", "휴약기간은 먹는 사람의 건강을 지키기 위한 약속입니다",
           "가축에게 약을 쓴 뒤, 고기·우유·알에 남은 약 성분이 기준 아래로 빠져나갈 때까지 팔지 않고 기다리는 기간입니다.")
    items = [
        ("왜 필요한가요", "사람의 건강 때문입니다",
         "고기에 남은 약 성분은 대부분 해가 될 정도는 아닙니다. 하지만 일부 사람에게는 해가 될 수 있어 나라가 기준을 정해 관리합니다. 약은 몸에서 빠져나가는 속도가 달라 약마다 휴약기간이 따로 정해져 있습니다."),
        ("2024년 1월부터", "기준이 훨씬 엄격해졌습니다",
         "소·돼지·닭고기, 우유, 달걀에서 기준이 따로 정해지지 않은 약은 고기 1kg에 0.01mg 이하만 허용됩니다. 사실상 허가받은 약을 정해진 방법대로만 써야 합니다."),
        ("어기면", "농가 손해가 큽니다",
         "휴약기간 등 사용 기준을 어기면 100만 원 이하 과태료가 나옵니다. 기준을 넘은 고기는 전부 폐기되고, 그 농가는 6개월 동안 팔 때마다 검사를 받습니다."),
    ]
    bw = (W - 2 * M - Inches(0.3) * 2) / 3
    for i, (tag, head, body) in enumerate(items):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(2.1), bw, Inches(3.75), fill=BG, lines=[
            (tag, 12, True, ACC), (head, 17, True, INK), ("", 5, False, INK), (body, 13, False, SUB)])
    box(s, M, Inches(6.05), W - 2 * M, Inches(0.8), fill=ACC_SOFT, lines=[
        ("지금은 농장에서 사람이 약품 기록부를 보고 휴약기간을 지켰는지 확인합니다. "
         "한 번 놓치면 폐기와 과태료로 이어지기 때문에, 이 확인을 자동으로 돕는 장치가 필요합니다.", 13, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE)
    source(s, "[참고 30, 31, 32] 식품안전나라 잔류물질정보, 축·수산물 동물용의약품 PLS 제도(2024-01-01 시행), 농림축산식품부 동물용의약품 안전사용 안내")


def s_grading(prs):
    s = blank(prs)
    header(s, "배경 2", "한우 등급은 소를 잡은 뒤에 매깁니다",
           "축산물품질평가원의 품질평가사가 도축장에서, 반으로 가른 소의 등 부위를 잘라 드러난 고기 단면을 보고 판정합니다.")
    box(s, M, Inches(2.05), Inches(5.95), Inches(2.75), fill=BG, lines=[
        ("① 고기 양 등급 · A / B / C", 17, True, INK), ("", 5, False, INK),
        ("고기가 얼마나 많이 나오는지 봅니다.", 13, True, SUB),
        ("소를 잡고 머리·가죽·내장·발을 뗀 몸통 무게, 등 지방 두께, 등심 크기를 계산식에 넣어 정합니다. "
         "품종과 성별에 따라 계산식이 6가지로 나뉩니다.", 13, False, SUB)])
    box(s, Inches(6.75), Inches(2.05), Inches(5.95), Inches(2.75), fill=BG, lines=[
        ("② 고기 품질 등급 · 1++ / 1+ / 1 / 2 / 3", 17, True, INK), ("", 5, False, INK),
        ("고기 품질이 어떤지 봅니다.", 13, True, SUB),
        ("고기 속 하얀 지방 무늬, 고기 색, 지방 색, 고기 결, 소의 나이 흔적을 봅니다. "
         "어느 등급에도 넣기 어려운 경우는 '등외'로 판정합니다.", 13, False, SUB)])
    text(s, M, Inches(5.0), Inches(7.9), Inches(1.9), [
        ("2019년 12월에 바뀐 점", 13, True, INK),
        ("• 최고 등급(1++)을 받는 지방 무늬 기준을 낮췄습니다.", 12, False, SUB),
        ("• 지방 무늬·고기 색·지방 색·결을 따로 평가해 가장 낮은 결과를 최종 등급으로 정합니다.", 12, False, SUB),
        ("• 전국 51개 도축장에는 사진으로 등심 크기와 지방 무늬를 재는 AI 장비가 있고, 최종 판정은 사람이 합니다.", 12, False, SUB)])
    box(s, Inches(8.75), Inches(5.05), Inches(3.95), Inches(1.75), fill=ACC_SOFT, lines=[
        ("핵심", 12, True, ACC),
        ("등급은 소를 잡은 뒤에야 알 수 있습니다. 농가는 팔기 전까지 어떤 등급이 나올지 모릅니다.", 13, True, ACC)])
    source(s, "[참고 33, 34, 35, 36] 찾기쉬운 생활법령 축산물 표시기준, 경향신문(2019-11-25), 한국축산물처리협회 등급 기준, 축산신문 AI 등급판정 장비")


# ─────────────────────────────── EDA
def s_data(prs):
    s = blank(prs)
    header(s, "EDA", "분석에 사용한 데이터",
           "한국식품안전관리인증원이 공개한 데이터 3종을 가축 통계·예산·정책 자료와 연결했습니다.")
    rows = [
        ["데이터", "제공 기관", "규모", "활용 내용"],
        ["HACCP 인증업체정보", "한국식품안전관리인증원", "41,069건", "축종·업종별 인증 현황, 인증일"],
        ["축산물 스마트HACCP 인증업체", "한국식품안전관리인증원", "298건", "스마트HACCP 등록 현황, 중요관리점"],
        ["스마트HACCP 기술지원 업체", "한국식품안전관리인증원", "2,719건", "스마트HACCP 등록 전 준비 단계"],
        ["가축동향조사 (2026년 2분기)", "국가데이터처", "축종별 농장 수", "인증률 계산의 분모"],
        ["축산물 HACCP 컨설팅 시행지침", "농림축산식품부", "2019~2024년 예산", "농장 인증 지원 규모"],
        ["생산단계 스마트HACCP 추진계획", "농림축산식품부", "2024·2025년", "계획과 실제 사업 비교"],
    ]
    table(s, M, Inches(2.0), W - 2 * M, rows,
          [Inches(3.6), Inches(3.2), Inches(2.2), Inches(3.13)], row_h=Inches(0.55), size=14)
    source(s, "[참고 1~7] 축산물 스마트HACCP 데이터는 원본이 563행이지만, 한 업체가 여러 행으로 나뉘어 있어 인증번호 기준으로 정리하면 298건입니다.")


def s_rate(prs):
    s = blank(prs)
    header(s, "EDA 1", "한우 농장의 HACCP 인증률은 2.2%입니다",
           "인증률 = HACCP 인증을 받은 농장 수 ÷ 전국 사육 농장 수. 다섯 축종 중 가장 낮습니다.")
    bars(s, M, Inches(1.95), Inches(7.4), Inches(4.8),
         ["닭", "돼지", "오리", "젖소", "한우"], [52.9, 22.1, 20.9, 11.1, 2.2], hot={4}, vmax=65)
    rows = [["축종", "인증 농장", "전국 농장"],
            ["닭", "1,651", "3,123"], ["돼지", "1,209", "5,465"], ["오리", "105", "502"],
            ["젖소", "591", "5,342"], ["한우", "1,607", "72,389"]]
    table(s, Inches(8.4), Inches(2.1), Inches(4.3), rows,
          [Inches(1.3), Inches(1.5), Inches(1.5)], hot_rows=(5,), num_cols=(1, 2))
    box(s, Inches(8.4), Inches(4.85), Inches(4.3), Inches(1.95), fill=BG, lines=[
        ("비교하실 때 주의할 점", 12, True, INK),
        ("닭은 3천 마리 이상, 오리는 2천 마리 이상 기르는 농가만 통계에 포함됩니다. "
         "반면 한우는 한두 마리를 기르는 농가까지 모두 포함되어 있어, 인증률이 상대적으로 더 낮게 계산됩니다.", 11, False, SUB)])
    source(s, "[참고 1, 4] 한국식품안전관리인증원 HACCP 인증업체정보(2026-09-08 수집), 국가데이터처 가축동향조사(2026-06-01 기준)")


def s_why(prs):
    s = blank(prs)
    header(s, "EDA 1 · 이유", "농장은 HACCP을 받지 않아도 고기를 팔 수 있습니다",
           "HACCP은 법으로 반드시 받아야 하는 곳과, 원하면 신청하는 곳으로 나뉩니다. 한우 농장은 신청하는 쪽입니다.")
    steps = [
        ("① 농장", "소를 기르는 곳\n(1차 생산)", "원하면 신청", False),
        ("② 도축장", "소를 잡아\n고기로 만드는 곳", "법으로 의무", True),
        ("③ 가공·포장 공장", "고기를 자르고\n포장하는 곳", "법으로 의무", True),
        ("④ 정육점·마트", "고기를 파는 곳", "원하면 신청", False),
    ]
    bw, gap = Inches(2.7), Inches(0.44)
    for i, (head, body, tag, must) in enumerate(steps):
        x = M + (bw + gap) * i
        box(s, x, Inches(2.05), bw, Inches(2.2), fill=ACC_SOFT if i == 0 else BG, lines=[
            (head, 17, True, ACC if i == 0 else INK), ("", 4, False, INK), (body, 13, False, SUB)],
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        box(s, x, Inches(4.35), bw, Inches(0.5), fill=GREEN if must else WHITE, line=None if must else LINE, lines=[
            (tag, 13, True, WHITE if must else MUTE)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < len(steps) - 1:
            arrow(s, x + bw + Inches(0.06), Inches(2.95))
    box(s, M, Inches(5.05), Inches(4.9), Inches(1.85), fill=BG, lines=[
        ("왜 2.2%밖에 안 될까요", 13, True, INK),
        ("농장은 소를 기르기만 하고, 직접 잡거나 가공하지 않는 1차 생산자입니다. "
         "소비자가 사는 한우고기는 이미 HACCP이 의무인 도축장과 가공 공장을 거치기 때문에, "
         "농장은 인증이 없어도 소를 팔 수 있습니다.", 11, False, SUB)])
    box(s, Inches(5.7), Inches(5.05), Inches(7.0), Inches(1.85), fill=ACC_SOFT, lines=[
        ("그래도 농장 단계가 필요한 세 가지 이유", 13, True, ACC),
        ("① 먹는 사람의 건강: 약이 남은 고기를 막는 일은 농장에서만 할 수 있습니다. 도축장에서는 폐기밖에 방법이 없습니다.", 11, False, INK),
        ("② 농가의 손해 방지: 휴약기간을 놓치면 전량 폐기와 과태료를 물고, 아팠던 소는 몸통 무게가 가볍게 나옵니다.", 11, False, INK),
        ("③ 소의 건강: 약 기록이 쌓이면 자주 아픈 소를 일찍 알아챕니다. 일본 농장 HACCP 기준은 동물의 건강과 복지도 목표로 합니다.", 11, False, INK)])
    source(s, "[참고 44, 45, 46] 축산물 위생관리법 제9조·시행규칙 제7조. 의무 대상: 도축업·집유업·식육가공업·유가공업·알가공업·식용란선별포장업·식육포장처리업(울릉도·백령도 제외)")


def s_same(prs):
    s = blank(prs)
    header(s, "EDA 2", "한우 농가가 인증을 덜 받은 것은 아닙니다",
           "인증받은 농가 수는 닭과 비슷하지만, 전국 한우 농가 수가 닭의 약 23배여서 비율이 낮게 나옵니다.")
    text(s, M, Inches(1.9), Inches(6), Inches(0.35), [("지금까지 인증받은 농가 수", 12, True, SUB)])
    bars(s, M, Inches(2.2), Inches(6.0), Inches(4.0),
         ["닭", "한·육우", "돼지", "젖소", "오리"], [1651, 1612, 1209, 591, 105],
         hot={1}, fmt='#,##0"곳"', vmax=2100)
    text(s, Inches(6.9), Inches(1.9), Inches(5.8), Inches(0.35), [("최근 5년간 한 해 평균 새로 인증받은 농가 수", 12, True, SUB)])
    bars(s, Inches(6.9), Inches(2.2), Inches(5.8), Inches(4.0),
         ["닭", "한·육우", "돼지", "젖소", "오리"], [60, 90, 41, 44, 1],
         hot={1}, fmt='"연 "0"곳"', vmax=115)
    box(s, M, Inches(6.15), W - 2 * M, Inches(0.72), fill=ACC_SOFT, lines=[
        ("최근에는 한우 농가가 닭보다 더 많이 인증을 받고 있습니다. 그러나 한 해 90곳 속도로는 7만여 곳을 인증하기 어렵습니다. "
         "농가를 설득하는 것만으로는 풀리지 않는 구조적인 문제입니다.", 12, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE)
    source(s, "[참고 1, 4] '닭은 계열사가 여러 농가를 묶어 신청해 인증이 쉽다'는 가설도 검토했습니다. 같은 날 5곳 이상이 함께 인증받은 비중이 한·육우 14.3%, 닭 4.5%로 오히려 한우가 높아 가설과 맞지 않았습니다.")


def s_trend(prs):
    s = blank(prs)
    header(s, "EDA 3", "HACCP 신규 인증은 두 배로 늘었지만, 농장만 줄었습니다",
           "한 해 동안 새로 HACCP 인증을 받은 곳(식품·축산물 전체): 2015년 990곳 → 2025년 2,169곳 / 그중 축산물 농장: 407곳 → 194곳")
    t = pd.read_csv(OUT + "/부문별_연도별신규.csv", index_col=0).loc[2015:2025]
    d = CategoryChartData()
    d.categories = [str(y) for y in t.index]
    for c in ["식품", "축산물 가공·유통", "축산물 농장"]:
        d.add_series(c, [int(v) for v in t[c]])
    ch = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, M, Inches(1.9), W - 2 * M, Inches(4.15), d).chart
    ch.has_title = False
    ch.font.name = FONT
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(12)
    colors = [RGBColor(0x6B, 0x75, 0x71), RGBColor(0xB9, 0xC1, 0xBD), ACC]
    widths = [2.0, 2.0, 3.5]
    for ser, col, wd in zip(ch.plots[0].series, colors, widths):
        ser.format.line.color.rgb = col
        ser.format.line.width = Pt(wd)
        ser.smooth = False
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = col
        ser.marker.format.line.color.rgb = col
    style_axis(ch.category_axis, 11)
    style_axis(ch.value_axis, 11)
    ch.value_axis.major_gridlines.format.line.color.rgb = LINE
    box(s, M, Inches(6.1), W - 2 * M, Inches(0.75), fill=BG, lines=[
        ("인증 전체는 크게 늘었으므로 심사 여력이 부족해서 생긴 현상은 아닙니다. 식품과 축산물 가공·유통은 늘었는데, "
         "농장 쪽으로 새로 들어오는 흐름만 줄어든 것입니다.", 12, True, INK)], anchor=MSO_ANCHOR.MIDDLE)
    source(s, "[참고 1] 각 업소가 처음 인증받은 연도 기준입니다. 인증이 만료된 곳은 공개 데이터에서 빠져 과거 연도가 실제보다 작게 잡히는데, 그럼에도 농장은 2015년이 가장 많았습니다.")


def s_smart(prs):
    s = blank(prs)
    header(s, "EDA 4", "스마트HACCP에 등록한 농장은 한 곳도 없습니다",
           "업종별로 HACCP 인증을 받은 곳 중 스마트HACCP까지 등록한 비율입니다.")
    cats = ["유가공업", "알가공업", "식육포장처리업", "식육가공업", "식육판매업", "식용란선별포장업", "식용란수집판매업", "농장"]
    vals = [7.59, 3.95, 3.28, 2.75, 1.64, 1.02, 0.50, 0.0]
    bars(s, M, Inches(1.9), Inches(7.6), Inches(5.0), cats, vals, hot={7}, fmt='0.00"%"', vmax=9.5, label_size=11)
    box(s, Inches(8.6), Inches(1.95), Inches(4.1), Inches(1.35), fill=WHITE, line=LINE, lines=[
        ("HACCP 인증을 받은 농장", 12, False, SUB), ("6,294곳 중 0곳", 24, True, ACC)])
    box(s, Inches(8.6), Inches(3.45), Inches(4.1), Inches(1.35), fill=WHITE, line=LINE, lines=[
        ("스마트HACCP 기술지원 요청", 12, False, SUB), ("2,719건 중 농장 1건", 22, True, ACC)])
    box(s, Inches(8.6), Inches(4.95), Inches(4.1), Inches(1.85), fill=BG, lines=[
        ("기술지원은 스마트HACCP 등록 전에 받는 준비 단계입니다. 축산물 가공·유통 업소는 1,154건을 요청했지만 "
         "농장은 1건뿐이어서, 앞으로 등록할 농장도 거의 없는 상황입니다.", 12, False, SUB)])
    source(s, "[참고 2, 3] 농장은 가축사육업·종축업·부화업을 합친 수치입니다. 출처: 한국식품안전관리인증원 축산물 스마트HACCP 인증업체 정보, 스마트HACCP 기술지원 업체 정보")


def s_ccp(prs):
    s = blank(prs)
    header(s, "EDA 5", "스마트HACCP은 기계로 재는 항목을 중심으로 운영됩니다",
           "스마트HACCP 등록 업체의 중요관리점 668개 중 90.3%가 냉장 온도·금속 검사·가열처럼 기계로 재는 항목입니다.")
    bars(s, M, Inches(1.95), Inches(6.2), Inches(4.3),
         ["냉장·냉동 보관", "금속 검사", "가열·살균", "이물질 검사", "그 외"], [379, 159, 55, 10, 65],
         hot=set(), fmt='#,##0"개"', vmax=460)
    y = Inches(2.0)
    box(s, Inches(7.2), y, Inches(2.65), Inches(4.0), fill=BG, lines=[
        ("스마트HACCP 등록 업체", 13, True, INK), ("", 6, False, INK),
        ("무엇을 관리하나요", 11, False, MUTE), ("냉장 온도, 금속 검사, 살균 온도", 13, False, INK), ("", 6, False, INK),
        ("누가 기록하나요", 11, False, MUTE), ("온도계·검사 기계가 자동으로 기록합니다", 13, False, INK)])
    box(s, Inches(10.05), y, Inches(2.65), Inches(4.0), fill=ACC_SOFT, lines=[
        ("한우 농장", 13, True, ACC), ("", 6, False, INK),
        ("무엇을 관리하나요", 11, False, MUTE), ("팔기 전 휴약기간을 지켰는지, 주사 바늘이 남지 않았는지", 13, False, INK), ("", 6, False, INK),
        ("누가 기록하나요", 11, False, MUTE), ("사람이 약품 사용 기록부를 보고 확인합니다", 13, False, INK)])
    box(s, M, Inches(6.3), W - 2 * M, Inches(0.6), fill=WHITE, line=LINE, lines=[
        ("농장의 핵심 관리 항목은 기계로 잴 수 없습니다. 그래서 지금의 스마트HACCP 방식으로는 농장이 참여하기 어렵습니다.", 13, True, INK)],
        anchor=MSO_ANCHOR.MIDDLE)
    source(s, "[참고 2, 20, 21] '완제품 냉장보관', '완제품냉장보관공정'처럼 표기가 다른 항목은 핵심 단어 기준으로 묶어 집계했습니다.")


def s_plan(prs):
    s = blank(prs)
    header(s, "EDA 6", "계획에는 농장이 있었지만, 실제 사업과 예산에서는 빠졌습니다")
    text(s, M, Inches(1.45), Inches(6.3), Inches(0.4), [("농림축산식품부 '생산단계 스마트HACCP' 계획과 실제 사업", 13, True, SUB)])
    rows = [["연도", "계획", "실제 사업"],
            ["2024", "집유장", "집유장 2곳"],
            ["2025", "도축장, 농장(젖소)", "도축장, 집유장"],
            ["2026", "닭·오리 도축장, 농장(소·돼지·닭)", "닭·오리 도축장, 작업장 8곳"]]
    table(s, M, Inches(1.9), Inches(6.3), rows, [Inches(0.9), Inches(2.9), Inches(2.5)],
          row_h=Inches(0.6), hot_rows=(2, 3), size=13)
    text(s, M, Inches(4.45), Inches(6.3), Inches(2.2), [
        ("2025년과 2026년 계획에 농장이 포함되어 있었지만, 실제 사업은 두 해 모두 도축장과 집유장만 대상으로 했습니다.", 13, True, ACC),
        ("2025년 사업 신청서 양식에는 '농장' 항목까지 마련되어 있었습니다.", 12, False, SUB),
        ("※ 집유장은 농장에서 짠 우유를 모으는 곳, 도축장은 가축을 도축하는 곳입니다.", 11, False, MUTE)])
    text(s, Inches(7.3), Inches(1.45), Inches(5.4), Inches(0.4), [("축산물 HACCP 컨설팅 지원 예산 (백만 원)", 13, True, SUB)])
    bars(s, Inches(7.2), Inches(1.85), Inches(5.5), Inches(3.3),
         ["2019", "2020", "2021", "2022", "2023", "2024"], [2800, 2800, 2593, 2593, 2033, 1015],
         hot={5}, horizontal=False, fmt='#,##0', vmax=3300)
    box(s, Inches(7.3), Inches(5.25), Inches(5.4), Inches(1.6), fill=ACC_SOFT, lines=[
        ("5년 사이 64% 감소했습니다", 14, True, ACC),
        ("2024년 예산을 모두 농장 컨설팅에 쓰더라도 126~169곳만 지원할 수 있습니다(농장 1곳당 600만~800만 원). "
         "같은 해 새로 인증받은 농장은 176곳이었습니다.", 11, False, INK)])
    source(s, "[참고 5, 6, 7, 8, 9] 농림축산식품부 2024·2025 생산단계 축산물 스마트 HACCP 현장구축사업 추진계획, 2024 축산물 HACCP 컨설팅 지원사업 시행지침, 대한급식신문(2026-04-02). 2026년은 보도 기준입니다.")


def s_people(prs):
    s = blank(prs)
    header(s, "EDA 7", "한우 농가는 규모가 작고 농장주 나이가 많습니다")
    box(s, M, Inches(1.7), Inches(5.9), Inches(3.2), fill=BG, lines=[
        ("농장주가 60세 이상인 한우 농가", 16, False, SUB), ("63%", 80, True, ACC), ("통계청 2020 농림어업총조사 기준", 12, False, MUTE)],
        anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    box(s, Inches(6.8), Inches(1.7), Inches(5.9), Inches(3.2), fill=BG, lines=[
        ("소를 50마리 미만 기르는 농가", 16, False, SUB), ("72.6%", 80, True, ACC), ("한·육우 농장 75,647곳 중 54,884곳", 12, False, MUTE)],
        anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    text(s, M, Inches(5.25), W - 2 * M, Inches(1.6), [
        ("HACCP 인증을 받으려면 약을 쓸 때마다 기록을 남기고, 관리 기준서를 계속 관리해야 합니다.", 15, False, INK),
        ("소 수십 마리를 기르는 나이 든 농장주에게는 이 기록 작업 자체가 가장 큰 부담입니다.", 15, True, INK)])
    source(s, "[참고 4, 11] 한국농어민신문(2022-04-05, 통계청 2020 농림어업총조사 인용: 60~69세 41.4%, 70세 이상 21.6%), 국가데이터처 가축동향조사(2026-06-01, 사육 규모별 농장 수)")


def s_research(prs):
    s = blank(prs)
    header(s, "EDA 8", "해외 연구: 아팠던 소는 몸통 무게가 가볍고 등급이 낮았습니다",
           "치료 기록과 도축 결과를 함께 본 연구들입니다. 한우를 대상으로 한 연구는 찾지 못했습니다.")
    rows = [["연구", "대상", "결과"],
            ["일본 · Akaike 외 (2023)", "흑소 농장 1곳, HACCP 도입 전후 5년", "호흡기 병 치료 78마리 → 28~39마리, 최고 등급 비율 36.5% → 73.1% (EDA 9에서 자세히)"],
            ["미국 · Schneider 외 (2009)", "소 5,976마리", "호흡기 병을 치료한 소는 몸통 무게가 8.2kg 가볍고 지방 무늬 점수도 낮았습니다. 치료 횟수가 많을수록 더 나빴습니다."],
            ["미국 · Wilson 외 (2017)", "소 174마리", "상위 두 등급 비율이 치료 0회 70.3%, 1회 56.5%, 3~4회 36.2%였습니다."],
            ["호주 · Blakebrough-Hall 외 (2020)", "소 898마리", "3번 이상 치료한 소는 몸통 무게가 39.6kg 가볍고, 한 마리당 약 385호주달러 손해였습니다."],
            ["폴란드 · Przyborowska-Zhalniarovich 외 (2023)", "돼지 3,060마리", "폐에 병이 있던 돼지는 살코기 비율이 60.0%에서 58%대로 낮았습니다."]]
    table(s, M, Inches(1.9), Inches(8.3), rows, [Inches(2.5), Inches(2.0), Inches(3.8)], row_h=Inches(0.72), size=11)
    box(s, Inches(9.1), Inches(1.9), Inches(3.6), Inches(4.35), fill=BG, lines=[
        ("함께 봐야 할 점", 13, True, INK), ("", 4, False, INK),
        ("• 몸통 무게가 준다는 결과는 연구마다 같지만, 지방 무늬는 엇갈립니다. 미국 소 6만 마리 연구에서는 줄지 않았습니다.", 11, False, SUB),
        ("• 일본 연구에서는 지방 무늬가 좋은 혈통이 오히려 일부 병에 더 잘 걸리는 경향이 나왔습니다.", 11, False, SUB),
        ("• 일본 HACCP 연구는 농장 1곳이고 사료 등 다른 개선도 함께 해, HACCP만의 효과로 단정하기 어렵습니다.", 11, False, SUB)])
    box(s, M, Inches(6.35), W - 2 * M, Inches(0.55), fill=ACC_SOFT, lines=[
        ("휴약기간 자체가 등급을 바꾼다는 연구는 없습니다. 약을 쓰게 만든 병이 몸통 무게와 등급에 영향을 주는 구조입니다.", 13, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE)
    source(s, "[참고 23~29] 각 논문의 원문으로 수치를 확인했습니다. 몸통 무게 = 소를 잡고 머리·가죽·내장·발을 뗀 무게입니다.")


def s_japan_beef(prs):
    s = blank(prs)
    header(s, "EDA 9", "일본 흑소 농장: HACCP 뒤 아픈 소는 줄고 최고 등급은 두 배가 됐습니다",
           "흑소 280마리를 키우는 가족 농장(4명)이 2018년 4월 준비를 시작해 2019년 8월 운영, 2020년 8월 인증을 받았습니다.")
    box(s, M, Inches(1.95), Inches(3.85), Inches(3.35), fill=BG, lines=[
        ("이 농장의 관리 방법", 14, True, INK), ("", 4, False, INK),
        ("중요관리점: 도축장에 보낼 소 고르기", 12, True, SUB),
        ("• 출하할 소의 귀표 번호를 약 사용 기록·주사 바늘 기록과 대조", 12, False, SUB),
        ("• 휴약기간이 남은 소는 출하를 미룸", 12, False, SUB),
        ("• 2019~2021년 기준 위반 0건", 12, True, GREEN), ("", 4, False, INK),
        ("시작한 이유: 소를 보내는 가공 공장이 싱가포르 수출 기준에 맞춰 항생제·바늘 관리를 요청했습니다.", 11, False, MUTE)])
    text(s, Inches(4.7), Inches(1.9), Inches(4.0), Inches(0.35), [("들여온 지 6개월 안에 호흡기 병으로 치료받은 소", 12, True, SUB)])
    bars(s, Inches(4.6), Inches(2.25), Inches(4.0), Inches(3.05),
         ["2017(전)", "2018", "2019", "2020", "2021"], [78, 28, 39, 29, 38],
         hot={0}, horizontal=False, fmt='0"마리"', vmax=95, label_size=11)
    text(s, Inches(8.75), Inches(1.9), Inches(4.0), Inches(0.35), [("최고 등급(5등급)을 받은 소의 비율", 12, True, SUB)])
    bars(s, Inches(8.65), Inches(2.25), Inches(4.05), Inches(3.05),
         ["2017(전)", "2018", "2019", "2020", "2021"], [36.5, 66.0, 77.5, 77.1, 73.1],
         hot={1, 2, 3, 4}, horizontal=False, fmt='0.0"%"', vmax=95, label_size=11)
    box(s, M, Inches(5.45), Inches(5.3), Inches(1.4), fill=WHITE, line=LINE, lines=[
        ("매출도 늘었습니다", 13, True, INK),
        ("최고 등급 시세는 kg당 2,798엔(2017)에서 2,666엔(2021)으로 내렸지만, "
         "최고 등급 소가 늘어 전체 매출은 2021년이 더 높았습니다. 죽은 소도 4~5마리에서 1~2마리로 줄었습니다(통계적 차이는 아님).", 11, False, SUB)])
    box(s, Inches(6.05), Inches(5.45), Inches(6.65), Inches(1.4), fill=ACC_SOFT, lines=[
        ("주의할 점", 13, True, ACC),
        ("도입 전 비교 기록이 2017년 1년뿐이고, 비교할 농장이 없습니다. 같은 시기에 감시 카메라·겨울 바람막이·물통 소독·사료 첨가제 변경도 함께 했습니다. "
         "연구진도 '주로 기본 위생 관리 덕분'이라고 봤습니다.", 11, False, INK)])
    source(s, "[참고 23] Akaike 외(2023) 본문. 호흡기 병 치료 수와 5등급 비율은 2018~2021년 모두 2017년보다 통계적으로 의미 있게 좋았습니다. 일본 등급은 5등급이 최고입니다.")


def s_japan_compare(prs):
    s = blank(prs)
    header(s, "EDA 10", "일본 소 농장 2곳은 약 기록 대조를 사람이 손으로 했습니다",
           "두 농장 모두 휴약기간 위반이 한 번도 없었습니다. 우리 제안은 이 대조를 자동으로 하자는 것입니다.")
    rows = [["", "흑소 농장 (2023년 논문)", "젖소 농장 (2022년 논문)"],
            ["규모", "흑소 280마리, 가족 4명", "우유 짜는 소 약 110마리, 어린 암소 약 100마리"],
            ["시작한 이유", "수출하는 가공 공장이 항생제·바늘 관리를 요청", "2015년 항생제 섞인 우유를 내보낼 뻔한 사고"],
            ["확인 방법", "출하할 소의 귀표 번호 ↔ 약·주사 바늘 기록(종이)", "빨간 테이프를 감은 소 ↔ 화이트보드의 귀표 번호"],
            ["휴약기간 위반", "0건 (2019~2021년)", "0건 (2018~2020년)"],
            ["소의 건강", "호흡기 병 치료 78마리 → 28~39마리", "체세포 22.8만 → 16.3만(좋아짐), 유방염 월 3.8 → 7.3마리(늘어남)"],
            ["생산", "최고 등급 비율 36.5% → 73.1%", "하루 우유량 33.8kg → 35.1kg"],
            ["한계", "도입 전 기록 1년뿐, 사료 등 동시 개선", "도입 전 기록 없음, 번식 관리 개선 영향이 큼"]]
    table(s, M, Inches(1.85), W - 2 * M, rows, [Inches(1.9), Inches(4.9), Inches(5.33)],
          row_h=Inches(0.47), hot_rows=(4,), size=12)
    box(s, M, Inches(5.75), Inches(5.95), Inches(1.1), fill=BG, lines=[
        ("일본 농장", 13, True, INK),
        ("사람이 종이·화이트보드·빨간 테이프로 매번 대조했습니다.", 12, False, SUB)], anchor=MSO_ANCHOR.MIDDLE)
    box(s, Inches(6.75), Inches(5.75), Inches(5.95), Inches(1.1), fill=GREEN, lines=[
        ("우리 제안", 13, True, WHITE),
        ("약 판매기록과 소 번호를 연결해 컴퓨터가 대조합니다. 나이 든 농장주도 손으로 쓰지 않고 같은 확인을 할 수 있습니다.", 12, False, WHITE)],
        anchor=MSO_ANCHOR.MIDDLE)
    source(s, "[참고 23, 46] Akaike 외(2023, 흑소)·(2022, 젖소) 본문. 체세포 = 우유 속 소의 세포 수로, 낮을수록 젖이 건강하다는 뜻입니다. 유방염 = 소의 젖에 생기는 염증입니다.")


def s_timing(prs):
    s = blank(prs)
    header(s, "EDA 11", "언제 팔지가 농가의 실제 고민입니다",
           "거세우는 최고 등급이 가장 많이 나오는 시기보다 3~4개월 일찍 팝니다. 더 키우면 등급은 오를 수 있지만 사료비가 듭니다.")
    rows = [["구분 (2024년)", "평균적으로 파는 시기", "최고 등급(1++)이 가장 많이 나온 시기"],
            ["거세우", "31.6개월", "35개월 · 판 소의 43.9%가 최고 등급"],
            ["암소", "53.4개월", "33개월 · 판 소의 25.4%가 최고 등급"]]
    table(s, M, Inches(2.0), W - 2 * M, rows, [Inches(2.6), Inches(3.8), Inches(5.73)], row_h=Inches(0.7), size=15, hot_rows=(1,))
    box(s, M, Inches(4.35), Inches(5.95), Inches(2.4), fill=BG, lines=[
        ("일찍 팔면", 15, True, INK), ("", 4, False, INK),
        ("강원대 연구팀은 30개월 대신 24개월에 팔면 사료비가 약 32%, 온실가스가 약 25% 줄어든다고 밝혔습니다.", 13, False, SUB),
        ("2025년에도 거세우를 파는 시기는 0.1개월 늘었습니다.", 13, False, SUB)])
    box(s, Inches(6.75), Inches(4.35), Inches(5.95), Inches(2.4), fill=ACC_SOFT, lines=[
        ("그래서", 15, True, ACC), ("", 4, False, INK),
        ("팔기 전에 '지금 팔면 몇 등급, 한 달 더 키우면 몇 등급'을 알 수 있다면 사료비와 등급 사이에서 더 나은 결정을 할 수 있습니다. "
         "해결 방안 B가 이 부분을 다룹니다.", 13, True, ACC)])
    source(s, "[참고 39, 40, 41] 축산물품질평가원 2024·2025 축산물등급판정 통계연보(한국농어민신문 2025-04-04, 농민신문 2026-04-08 인용), 농림축산식품부 보도자료(2024-04-03)")


# ─────────────────────────────── 문제 인식
def s_problem(prs):
    s = blank(prs)
    header(s, "문제 인식", "문제를 한 문장으로 정리하면 다음과 같습니다")
    box(s, M, Inches(1.5), W - 2 * M, Inches(1.35), fill=ACC_SOFT, lines=[
        ("축산물 안전의 첫 단계인 한우 농장은 HACCP과 스마트HACCP 모두에서 가장 크게 비어 있습니다. "
         "제도는 공장을 기준으로 설계되어 있고, 나이 든 농장주에게 인증은 매일 손으로 기록해야 하는 부담으로 남아 있습니다.", 16, True, ACC)],
        anchor=MSO_ANCHOR.MIDDLE)
    items = [
        ("① 구조", "농가 수가 많고, 규모가 작고, 농장주 나이가 많습니다",
         "한·육우 농장 7만여 곳 중 73%가 50마리 미만이고, 농장주의 63%가 60세 이상입니다. 한 곳씩 설득하는 방식으로는 한계가 있습니다."),
        ("② 설계", "스마트HACCP은 기계형, 농장은 기록형입니다",
         "스마트HACCP 중요관리점의 90%가 기계로 재는 항목입니다. 농장의 휴약기간은 기계로 잴 수 없습니다."),
        ("③ 지원", "농장을 인증으로 이끄는 지원이 줄었습니다",
         "컨설팅 예산이 5년 사이 64% 줄었고, 스마트HACCP 계획의 농장 단계는 두 해 연속 실행되지 않았습니다."),
        ("④ 보상", "인증을 받아도 이득이 잘 보이지 않습니다",
         "현재 혜택은 정기 점검 면제 수준입니다. 등급은 소를 잡은 뒤에야 알 수 있어, 기록이 판매 결정에 도움이 되지 않습니다."),
    ]
    bw = (W - 2 * M - Inches(0.3) * 3) / 4
    for i, (tag, head, body) in enumerate(items):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(3.1), bw, Inches(3.7), fill=WHITE, line=LINE, lines=[
            (tag, 13, True, ACC), (head, 16, True, INK), ("", 6, False, INK), (body, 13, False, SUB)])


# ─────────────────────────────── 해결 방안
def s_principle(prs):
    s = blank(prs)
    header(s, "해결 방안", "해결 방안의 세 가지 원칙",
           "새로운 시스템을 만들기보다, 이미 운영 중인 제도와 데이터를 서로 연결합니다.")
    items = [
        ("1", "이미 있는 기록을 연결합니다",
         "수의사 전자처방, 동물 약 판매기록, 축산물이력제, 축사로가 이미 운영되고 있습니다. 지금은 서로 연결되어 있지 않을 뿐입니다."),
        ("2", "직접 쓰지 않고, 한 번만 선택합니다",
         "농가는 '이 약을 몇 번 소에 쓰셨나요?'라는 질문에 한 번만 답하시면 됩니다. 나머지 기록은 자동으로 채워집니다."),
        ("3", "기록한 만큼 등급 정보로 돌려드립니다",
         "팔기 전에 휴약기간이 끝났는지와 예상 등급을 함께 보여드립니다. 기록이 곧 판매 결정에 도움이 되는 정보가 됩니다."),
    ]
    bw = (W - 2 * M - Inches(0.35) * 2) / 3
    for i, (no, head, body) in enumerate(items):
        x = M + (bw + Inches(0.35)) * i
        box(s, x, Inches(2.2), bw, Inches(4.3), fill=BG, lines=[
            (no, 40, True, GREEN), (head, 19, True, INK), ("", 6, False, INK), (body, 14, False, SUB)])


def s_flow(prs):
    s = blank(prs)
    header(s, "해결 방안 A", "약 관리를 자동으로 기록합니다",
           "동물 약을 파는 곳의 판매기록을 농장의 스마트HACCP 기록으로 바로 연결합니다.")
    steps = [
        ("약 판매처", "항생제 등을 팔 때\n농장번호와 함께\n전자로 등록"),
        ("앱 알림", "\"몇 번 소에\n쓰셨나요?\"\n한 번 선택"),
        ("자동 계산", "약마다 정해진\n휴약기간으로\n팔 수 있는 날 계산"),
        ("판매 제한", "휴약기간이 남은 소는\n출하 버튼이\n눌리지 않음"),
        ("도축 후 확인", "축산물이력제 결과로\nHACCP 기록\n자동 확인"),
    ]
    bw, gap = Inches(2.15), Inches(0.42)
    for i, (head, body) in enumerate(steps):
        x = M + (bw + gap) * i
        box(s, x, Inches(2.15), bw, Inches(2.4), fill=GREEN if i == 1 else BG, lines=[
            (head, 16, True, WHITE if i == 1 else INK), ("", 4, False, INK),
            (body, 13, False, WHITE if i == 1 else SUB)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < len(steps) - 1:
            arrow(s, x + bw + Inches(0.04), Inches(3.17))
    text(s, M, Inches(4.85), Inches(6.0), Inches(2.1), [
        ("이미 마련되어 있는 근거", 13, True, INK),
        ("• 처방이 필요한 약은 2020년 2월 28일부터 전자처방(eVET)으로만 발급됩니다.", 12, False, SUB),
        ("• 항생제·호르몬제·구충제 등은 처방이 없어도 판매기록 작성이 의무입니다.", 12, False, SUB),
        ("• 축산물이력제에는 소마다 도축한 날·도축 검사 결과·등급이 기록됩니다.", 12, False, SUB),
        ("• 일본 소 농장 2곳은 같은 대조를 손으로 해 휴약기간 위반이 0건이었습니다.", 12, False, SUB)])
    text(s, Inches(6.9), Inches(4.85), Inches(5.8), Inches(2.1), [
        ("주의할 점", 13, True, INK),
        ("• 약 판매처는 HACCP 인증 대상이 아닙니다. 판매처의 기존 기록을 농장 HACCP에 연결하는 방식입니다.", 12, False, SUB),
        ("• 약을 산 날과 쓴 날이 다를 수 있어, 어느 소에 썼는지는 한 번 선택하셔야 합니다.", 12, False, SUB),
        ("• 선택하지 않으시면 휴약기간이 끝날 때까지 농장 전체를 '출하 주의'로 표시합니다.", 12, False, SUB)])
    source(s, "[참고 12, 13, 14, 23, 46] 동물용 의약품등 취급규칙(판매기록 작성 대상), 수의사 전자처방전 발급 의무화(2020-02-28 시행), 축산물품질평가원 쇠고기이력정보 공개 API")


def s_grade(prs):
    s = blank(prs)
    header(s, "해결 방안 B", "팔기 전에 예상 등급을 함께 보여드립니다",
           "팔기 전 등급 예측은 아직 연구·대회 단계이고, 약 쓴 기록이나 아팠던 기록을 넣은 곳은 찾지 못했습니다.")
    rows = [["이미 있는 방법", "누가", "언제", "한계"],
            ["AI 등급 판정 장비", "축산물품질평가원", "소를 잡은 뒤", "판정을 돕는 도구, 예측 아님"],
            ["등급 예측 프로그램 (2019)", "축산물품질평가원", "소를 잡은 뒤", "몸통 무게 등 잡은 뒤에 아는 값이 필요"],
            ["날씨·혈통으로 등급 예측", "날씨 빅데이터 대회 수상작 (2026)", "팔기 전", "대회 수상작, 서비스 아님"],
            ["초음파 검사", "한국육류연구소 등", "팔기 전", "고기 양 C등급 시기 83.8% 적중, 2천만 원"],
            ["유전자 검사", "농협", "송아지 때", "타고난 체질만 확인"],
            ["활동량·피검사로 예측", "일본 연구 (검토 전 논문)", "팔기 전", "연구 단계"]]
    table(s, M, Inches(1.9), Inches(7.5), rows,
          [Inches(2.05), Inches(2.2), Inches(1.05), Inches(2.2)], row_h=Inches(0.47), size=11)
    box(s, Inches(8.4), Inches(1.9), Inches(4.3), Inches(2.95), fill=GREEN, lines=[
        ("이번 제안이 더하는 것", 14, True, WHITE), ("", 4, False, WHITE),
        ("태어난 날·성별·아비 소 정보에", 13, False, WHITE),
        ("+ 약 쓴 기록 · 아팠던 기록", 16, True, WHITE), ("", 4, False, WHITE),
        ("→ 팔기 전 화면에서 휴약기간이 끝났는지와 예상 등급을 함께 확인", 13, False, WHITE)])
    box(s, Inches(8.4), Inches(5.0), Inches(4.3), Inches(1.8), fill=BG, lines=[
        ("화면 예시 (설명용 가상 수치)", 12, True, INK),
        ("352번 소 · 휴약기간 끝남", 12, False, SUB),
        ("지금 팔면: 1++ 40% · 1+ 35%", 12, False, SUB),
        ("한 달 더 키우면: 1++ 48% · 사료비 +32만 원", 12, False, SUB)])
    text(s, M, Inches(5.3), Inches(7.5), Inches(1.6), [
        ("등급은 하나로 단정하지 않고 확률로 보여드립니다. 고기 속 지방 무늬는 겉으로 보이지 않아 예측에 오차가 있기 때문입니다.", 12, False, SUB),
        ("새 앱을 만들지 않고, 농가가 이미 등급 결과를 확인하는 '축산물원패스'에 기능을 더하는 방식입니다.", 12, False, SUB)])
    source(s, "[참고 16, 17, 18, 36, 37, 38, 43] 초음파 83.8%는 고기 '양' 등급 기준입니다. 일본 연구는 다른 학자 검토 전 논문으로, 요약만 확인했습니다.")


def s_verify(prs):
    s = blank(prs)
    header(s, "해결 방안 C", "HACCP이 등급에 도움이 되는지 검증합니다",
           "단순 평균 비교에는 원래 규모가 크고 관리를 잘하는 농가의 효과가 섞여 정확하지 않습니다.")
    box(s, M, Inches(2.0), Inches(5.9), Inches(3.0), fill=BG, lines=[
        ("❌ 단순 비교", 16, True, MUTE), ("", 4, False, INK),
        ("인증 농가 평균 등급과 비인증 농가 평균 등급을 비교합니다.", 14, False, SUB), ("", 4, False, INK),
        ("인증 농가가 원래 규모가 크고 관리를 잘해서 등급이 좋은 것인지, HACCP 효과인지 구분할 수 없습니다.", 13, False, SUB)])
    box(s, Inches(6.8), Inches(2.0), Inches(5.9), Inches(3.0), fill=WHITE, line=GREEN, lines=[
        ("✅ 같은 농장의 인증 전후 비교", 16, True, GREEN), ("", 4, False, INK),
        ("같은 농장의 인증 전 등급과 인증 후 등급을 비교합니다.", 14, False, INK), ("", 4, False, INK),
        ("같은 농장이므로 규모·관리 수준의 차이가 사라집니다. HACCP 데이터에 인증일이 있어 전후를 나눌 수 있습니다. "
         "일본 흑소 농장 1곳의 사례(EDA 9)를 전국 규모로 확인하는 방법입니다.", 13, False, SUB)])
    text(s, M, Inches(5.3), W - 2 * M, Inches(1.6), [
        ("필요한 준비", 13, True, INK),
        ("• 축산물이력제의 소별 등급과 농장번호: 현재 공개 API는 한 마리씩만 조회되어, 축산물품질평가원에 데이터 제공을 요청해야 합니다.", 12, False, SUB),
        ("• 농장 이름만으로 연결하면 안 됩니다: 한우 HACCP 농장 이름의 17.6%가 다른 농장과 겹쳐 주소까지 함께 맞춰야 합니다.", 12, False, SUB)])
    source(s, "[참고 1, 14, 23] 인증일은 HACCP 인증업체정보, 소별 등급·농장번호는 축산물품질평가원 쇠고기이력정보 API 명세 기준입니다.")


def s_effect(prs):
    s = blank(prs)
    header(s, "기대 효과", "누구에게 무엇이 달라지는지 정리했습니다")
    items = [
        ("농가", "매일 손으로 쓰던 약 기록이 줄고, 언제 팔지 판단할 정보가 생깁니다. 휴약기간을 놓쳐 고기를 폐기하는 일도 줄어듭니다."),
        ("한국식품안전관리인증원", "농장 심사를 종이 기록부 대신 전산 기록으로 확인할 수 있습니다. 스마트HACCP 등록 농장이 0곳인 문제를 풀 수 있습니다."),
        ("소비자", "농장부터 도축·가공까지 안전 관리가 하나로 이어진 고기를 받게 됩니다."),
    ]
    bw = (W - 2 * M - Inches(0.35) * 2) / 3
    for i, (who, body) in enumerate(items):
        x = M + (bw + Inches(0.35)) * i
        box(s, x, Inches(1.6), bw, Inches(2.45), fill=BG, lines=[(who, 17, True, GREEN), ("", 4, False, INK), (body, 13, False, INK)])
    rows = [["성과 지표", "현재", "목표"],
            ["농장 스마트HACCP 등록", "0곳", "젖소 농장 시범 사업으로 첫 등록"],
            ["농장 스마트HACCP 기술지원 요청", "1건", "672건 (축산물 가공·유통 업소의 요청률 10.7% 수준)"],
            ["휴약기간 자동 확인 비율", "측정되지 않음", "시범 농장의 판매 건 기준으로 새로 측정"],
            ["예상 등급 정확도", "기준 없음", "약 쓴 기록·아팠던 기록을 넣었을 때 얼마나 나아지는지 측정"]]
    table(s, M, Inches(4.3), W - 2 * M, rows, [Inches(4.2), Inches(2.3), Inches(5.63)], row_h=Inches(0.48), size=13)


def s_limit(prs):
    s = blank(prs)
    header(s, "확인 필요", "아직 확인이 필요한 부분입니다")
    items = [
        ("'HACCP을 받으면 등급이 좋아진다'", "일본 흑소 농장 1곳에서 최고 등급 비율이 두 배가 됐지만, 비교 농장이 없고 사료 등도 함께 바꿔 단정하기 어렵습니다. 국내 연구는 못 찾았습니다."),
        ("'약 쓴 기록·아팠던 기록이 등급에 영향을 준다'", "해외 연구에서 몸통 무게와 가치가 준다는 결과는 일관되지만, 고기 품질 등급은 연구마다 엇갈립니다. 한우 연구는 찾지 못했습니다."),
        ("'휴약기간이 등급을 바꾼다'", "관련 연구를 찾지 못했습니다. '약을 쓰게 만든 병이 영향을 준다'고 설명하는 것이 정확합니다."),
        ("사업·예산·통계 일부", "2026년 스마트HACCP 사업 대상은 기사로만 확인했습니다. 컨설팅 예산은 2024년까지, 2025년 거세우 판매 시기는 변화량만 확인했습니다."),
        ("공모전 공식 제출 양식", "아직 확인하지 못했습니다. 양식을 확인한 뒤 목차를 맞춰야 합니다."),
    ]
    y = Inches(1.55)
    for head, body in items:
        box(s, M, y, W - 2 * M, Inches(0.95), fill=BG, lines=[(head, 15, True, INK), (body, 13, False, SUB)],
            anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(1.07)


# ─────────────────────────────── 참고문헌
REFS = [
    ("한국식품안전관리인증원", "HACCP 인증업체정보 (공공데이터포털 15033311)", "2026-09-08 수집", "https://www.data.go.kr/data/15033311/openapi.do"),
    ("한국식품안전관리인증원", "축산물 스마트HACCP 인증업체 정보 서비스 (공공데이터포털 15118113)", "2026-09-14 수집", "https://www.data.go.kr/data/15118113/openapi.do"),
    ("한국식품안전관리인증원", "스마트HACCP 기술지원 업체 정보 (공공데이터포털 15131084)", "2026-09-14 수집", "https://www.data.go.kr/data/15131084/openapi.do"),
    ("국가데이터처", "2026년 2분기 가축동향조사 결과 (보도자료, 2026-06-01 기준)", "2026-07-24", "https://www.korea.kr/briefing/pressReleaseView.do?newsId=156772070"),
    ("농림축산식품부", "2024년 축산물 HACCP 컨설팅 지원사업 시행지침 (한국한우협회 게시본)", "2024", "https://www.ihanwoo.org/19324"),
    ("농림축산식품부", "2024년 생산단계 축산물 스마트 HACCP 현장구축사업 추진계획 (공고 제2024-321호)", "2024-07-26", "https://www.mafra.go.kr/bbs/home/791/570945/artclView.do"),
    ("농림축산식품부", "2025년 생산단계 축산물 스마트 HACCP 현장구축사업 추진계획 (공고 제2024-509호)", "2024-12-02", "https://www.mafra.go.kr/bbs/home/791/572428/artclView.do"),
    ("대한급식신문", "축산물 안전 사각지대 해소된다", "2026-04-02", "https://www.fsnews.co.kr/news/articleView.html?idxno=60482"),
    ("뉴스프리존", "해썹인증원, 축산물 생산단계 안전관리 'AI 혁신' 이끈다", "2026-04-13", "https://www.newsfreezone.co.kr/news/articleView.html?idxno=684800"),
    ("푸드투데이", "도축장·집유장도 '스마트 해썹' 도입...축산물 안전관리 디지털화", "2026-03-31", "https://www.foodtoday.or.kr/news/article.html?no=203585"),
    ("한국농어민신문", "청년농 육성한다지만…높은 진입장벽 탓 후계인력 확보 '막막' (통계청 2020 농림어업총조사 인용)", "2022-04-05", "http://www.agrinet.co.kr/news/articleView.html?idxno=308392"),
    ("국가법령정보센터", "동물용 의약품등 취급규칙 (판매기록 작성 대상)", "현행", "https://www.law.go.kr/lsInfoP.do?lsiSeq=238203"),
    ("이엠디", "수의사 전자처방전 발급 의무화 2020년 2월 28일부터 시행 (농림축산식품부 발표)", "2020-02-28 시행", "https://www.mdon.co.kr/news/article.html?no=25680"),
    ("축산물품질평가원", "쇠고기이력정보 (공공데이터포털 15056898)", "명세 확인 2026-09-17", "https://www.data.go.kr/data/15056898/openapi.do"),
    ("축산신문", "기상·축산 데이터 결합해 한우 생산성 높인다", "2026-08-12", "https://www.chuksannews.co.kr/news/article.html?no=273018"),
    ("축산신문", "축평원, AI·빅데이터 활용 한우·돼지 농가 맞춤형 분석 강화", "2026-09-02", "https://www.chuksannews.co.kr/news/article.html?no=273356"),
    ("농기자재신문", "거세한우의 지방 축적 과다 방지를 위한 초음파 진단시스템 개발", "2021-12-23", "http://newsam.co.kr/mobile/article.html?no=33861"),
    ("축산경제신문", "한우 암소 유전체 정보 이용하면 농가 맞춤형 개량 가능", "날짜 미확인", "http://www.chukkyung.co.kr/news/articleView.html?idxno=59178"),
    ("축산경제신문", "창간 36주년 특집: AI가 바꾸는 축산", "2026-09-04", "https://www.chukkyung.co.kr/news/articleView.html?idxno=82768"),
    ("국립축산과학원", "한우 농장 정보(데이터) 관리, 활용은 '축사로'에서 (보도자료)", "2022-03-30", "https://www.nias.go.kr/front/soboarddown.do?cmCode=M090814150850297&boardSeqNum=3863&fileSeqNum=3145"),
    ("한국농어민신문", "생산단계 HACCP 우수현장을 가다: 거창축협 생축사업장", "2020-11-20", "http://www.agrinet.co.kr/news/articleView.html?idxno=180983"),
    ("식품의약품안전처", "식품 및 축산물 안전관리인증기준 일부개정고시 (제2025-8호)", "2025-02-18", "https://www.mfds.go.kr/brd/m_207/view.do?seq=15051"),
    ("Akaike H, Nagai M, Okatani AT, Morita Y", "Food safety, livestock health, and productivity of a farm fatting Japanese black cattle following implementation of a certificated HACCP system. J Vet Med Sci", "2023", "https://doi.org/10.1292/jvms.23-0034"),
    ("Schneider MJ, Tait RG, Busby WD, Reecy JM", "An evaluation of bovine respiratory disease complex in feedlot cattle. J Anim Sci 87", "2009", "https://doi.org/10.2527/jas.2008-1283"),
    ("Wilson BK, Step DL, Maxwell CL 외", "Effect of bovine respiratory disease during the receiving period on steer finishing performance, efficiency, carcass characteristics, and lung scores. Prof Anim Sci", "2017", "https://pmc.ncbi.nlm.nih.gov/articles/PMC7147665/"),
    ("Blakebrough-Hall C, McMeniman JP, González LA", "An evaluation of the economic effects of bovine respiratory disease on animal performance, carcass traits, and economic outcomes in feedlot cattle. J Anim Sci 98", "2020", "https://doi.org/10.1093/jas/skaa005"),
    ("Grimes BB, McEvers TJ, Tennant TC, Lawrence TE", "Association of bovine lung lesions at slaughter with carcass performance, value and liver-abscess severity. Transl Anim Sci", "2025", "https://pmc.ncbi.nlm.nih.gov/articles/PMC12405687/"),
    ("Inoue K, Honda T, Oyama K", "Genetic relationships between internal diseases diagnosed at slaughter and carcass traits in Japanese Black cattle. J Anim Sci 93", "2015", "https://doi.org/10.2527/jas.2014-8497"),
    ("Przyborowska-Zhalniarovich P 외", "Association between Enzootic Pneumonia-Like Lung Lesions and Carcass Quality and Meat pH Value in Slaughter Pigs. Animals", "2023", "https://pmc.ncbi.nlm.nih.gov/articles/PMC10340002/"),
    ("식품의약품안전처 식품안전나라", "잔류물질정보: 식품에 남아있는 동물용 의약품 우리몸에 위험할까요?", "현행", "http://www.foodsafetykorea.go.kr/residue/contents/view.do?contentsKey=6"),
    ("KDI 경제교육·정보센터", "축·수산물 동물용의약품 PLS 제도 2024년 1월 1일 본격 시행", "2023", "https://eiec.kdi.re.kr/policy/materialView.do?num=246571"),
    ("농림축산식품부", "동물용의약(외)품 안전사용 안내 (휴약기간·과태료)", "현행", "https://www.mafra.go.kr/bbs/PLS/830/571195/download.do"),
    ("찾기쉬운 생활법령정보", "축산물 표시기준 등 (등급 표시)", "현행", "https://easylaw.go.kr/CSP/CnpClsMainBtr.laf?popMenu=ov&csmSeq=1011&ccfNo=4&cciNo=3&cnpClsNo=2"),
    ("경향신문", "쇠고기 '마블링' 적어도 1++ 등급 판정...12월부터 등급제 개편", "2019-11-25", "https://www.khan.co.kr/article/201911251104001"),
    ("한국축산물처리협회", "등급 기준 (육량지수 계산식)", "현행", "https://cleanmeat.or.kr/data/03.php"),
    ("축산신문", "한우 온라인 구매도 '마블링' 보고 고른다…AI 등급판정 이미지 민간 개방", "2026", "https://www.chuksannews.co.kr/news/article.html?no=273287"),
    ("부산일보", "한우 도축전 AI로 도체등급 예측…날씨 빅데이터 콘테스트 최우수상", "2026-08-06", "https://www.busan.com/view/busan/view.php?code=2026080617425417698"),
    ("이데일리", "축평원, 소고기 등급 예측 가능한 진단 프로그램 개발", "2019-11-04", "https://edaily.co.kr/News/Read?mediaCodeNo=257&newsId=03955686622682440"),
    ("한국농어민신문", "한우 1++ 등급 비중이 가장 높은 사육월령은? (2024 축산물등급판정 통계연보 인용)", "2025-04-04", "https://www.agrinet.co.kr/news/articleView.html?idxno=335756"),
    ("농민신문", "지난해 한우고기 78% 1등급 이상 받았다 (2025 축산물등급판정 통계연보 인용)", "2026-04-08", "https://www.nongmin.com/article/20260406500780"),
    ("농림축산식품부", "한우 사육기간 단축으로 두 마리 토끼 잡는다 (보도자료)", "2024-04-03", "https://www.mafra.go.kr/bbs/home/792/569966/artclView.do"),
    ("조현진, 정연복, 김병도, 정사무엘, 서성원", "개정된 소 도체 등급판정 세부기준에서 도체성적과 도축월령이 등급판정에 미치는 영향력 평가. 농업생명과학연구", "2020", "https://www.kci.go.kr/kciportal/ci/sereArticleSearch/ciSereArtiView.kci?sereArticleSearchBean.artiId=ART002578586"),
    ("Research Square (프리프린트)", "A machine learning approach to predict meat production factors in Japanese black cattle fattening", "검토 전", "https://www.researchsquare.com/article/rs-6653315/v1"),
    ("찾기쉬운 생활법령정보", "가공식품(농축수산물) > 인증제도 > 축산물 안전관리인증기준(HACCP)", "현행", "https://easylaw.go.kr/CSP/CnpClsMain.laf?popMenu=ov&csmSeq=1007&ccfNo=2&cciNo=3&cnpClsNo=2"),
    ("국가법령정보센터", "축산물 위생관리법 제9조(안전관리인증기준 등)", "현행", "https://www.law.go.kr/lsInfoP.do?lsId=001507&ancYnChk=0"),
    ("Akaike H, Nagai M, Okatani AT, Morita Y", "Food safety, livestock health, and productivity of a dairy farm following implementation of a certificated Hazard Analysis and Critical Control Points system. J Vet Med Sci 84(7): 924-928", "2022", "https://doi.org/10.1292/jvms.21-0563"),
]


def s_refs(prs):
    per = 12
    pages = [REFS[i:i + per] for i in range(0, len(REFS), per)]
    for pi, chunk in enumerate(pages):
        s = blank(prs)
        header(s, "참고문헌", "자료 출처 ({}/{})".format(pi + 1, len(pages)),
               "각 슬라이드 아래의 [참고 번호]가 이 목록의 번호와 같습니다.")
        tb = s.shapes.add_textbox(M, Inches(1.85), W - 2 * M, Inches(5.4))
        tf = tb.text_frame
        tf.word_wrap = True
        for k, (org, title, date, url) in enumerate(chunk):
            no = pi * per + k + 1
            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            p.space_after = Pt(1)
            r = p.add_run()
            r.text = "[{}] {}, 「{}」, {}".format(no, org, title, date)
            r.font.name = FONT
            r.font.size = Pt(10)
            r.font.bold = True
            r.font.color.rgb = INK
            p2 = tf.add_paragraph()
            p2.space_after = Pt(3)
            r2 = p2.add_run()
            r2.text = "      " + url
            r2.font.name = FONT
            r2.font.size = Pt(8.5)
            r2.font.color.rgb = MUTE
            r2.hyperlink.address = url


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    cover(prs)
    s_terms(prs); s_withdrawal(prs); s_grading(prs)
    divider(prs, "01", "EDA", "데이터로 살펴본 한우 농장 HACCP 현황")
    s_data(prs); s_rate(prs); s_why(prs); s_same(prs); s_trend(prs); s_smart(prs); s_ccp(prs); s_plan(prs); s_people(prs)
    s_research(prs); s_japan_beef(prs); s_japan_compare(prs); s_timing(prs)
    divider(prs, "02", "문제 인식", "왜 한우 농장만 비어 있을까요")
    s_problem(prs)
    divider(prs, "03", "해결 방안", "기록은 자동으로, 결과는 등급으로")
    s_principle(prs); s_flow(prs); s_grade(prs); s_verify(prs); s_effect(prs); s_limit(prs)
    s_refs(prs)
    # python-pptx 기본 서식에 박힌 문서 정보(2013-01-27, 'Steve Canny')를 지금 값으로 바꾼다
    import datetime
    _now = datetime.datetime.now()
    _cp = prs.core_properties
    _cp.title = "한우 HACCP 발표 초안 v4"
    _cp.subject = "2026 한국식품안전관리인증원 공공데이터 분석·활용 경진대회"
    _cp.author = ""
    _cp.last_modified_by = ""
    _cp.created = _now
    _cp.modified = _now
    _cp.revision = 1
    p = OUT + "/한우HACCP_발표초안_v4.pptx"
    try:
        prs.save(p)
    except PermissionError:
        p = OUT + "/한우HACCP_발표초안_v5.pptx"
        prs.save(p)
    print("[저장]", p, len(prs.slides), "장")


if __name__ == "__main__":
    main()
