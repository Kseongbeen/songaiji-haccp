# -*- coding: utf-8 -*-
"""
41. 한우 출하노트(가칭) — 휴대폰 앱 프로토타입 구상 (PPT)

v6 발표 초안에서 '구상'만 따로 뺀 자료. 내용 기준은 output/_앱_프로토타입_구상.md.
디자인 도구는 40_한우_PPT_v6.py에서 가져온다. 화면은 휴대폰 모양으로 그린다.

표기: [확인됨] v6 검증 사실 / [가정] 설계 가정 / [가상] 시연용 숫자·인물

산출: output/한우출하노트_앱구상_v1.pptx
"""

import datetime
import importlib.util
import os

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("ppt40", os.path.join(HERE, "40_한우_PPT_v6.py"))
v6 = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(v6)

INK, SUB, MUTE, LINE, BG, WHITE = v6.INK, v6.SUB, v6.MUTE, v6.LINE, v6.BG, v6.WHITE
ACC, ACC_SOFT, GREEN, GREEN_SOFT, GRAY = v6.ACC, v6.ACC_SOFT, v6.GREEN, v6.GREEN_SOFT, v6.GRAY
AMBER = RGBColor(0xB7, 0x6E, 0x00)
AMBER_SOFT = RGBColor(0xFB, 0xEE, 0xD5)
W, H, M = v6.W, v6.H, v6.M
blank, box, pill, chevron, header, table, text, notes = (
    v6.blank, v6.box, v6.pill, v6.chevron, v6.header, v6.table, v6.text, v6.notes)
OUT = os.path.join(HERE, "output")


def foot(slide, t):
    text(slide, M, Inches(7.02), W - 2 * M, Inches(0.3), [(t, 9, False, MUTE)])


# ─────────────────────────────────────────────────────────────
# 휴대폰 화면 그리기
# items: (종류, 내용)
#   "bar"   : 화면 제목 줄
#   "big"   : 큰 글씨 한 줄 (글, 색)
#   "t"     : 보통 글 (글, 색)
#   "s"     : 작은 회색 글
#   "card"  : 카드 [(글, 크기, 굵게, 색), ...], 배경색
#   "chips" : 알약 여러 개 [(글, 배경, 글색), ...]
#   "btn"   : 주 버튼 글
#   "btn2"  : 보조 버튼 글
#   "row"   : 목록 한 줄 (왼쪽 글, 오른쪽 상태, 상태색, 상태배경)
#   "gap"   : 빈칸 (인치)
# ─────────────────────────────────────────────────────────────
def _tw(t, wide, narrow):
    return sum(wide if ord(ch) > 0x2E80 else narrow for ch in t)


def phone(s, x, y, items, w=Inches(2.95), h=Inches(5.15), label=None):
    frame = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    frame.adjustments[0] = 0.12
    frame.fill.solid(); frame.fill.fore_color.rgb = INK
    frame.line.fill.background(); frame.shadow.inherit = False
    pad = Inches(0.09)
    sx, sy, sw, sh = x + pad, y + pad, w - 2 * pad, h - 2 * pad
    scr = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, sw, sh)
    scr.adjustments[0] = 0.1
    scr.fill.solid(); scr.fill.fore_color.rgb = WHITE
    scr.line.fill.background(); scr.shadow.inherit = False
    ix, iw = sx + Inches(0.14), sw - Inches(0.28)
    cy = sy + Inches(0.22)
    for kind, val in items:
        if kind == "bar":
            text(s, ix, cy, iw, Inches(0.3), [(val, 12, True, INK)])
            ln = s.shapes.add_connector(1, ix, cy + Inches(0.36), ix + iw, cy + Inches(0.36))
            ln.line.color.rgb = LINE
            cy += Inches(0.45)
        elif kind == "big":
            t, c = val
            text(s, ix, cy, iw, Inches(0.36), [(t, 15, True, c)])
            cy += Inches(0.38)
        elif kind == "t":
            t, c = val
            n = max(1, -(-len(t) // 17))
            text(s, ix, cy, iw, Inches(0.22) * n, [(t, 10, False, c)])
            cy += Inches(0.22) * n + Inches(0.04)
        elif kind == "s":
            n = max(1, -(-len(val) // 22))
            text(s, ix, cy, iw, Inches(0.19) * n, [(val, 8.5, False, MUTE)])
            cy += Inches(0.19) * n + Inches(0.03)
        elif kind == "card":
            lines, fill = val
            hh = Inches(0.3) + sum(Inches(0.25) if sz < 12 else Inches(0.31) for _, sz, _, _ in lines)
            box(s, ix, cy, iw, hh, fill=fill, lines=lines, r=Inches(0.1), pad=0.1)
            cy += hh + Inches(0.08)
        elif kind == "chips":
            cx = ix
            for t, f, c in val:
                cw = Inches(0.34) + _tw(t, Inches(0.14), Inches(0.075))
                if cx + cw > ix + iw:
                    cx = ix; cy += Inches(0.34)
                pill(s, cx, cy, cw, Inches(0.28), t, f, c, size=9, line=LINE if f == WHITE else None)
                cx += cw + Inches(0.06)
            cy += Inches(0.36)
        elif kind == "row":
            left, st, c, f = val
            text(s, ix, cy + Inches(0.02), Inches(1.6), Inches(0.26), [(left, 10, True, INK)])
            pw = Inches(0.32) + _tw(st, Inches(0.13), Inches(0.07))
            pill(s, ix + iw - pw, cy, pw, Inches(0.26), st, f, c, size=8.5)
            cy += Inches(0.32)
        elif kind == "btn":
            pill(s, ix, cy, iw, Inches(0.36), val, GREEN, WHITE, size=11)
            cy += Inches(0.43)
        elif kind == "btn2":
            pill(s, ix, cy, iw, Inches(0.34), val, WHITE, SUB, size=10, line=LINE)
            cy += Inches(0.41)
        elif kind == "gap":
            cy += Inches(val)
    if label:
        pill(s, x + (w - Inches(1.9)) / 2, y + h + Inches(0.08), Inches(1.9), Inches(0.3), label, BG, SUB, size=10)


# ─────────────────────────────────────────────────────────────
# 슬라이드
# ─────────────────────────────────────────────────────────────
def cover(prs):
    s = blank(prs)
    box(s, Inches(0.6), Inches(0.6), W - Inches(1.2), H - Inches(1.2), fill=BG, r=Inches(0.4))
    pill(s, Inches(1.2), Inches(1.5), Inches(4.6), Inches(0.46), "휴대폰 앱 프로토타입 구상 v1", WHITE, SUB, size=13)
    text(s, Inches(1.2), Inches(2.3), Inches(7.2), Inches(2.4),
         [("한우 출하노트", 46, True, INK), ("(가칭)", 18, False, MUTE)])
    text(s, Inches(1.2), Inches(3.9), Inches(7.0), Inches(1.4),
         [("HACCP으로 위생과 안전을 지키고,", 22, True, SUB), ("등급 예측으로 돌려줍니다", 22, True, ACC)])
    pill(s, Inches(1.2), Inches(5.6), Inches(6.4), Inches(0.42),
         "클릭 프로토타입 제작 완료(Figma) · 기준 문서: _앱_프로토타입_구상.md", WHITE, MUTE, size=11)
    phone(s, Inches(9.1), Inches(0.95), [
        ("bar", "9월 22일 (화)  박영수 님"),
        ("card", ([("🔔 확인할 약 기록 2건", 13, True, ACC)], ACC_SOFT)),
        ("row", ("휴약 중", "5마리", AMBER, AMBER_SOFT)),
        ("row", ("출하 가능", "57마리", GREEN, GREEN_SOFT)),
        ("row", ("출하 주의", "2마리", ACC, ACC_SOFT)),
        ("row", ("2주 안 출하 예정", "4마리", SUB, BG)),
        ("gap", 0.05),
        ("btn", "📝 오늘 한 일 기록"), ("btn2", "🐄 우리 소"), ("btn2", "📊 9월 농장 성적표"),
    ], label="[가상] 홈 화면")
    notes(s, "이 자료는 발표 초안 v6에서 '무엇을 만들지'만 따로 뺀 구상이고, 지금은 실제로 만든 클릭 프로토타입 내용까지 담았습니다.")


def s_summary(prs):
    s = blank(prs)
    header(s, "한 장 요약", "기획서 2-1 아이디어 핵심 개요와 같은 내용")
    rows = [["항목", "내용"],
            ["핵심 개념", "한우 농장의 약품·일상 관리 기록을 간소화하고, 축적된 기록으로 휴약 확인과 출하 판단을 지원하는 앱. "
                        "안전 관리에 필요한 기록을 농가의 경영 정보로 돌려주어 지속적인 기록 참여를 유도"],
            ["주 사용자", "수기 기록에 익숙한 고령 한우 농장주. 초기에는 50~100마리 규모 농장을 대상으로 문자 입력보다 선택·확인 중심으로 설계"],
            ["이용 흐름", "처방·판매 기록 불러오기 → 실제 투약 개체·시점 확인 → 휴약 계산·출하 전 점검 → HACCP 장부 작성"],
            ["설계 원칙", "구매·처방 기록은 초안으로 수집하고 농가 확인 후 사용 기록으로 확정. 불완전한 기록은 '확인 필요'로 구분"],
            ["역할 구분", "앱은 기록 작성·대조·장부 정리를 대신하고, 위생·사양 관리의 실행, 투약 개체 확정, 출하 결정, 인증 심사 대응은 농장주가 수행"],
            ["차별성", "개별 기능의 신규성이 아니라 외부 기록 수집 → 농가 확인 → 안전 관리 → 장부 작성·출하 판단을 연결하는 범위"]]
    table(s, M + Inches(0.08), Inches(1.65), W - 2 * M - Inches(0.16), rows,
          [Inches(1.5), Inches(10.47)], row_h=Inches(0.78), size=14)
    notes(s, "【설명】 기획서 2-1을 그대로 옮긴 장입니다. 앱이 하는 일과 농장주가 하는 일을 갈라 둔 것이 핵심입니다.")


def s_why(prs):
    s = blank(prs)
    header(s, "왜 이 앱인가", "기록 부담은 과중하고, 보상은 없습니다",
           "기획서 1-1 현황 및 문제점. 숫자는 모두 공공데이터·공식 통계에서 확인한 값입니다.")
    stats = [("2.5%", "한우 농장 HACCP 인증률", "72,389곳 중 1,835곳 · 축종 중 최저"),
             ("43.30%", "HACCP 적용 애로 1순위", "'기록 작성' (남인식 외 2014, 산란계 15개소)"),
             ("63%", "축산 경영주 60세 이상 비율", "한·육우 농장 72.6%는 50마리 미만")]
    bw = (W - 2 * M - Inches(0.3) * 2) / 3
    for i, (big, head, small) in enumerate(stats):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(2.0), bw, Inches(1.8), fill=BG, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, lines=[
            (big, 38, True, ACC), (head, 14, True, INK), (small, 11, False, MUTE)])
    causes = [("원인 1 · 구조", "인증이 소 값에 미반영",
               "농장 HACCP은 자율 인증(축산물 위생관리법 제9조). 등급판정 세부기준에 인증 항목이 없어 수취 가격 변동 없음"),
              ("원인 2 · 설계", "HACCP 전산화가 농장 항목과 불일치",
               "등록 업체 중요관리점 668개 중 603개(90.3%)가 센서 자동 측정 항목. 농장 핵심 관리점은 투약 기록과 개체 번호 대조"),
              ("원인 3 · 지원", "방문 컨설팅의 도달 범위 한계",
               "컨설팅 예산 2019년 28억 → 2024년 10억 1,500만 원(64% 감소). 연 169~330곳, 2만여 곳에 닿으려면 약 63년")]
    for i, (tag, head, body) in enumerate(causes):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(4.05), bw, Inches(1.95), fill=WHITE, line=LINE, pad=0.25, lines=[
            (tag, 12, True, ACC), (head, 14, True, INK), ("", 3, False, INK), (body, 12, False, SUB)])
    box(s, M, Inches(6.2), W - 2 * M, Inches(0.8), fill=GREEN_SOFT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, lines=[
        ("인증률 2.5%는 의지 문제가 아니라 비용·편익 계산의 결과 → 부담을 줄이고 보상을 만들어야 함", 16, True, GREEN)])


def s_existing(prs):
    s = blank(prs)
    header(s, "연결할 것", "새로 만들지 않고 이미 있는 기록을 잇습니다", "기획서 1-2 [표 3] 기존 보유 데이터와 공백")
    rows = [["데이터", "담긴 정보", "근거", "공백"],
            ["수의사 전자처방", "농장명, 동물 종류, 약 성분·용량, 처방일수, 판매 수량", "수의사법 제12조의2, 시행규칙 제11조",
             "소 이력번호 기재 의무 없음(동일 축사는 처방전 1건)"],
            ["약 판매기록", "판매일, 제품명, 수량, 용도, 구매자", "동물용 의약품등 취급규칙 제22조",
             "서식·전산 제출 규정 없음, 농장번호 없음. 약국은 일부 처방대상 약을 처방전 없이 판매 가능"],
            ["축산물이력제", "개체별 출생·이동·도축·등급, 출하 신고", "가축 및 축산물 이력관리에 관한 법률", "약 사용 정보 없음"],
            ["축사로(국립축산과학원)", "한우 농장 HACCP 기록, 약품 구매·사용·처방 기록 칸, 이력제 자동 연동", "2023년 기능 추가",
             "약 기록은 농장주가 수기 입력"]]
    table(s, M + Inches(0.08), Inches(2.0), W - 2 * M - Inches(0.16), rows,
          [Inches(1.9), Inches(3.6), Inches(2.7), Inches(3.77)], row_h=Inches(0.9), size=12)
    box(s, M, Inches(5.9), W - 2 * M, Inches(1.0), fill=GREEN_SOFT, pad=0.3, anchor=MSO_ANCHOR.MIDDLE, lines=[
        ("연계 효과", 14, True, GREEN),
        ("네 가지를 연계하면 '약 기록과 소 번호 대조'를 전산이 대체. 센서가 아닌 기록 연결 방식의 농장 HACCP 전산화", 15, False, INK)])


def s_user(prs):
    s = blank(prs)
    header(s, "사용자와 역할", "앱이 할 일과 농장주가 할 일을 나눴습니다")
    box(s, M, Inches(1.9), Inches(6.1), Inches(2.3), fill=BG, pad=0.32, lines=[
        ("주 사용자", 16, True, INK),
        ("수기 기록에 익숙한 고령 한우 농장주. 축산 경영주의 63%가 60세 이상", 14, False, INK), ("", 4, False, INK),
        ("첫 확산 대상", 16, True, INK),
        ("50~100마리 규모 한·육우 농장 12,461곳. 기록 건수는 수기 한계를 넘지만 전담 인력은 없는 규모", 14, False, INK)])
    box(s, M, Inches(4.35), Inches(6.1), Inches(2.5), fill=GREEN_SOFT, pad=0.32, lines=[
        ("설계 원칙", 16, True, GREEN), ("", 3, False, INK),
        ("• 구매·처방 기록은 초안으로 수집, 농가 확인 후 사용 기록으로 확정", 13, False, INK),
        ("• 불완전한 기록은 '확인 필요'로 구분", 13, False, INK),
        ("• 일상 기록은 버튼 선택 + 필요한 세부 정보 입력", 13, False, INK),
        ("• 문자 입력보다 선택·확인 중심", 13, False, INK)])
    rows = [["구분", "하는 일"],
            ["앱", "기록 작성·대조·장부 정리"],
            ["앱", "처방·판매 기록 불러오기, 휴약 종료 시각 계산, 출하 예정 개체 대조"],
            ["앱", "HACCP 장부 출력, 성적표·출하 판단 정보 제공"],
            ["농장주", "위생·사양 관리의 실행"],
            ["농장주", "투약 개체 확정"],
            ["농장주", "출하 결정"],
            ["농장주", "인증 심사 대응"]]
    table(s, Inches(7.05), Inches(1.95), Inches(5.6), rows, [Inches(1.2), Inches(4.4)], row_h=Inches(0.58), size=13,
          hot_rows=(4, 5, 6, 7))
    foot(s, "기획서 2-1 '역할 구분'과 같습니다. 앱이 관리를 대신한다고 말하지 않는 것이 중요합니다.")


def s_flow(prs):
    s = blank(prs)
    header(s, "이용 흐름", "기록이 쌓여 출하 판단으로 돌아옵니다",
           "농장주가 직접 하는 일은 '투약 개체 확인'과 '출하 결정' 두 가지입니다.")
    steps = [("① 기록 수집", "처방·판매 기록\n불러오기", "A", False),
             ("② 개체 확인", "실제 투약 개체와\n시점 확인", "A", True),
             ("③ 휴약 계산", "종료 시각 자동 계산\n(규칙 기반)", "A", False),
             ("④ 출하 전 점검", "출하 예정 개체와\n대조·경고", "A", True),
             ("⑤ 장부 작성", "HACCP 장부\n자동 정리", "A", False),
             ("⑥ 성적표", "실제 등급·건강\n지표 환류", "B", False),
             ("⑦ 출하 판단", "예상 등급·사료비\n시나리오 비교", "B", False)]
    bw, gap = Inches(1.6), Inches(0.13)
    for i, (head, body, axis, user) in enumerate(steps):
        x = M + (bw + gap) * i
        fill = GREEN if user else (BG if axis == "A" else ACC_SOFT)
        box(s, x, Inches(2.3), bw, Inches(2.3), fill=fill, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, pad=0.08, lines=[
            (head, 14, True, WHITE if user else INK), ("", 3, False, INK), (body, 10, False, WHITE if user else SUB)])
    pill(s, M, Inches(4.8), (bw + gap) * 5 - gap, Inches(0.46), "A. 기록 부담 제거와 안전 확인 (①~⑤)", BG, INK, size=13)
    pill(s, M + (bw + gap) * 5, Inches(4.8), (bw + gap) * 2 - gap, Inches(0.46), "B. 기록의 수익 환류 (⑥~⑦)", ACC_SOFT, ACC, size=13)
    box(s, M, Inches(5.55), W - 2 * M, Inches(1.3), fill=WHITE, line=LINE, pad=0.3, anchor=MSO_ANCHOR.MIDDLE, lines=[
        ("초록 칸 = 농장주가 판단하는 곳", 14, True, GREEN),
        ("과거 도축 결과(이력제 등급·도체중)는 출하 판단 모델의 정답 자료로 활용. 초기에는 혈통·거세·월령 기반 기본 모델로 시작하고, "
         "HACCP 기록의 추가 효과는 앱 기록이 있는 개체가 출하된 뒤 검증", 12, False, SUB),
        ("모든 연계는 농가 동의 범위에서만 이루어지며 항목별 철회를 지원", 12, False, SUB)])


def s_features(prs):
    s = blank(prs)
    header(s, "핵심 기능", "기능 4개와 농가가 얻는 가치", "기획서 2-1 [표 6]과 같은 구성입니다.")
    rows = [["핵심 기능", "주요 내용", "농가가 얻는 가치"],
            ["약 기록 수집·확정", "처방·판매 기록을 불러오고 실제 투약 개체·시점을 확인. 연계되지 않은 약은 포장 인식·목록 선택으로 입력",
             "재입력 부담 완화, 실제 투약 기록 확보"],
            ["휴약 확인·출하 점검", "확정 기록으로 휴약 종료 시각을 계산하고 출하 예정 개체와 대조. 미경과·미확정 상태를 알리고 조치 기록 저장",
             "출하 전 확인 누락 방지"],
            ["일상 기록·HACCP 장부", "사양·질병·위생 기록을 항목별로 입력하고 약품·출하 점검 기록과 함께 장부로 출력. 축사로 전송은 기관 협의 후 추진",
             "중복 기록 감소, 심사 준비 지원"],
            ["성적표·출하 판단", "초기에는 실제 등급·도체중·건강 지표를 제공. 데이터 축적 후 예상 등급 확률·추가 사육비·시나리오별 손익 비교로 확장",
             "기록을 경영 정보로 활용"]]
    table(s, M + Inches(0.08), Inches(2.0), W - 2 * M - Inches(0.16), rows,
          [Inches(2.2), Inches(6.7), Inches(3.07)], row_h=Inches(1.0), size=13, hot_rows=(1, 4))
    box(s, M, Inches(6.3), W - 2 * M, Inches(0.7), fill=BG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, lines=[
        ("기능을 늘리지 않고, 기록 부담을 줄이는 것(A)과 기록을 수익 정보로 돌려주는 것(B)에만 집중", 14, True, INK)])


def s_principles(prs):
    s = blank(prs)
    header(s, "화면 원칙", "60대 이상이 도움 없이 쓸 수 있게")
    rules = [("크게", "글자 18pt 이상, 핵심 숫자 28pt 이상"),
             ("적게", "한 화면에 버튼 3개 이하"),
             ("글자와 색 함께", "초록 '출하 가능' · 주황 '휴약 중 D-7' · 빨강 '출하 주의'"),
             ("귀표 뒷 4자리", "0352처럼 크게, 12자리 이력번호는 작게"),
             ("풀어 쓰기", "휴약기간(약 성분이 빠질 때까지 출하하지 않고 기다리는 기간)"),
             ("이유 보여 주기", "판단 화면마다 '왜 이렇게 나왔나요?'"),
             ("되돌리기", "잘못 고른 소는 7일 안에 고침 [가정]"),
             ("앱이 어려우면 문자", "귀표 뒷번호를 문자로 답장해도 같은 처리")]
    cw = (W - 2 * M - Inches(0.25) * 3) / 4
    for i, (head, body) in enumerate(rules):
        x = M + (cw + Inches(0.25)) * (i % 4)
        y = Inches(1.7) + Inches(2.05) * (i // 4)
        box(s, x, y, cw, Inches(1.85), fill=BG if i % 2 == 0 else GREEN_SOFT, pad=0.25, lines=[
            (head, 16, True, GREEN if i % 2 else INK), ("", 4, False, INK), (body, 13, False, SUB)])
    rows = [["화면", "03 동의", "04 홈", "05 약 기록", "18~21 목록", "02 소 상세", "10 출하 판단", "07 출하 점검", "08 오늘 한 일", "09 성적표", "17 장부"]]
    rows.append(["만든 상태", "완료", "완료", "완료", "완료", "완료", "완료", "완료", "완료", "완료", "완료"])
    table(s, M + Inches(0.08), Inches(5.95), W - 2 * M - Inches(0.16), rows,
          [Inches(0.8)] + [Inches(1.117)] * 10, row_h=Inches(0.42), size=10)


def s_pick(prs):
    s = blank(prs)
    header(s, "핵심 화면 · 05 → 06", "약 기록 확인: 쓴 소 고르기 (터치 세 번)")
    phone(s, M + Inches(0.2), Inches(1.45), [
        ("bar", "‹ 이전                    ⌂ 홈"),
        ("card", ([("9월 20일 구매 · 한빛동물약품", 10, False, SUB), ("옥시테트라사이클린 주사 1병", 12, True, INK),
                   ("처방: 새들동물병원 · 휴약 28일", 10, False, SUB)], BG)),
        ("t", ("어느 소에 쓰셨나요?", INK)),
        ("s", "3번 칸"),
        ("chips", [("0352 ✓", GREEN, WHITE), ("0353", WHITE, INK), ("0360", WHITE, INK)]),
        ("s", "4번 칸"),
        ("chips", [("0411", WHITE, INK), ("0412", WHITE, INK), ("0415", WHITE, INK)]),
        ("t", ("언제 쓰셨나요?", INK)),
        ("chips", [("오늘", WHITE, INK), ("어제 ✓", GREEN, WHITE), ("날짜 고르기", WHITE, INK)]),
        ("btn", "확인"), ("btn2", "아직 안 썼어요"),
    ], label="[가상] 05 약 기록 확인")
    phone(s, M + Inches(3.55), Inches(1.45), [
        ("bar", "‹ 이전                    ⌂ 홈"),
        ("gap", 0.1),
        ("big", ("0352번 소", INK)),
        ("card", ([("휴약 중", 10, False, AMBER), ("10월 18일 오전 9시 이후 출하할 수 있어요", 12, True, AMBER),
                   ("휴약: 약 성분이 소 몸에서 빠질 때까지 기다리는 기간", 9, False, AMBER)], AMBER_SOFT)),
        ("card", ([("HACCP 약품 사용 기록이 자동으로 채워졌어요", 10, True, GREEN),
                   ("평가표 29번 · 약품 입출고 기록", 9, False, SUB)], GREEN_SOFT)),
        ("gap", 0.1),
        ("btn", "홈으로"), ("btn2", "잘못 골랐어요"),
    ], label="[가상] 06 휴약 종료 안내")
    box(s, Inches(7.15), Inches(1.45), Inches(5.55), Inches(5.45), fill=BG, pad=0.3, lines=[
        ("휴약 종료 계산 규칙", 15, True, INK), ("", 3, False, INK),
        ("• 마지막 투약 시각 + 휴약일수 × 24시간 [확인됨]", 12, False, INK),
        ("   예) 금요일 오전 9시 + 5일 → 다음 주 수요일 오전 9시 이후", 11, False, SUB),
        ("• 여러 약을 썼으면 가장 늦게 끝나는 날짜 [확인됨]", 12, False, INK),
        ("• 같은 성분도 제조사마다 휴약기간이 달라 제품 단위로 저장 [확인됨]", 12, False, INK),
        ("• 투약 시각을 모르면 그날 23시 59분으로 계산 [가정]", 12, False, INK), ("", 6, False, INK),
        ("이 화면의 규칙", 15, True, INK), ("", 3, False, INK),
        ("• 처방전에 축사 번호가 있으면 그 칸을 먼저 펼침", 12, False, INK),
        ("• '아직 안 씀' → 3일 뒤 다시 묻기 [가정]", 12, False, INK),
        ("• 7일 동안 안 고르면 농장 전체 '휴약 확인 필요' [가정]", 12, False, INK),
        ("• 칸 이름을 누르면 칸 전체 선택", 12, False, INK)])
    foot(s, "계산 규칙 출처: 농림축산식품부·농림축산검역본부 「동물용의약(외)품 안전사용 안내」")
    notes(s, "【설명】 이 화면이 앱의 심장입니다. 농장주가 하는 일은 소 번호를 누르고, 언제 썼는지 누르고, 확인을 누르는 세 번뿐입니다. 나머지 계산과 장부 작성은 앱이 합니다.")


def s_screens_a(prs):
    s = blank(prs)
    header(s, "화면 · 04 · 18 · 02", "홈에서 목록으로, 목록에서 소 상세로")
    phone(s, M + Inches(0.3), Inches(1.4), [
        ("bar", "9월 22일 (화)  박영수 님"),
        ("card", ([("🔔 확인할 약 기록 2건", 13, True, ACC),
                   ("어느 소에 쓰셨는지 골라 주세요.", 9, False, ACC)], ACC_SOFT)),
        ("row", ("휴약 중", "5마리", AMBER, AMBER_SOFT)),
        ("row", ("출하 가능", "57마리", GREEN, GREEN_SOFT)),
        ("row", ("출하 주의", "2마리", ACC, ACC_SOFT)),
        ("row", ("2주 안 출하 예정", "4마리", SUB, BG)),
        ("gap", 0.05),
        ("btn", "📝 오늘 한 일 기록"), ("btn2", "🐄 우리 소"), ("btn2", "📊 9월 농장 성적표"),
    ], label="04 홈 (숫자를 누르면 목록)")
    phone(s, M + Inches(4.3), Inches(1.4), [
        ("bar", "‹ 이전                    ⌂ 홈"),
        ("big", ("휴약 중 5마리", INK)),
        ("s", "끝나는 날이 빠른 순서입니다."),
        ("row", ("0603 · 18개월", "휴약 ~9/27", AMBER, AMBER_SOFT)),
        ("row", ("0511 · 24개월", "휴약 ~9/30", AMBER, AMBER_SOFT)),
        ("row", ("0508 · 27개월", "휴약 ~10/5", AMBER, AMBER_SOFT)),
        ("row", ("0421 · 29개월", "출하 주의", ACC, ACC_SOFT)),
        ("row", ("0352 · 30개월", "출하 주의", ACC, ACC_SOFT)),
        ("gap", 0.05),
        ("s", "우리 소 62마리는 칸별로 모아서 보여 줍니다."),
    ], label="18 휴약 중 목록")
    phone(s, M + Inches(8.3), Inches(1.4), [
        ("bar", "‹ 이전                  다음 ›"),
        ("big", ("0353번 소", INK)),
        ("s", "1 / 3  이 소의 상태     3번 칸 · 거세 · 31개월"),
        ("card", ([("기록상 휴약 제한 없음", 11, True, GREEN),
                   ("확인된 약 사용 기록에 남은 휴약 기간이 없습니다.", 9, False, GREEN)], GREEN_SOFT)),
        ("s", "오늘 기준 · 2026년 9월 22일"),
        ("t", ("출하 예정일 · 10월 2일", INK)),
        ("s", "등급·사료비를 비교할 수 있어요."),
        ("btn", "언제 팔지 보기 (출하 판단)"),
    ], label="02 소 상세 1/3 (팀원 화면)")
    foot(s, "모든 숫자·번호는 [가상]입니다. 소 62마리를 한 벌의 자료로 정해 화면마다 숫자가 어긋나지 않게 맞췄습니다.")


def s_screens_b(prs):
    s = blank(prs)
    header(s, "화면 · 10 · 07", "출하 판단과 출하 전 점검 — A와 B가 만나는 곳")
    phone(s, M + Inches(0.3), Inches(1.4), [
        ("bar", "‹ 이전                  다음 ›"),
        ("big", ("0353번 소", INK)),
        ("s", "2 / 3  언제 팔까요?"),
        ("card", ([("지금 팔면 → 1++ 44%", 10, True, INK),
                   ("1개월 더 → 1++ 50%", 10, True, INK),
                   ("사료비 +32만 원", 9, False, SUB)], BG)),
        ("t", ("이렇게 본 이유", INK)),
        ("row", ("혈통 KPN1180", "높임", GREEN, GREEN_SOFT)),
        ("row", ("거세 · 31개월", "높임", GREEN, GREEN_SOFT)),
        ("row", ("사료 단계 22·23번", "높임", GREEN, GREEN_SOFT)),
        ("row", ("치료 2회 · 31번", "낮춤", ACC, ACC_SOFT)),
        ("row", ("여름 고온 · 기상청", "낮춤", ACC, ACC_SOFT)),
        ("s", "※ 참고용 예상치입니다."),
        ("btn", "출하 예정에 넣기"),
    ], label="10 출하 판단 (2/3)")
    phone(s, M + Inches(4.3), Inches(1.4), [
        ("bar", "‹ 이전                    ⌂ 홈"),
        ("big", ("출하 전 점검", INK)),
        ("s", "출하 준비 · 10월 2일 · 휴약이 끝났는지 자동 확인"),
        ("row", ("0353번 소", "휴약 끝남", GREEN, GREEN_SOFT)),
        ("row", ("0360번 소", "휴약 끝남", GREEN, GREEN_SOFT)),
        ("row", ("0352번 소", "16일 남음", ACC, ACC_SOFT)),
        ("card", ([("0352번 소는 잔류 기준을 넘을 수 있어요", 10, True, ACC),
                   ("출하를 막지는 않습니다. 빼거나 수의사와 상담해 주세요.", 9, False, ACC)], ACC_SOFT)),
        ("btn", "0352번 빼고 진행"),
        ("btn2", "수의사에게 전화"),
        ("btn2", "그래도 진행 (사유 입력)"),
    ], label="07 출하 전 점검")
    box(s, Inches(8.45), Inches(1.4), Inches(4.25), Inches(5.5), fill=BG, pad=0.28, lines=[
        ("10 출하 판단", 14, True, INK),
        ("• 등급은 하나로 단정하지 않고 확률로", 11, False, INK),
        ("• 실제 판정은 축산물품질평가사 [확인됨]", 11, False, INK),
        ("• 사료비 = 하루 급여량 × 단가 × 일수 (농장주 입력)", 11, False, INK),
        ("• 근거는 약 기록만이 아니라 HACCP 사양·건강 기록과 혈통·월령·기상까지", 11, False, INK), ("", 6, False, INK),
        ("07 출하 전 점검 → 12·13 마무리 화면", 14, True, INK),
        ("• 막지 않고 경고만 [확인됨: 차단 근거 없음]", 11, False, INK),
        ("• 빼고 진행 → '출하 명단 확정' 화면, HACCP 출하 관리(36번) 자동 작성", 11, False, INK),
        ("• 그래도 진행 → '진행 사유 입력' 화면에서 이유를 고르고 장부에 남김", 11, False, INK),
        ("  (일본 인증 농장의 중요관리점과 같은 역할)", 10, False, SUB),
        ("• 수의사에게 전화는 팝업으로, 이력제 출하 신고는 기존대로", 11, False, INK)])
    foot(s, "출하 판단의 모든 숫자는 [가상]입니다. 시범 사업에서 실제 예측 모델을 만듭니다.")


def s_screens_c(prs):
    s = blank(prs)
    header(s, "화면 · 03 · 17 · 문자", "동의, HACCP 장부, 앱이 어려운 분을 위한 문자")
    phone(s, M + Inches(0.3), Inches(1.4), [
        ("bar", "한우 출하노트"),
        ("t", ("박영수 님 농장을 찾았습니다", INK)),
        ("s", "농장식별번호 402517 · 소 62마리"),
        ("t", ("가져올 정보를 골라 주세요", INK)),
        ("row", ("수의사 처방 기록", "✓ 동의", WHITE, GREEN)),
        ("row", ("약 구매 기록", "✓ 동의", WHITE, GREEN)),
        ("row", ("도축·등급 결과", "✓ 동의", WHITE, GREEN)),
        ("s", "목적·항목·보관 기간·거부할 권리와 불이익을 알려드립니다. 자세히 보기 ›"),
        ("gap", 0.1),
        ("btn", "동의하고 시작"), ("btn2", "나중에"),
    ], label="03 처음 시작 · 동의")
    phone(s, M + Inches(4.3), Inches(1.4), [
        ("bar", "‹ 이전                    ⌂ 홈"),
        ("big", ("HACCP 장부 미리보기", INK)),
        ("card", ([("HACCP 개체 기록부 · 0353번", 11, True, INK),
                   ("26.09.10  36번 출하 관리", 9, False, INK),
                   ("26.08.02  31번 질병예방", 9, False, INK),
                   ("26.06.15  22번 사육단계", 9, False, INK),
                   ("25.03.11  34번 입식 310kg", 9, False, INK)], BG)),
        ("s", "심사 때 이 장부를 그대로 보여 주면 됩니다."),
        ("btn", "📤 카카오톡으로 보내기"), ("btn2", "💾 휴대폰에 저장"),
    ], label="17 HACCP 장부")
    box(s, Inches(8.45), Inches(1.4), Inches(4.25), Inches(2.75), fill=GREEN_SOFT, pad=0.25, lines=[
        ("문자 대안 (권장)", 14, True, GREEN),
        ("[한우 출하노트] 9/20 한빛동물약품에서 옥시테트라사이클린 주사 1병을 사셨습니다. 쓰신 소의 귀표 뒷번호를 답장해 주세요. 예) 0352 0353 / 안 썼으면 0", 11, False, INK),
        ("→ 답장: \"0352 확인. 10/18 09시 이후 출하 가능.\"", 11, True, INK)])
    box(s, Inches(8.45), Inches(4.3), Inches(4.25), Inches(2.6), fill=BG, pad=0.25, lines=[
        ("동의·설정 규칙", 14, True, INK),
        ("• 동의 시 목적·항목·보유 기간·거부 권리와 불이익을 알림 [확인됨: 개인정보 보호법 제15조②]", 11, False, INK),
        ("• 하나를 거부해도 나머지 기능은 씀", 11, False, INK),
        ("• 설정에서 항목별 철회, 가족 초대, 글씨 크기, 알림 방식(앱/문자)", 11, False, INK),
        ("• 장부 양식은 축사로·인증원 양식에 맞춤 [가정: 확인 필요]", 11, False, INK)])


def s_data(prs):
    s = blank(prs)
    header(s, "데이터", "현황 분석에 쓴 데이터와, 서비스에 필요한 데이터", "기획서 2-2 [표 7]·[표 8]")
    rows = [["현황 분석에 활용 (완료)", "규모·기준", "분석 활용"],
            ["HACCP 인증업체정보 / 인증원", "41,069건 / 2026.9.8. 수집", "축종·시도별 인증 농장 수, 신규 인증 추이"],
            ["축산물 스마트HACCP 인증업체 정보 / 인증원", "563행 → 298곳", "업종·중요관리점 구성, 농장 적용 공백 진단"],
            ["가축동향조사 / 국가데이터처", "2026년 2분기", "인증률 분모, 규모 보정, 시범 지역 선정"]]
    table(s, M + Inches(0.08), Inches(1.95), W - 2 * M - Inches(0.16), rows,
          [Inches(4.2), Inches(3.0), Inches(4.77)], row_h=Inches(0.5), size=12)
    rows2 = [["서비스 구축 시 확보", "필요한 데이터", "확보 상태·단계별 방법"],
             ["농가 입력 / 앱·축사로", "사료 단계·입고, 질병·치료, 입식, 소독·깔짚·폐사", "시범 사업부터 항목별 수집. 기존 축사로 기록은 동의·협의 후 활용"],
             ["축산물이력제 / 축평원", "출생·성별·혈통·이동, 도축일·등급·도체중", "개체번호별 공개 조회 활용 검토, 시범에서 동의 기반 마이데이터 연계"],
             ["전자처방(eVET) / 대한수의사회", "농장명, 약 성분·용량·수량, 처방·직접 투약", "본인 조회 서비스 존재, 공개 API 미확인 → 연계 협의 필요"],
             ["약 판매처", "판매일, 제품명·수량, 구매자", "초기에는 포장 인식·목록 선택. 참여 판매처와 전산 연계 추진"],
             ["축사로 / 국립축산과학원", "유전능력 평가값, 초음파 측정값, 기존 HACCP 기록", "공개 API 미확인. 장부 내보내기부터 시작, 기관 협의"],
             ["약 포장·첨부문서 / 제조사", "제품별 휴약기간", "포장 기재 의무 있으나 공개 데이터 없음 → 시범 약품 기준표 구축"]]
    table(s, M + Inches(0.08), Inches(4.0), W - 2 * M - Inches(0.16), rows2,
          [Inches(3.0), Inches(3.6), Inches(5.37)], row_h=Inches(0.52), size=12, hot_rows=(3, 6))
    foot(s, "인증원 데이터 2종은 이용허락범위 제한 없음. 제출 자료에는 개별 농장명·대표자명 없이 집계값만 사용. 현재 프로토타입은 가상 데이터로 구성")


def s_model(prs):
    s = blank(prs)
    header(s, "출하 판단 모델", "HACCP 기록이 예측을 개선하는지부터 측정", "기획서 2-3 (3)")
    box(s, M, Inches(1.9), Inches(6.1), Inches(2.6), fill=BG, pad=0.32, lines=[
        ("모델 구성", 16, True, INK), ("", 3, False, INK),
        ("기본 모델 = 이력제·축사로의 유전·성별·월령", 13, False, INK),
        ("HACCP 모델 = 기본 + 사양·건강·환경 기록", 13, True, INK),
        ("두 모델의 예측력 차이가 곧 'HACCP 기록의 값'", 13, True, GREEN), ("", 3, False, INK),
        ("방법: 등급 구간별 확률을 내는 분류 모델(다항 로지스틱 회귀, 그래디언트 부스팅 등) + 확률 보정. "
         "로그 손실·브라이어 점수, 도체중은 MAE·RMSE", 12, False, SUB)])
    box(s, M, Inches(4.65), Inches(6.1), Inches(2.2), fill=WHITE, line=LINE, pad=0.3, lines=[
        ("설계 원칙", 15, True, INK),
        ("• 출하 전 알 수 없는 도체형질·도축 후 정보는 입력에서 제외", 12, False, INK),
        ("• 학습에 쓰지 않은 농장·이후 시기 개체로 검증(농장·시간 분리)", 12, False, INK),
        ("• '치료 없음'과 '기록 없음'을 구분", 12, False, INK),
        ("• 치료 이력은 관리 상황 지표이며 인과 효과로 해석하지 않음", 12, False, INK)])
    box(s, Inches(7.05), Inches(1.9), Inches(5.6), Inches(2.1), fill=ACC_SOFT, pad=0.3, lines=[
        ("정확도 전제", 15, True, ACC), ("", 3, False, INK),
        ("한우 근내지방도 예측의 상관계수는 선행 연구에서도 0.64 수준(Shahinfar 외 2020, 3,989두)", 13, False, INK),
        ("→ 효과를 약속하지 않고 측정부터 수행. 도입 초기 제공 값은 예측이 아닌 실측 결과(월간 성적표)", 13, True, INK)])
    box(s, Inches(7.05), Inches(4.15), Inches(5.6), Inches(1.35), fill=GREEN_SOFT, pad=0.3, lines=[
        ("경제성 비교", 15, True, GREEN),
        ("예상 등급 확률과 함께 예상 도체중·판매금액·추가 사육비를 비교. 확정적 '최적 출하일'이 아니라 "
         "시나리오별 예상 손익과 가정·예측 범위를 함께 제시", 13, False, INK)])
    box(s, Inches(7.05), Inches(5.65), Inches(5.6), Inches(1.2), fill=WHITE, line=LINE, pad=0.3, lines=[
        ("결과 설명", 15, True, INK),
        ("XAI 기법으로 요인별 기여도(높임/낮춤)를 산출해 정해진 쉬운 문장 틀로 제시. 프로토타입은 설명용 모의 규칙", 13, False, SUB)])


def s_demo(prs):
    s = blank(prs)
    header(s, "시연 순서", "기획서 2-1 이용 흐름 그대로 눌러 가며 보여 줍니다")
    rows = [["순서", "화면", "말할 내용"],
            ["1", "04 홈", "박영수 씨 농장 62마리. 확인할 약 기록 2건 알림, 휴약 중 5 · 출하 가능 57 · 출하 주의 2."],
            ["2", "05 약 기록 확인", "처방·판매로 들어온 기록을 확인. 연계되지 않은 약은 22 약 포장 인식으로 불러옴"],
            ["3", "06 휴약 종료 안내", "10월 18일 오전 9시 이후 출하 가능. 계산도 HACCP 기록(29번)도 앱이 합니다."],
            ["4", "08 오늘 한 일", "물통 소독을 눌러 HACCP 일지에 기록. 여기까지가 위생·안전입니다."],
            ["5", "09·15 성적표", "이력제 실제 등급과 호흡기 치료율을 지역 평균과 비교합니다."],
            ["6", "10 출하 판단", "지금 출하와 추가 사육의 예상 등급·사료비 비교(시범 후반 기능, 화면은 설명용 모의 규칙)"],
            ["7", "07·12 출하 점검", "휴약이 남은 0352번을 넣자 경고. 빼고 진행하면 출하 기록(36번)이 자동으로 남습니다."]]
    table(s, M + Inches(0.08), Inches(1.55), W - 2 * M - Inches(0.16), rows,
          [Inches(0.9), Inches(2.3), Inches(8.77)], row_h=Inches(0.62), size=13)
    box(s, M, Inches(6.6), W - 2 * M, Inches(0.45), fill=GREEN_SOFT, anchor=MSO_ANCHOR.MIDDLE, lines=[
        ("화면의 수치·농장명·제품명은 가상 자료입니다. 실제 연계·계산은 후속 구현 대상입니다.", 12, True, GREEN)])


def s_legal(prs):
    s = blank(prs)
    header(s, "지켜야 할 조건", "법령과 제도 안에서 설계했습니다", "기획서 1-1·2-1·2-3에 흩어진 조건을 모은 장입니다.")
    rows = [["쟁점", "근거", "설계 반영"],
            ["개인정보", "개인정보 보호법 제15조·제17조", "농가 동의 범위에서만 수집·연계, 항목별 철회 지원"],
            ["학습 자료", "—", "분석·학습 자료는 농장을 식별할 수 없도록 가명 처리"],
            ["이상 탐지 결과", "—", "농가에 먼저 통보, 인증원에는 농장 단위 위험도만 전달"],
            ["출하 제한", "차단 근거 없음", "차단하지 않고 경고. 강행 시 사유 기록"],
            ["예측 정보", "등급판정은 축산물품질평가사", "확률로 표시하고 참고용임을 고지"],
            ["잔류 규제", "PLS 2024.1.1. 시행", "휴약 미경과 개체는 출하 전 점검에서 경고, 점검 결과를 장부에 저장"],
            ["휴약기간 표기", "동물용 의약품등 취급규칙 제42조", "포장·첨부문서 기재 의무를 근거로 제품별 기준표 구축"]]
    table(s, M + Inches(0.08), Inches(2.0), W - 2 * M - Inches(0.16), rows,
          [Inches(2.4), Inches(3.6), Inches(5.97)], row_h=Inches(0.62), size=13)


def s_kpi(prs):
    s = blank(prs)
    header(s, "성과 지표", "시범 사업에서 처음 재는 것", "기획서 3-1 [표 11] 성과 지표와 시범 사업 목표(안)")
    rows = [["지표", "현재", "시범 사업 목표(안)", "측정 방법"],
            ["HACCP 인증 시범 농장", "한우 인증률 2.5%", "참여 농장 인증 취득·유지", "인증원 HACCP 인증업체정보"],
            ["출하 개체 휴약 자동 확인 비율", "0% (전산 대조 수단 없음)", "시범 농장 출하 100%", "앱 출하 점검 기록 ÷ 이력제 출하 신고"],
            ["약 기록 확인률", "수기 기록, 전산 확인 불가", "7일 이내 80% 이상", "수집된 처방·판매 기록 중 투약 개체가 지정된 비율"],
            ["투약 1건 기록 소요", "수기 기록 (도입 전 측정)", "터치 3회·1분 이내", "도입 전 수기 기록과 앱 입력 시간 비교"],
            ["경고 후 조치율", "경고 체계 없음", "경고 개체 100% 조치 또는 사유 기록", "출하 제외·수의사 상담·사유 기록 비율"],
            ["투약 개체 지정 단독 완료율", "해당 없음", "60세 이상 농가 80% 이상", "사용성 시험·관찰"],
            ["HACCP 기록의 예측 기여", "기준 없음", "기본 모델 대비 개선 여부 검증", "로그 손실·브라이어 점수, 도체중 MAE·RMSE 비교"],
            ["지속 사용률", "해당 없음", "6개월 후 70% 이상", "월별 기록 입력 농가 비율, 농가·수의사 의견"]]
    table(s, M + Inches(0.08), Inches(1.95), W - 2 * M - Inches(0.16), rows,
          [Inches(2.9), Inches(2.6), Inches(3.0), Inches(3.47)], row_h=Inches(0.56), size=12, hot_rows=(2, 4, 6))


def s_build(prs):
    s = blank(prs)
    header(s, "추진 계획", "기획서 3-2 [표 13]·[표 14]", "1단계는 끝났고, 다음은 실제로 계산이 도는 모바일 웹앱입니다.")
    opts = [("1. 프로토타입 (완료)", "클릭 프로토타입", "화면 84개·팝업 11개·연결 613곳 제작 완료. 후속으로 휴약 계산·출하 판단이 실제 동작하는 모바일 웹앱 구현", "결과물: 휴대폰 구동 시연 앱", GREEN_SOFT, GREEN),
            ("2. 시범 사업", "1~2개 시군", "한우 농장, 참여 수의사·판매처와 전자처방·판매기록·축사로 연계", "결과물: 성과 지표 첫 측정, 출하 판단 모델 검증", BG, INK),
            ("3. 확산", "제도·사업 편입", "판매기록 전산 제출 규정 검토, 농장 HACCP 지원 사업에 기록 연계 모듈 편입", "결과물: 농장 HACCP 전산 기록 표준 모델", BG, INK)]
    bw = (W - 2 * M - Inches(0.3) * 2) / 3
    for i, (name, what, good, bad, fill, col) in enumerate(opts):
        x = M + (bw + Inches(0.3)) * i
        box(s, x, Inches(1.6), bw, Inches(2.45), fill=fill, pad=0.25, lines=[
            (name, 15, True, col), (what, 11, False, SUB), ("", 3, False, INK),
            ("장점  " + good, 11, False, INK), ("단점  " + bad, 11, False, ACC)])
    rows = [["단계", "할 일", "끝났다고 보는 기준", "상태"],
            ["1", "가상 데이터: 농장 1곳, 소 62마리(칸별 구성·휴약·출하 예정), 약 제품, 처방·판매, 도축 결과", "전 화면 수치가 일치하는 단일 데이터셋", "완료"],
            ["2", "클릭 프로토타입: 동의·홈·약 기록·휴약 안내·출하 점검·기록·성적표·출하 판단·장부", "화면 84개·연결 613곳, 동의부터 출하 명단 확정까지 연속 조작", "완료"],
            ["3", "고령 사용자 설계 기준 적용: 대형 글자, 이전·홈 버튼, 색상·문자 병기, 달력 선택, 용어 해설", "화면별 휴대폰 한 화면 표시(소 목록 제외), 본문 16pt 이상", "완료"],
            ["4", "휴약 계산 규칙 구현·검증", "농림축산식품부 안내문 예시(금 9시 + 5일 → 수 9시) 검사 통과", "예정"],
            ["5", "모바일 웹앱 전환, 출하 판단(설명용 모의 규칙) 연결", "링크·QR 접속 후 핵심 흐름 전 구간 무중단 조작", "예정"],
            ["6", "60세 이상 사용자 3명 대상 사용성 시험", "투약 개체 지정 단독 완료 여부와 소요 시간 기록", "예정"]]
    table(s, M + Inches(0.08), Inches(4.3), W - 2 * M - Inches(0.16), rows,
          [Inches(0.7), Inches(6.0), Inches(4.37), Inches(0.9)], row_h=Inches(0.38), size=11)


def s_open(prs):
    s = blank(prs)
    header(s, "한계와 확인 과제", "모르는 것은 모른다고 적었습니다", "기획서 3-2 '한계와 확인 과제'")
    items = [("데이터 연계", "eVET은 공개 API 부재 → 대한수의사회·농식품부와 동의 기반 연계 협의 필요. 처방 없이 판매되는 약은 "
                           "판매자 보고 의무가 없어 초기에는 농가 입력(포장 인식·목록 선택)에 의존. 동물용의약품 포장 바코드 표준 여부도 확인 필요", ACC_SOFT, ACC),
             ("효과 근거", "HACCP의 등급 개선 근거는 일본 농장 1곳 사례뿐이며 비교 농장 부재 → 2-3 (5)의 이중차분 설계로 검증", BG, INK),
             ("차별성 한계", "개별 기능(외부 의약품 기록 연계, 개체별 휴약 관리, 도체형질 예측)은 국내외 선행 사례 존재. "
                           "본 조사만으로 동일한 통합 서비스가 없다고 단정하지 않으며, 차별성은 연계 범위와 실증 성과로 제시", BG, INK)]
    y = Inches(2.0)
    for head, body, fill, col in items:
        box(s, M, y, W - 2 * M, Inches(1.35), fill=fill, pad=0.32, anchor=MSO_ANCHOR.MIDDLE, lines=[
            (head, 16, True, col), (body, 13, False, INK)])
        y += Inches(1.5)
    box(s, M, y + Inches(0.15), W - 2 * M, Inches(0.9), fill=GREEN_SOFT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, lines=[
        ("확인된 것과 확인해야 할 것을 나눠 적는 것이 이 기획서의 원칙입니다.", 15, True, GREEN)])


def s_proto(prs):
    s = blank(prs)
    header(s, "프로토타입 현황", "설계만이 아니라 눌러 볼 수 있게 만들었습니다",
           "동의 화면에서 시작해 출하 명단 확정까지 조작할 수 있습니다. 실제 API 연계와 계산·예측 기능은 후속 구현 대상입니다.")
    stats = [("84", "화면", "팀원 화면 2개 포함"),
             ("11", "눌렀을 때 뜨는 창", "달력·전화·저장 등"),
             ("613", "화면 사이 연결", "안 되는 버튼 없음"),
             ("62", "소 마릿수", "칸별 구성까지 일치")]
    bw = (W - 2 * M - Inches(0.25) * 3) / 4
    for i, (big, head, small) in enumerate(stats):
        x = M + (bw + Inches(0.25)) * i
        box(s, x, Inches(2.0), bw, Inches(1.5), fill=BG if i % 2 == 0 else GREEN_SOFT,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, lines=[
                (big, 38, True, ACC if i % 2 == 0 else GREEN), (head, 14, True, INK), (small, 11, False, MUTE)])
    rows = [["묶음", "화면", "무엇을 하나"],
            ["팀원 원본", "01 우리 소, 02 소 상세 1/3", "팀원이 만든 화면을 그대로 옮기고 다음 단계만 이어 붙임"],
            ["A 위생·안전", "03 동의, 04 홈, 05 약 기록 확인, 22 약 포장 인식, 06 휴약 종료 안내, 07 출하 전 점검", "약 기록을 채우고 출하 전에 휴약을 확인"],
            ["A 마무리", "12 출하 명단 확정, 12-2 사유 기록, 13 진행 사유 입력, 14 기록 저장 완료", "경고 뒤 두 갈래를 모두 화면으로 만듦"],
            ["B 보상", "08 오늘 한 일, 09 농장 성적표, 15 결과 자세히, 16 지난달과 비교, 10 출하 판단, 11 기록 이력, 17 HACCP 장부", "HACCP 기록을 성적표·예상 등급·심사용 장부로 돌려줌"],
            ["목록", "18 휴약 중 5, 19 출하 가능 57, 20 출하 주의 2, 21 2주 안 출하 예정 4", "홈의 숫자를 누르면 그 수만큼 목록이 열림"],
            ["소마다", "02-0352 ~ 02-0815 (61개)", "어느 목록에서 어느 소를 눌러도 그 소의 화면이 열림"]]
    table(s, M + Inches(0.08), Inches(3.62), W - 2 * M - Inches(0.16), rows,
          [Inches(1.3), Inches(5.2), Inches(5.47)], row_h=Inches(0.45), size=12)
    foot(s, "화면 수·연결 수는 파일에서 세어 확인한 값입니다. 화면 속 숫자·이름·전화번호는 모두 [가상]입니다.")
    notes(s, "【설명】 기획서에 그림만 넣은 것이 아니라, 심사위원이 직접 눌러 볼 수 있는 프로토타입을 만들었습니다.")


def s_senior(prs):
    s = blank(prs)
    header(s, "어르신이 쓰는 앱", "60대 이상 농장주 기준으로 다시 손봤습니다")
    rows = [["바꾼 것", "전", "후"],
            ["글자 크기", "작은 설명 글 12~14pt", "본문 16pt 이상, 버튼 글자 19pt, 제목 26~30pt"],
            ["버튼", "화면마다 다름", "버튼 높이를 키우고 손가락으로 누르기 쉽게 통일"],
            ["돌아가기", "작은 글씨 '이전 화면'", "모든 화면 위쪽에 큰 '‹ 이전'과 '⌂ 홈' 버튼"],
            ["그림 표시", "글자만", "📝 오늘 한 일, 🐄 우리 소, 📊 성적표, 💧 물통 소독처럼 그림과 함께"],
            ["날짜 입력", "며칠인지 골라 적기", "달력에서 직접 고르기 (지난 달·다음 달로 넘김, 못 고르는 날은 흐리게)"],
            ["어려운 말", "'휴약'만 표시", "'휴약: 약 성분이 소 몸에서 빠질 때까지 기다리는 기간' 뜻풀이 함께"],
            ["색과 글자", "색으로만 구분", "'휴약 끝남' / '휴약 ~10/5' / '출하 주의'처럼 글자를 함께 표시"],
            ["한 화면", "스크롤해야 보이는 화면", "긴 소 목록을 빼면 모든 화면이 휴대폰 한 화면에 들어옴"]]
    table(s, M + Inches(0.08), Inches(1.6), W - 2 * M - Inches(0.16), rows,
          [Inches(1.6), Inches(3.4), Inches(6.97)], row_h=Inches(0.5), size=13, hot_rows=(1, 3, 5))
    box(s, M, Inches(6.5), W - 2 * M, Inches(0.55), fill=GREEN_SOFT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, lines=[
        ("설계 기준은 그대로입니다 — 박영수 씨가 자녀 도움 없이 '약 쓴 소 고르기'를 끝낼 수 있는가.", 15, True, GREEN)])


def s_dataset(prs):
    s = blank(prs)
    header(s, "시연 데이터", "소 62마리를 한 벌로 정하고 모든 화면을 맞췄습니다",
           "홈에 '출하 가능 57마리'라고 뜨면 실제로 57마리 목록이 열립니다. [가상] 자료입니다.")
    rows = [["칸", "마릿수", "월령", "설명"],
            ["1번 칸", "11마리", "12~15개월", "육성 단계"],
            ["2번 칸", "11마리", "16~20개월", "비육 전기 (0603번 휴약 중)"],
            ["3번 칸", "10마리", "28~31개월", "출하 임박 (0352·0353·0360번)"],
            ["4번 칸", "10마리", "22~30개월", "0411 약 기록 확인, 0415 자료 부족, 0421 출하 주의"],
            ["5번 칸", "10마리", "24~27개월", "0508·0511번 휴약 중"],
            ["6번 칸", "10마리", "8~11개월", "입식 직후"]]
    table(s, M + Inches(0.08), Inches(2.0), Inches(6.3), rows,
          [Inches(1.0), Inches(1.0), Inches(1.3), Inches(3.0)], row_h=Inches(0.55), size=12)
    rows2 = [["홈에 뜨는 숫자", "뜻", "목록"],
             ["휴약 중 5마리", "휴약이 안 끝난 소", "0603·0511·0508 0421·0352"],
             ["출하 가능 57마리", "남은 휴약 기간이 없는 소", "칸별로 모아서 표시"],
             ["출하 주의 2마리", "예정일에 휴약이 안 끝남", "0352(16일)·0421(4일)"],
             ["2주 안 출하 예정 4마리", "9/22~10/6", "0412·0353·0360·0352"]]
    table(s, Inches(7.0), Inches(2.0), Inches(5.6), rows2,
          [Inches(1.9), Inches(1.9), Inches(1.8)], row_h=Inches(0.58), size=12)
    box(s, Inches(7.0), Inches(5.05), Inches(5.6), Inches(1.85), fill=BG, pad=0.28, lines=[
        ("맞춰 둔 것", 14, True, INK),
        ("• 5마리 + 57마리 = 62마리 (출하 주의 2마리는 휴약 중 5마리에 포함)", 11, False, INK),
        ("• 출하 전 점검(10월 2일) 명단과 소별 출하 예정일이 일치", 11, False, INK),
        ("• 성적표의 호흡기 치료율 10%는 1·6번 칸 21마리 중 2마리", 11, False, INK),
        ("• 2026년 9월 22일은 화요일 — 달력 요일까지 실제와 맞춤", 11, False, INK)])
    foot(s, "화면마다 숫자가 어긋나면 심사 때 바로 드러납니다. 자료를 한 벌로 고정한 이유입니다.")


def s_ai(prs):
    s = blank(prs)
    header(s, "인공지능", "도입 범위와 순서", "기획서 2-1 'AI 도입 범위'와 2-3 (3)·(4)를 옮긴 장입니다.")
    rows = [["단계", "적용", "내용", "판단 근거"],
            ["시범 초기", "약 라벨·포장 인식", "바코드 또는 라벨 문자를 인식해 제품명·성분·휴약기간 입력을 보조",
             "처방 없이 구매한 약은 전산 기록이 없어 농가 입력에 의존"],
            ["데이터 축적 후", "출하 판단 모델 검증", "기본 모델(혈통·거세·월령)과 HACCP 모델(사양·건강·위생 추가)의 예측력 차이를 측정",
             "추가 기록의 예측 기여가 확인돼야 기능으로 제공"],
            ["확산 단계", "기록 이상 탐지 (검토)", "일괄 기재·누락·휴약 위반 위험 탐지 → 농가 통보, 인증원에는 농장 단위 위험도",
             "제한된 심사 인력으로 고위험 농장 우선 점검"]]
    table(s, M + Inches(0.08), Inches(2.0), Inches(8.4), rows,
          [Inches(1.3), Inches(1.9), Inches(3.0), Inches(2.2)], row_h=Inches(1.15), size=12, hot_rows=(1,))
    box(s, Inches(9.05), Inches(1.95), Inches(3.65), Inches(2.3), fill=BG, pad=0.28, lines=[
        ("인공지능을 쓰지 않는 곳", 14, True, INK), ("", 3, False, INK),
        ("• 휴약 계산은 규칙 기반 처리 (마지막 투약 시각 + 휴약일수 × 24시간)", 12, False, INK),
        ("• 결과 설명은 XAI 기여도 + 정해진 문장 틀", 12, False, INK),
        ("• 생성형 언어모델은 오해석 위험으로 사용하지 않음", 12, True, ACC)])
    box(s, Inches(9.05), Inches(4.4), Inches(3.65), Inches(2.45), fill=GREEN_SOFT, pad=0.28, lines=[
        ("구현 조건과 한계", 14, True, GREEN), ("", 3, False, INK),
        ("• 동물용의약품 포장 바코드 표준 미확인 → 라벨 문자 인식·목록 선택 병행", 12, False, INK),
        ("• 제품별 휴약기간 표는 시범 대상 약품 설명서로 구축, 인식값은 기준표 대조와 농가 확인", 12, False, INK),
        ("• 기록·사진은 농가 동의 범위에서만 사용, 학습 시 가명 처리", 12, False, INK)])
    foot(s, "정확도 전제: 한우 근내지방도 예측의 상관계수는 선행 연구에서도 0.64 수준(Shahinfar 외 2020) — 효과를 약속하지 않고 측정부터 수행")


def s_policy(prs):
    s = blank(prs)
    header(s, "정책 제언", "농장 HACCP 참여를 넓히려면", "기획서 3-2 '농장 HACCP 참여 확대를 위한 제언'과 같습니다.")
    items = [("농장형 관리 기준", "농장 HACCP의 핵심 관리점인 '휴약기간 경과 확인'을 기록 연계로 관리하는 방식을 인정하고, "
                                "앱·축사로 전산 장부를 심사 자료로 활용"),
             ("데이터 기반 정비", "약 판매기록의 전산 제출·농장번호 기재 규정 마련(동물용 의약품등 취급규칙 제22조 보완), "
                                "제품별 휴약기간의 공개 데이터화"),
             ("지원 방식 전환", "방문 컨설팅(2024년 예산 기준 169~330곳) 중심에서 기존 전자 기록 연계형 지원으로 전환해 "
                              "50마리 이상 한·육우 2만여 곳 규모로 확장"),
             ("확장 방향", "인증 농장이 늘어난 뒤, 정부 로드맵에 있으나 아직 실행되지 않은 생산단계 스마트HACCP 농장 단계와 연결")]
    y = Inches(1.95)
    for i, (head, body) in enumerate(items):
        box(s, M, y, W - 2 * M, Inches(1.05), fill=BG if i % 2 == 0 else WHITE, line=None if i % 2 == 0 else LINE,
            pad=0.28, anchor=MSO_ANCHOR.MIDDLE, lines=[(head, 15, True, INK), (body, 13, False, SUB)])
        y += Inches(1.15)
    box(s, M, y + Inches(0.1), W - 2 * M, Inches(1.0), fill=GREEN_SOFT, pad=0.3, anchor=MSO_ANCHOR.MIDDLE, lines=[
        ("본 제안의 위치", 15, True, GREEN),
        ("「한우 출하노트」는 위 방향이 농장 현장에서 작동함을 보여 주는 시범 모델. 인증 농장이 늘면 식품의약품안전처의 "
         "스마트HACCP 확대 방향(2026년 650개소 → 2030년 1,050개소)과 연결", 13, False, INK)])


def main():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    cover(prs); s_summary(prs); s_why(prs); s_existing(prs); s_user(prs)
    s_flow(prs); s_features(prs); s_principles(prs)
    s_proto(prs); s_pick(prs); s_screens_a(prs); s_screens_b(prs); s_screens_c(prs)
    s_senior(prs); s_dataset(prs)
    s_ai(prs); s_policy(prs); s_data(prs); s_model(prs); s_demo(prs); s_legal(prs); s_kpi(prs); s_build(prs); s_open(prs)
    now = datetime.datetime.now()
    cp = prs.core_properties
    cp.title = "한우 출하노트(가칭) 앱 프로토타입 구상 v1"
    cp.subject = "2026 한국식품안전관리인증원 공공데이터 분석·활용 경진대회"
    cp.author = ""
    cp.last_modified_by = ""
    cp.created = now
    cp.modified = now
    cp.revision = 1
    p = os.path.join(OUT, "한우출하노트_앱구상_v1.pptx")
    prs.save(p)
    print("[저장]", p, len(prs.slides), "장")


if __name__ == "__main__":
    main()
