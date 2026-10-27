# -*- coding: utf-8 -*-
"""
35. "닭은 묶어서 신청하니까 쉽고 한우는 하나씩이라 어렵다" — 틀렸다

【검사하는 주장】
34번이 축종별 농장 인증률을 내놓자 바로 나온 해석이다.
닭·오리는 계열주체가 위탁 농가를 묶어 관리하니 인증도 한 번에 들어가고,
한우는 농가마다 개별로 신청해야 하니 느리다는 것.
그럴듯하다. 그래서 검사한다.

【어떻게 재는가】
묶어서 신청하면 **같은 날 무더기로 인증일이 찍힌다.** 그게 관측 가능한 흔적이다.
축종별로 인증 농가의 최초 인증일을 모아, 같은 날에 몇 곳이 함께 들어왔는지 센다.
계열화 가설이 맞다면 닭의 동일자 묶음이 한우보다 뚜렷해야 한다.

【결과 — 정반대다】
한우가 닭보다 더 뭉쳐 있다(같은 날 5곳 이상 들어온 비중은 실행 결과 참고).
묶음 신청은 닭의 높은 인증률을 설명하지 못한다.

【그러면 진짜 이유는】
분자를 보면 답이 나온다. 인증받은 농장 수 자체가 축종 사이에 크게 다르지 않다.
연간 신규도 최근에는 한우가 닭보다 많다.
**한우의 낮은 인증률은 인증을 덜 받아서가 아니라, 비슷한 수를 받았는데 모집단이 23배라서다.**

【집계 기준】 — 34번·기획서 2-3절과 같은 규칙
'한우,육우'처럼 축종이 여럿 적힌 건은 쪼개서 각 축종에 넣고,
업체명+시도+시군구가 모두 같을 때만 같은 농장으로 본다.
부문별 신규 인증도 같은 키(업체명+시도+시군구)로 업소를 가른다.

산출: output/축종_동일자묶음.csv, output/축종_연도별신규.csv, output/부문별_연도별신규.csv,
      output/기획서_그림/그림_신규추이.png (기획서 [그림 1])
"""

import io
import os
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

OUT = "output"
FIG = os.path.join(OUT, "기획서_그림")

# 같은 농장(업소)으로 보는 기준. 34번과 같다.
농장키 = ["업체명", "시도", "시군구"]
축종군 = {"한우": "한우", "육우": "한우", "젖소": "젖소", "돼지": "돼지",
       "산란계": "닭", "육계": "닭", "오리": "오리"}
순서 = ["닭", "한우", "돼지", "젖소", "오리"]


def load():
    h = pd.read_csv(OUT + "/haccp_전처리.csv", dtype=str)
    a = h[(h["식품구분"] == "축산물") & (h["업태"] == "가축사육업")].copy()
    for c in 농장키:
        a[c] = a[c].fillna("").str.strip()
    # 복수 축종은 쪼개서 각 축종에 넣는다.
    a["축종"] = a["품목"].astype(str).str.replace(" ", "", regex=False).str.split(",")
    a = a.explode("축종", ignore_index=True)
    a["인증일"] = pd.to_datetime(a["인증일"], errors="coerce")
    a["군"] = a["축종"].map(축종군)
    a = a.dropna(subset=["군", "인증일"])
    # 농장 단위 최초 인증일. 갱신 건이 섞이면 묶음이 과장된다.
    # 한우·육우를 함께 받은 농장은 '한우' 군에서 한 번만 센다.
    f = a.sort_values("인증일").drop_duplicates(["군"] + 농장키).copy()
    f["연"] = f["인증일"].dt.year
    print("[로드] 가축사육업 농장 {:,}곳 (축종군별)".format(len(f)))
    return f


def 묶음(f):
    rows = []
    for g in 순서:
        s = f[f["군"] == g]
        c = s.groupby("인증일").size()
        rows.append({
            "축종": g, "인증농가": len(s), "고유인증일": len(c),
            "최대동일자": int(c.max()),
            "동일자2곳이상%": round(c[c >= 2].sum() / len(s) * 100, 1),
            "동일자5곳이상%": round(c[c >= 5].sum() / len(s) * 100, 1),
        })
    t = pd.DataFrame(rows)
    print()
    print("=" * 78)
    print(" 묶음 신청의 흔적 — 같은 날 몇 곳이 함께 들어왔는가")
    print("=" * 78)
    print(t.to_string(index=False))
    t.to_csv(OUT + "/축종_동일자묶음.csv", index=False, encoding="utf-8-sig")

    닭 = t[t["축종"] == "닭"].iloc[0]
    한우 = t[t["축종"] == "한우"].iloc[0]
    print()
    print(" 같은 날 5곳 이상: 한우 {:.1f}%  vs  닭 {:.1f}%".format(
        한우["동일자5곳이상%"], 닭["동일자5곳이상%"]))
    print(" 한우가 더 뭉쳐 있다. **계열화 묶음 가설은 지지되지 않는다.**")
    print(" 두 축종 모두 동일자 묶음이 있는데, 이건 계열화보다 심사 일정의 흔적으로 보인다.")
    return t


def 추세(f):
    t = (f.pivot_table(index="연", columns="군", values="업체명", aggfunc="count")
         .fillna(0).astype(int).reindex(columns=순서).loc[2015:2025])
    print()
    print("=" * 78)
    print(" 연도별 신규 인증 농가")
    print("=" * 78)
    print(t.to_string())
    t.to_csv(OUT + "/축종_연도별신규.csv", encoding="utf-8-sig")

    최근 = t.loc[2021:2025].sum()
    print()
    print(" [최근 5년 합계]")
    for g in 순서:
        print("   {:<4} {:>5,}곳  (연 {:.1f}곳)".format(g, 최근[g], 최근[g] / 5))
    한우 = t.loc[2021:2025, "한우"]
    print()
    print(" [기획서 1-1] 한우·육우 신규 인증 2021~2025년 {}곳".format(
        "·".join(str(v) for v in 한우)))
    print("   연평균 {:.1f}곳 = 한우 농장 72,389곳의 연 {:.2f}%".format(
        한우.mean(), 한우.mean() / 72389 * 100))
    print()
    print(" 최근 5년만 보면 한우가 닭보다 많다. 한우가 느린 게 아니다.")
    return t


def 부문별():
    """인증원 전체가 막힌 것인지, 농장만 막힌 것인지 가른다.

    식품 / 축산물 가공·유통 / 축산물 농장 세 부문의 연도별 신규 인증 업소를 센다.
    업소 단위 최초 인증 연도라 과거 연도는 생존편향으로 작게 잡힌다(8장).
    그런데도 농장은 2015년이 가장 크다. 실제 감소는 표보다 가파르다."""
    h = pd.read_csv(OUT + "/haccp_전처리.csv", dtype=str)
    h["인증일"] = pd.to_datetime(h["인증일"], errors="coerce")
    h = h.dropna(subset=["인증일"])
    h["연"] = h["인증일"].dt.year
    h["부문"] = "축산물 가공·유통"
    h.loc[h["식품구분"] == "식품", "부문"] = "식품"
    h.loc[(h["식품구분"] == "축산물") & (h["업태"] == "가축사육업"), "부문"] = "축산물 농장"
    for c in 농장키:
        h[c] = h[c].fillna("").str.strip()
    f = h.sort_values("인증일").drop_duplicates(["부문"] + 농장키)
    t = (f.pivot_table(index="연", columns="부문", values="업체명", aggfunc="count")
         .fillna(0).astype(int).reindex(columns=["식품", "축산물 가공·유통", "축산물 농장"]))
    t["합계"] = t.sum(axis=1)
    t = t.loc[2015:2025]
    print()
    print("=" * 78)
    print(" 부문별 신규 인증 업소 — 인증원 전체가 막혔나, 농장만 막혔나")
    print("=" * 78)
    print(t.to_string())
    t.to_csv(OUT + "/부문별_연도별신규.csv", encoding="utf-8-sig")
    print()
    print(" 합계 {:,} -> {:,}  (처리 역량은 두 배 넘게 커졌다)".format(
        t.loc[2015, "합계"], t.loc[2025, "합계"]))
    print(" 농장 {:,} -> {:,}  (농장만 줄었다)".format(
        t.loc[2015, "축산물 농장"], t.loc[2025, "축산물 농장"]))
    print(" 2019년 이후 농장 연 {:,}~{:,}개".format(
        t.loc[2019:, "축산물 농장"].min(), t.loc[2019:, "축산물 농장"].max()))
    return t


def 그림_신규추이(t):
    """기획서 [그림 1] 부문별 HACCP 신규 인증 추이. 기획서에 넣는 크기(1482×722)로 그린다."""
    for _f in ["Malgun Gothic", "NanumGothic", "AppleGothic"]:
        if any(_f in x.name for x in font_manager.fontManager.ttflist):
            plt.rcParams["font.family"] = _f
            break
    plt.rcParams["axes.unicode_minus"] = False
    색 = {"식품": "#a9b1ad", "축산물 가공·유통": "#5e6863", "축산물 농장": "#b8411a"}
    글자, 격자 = "#4c5752", "#dee4e0"

    fig, ax = plt.subplots(figsize=(14.82, 7.22), dpi=100)
    x = list(t.index)
    for 부문 in ["식품", "축산물 가공·유통", "축산물 농장"]:
        농장 = 부문 == "축산물 농장"
        ax.plot(x, t[부문], color=색[부문], lw=5 if 농장 else 3.2, marker="o",
                ms=10 if 농장 else 9, zorder=3 if 농장 else 2)
    끝 = x[-1]
    가공, 식품 = t.loc[끝, "축산물 가공·유통"], t.loc[끝, "식품"]
    위, 아래 = max(가공, 식품), min(가공, 식품)
    if 위 - 아래 < 110:          # 라벨이 겹치지 않게 벌린다
        위, 아래 = (위 + 아래) / 2 + 55, (위 + 아래) / 2 - 55
    ax.text(끝 + .25, 위 if 가공 >= 식품 else 아래, "축산물 가공·유통", color=색["축산물 가공·유통"],
            fontsize=19, va="center")
    ax.text(끝 + .25, 아래 if 가공 >= 식품 else 위, "식품", color=색["축산물 가공·유통"], fontsize=19,
            va="center")
    ax.text(끝 + .25, t.loc[끝, "축산물 농장"], "축산물 농장", color=색["축산물 농장"],
            fontsize=19, va="center", fontweight="bold")
    for 연 in (x[0], 끝):
        v = int(t.loc[연, "축산물 농장"])
        ax.text(연 + (.3 if 연 == x[0] else -.2), v + 70, "{:,}".format(v),
                color=색["축산물 농장"], fontsize=17, ha="center", va="bottom")

    ax.set_xlim(x[0] - .4, 끝 + 2.3)
    ax.set_xticks(x[::2])
    ax.set_ylim(100, ax.get_ylim()[1])
    ax.grid(axis="y", color=격자, lw=1.2)
    ax.set_axisbelow(True)
    ax.tick_params(colors=글자, labelsize=18, length=0, pad=10)
    for sp in ["top", "right", "left"]:
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color(격자)
    fig.subplots_adjust(left=.07, right=.97, top=.97, bottom=.1)
    os.makedirs(FIG, exist_ok=True)
    fig.savefig(os.path.join(FIG, "그림_신규추이.png"), facecolor="white")
    plt.close(fig)
    print("[저장] output/기획서_그림/그림_신규추이.png  (기획서 [그림 1])")


if __name__ == "__main__":
    f = load()
    묶음(f)
    추세(f)
    누적 = f.groupby("군").size().reindex(순서)
    print()
    print("=" * 78)
    print(" 결론")
    print("=" * 78)
    print(" 인증받은 농가 수:  " + " · ".join(
        "{} {:,}".format(g, 누적[g]) for g in 순서))
    print(" 분자는 축종 사이에 큰 차이가 없다. 차이는 대부분 분모(산업 구조)에 있다.")
    print(" 닭은 전국 농가가 3,123곳뿐이고 한우는 72,389곳이다. 23배다.")
    print()
    print(" 한우의 {:.1f}%는 인증을 덜 받아서가 아니다.".format(누적["한우"] / 72389 * 100))
    print(" 비슷한 수를 받았는데 나눌 모집단이 23배 크기 때문이다.")
    print(" 그래서 '농가를 더 설득하자'가 아니라 '단위를 바꾸자'가 답이 된다.")
    그림_신규추이(부문별())
