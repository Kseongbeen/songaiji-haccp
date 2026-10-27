# -*- coding: utf-8 -*-
"""
17. 우선명단 점수모형의 후향(prospective) 검증

【왜 필요한가】
15_유입갭.py 가 만든 '유입점수'는 종업원 수를 못 구해서 업력·다사업장·지역인프라
세 가지 대리지표로 '받을 여력'을 쟀다. 그런데 이 대리지표가 실제로 인증 취득을
설명하는지는 확인된 적이 없다. 검증 없는 명단은 제출본에서 가장 먼저 무너진다.

【검증 설계 — 과거로 돌아가서 미래를 맞히기】
  기준시점 T0 = 2021-01-01 (2019·2020·2022 로도 반복)
  1) T0 시점에 '영업 중이면서 아직 HACCP 미인증'이던 식품제조가공업소를 모은다
  2) 그 시점의 정보만으로 유입점수를 계산한다 (업력·사업장수·지역인증률 모두 T0 기준)
  3) T0 이후 2026-09-08 까지 실제로 인증을 받았는가를 정답(y)으로 놓는다
  4) 점수가 y를 얼마나 맞히는지 AUC·리프트·구간별 실제 전환율로 잰다

이러면 미래 정보가 새지 않는다. 지역인증률도 T0 시점 값이라 누수가 없다.

【해석 기준】
  AUC 0.5       무작위. 명단이 의미 없다.
  AUC 0.60~0.65 약하지만 실재하는 신호. 대량 접촉 우선순위로는 쓸 만하다.
  상위 10% 리프트 1.5배 이상이면 "같은 인력으로 1.5배 많이 전환시킨다"고 말할 수 있다.

【선행】 14 → 15 실행되어 data/식품제조가공업_인허가.csv 가 있을 것
"""

import os
import re
import sys
import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

SRC = "data/식품제조가공업_인허가.csv"
OUT = "output"
BASE = pd.Timestamp("2026-09-08")
T0S = ["2019-01-01", "2020-01-01", "2021-01-01", "2022-01-01"]

for _f in ["Malgun Gothic", "NanumGothic", "AppleGothic"]:
    if any(_f in f.name for f in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = _f
        break
plt.rcParams["axes.unicode_minus"] = False


def 정규화(s):
    s = str(s)
    s = re.sub(r"\(주\)|주식회사|㈜|\(유\)|유한회사|\(합\)|합자회사|농업회사법인|"
               r"영어조합법인|영농조합법인|협동조합|사회적기업", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"[\s\-·.,]", "", s)
    return s


def 시군구(addr):
    p = str(addr).split()
    return " ".join(p[:2]) if len(p) >= 2 else ""


def load():
    if not os.path.exists(SRC):
        print("[없음] " + SRC + " — 먼저 14_식품안전나라_수집.py 실행")
        sys.exit(0)
    lic = pd.read_csv(SRC, dtype=str)
    lic["인허가일자"] = pd.to_datetime(lic["인허가일자"], errors="coerce")
    lic["폐업일자"] = pd.to_datetime(lic["폐업일자"], errors="coerce")
    lic["키"] = lic["사업장명"].map(정규화)
    lic["시군구"] = lic["소재지전체주소"].map(시군구)

    h = pd.read_csv(OUT + "/haccp_전처리.csv", dtype=str)
    h = h[(h["식품구분"] == "식품") & h["업태"].str.contains("식품제조가공업", na=False)].copy()
    h["키"] = h["업체명"].map(정규화)
    h["인증일"] = pd.to_datetime(h["인증일"], errors="coerce")
    최초 = h.groupby("키")["인증일"].min()
    print("[로드] 인허가 {:,}건 / HACCP 식품제조가공 업소 {:,}개".format(len(lic), len(최초)))
    print("       최초인증일 분포: {} ~ {}".format(최초.min().date(), 최초.max().date()))
    return lic, 최초


def 코호트(lic, 최초, T0):
    """T0 시점 영업 중 · 미인증 업소를 모으고 T0 시점 특징을 만든다."""
    T0 = pd.Timestamp(T0)
    영업 = lic[(lic["인허가일자"] <= T0) &
              (lic["폐업일자"].isna() | (lic["폐업일자"] > T0))].copy()
    영업["인증일"] = 영업["키"].map(최초)
    영업["T0인증"] = 영업["인증일"].notna() & (영업["인증일"] <= T0)

    지역 = 영업.groupby("시군구")["T0인증"].agg(["size", "sum"])
    지역 = 지역[지역["size"] >= 30]
    지역["율"] = 지역["sum"] / 지역["size"]
    중앙 = 지역["율"].median()
    밀집 = set(지역[지역["율"] >= 중앙].index)

    보유 = 영업.groupby("키").size()

    d = 영업[~영업["T0인증"]].copy()
    d["업력"] = (T0 - d["인허가일자"]).dt.days / 365.25
    d["사업장수"] = d["키"].map(보유)
    d["업력5년"] = d["업력"] >= 5
    d["다사업장"] = d["사업장수"] >= 2
    d["인프라"] = d["시군구"].isin(밀집)
    d["유입점수"] = d["업력5년"] * 40 + d["다사업장"] * 35 + d["인프라"] * 25

    d["y"] = d["인증일"].notna() & (d["인증일"] > T0) & (d["인증일"] <= BASE)
    return d, T0, 중앙


def auc(score, y):
    """순위 기반 AUC (동점은 평균순위). scipy 없이 계산."""
    y = np.asarray(y, dtype=bool)
    s = np.asarray(score, dtype=float)
    n1, n0 = y.sum(), (~y).sum()
    if n1 == 0 or n0 == 0:
        return np.nan
    r = pd.Series(s).rank().values
    return (r[y].sum() - n1 * (n1 + 1) / 2) / (n1 * n0)


def 평가(d, T0, label, 상세=True):
    base = d["y"].mean()
    a = auc(d["유입점수"], d["y"])

    tab = (d.groupby("유입점수")
             .agg(대상=("y", "size"), 전환=("y", "sum"))
             .assign(전환율=lambda t: t["전환"] / t["대상"]))
    tab["리프트"] = tab["전환율"] / base

    dd = d.sort_values(["유입점수", "업력"], ascending=False)
    k = max(1, int(len(dd) * 0.10))
    top10 = dd.head(k)["y"].mean()

    최우선 = d[d["유입점수"] >= 90]
    row = dict(기준시점=str(T0.date()), 대상=len(d), 전환=int(d["y"].sum()),
               기저전환율=round(base * 100, 2), AUC=round(a, 4),
               상위10퍼_전환율=round(top10 * 100, 2), 상위10퍼_리프트=round(top10 / base, 2),
               최우선_건수=len(최우선),
               최우선_전환율=round(최우선["y"].mean() * 100, 2) if len(최우선) else np.nan,
               최우선_리프트=round(최우선["y"].mean() / base, 2) if len(최우선) else np.nan)

    if 상세:
        print("\n" + "=" * 66)
        print("[{}] T0={} → {}".format(label, T0.date(), BASE.date()))
        print("  대상(T0 영업·미인증) {:,}개 중 {:,}개가 이후 인증".format(len(d), int(d["y"].sum())))
        print("  기저 전환율 {:.2f}%   AUC {:.4f}".format(base * 100, a))
        print("  상위 10% 전환율 {:.2f}% (리프트 {:.2f}배)".format(top10 * 100, top10 / base))
        print("\n  [점수 구간별 실제 전환율]")
        t = tab.copy()
        t["전환율"] = (t["전환율"] * 100).round(2)
        t["리프트"] = t["리프트"].round(2)
        print(t.to_string())
    return row, tab


def 단변량(d):
    """세 대리지표를 하나씩 떼어서 각각의 판별력을 본다."""
    base = d["y"].mean()
    rows = []
    for c, nm in [("업력5년", "업력 5년 이상"), ("다사업장", "사업장 2개 이상"), ("인프라", "인증 밀집 지역")]:
        g = d.groupby(c)["y"].agg(["size", "mean"])
        참 = g.loc[True, "mean"] if True in g.index else np.nan
        거짓 = g.loc[False, "mean"] if False in g.index else np.nan
        rows.append(dict(지표=nm,
                         해당=int(g.loc[True, "size"]) if True in g.index else 0,
                         해당_전환율=round(참 * 100, 2),
                         비해당_전환율=round(거짓 * 100, 2),
                         리프트=round(참 / base, 2),
                         단독AUC=round(auc(d[c].astype(float), d["y"]), 4)))
    t = pd.DataFrame(rows)
    print("\n  [대리지표 단독 성능]")
    print(t.to_string(index=False))
    return t


def 가중치탐색(d):
    """40/35/25 배점이 최적인지, 데이터가 말하는 배점은 얼마인지."""
    best, rows = None, []
    for w1 in range(0, 101, 5):
        for w2 in range(0, 101 - w1, 5):
            w3 = 100 - w1 - w2
            s = d["업력5년"] * w1 + d["다사업장"] * w2 + d["인프라"] * w3
            a = auc(s, d["y"])
            rows.append((w1, w2, w3, a))
            if best is None or a > best[3]:
                best = (w1, w2, w3, a)
    현재 = auc(d["업력5년"] * 40 + d["다사업장"] * 35 + d["인프라"] * 25, d["y"])
    print("\n  [배점 탐색] 현재 40/35/25 → AUC {:.4f}".format(현재))
    print("              최적 {}/{}/{} → AUC {:.4f}  (차이 {:+.4f})".format(
        best[0], best[1], best[2], best[3], best[3] - 현재))
    return pd.DataFrame(rows, columns=["업력", "다사업장", "인프라", "AUC"]), best, 현재


if __name__ == "__main__":
    lic, 최초 = load()

    요약들, 주tab, 주d, 주T0 = [], None, None, None
    for t0 in T0S:
        d, T0, 중앙 = 코호트(lic, 최초, t0)
        상세 = (t0 == "2021-01-01")
        row, tab = 평가(d, T0, "후향검증", 상세=상세)
        요약들.append(row)
        if 상세:
            uni = 단변량(d)
            grid, best, 현재 = 가중치탐색(d)
            주tab, 주d, 주T0 = tab, d, T0
            uni.to_csv(OUT + "/검증_대리지표.csv", index=False, encoding="utf-8-sig")
            grid.sort_values("AUC", ascending=False).head(30).to_csv(
                OUT + "/검증_배점탐색.csv", index=False, encoding="utf-8-sig")

    S = pd.DataFrame(요약들)
    print("\n" + "=" * 66)
    print("[기준시점을 바꿔도 같은 결론인가]")
    print(S.to_string(index=False))
    S.to_csv(OUT + "/검증_요약.csv", index=False, encoding="utf-8-sig")

    주요약 = [r for r in 요약들 if r["기준시점"] == "2021-01-01"][0]
    fig, ax = plt.subplots(1, 3, figsize=(16, 4.8))

    t = 주tab.reset_index()
    ax[0].bar(t["유입점수"].astype(str), t["전환율"] * 100, color="#2b6cb0")
    ax[0].axhline(주d["y"].mean() * 100, color="#c53030", ls="--", lw=1.5,
                  label="기저 {:.1f}%".format(주d["y"].mean() * 100))
    ax[0].set_title("유입점수별 실제 인증 전환율\n(T0=2021-01-01 이후 2026-09-08까지)", fontsize=11)
    ax[0].set_xlabel("유입점수")
    ax[0].set_ylabel("전환율 (%)")
    ax[0].legend()

    dd = 주d.sort_values(["유입점수", "업력"], ascending=False)
    cum = dd["y"].cumsum() / dd["y"].sum()
    x = np.arange(1, len(dd) + 1) / len(dd)
    ax[1].plot(x * 100, cum * 100, color="#2b6cb0", lw=2, label="유입점수 순 접촉")
    ax[1].plot([0, 100], [0, 100], color="#a0aec0", ls="--", lw=1.2, label="무작위 접촉")
    ax[1].set_title("이득곡선 — AUC {:.3f}".format(주요약["AUC"]), fontsize=11)
    ax[1].set_xlabel("접촉한 업소 비율 (%)")
    ax[1].set_ylabel("확보한 신규 인증 비율 (%)")
    ax[1].legend()
    ax[1].grid(alpha=.25)

    ax[2].plot(S["기준시점"], S["AUC"], "o-", color="#2b6cb0", lw=2, label="AUC")
    ax[2].axhline(0.5, color="#a0aec0", ls="--", lw=1.2, label="무작위 (0.5)")
    ax[2].set_ylim(0.45, max(0.75, S["AUC"].max() + .05))
    ax[2].set_title("기준시점을 바꿔도 유지되는가", fontsize=11)
    ax[2].set_ylabel("AUC")
    ax[2].legend()
    ax[2].grid(alpha=.25)
    for _, r in S.iterrows():
        ax[2].annotate("{:.3f}".format(r["AUC"]), (r["기준시점"], r["AUC"]),
                       textcoords="offset points", xytext=(0, 8), ha="center", fontsize=9)

    fig.tight_layout()
    fig.savefig(OUT + "/그림30_명단검증.png", dpi=150, bbox_inches="tight")
    print("\n[저장] " + OUT + "/그림30_명단검증.png")
    print("[저장] 검증_요약.csv · 검증_대리지표.csv · 검증_배점탐색.csv")
