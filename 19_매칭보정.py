# -*- coding: utf-8 -*-
"""
19. 매칭 오류 진단과 보정 — 가장 강했던 변수가 가짜였다

【무엇을 발견했는가】
18번에서 '사업장 2개 이상' 변수가 압도적으로 강했다.
  사업장 1개 11.8% → 3개 29.1% → 4~5개 54.8% → 6~10개 73.6% → 11개 이상 100%
그런데 이 값들은 **동명이인 매칭의 산물**이었다.

  · HACCP 업소와 인허가 업소를 '정규화한 업체명'만으로 붙였다.
  · '우리식품'은 화성·영양·군산·영광에 각각 있다. 서로 다른 회사다.
    그런데 넷 중 하나라도 인증을 받으면 넷 모두 '인증받음'으로 찍힌다.
  · 동시에 그 넷은 '사업장 4개 보유'로도 찍힌다.
  → 같은 이름이 많을수록 (a) 다사업장 점수가 오르고 (b) 인증으로 찍힐 확률도 오른다.
    두 변수가 기계적으로 연결돼 있었다.

【귀무모형으로 확인】
이름이 n개 겹칠 때, 각자 독립적으로 확률 p로 인증받는다면
그 그룹의 누구 하나라도 인증일 확률은 1-(1-p)^n 이다.
실제 관측값이 이 곡선과 겹치면, 다사업장 효과는 전부 가공물이다.

  n=2  관측 25.09%  귀무 26.29%   (비 0.95)
  n=3  관측 29.05%  귀무 36.72%   (비 0.79)
  n=4~5 관측 54.76% 귀무 48.15%   (비 1.14)
  n=6~10 관측 73.60% 귀무 66.21%  (비 1.11)
겹친다. 그리고 '다사업장' 그룹의 78.5%는 서로 다른 시도에 흩어져 있다.
전국 체인이 아니라 그냥 흔한 이름이다.

【보정】
매칭 키를 '업체명'에서 '업체명 + 시도 + 시군구'로 바꾼다.
이러면 다른 지역의 동명 업체가 서로의 인증을 물려받지 않는다.
같은 시군구 안의 동명 충돌은 남지만, 대상이 1/200 수준으로 줄어든다.

【대가】
엄격 매칭은 인증률을 과소 추정한다. HACCP 등록 주소(공장)와
인허가 주소가 다른 업소를 놓치기 때문이다. 그래서 결과는 구간으로 말한다.
  느슨(이름만)  = 인증률 상한
  엄격(이름+지역) = 인증률 하한
"""

import os
import re
import sys
import importlib.util
import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import GradientBoostingClassifier

OUT = "output"
SRC = "data/식품제조가공업_인허가.csv"
BASE = pd.Timestamp("2026-09-08")
TRAIN_T0 = "2019-01-01"
TEST_T0S = ["2021-01-01", "2022-01-01"]

for _f in ["Malgun Gothic", "NanumGothic", "AppleGothic"]:
    if any(_f in f.name for f in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = _f
        break
plt.rcParams["axes.unicode_minus"] = False

_s = importlib.util.spec_from_file_location("v17", "17_명단검증.py")
v17 = importlib.util.module_from_spec(_s)
_s.loader.exec_module(v17)
auc, 정규화 = v17.auc, v17.정규화


# 두 데이터의 시도 표기가 다르다. 인허가 쪽 표기로 통일한다.
시도보정 = {"전라남도": "전남광주통합특별시", "광주광역시": "전남광주통합특별시",
         "서울시": "서울특별시", "세종시": "세종특별자치시", "세종특별시": "세종특별자치시",
         "충남": "충청남도", "충북": "충청북도", "경남": "경상남도", "경북": "경상북도"}


def 지역키(시도, 시군구):
    """'청주시 흥덕구' → '청주시'. 시도와 붙여 하나의 지역 키로.
    세종처럼 시군구가 없는 곳은 빈 문자열로 맞춘다(양쪽 모두 결측이 되도록)."""
    t = str(시군구)
    t = "" if t.strip() in ("", "nan", "None", "NaN") else t.split()[0]
    s = str(시도).strip()
    s = 시도보정.get(s, s)
    return s + "|" + t


def load():
    lic = pd.read_csv(SRC, dtype=str)
    lic["인허가일자"] = pd.to_datetime(lic["인허가일자"], errors="coerce")
    lic["폐업일자"] = pd.to_datetime(lic["폐업일자"], errors="coerce")
    lic["키"] = lic["사업장명"].map(정규화)
    부 = lic["소재지전체주소"].astype(str).str.split()
    lic["시도"] = 부.str[0]
    lic["시군구단"] = 부.str[1].fillna("")
    lic["지역"] = [지역키(a, b) for a, b in zip(lic["시도"], lic["시군구단"])]
    lic["엄격키"] = lic["키"] + "@" + lic["지역"]

    h = pd.read_csv(OUT + "/haccp_전처리.csv", dtype=str)
    h = h[(h["식품구분"] == "식품") & h["업태"].str.contains("식품제조가공업", na=False)].copy()
    h["키"] = h["업체명"].map(정규화)
    h["지역"] = [지역키(a, b) for a, b in zip(h["시도"], h["시군구"])]
    h["엄격키"] = h["키"] + "@" + h["지역"]
    h["인증일"] = pd.to_datetime(h["인증일"], errors="coerce")

    느슨 = h.groupby("키")["인증일"].min()
    엄격 = h.groupby("엄격키")["인증일"].min()
    print("[로드] 인허가 {:,}건 / HACCP 업소 느슨 {:,} · 엄격 {:,}".format(
        len(lic), len(느슨), len(엄격)))
    return lic, 느슨, 엄격


# ────────────────────────────────────────────────────────────────
def 코호트(lic, 최초, T0, 모드):
    """모드='느슨'이면 이름만, '엄격'이면 이름+지역으로 인증 여부를 붙인다."""
    T0 = pd.Timestamp(T0)
    col = "키" if 모드 == "느슨" else "엄격키"

    영업 = lic[(lic["인허가일자"] <= T0) &
              (lic["폐업일자"].isna() | (lic["폐업일자"] > T0))].copy()
    영업["인증일"] = 영업[col].map(최초)
    영업["T0인증"] = 영업["인증일"].notna() & (영업["인증일"] <= T0)

    g = 영업.groupby("지역")["T0인증"].agg(["size", "sum"])
    g["지역인증률"] = g["sum"] / g["size"]
    g.loc[g["size"] < 30, "지역인증률"] = np.nan

    전국 = 영업.groupby("키").size()                    # 전국 동명 수 (오염된 변수)
    동일지역 = 영업.groupby("엄격키").size()              # 같은 시군구 내 동명 수

    d = 영업[~영업["T0인증"]].copy()
    d["업력"] = (T0 - d["인허가일자"]).dt.days / 365.25
    d["동명수_전국"] = d["키"].map(전국)
    d["동명수_지역"] = d["엄격키"].map(동일지역)
    d["지역인증률"] = d["지역"].map(g["지역인증률"])
    d["지역업소수"] = d["지역"].map(g["size"])
    d["지역인증률"] = d["지역인증률"].fillna(d["지역인증률"].median())
    d["y"] = d["인증일"].notna() & (d["인증일"] > T0) & (d["인증일"] <= BASE)
    return d, g


def 귀무진단(d, 라벨):
    p = d["y"].mean()
    rows = []
    for lo, hi in [(1, 1), (2, 2), (3, 3), (4, 5), (6, 10), (11, 999)]:
        gsub = d[(d["동명수_전국"] >= lo) & (d["동명수_전국"] <= hi)]
        if len(gsub) < 10:
            continue
        nbar = gsub["동명수_전국"].mean()
        null = 1 - (1 - p) ** nbar
        rows.append(dict(매칭=라벨, 동명수="{}~{}".format(lo, hi), 건수=len(gsub),
                         평균n=round(nbar, 1),
                         관측=round(gsub["y"].mean() * 100, 2),
                         귀무=round(null * 100, 2),
                         비=round(gsub["y"].mean() / null, 2)))
    return pd.DataFrame(rows)


# ────────────────────────────────────────────────────────────────
def 설계행렬(d, 컷):
    X = pd.DataFrame(index=d.index)
    X["지역인증률"] = d["지역인증률"]
    X["log지역업소수"] = np.log1p(d["지역업소수"])
    X["신설1년"] = (d["업력"] <= 1).astype(float)
    X["업력"] = np.clip(d["업력"], 0, 40)
    X["업력제곱"] = np.clip(d["업력"], 0, 40) ** 2
    X["동일지역동명"] = (d["동명수_지역"] >= 2).astype(float)
    return X


def 점수표v3(d, 컷):
    """오염된 전국 동명수를 빼고, 지역 신호와 업력 U자만으로 다시 짠다."""
    s = pd.Series(0.0, index=d.index)
    q = pd.cut(d["지역인증률"], 컷, labels=False, include_lowest=True)
    s += q.fillna(2).map({0: 0, 1: 15, 2: 30, 3: 45, 4: 60}).astype(float)   # 최대 60
    dn = pd.qcut(d["지역업소수"].rank(method="first"), 4, labels=False)
    s += pd.Series(dn, index=d.index).map({0: 0, 1: 5, 2: 10, 3: 20}).astype(float)  # 최대 20
    a = d["업력"]
    s += np.select([a <= 1, a > 18, a > 12], [20, 12, 6], default=0)          # 최대 20
    return s


def 성능(score, y, 라벨, 상위=0.10):
    a = auc(score, y)
    base = y.mean()
    o = np.argsort(-np.asarray(score, dtype=float))
    k = max(1, int(len(y) * 상위))
    top = np.asarray(y)[o][:k].mean()
    return dict(모형=라벨, AUC=round(a, 4), 기저=round(base * 100, 2),
                상위10퍼전환율=round(top * 100, 2), 리프트=round(top / base, 2))


# ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    lic, 느슨, 엄격 = load()

    # ── (1) 진단 ────────────────────────────────────────────────
    진단 = []
    for 모드, 최초 in [("느슨(이름만)", 느슨), ("엄격(이름+지역)", 엄격)]:
        d, _ = 코호트(lic, 최초, "2021-01-01", "느슨" if "느슨" in 모드 else "엄격")
        진단.append(귀무진단(d, 모드))
    진단 = pd.concat(진단, ignore_index=True)
    print("\n" + "=" * 78)
    print("[진단] 동명 그룹 크기별 전환율 — 관측이 귀무(1-(1-p)^n)와 겹치면 가공물이다")
    print(진단.to_string(index=False))
    진단.to_csv(OUT + "/진단_동명매칭.csv", index=False, encoding="utf-8-sig")

    # ── (2) 인증률 구간 ─────────────────────────────────────────
    행 = []
    for 모드, 최초, m in [("느슨(상한)", 느슨, "느슨"), ("엄격(하한)", 엄격, "엄격")]:
        영업 = lic[lic["폐업일자"].isna() | (lic["폐업일자"] > BASE)]
        영업 = 영업[영업["인허가일자"] <= BASE]
        col = "키" if m == "느슨" else "엄격키"
        인증 = 영업[col].map(최초).notna()
        행.append(dict(매칭=모드, 영업중=len(영업), 인증=int(인증.sum()),
                      미인증=int((~인증).sum()),
                      인증률=round(인증.mean() * 100, 2)))
    C = pd.DataFrame(행)
    print("\n[인증률 구간] 기준일 {}".format(BASE.date()))
    print(C.to_string(index=False))
    C.to_csv(OUT + "/인증률_구간.csv", index=False, encoding="utf-8-sig")

    # ── (3) 엄격 매칭으로 모형 재검증 ────────────────────────────
    tr, _ = 코호트(lic, 엄격, TRAIN_T0, "엄격")
    컷 = list(np.unique(np.quantile(tr["지역인증률"], [0, .2, .4, .6, .8, 1.0])))
    컷[0], 컷[-1] = -1e9, 1e9
    중앙 = tr["지역인증률"].median()

    Xtr = 설계행렬(tr, 컷)
    lr = LogisticRegression(max_iter=2000).fit(Xtr, tr["y"])
    gb = GradientBoostingClassifier(n_estimators=150, max_depth=3,
                                    learning_rate=0.07, random_state=0).fit(Xtr, tr["y"])
    print("\n[로지스틱 계수 — 엄격 매칭]")
    for c, w in sorted(zip(Xtr.columns, lr.coef_[0]), key=lambda t: -abs(t[1])):
        print("   {:<12} {:+.4f}".format(c, w))

    표, 보관 = [], {}
    for t0 in [TRAIN_T0] + TEST_T0S:
        te = tr if t0 == TRAIN_T0 else 코호트(lic, 엄격, t0, "엄격")[0]
        Xte = 설계행렬(te, 컷)
        cand = {
            "규칙 v1 (업력40/전국동명35/인프라25)":
                (te["업력"] >= 5) * 40 + (te["동명수_전국"] >= 2) * 35
                + (te["지역인증률"] >= 중앙) * 25,
            "규칙 v3 (지역인프라60/밀도20/업력20)": 점수표v3(te, 컷),
            "로지스틱 회귀": lr.predict_proba(Xte)[:, 1],
            "그래디언트 부스팅": gb.predict_proba(Xte)[:, 1],
        }
        for nm, sc in cand.items():
            r = 성능(sc, te["y"], nm)
            r["코호트"] = ("학습 " if t0 == TRAIN_T0 else "검증 ") + t0
            r["대상"] = len(te)
            표.append(r)
        보관[t0] = (te, cand)

    R = pd.DataFrame(표)[["코호트", "대상", "모형", "기저", "AUC", "상위10퍼전환율", "리프트"]]
    print("\n" + "=" * 78)
    print("[엄격 매칭 재검증] 오염 변수를 뺀 뒤에도 신호가 남는가")
    print(R.to_string(index=False))
    R.to_csv(OUT + "/검증_모형비교_엄격.csv", index=False, encoding="utf-8-sig")

    # ── (4) 현재 기준 명단 v3 ────────────────────────────────────
    현재, 지역표 = 코호트(lic, 엄격, BASE, "엄격")
    현재 = 현재[현재["인증일"].isna()].copy()
    현재["유입점수_v3"] = 점수표v3(현재, 컷)
    현재["예측확률"] = lr.predict_proba(설계행렬(현재, 컷))[:, 1]
    현재["등급"] = pd.qcut(현재["예측확률"].rank(method="first"), [0, .5, .8, .95, .99, 1.0],
                       labels=["낮음", "보통", "관심", "높음", "최우선"])
    print("\n[{} 기준 미인증(엄격) {:,}개 — 등급 분포]".format(BASE.date(), len(현재)))
    print(현재.groupby("등급", observed=True).agg(
        건수=("키", "size"), 평균업력=("업력", "mean"),
        평균지역인증률=("지역인증률", "mean"),
        평균예측확률=("예측확률", "mean")).round(3).to_string())

    상위 = 현재.sort_values("예측확률", ascending=False)
    cols = [c for c in ["등급", "예측확률", "유입점수_v3", "사업장명", "시도", "시군구단",
                        "소재지전체주소", "인허가일자", "업력", "지역인증률"] if c in 상위.columns]
    상위[cols].to_csv(OUT + "/유입_우선명단_v3.csv", index=False, encoding="utf-8-sig")
    상위[상위["등급"] == "최우선"][cols].to_csv(
        OUT + "/유입_최우선명단_v3.csv", index=False, encoding="utf-8-sig")
    print("\n[저장] 유입_우선명단_v3.csv / 유입_최우선명단_v3.csv")

    # ── (5) 그림 ────────────────────────────────────────────────
    fig, ax = plt.subplots(1, 3, figsize=(16.5, 4.9))

    for 모드, col, mk in [("느슨(이름만)", "#c53030", "o"), ("엄격(이름+지역)", "#2b6cb0", "s")]:
        sub = 진단[진단["매칭"] == 모드]
        ax[0].plot(sub["동명수"], sub["관측"], mk + "-", color=col, lw=2, label="관측 · " + 모드)
        ax[0].plot(sub["동명수"], sub["귀무"], mk + "--", color=col, lw=1.2, alpha=.55,
                   label="귀무 · " + 모드)
    ax[0].set_title("동명 그룹이 클수록 인증으로 찍힌다\n느슨 매칭에서는 관측이 귀무와 겹친다", fontsize=11)
    ax[0].set_xlabel("같은 이름의 전국 업소 수")
    ax[0].set_ylabel("인증 전환율 (%)")
    ax[0].legend(fontsize=8)
    ax[0].grid(alpha=.25)

    te2, cand2 = 보관["2021-01-01"]
    y = np.asarray(te2["y"])
    for nm, col in [("규칙 v1 (업력40/전국동명35/인프라25)", "#a0aec0"),
                    ("규칙 v3 (지역인프라60/밀도20/업력20)", "#2b6cb0"),
                    ("로지스틱 회귀", "#dd6b20")]:
        sc = np.asarray(cand2[nm], dtype=float)
        o = np.argsort(-sc)
        ax[1].plot(np.arange(1, len(y) + 1) / len(y) * 100,
                   np.cumsum(y[o]) / y.sum() * 100, lw=2, color=col,
                   label="{}  AUC {:.3f}".format(nm.split(" (")[0], auc(sc, y)))
    ax[1].plot([0, 100], [0, 100], ls="--", lw=1.2, color="#cbd5e0", label="무작위")
    ax[1].set_title("엄격 매칭 · 검증 코호트(2021) 이득곡선", fontsize=11)
    ax[1].set_xlabel("접촉한 업소 비율 (%)")
    ax[1].set_ylabel("확보한 신규 인증 비율 (%)")
    ax[1].legend(fontsize=8.5)
    ax[1].grid(alpha=.25)

    b = te2["y"].mean()
    te2 = te2.copy()
    te2["q"] = pd.qcut(np.asarray(cand2["로지스틱 회귀"]).argsort().argsort(), 10, labels=False)
    dec = te2.groupby("q")["y"].mean() * 100
    ax[2].bar(dec.index + 1, dec.values, color="#2b6cb0")
    ax[2].axhline(b * 100, ls="--", color="#c53030", lw=1.5, label="기저 {:.1f}%".format(b * 100))
    ax[2].set_title("예측확률 십분위별 실제 전환율", fontsize=11)
    ax[2].set_xlabel("십분위 (10 = 상위 10%)")
    ax[2].set_ylabel("전환율 (%)")
    ax[2].legend()

    fig.tight_layout()
    fig.savefig(OUT + "/그림32_매칭보정.png", dpi=150, bbox_inches="tight")
    print("[저장] " + OUT + "/그림32_매칭보정.png")
