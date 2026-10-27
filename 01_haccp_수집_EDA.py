# -*- coding: utf-8 -*-
"""
2026 한국식품안전관리인증원 공공데이터 분석·활용 경진대회
01. HACCP 인증업체정보 수집 및 탐색적 데이터 분석(EDA)

목적: HACCP 인증 유지/이탈 분석의 데이터 타당성 검증
데이터: 공공데이터포털 한국식품안전관리인증원_HACCP 인증업체정보 (data.go.kr/data/15033311)
"""

import os
import time
import requests
import numpy as np
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

# ─────────────────────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────────────────────
SERVICE_KEY = os.environ.get(
    "HACCP_API_KEY",
    "",
)  # URL 인코딩된 일반 인증키
BASE_URL = "https://apis.data.go.kr/B553748/CertCompanyListService2/getCertCompanyListService2"

PAGE_SIZE = 100          # API 상한: 100건 초과 시 resultCode 13 오류
BASE_DATE = pd.Timestamp("2026-09-08")   # 분석 기준일(스냅샷 수집일)
RAW_PATH = "data/haccp_인증업체_raw.csv"

pd.set_option("display.width", 200)
pd.set_option("display.unicode.east_asian_width", True)


# ─────────────────────────────────────────────────────────────
# 1. 데이터 수집
# ─────────────────────────────────────────────────────────────
def fetch_page(session, page, retries=4):
    """단일 페이지 조회. 일시적 오류는 재시도한다."""
    url = f"{BASE_URL}?serviceKey={SERVICE_KEY}&pageNo={page}&numOfRows={PAGE_SIZE}&returnType=json"
    for _ in range(retries):
        try:
            body = session.get(url, timeout=60).json().get("body")
            if body:
                return [row["item"] for row in (body.get("items") or [])]
        except Exception:
            pass
        time.sleep(1.5)
    print(f"  [경고] page {page} 수집 실패")
    return []


def collect_all(use_cache=True):
    """전체 인증업체 목록을 수집한다."""
    if use_cache and os.path.exists(RAW_PATH):
        print(f"[수집] 캐시 사용: {RAW_PATH}")
        return pd.read_csv(RAW_PATH, dtype=str)

    session = requests.Session()
    total = int(
        session.get(f"{BASE_URL}?serviceKey={SERVICE_KEY}&pageNo=1&numOfRows=1&returnType=json",
                    timeout=30).json()["body"]["totalCount"]
    )
    pages = range(1, total // PAGE_SIZE + 2)
    print(f"[수집] 전체 {total:,}건 / {len(pages)}페이지")

    with ThreadPoolExecutor(6) as ex:
        chunks = list(ex.map(lambda p: fetch_page(session, p), pages))

    df = pd.DataFrame([r for c in chunks for r in c]).drop_duplicates()
    os.makedirs("data", exist_ok=True)
    df.to_csv(RAW_PATH, index=False, encoding="utf-8-sig")
    print(f"[수집] 완료: {df.shape[0]:,}행 × {df.shape[1]}열 → {RAW_PATH}")
    return df


# ─────────────────────────────────────────────────────────────
# 2. 전처리 및 파생변수
# ─────────────────────────────────────────────────────────────
def preprocess(df):
    """
    핵심 발견:
      issuedate    = 최초 인증일 (갱신해도 변하지 않음)
      issueenddate = 현재 유효기간 만료일 (갱신 시마다 3년씩 연장)
    → (만료일 - 최초인증일)이 3년의 배수로 누적되므로,
       이를 3으로 나누면 '인증 갱신 주기 수'를 복원할 수 있다.
    """
    df = df.copy()
    df["인증일"] = pd.to_datetime(df["issuedate"], errors="coerce")
    df["인증종료일"] = pd.to_datetime(df["issueenddate"], errors="coerce")

    # 사료 업종은 유효기간이 '2500-12-31'(사실상 무기한)로 기록되어 주기 계산에서 제외
    df["무기한여부"] = df["issueenddate"] == "2500-12-31"

    df["유지기간_년"] = (df["인증종료일"] - df["인증일"]).dt.days / 365.25
    df.loc[df["무기한여부"], "유지기간_년"] = np.nan

    df["갱신주기수"] = (df["유지기간_년"] / 3).round()          # 1 = 최초 인증만(갱신 0회)
    df["갱신횟수"] = (df["갱신주기수"] - 1).clip(lower=0)
    df["인증경과_년"] = (BASE_DATE - df["인증일"]).dt.days / 365.25
    df["만료까지_일"] = (df["인증종료일"] - BASE_DATE).dt.days
    df["최초인증연도"] = df["인증일"].dt.year

    df = df.rename(columns={
        "appointno": "인증번호", "company": "업체명", "productGb": "식품구분",
        "companykindNm": "업종", "businesstypeNm": "업태", "businessitemNm": "품목",
        "area1": "시도", "area2": "시군구",
    })
    return df


# ─────────────────────────────────────────────────────────────
# 3. EDA
# ─────────────────────────────────────────────────────────────
def eda(df):
    def head(title):
        print("\n" + "=" * 72); print(f" {title}"); print("=" * 72)

    head("1. 기본 구조")
    print(f"행×열: {df.shape[0]:,} × {df.shape[1]}")
    print(f"고유 인증번호: {df['인증번호'].nunique():,}건")
    print(f"고유 업체명  : {df['업체명'].nunique():,}개")
    print(f"업체당 평균 인증 수: {len(df)/df['업체명'].nunique():.2f}건")

    head("2. 결측치")
    miss = df[["인증번호", "업체명", "시도", "시군구", "인증일", "인증종료일"]].isna().sum()
    print(miss[miss >= 0].to_string())

    head("3. ★ 타당성 검증 — 이탈(만료) 업체가 데이터에 존재하는가")
    expired = (df["인증종료일"] < BASE_DATE).sum()
    print(f"인증종료일 < 기준일({BASE_DATE.date()}) : {expired:,}건")
    print(f"인증종료일 >= 기준일                  : {(df['인증종료일'] >= BASE_DATE).sum():,}건")
    print(f"인증종료일 최솟값                     : {df['인증종료일'].min().date()}")
    print("→ 만료 건이 0이면 본 API는 '현재 유효 인증'만 담은 스냅샷이며,")
    print("  이탈(포기) 업체를 직접 관측할 수 없다.")

    head("4. ★ 갱신주기 복원 가설 검증")
    v = df.dropna(subset=["갱신주기수"])
    print("갱신주기수별 실제 유지기간(년) — 평균이 3의 배수에 붙고 표준편차가 작으면 가설 성립")
    print(v.groupby("갱신주기수")["유지기간_년"].agg(건수="size", 평균="mean", 표준편차="std").round(3).to_string())

    head("5. 갱신횟수 분포")
    print(v["갱신횟수"].value_counts().sort_index().to_string())
    print(f"\n평균 갱신횟수: {v['갱신횟수'].mean():.2f}회")

    head("6. 최초인증 연도별 현재 잔존 건수 (코호트 생존)")
    print(df.groupby("최초인증연도").size().loc[2005:].to_string())
    print("→ 연도별 신규 인증 건수(인증원 통계연보)와 대조하면 코호트 생존율 산출 가능")

    head("7. 식품/축산물 구분")
    print(df["식품구분"].value_counts().to_string())

    head("8. 업종별 분포 및 평균 갱신횟수")
    print(df.groupby("업종").agg(건수=("인증번호", "size"),
                                평균갱신횟수=("갱신횟수", "mean"),
                                평균유지기간=("유지기간_년", "mean")).round(2).to_string())

    head("9. 업태 상위 15")
    print(df["업태"].value_counts().head(15).to_string())

    head("10. 시도별 분포 및 평균 갱신횟수")
    print(df.groupby("시도").agg(건수=("인증번호", "size"),
                                평균갱신횟수=("갱신횟수", "mean")).round(2)
            .sort_values("건수", ascending=False).to_string())

    head("11. 다수 인증 보유 업체 상위 10")
    print(df["업체명"].value_counts().head(10).to_string())

    head("12. 만료 임박 업체 (향후 90일 내)")
    soon = df[(df["만료까지_일"] >= 0) & (df["만료까지_일"] <= 90)]
    print(f"90일 내 만료 예정: {len(soon):,}건")
    print(soon.groupby("업종").size().to_string())


# ─────────────────────────────────────────────────────────────
def main():
    raw = collect_all()
    df = preprocess(raw)
    os.makedirs("output", exist_ok=True)
    df.to_csv("output/haccp_전처리.csv", index=False, encoding="utf-8-sig")
    eda(df)
    print("\n[완료] 전처리 결과 → output/haccp_전처리.csv")


if __name__ == "__main__":
    main()
