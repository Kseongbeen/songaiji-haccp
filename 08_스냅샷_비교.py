# -*- coding: utf-8 -*-
"""
08. 스냅샷 패널 구축 — 관측되지 않던 이탈을 '직접' 관측한다

【왜 필요한가】
01의 EDA에서 확인했듯 이 API는 '현재 유효한 인증'만 담은 스냅샷이다.
만료된 인증은 그 시점에 응답에서 사라진다. 그래서 이탈 업체를 볼 수 없었다.

뒤집어 말하면, 같은 API를 두 시점에 떠서 비교하면
    1차에는 있었는데 2차에는 없는 인증번호 = 그 사이 실제로 사라진 인증
이 된다. 개별 업체 단위로, 라벨이 붙은 진짜 이탈 사건이다.

03의 뺄셈 추정이 집단 크기만 주는 반면, 이 패널은 개별 관측치를 준다.
1차 스냅샷은 2026-09-08자로 이미 확보되어 있으므로, 오늘 2차를 뜨면 그날부터 관측이 시작된다.
관측 기간이 짧을수록 사건 수는 적지만, 만료 임박 코호트(90일 내 2,460건)가
실제로 갱신했는지 여부는 3개월이면 충분히 쌓인다.

【사용법】
    python 08_스냅샷_비교.py collect          # 오늘자 스냅샷 수집
    python 08_스냅샷_비교.py diff             # 가장 오래된 것과 최신 것을 비교
    python 08_스냅샷_비교.py diff 20260908 20261208
"""

import os
import sys
import glob
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

SERVICE_KEY = os.environ.get(
    "HACCP_API_KEY",
    "",
)
BASE_URL = "https://apis.data.go.kr/B553748/CertCompanyListService2/getCertCompanyListService2"
SNAP_DIR = "data/snapshots"
PAGE_SIZE = 100
KEY = "appointno"          # 인증번호 — 스냅샷 간 대조 키


# ─────────────────────────────────────────────────────────────
# 수집
# ─────────────────────────────────────────────────────────────
def fetch_page(session, page, retries=4):
    url = f"{BASE_URL}?serviceKey={SERVICE_KEY}&pageNo={page}&numOfRows={PAGE_SIZE}&returnType=json"
    for _ in range(retries):
        try:
            body = session.get(url, timeout=60).json().get("body")
            if body:
                return [row["item"] for row in (body.get("items") or [])]
        except Exception:
            pass
        time.sleep(1.5)
    print(f"  [경고] page {page} 실패")
    return []


def collect(tag=None):
    """오늘자(또는 지정 tag) 스냅샷을 뜬다. 41,069건 기준 약 410페이지."""
    os.makedirs(SNAP_DIR, exist_ok=True)
    tag = tag or pd.Timestamp.today().strftime("%Y%m%d")
    path = f"{SNAP_DIR}/haccp_{tag}.csv"
    if os.path.exists(path):
        print(f"[수집] 이미 존재: {path}")
        return path

    session = requests.Session()
    total = int(session.get(f"{BASE_URL}?serviceKey={SERVICE_KEY}&pageNo=1&numOfRows=1&returnType=json",
                            timeout=30).json()["body"]["totalCount"])
    pages = range(1, total // PAGE_SIZE + 2)
    print(f"[수집] {tag} — 전체 {total:,}건 / {len(pages)}페이지")

    with ThreadPoolExecutor(6) as ex:
        chunks = list(ex.map(lambda p: fetch_page(session, p), pages))

    df = pd.DataFrame([r for c in chunks for r in c]).drop_duplicates()
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"[수집] 완료 {len(df):,}건 → {path}")
    return path


def seed_from_raw():
    """01에서 받아 둔 2026-09-08 원본을 1차 스냅샷으로 등록한다."""
    src = "data/haccp_인증업체_raw.csv"
    dst = f"{SNAP_DIR}/haccp_20260908.csv"
    os.makedirs(SNAP_DIR, exist_ok=True)
    if os.path.exists(dst):
        return dst
    if not os.path.exists(src):
        print(f"[경고] 원본 없음: {src}")
        return None
    pd.read_csv(src, dtype=str).to_csv(dst, index=False, encoding="utf-8-sig")
    print(f"[등록] 1차 스냅샷 → {dst}")
    return dst



# ─────────────────────────────────────────────────────────────
# 영업일 보정
#
# 인증 등록·말소는 행정 처리라 주말·공휴일에는 일어나지 않는다.
# 실제로 2026-09-12(토)→09-13(일) 하루치는 이탈 3 · 신규 0 · 갱신 0 이었다.
# 관측 구간에 주말이 섞이면 달력일 연환산이 그만큼 낮게 나온다.
# 그래서 달력일과 영업일 두 가지를 함께 찍는다. 어느 쪽도 단정하지 않는다.
#
# 공휴일은 넣지 않았다. 연도별 공휴일 표를 외부에서 들여오면 재현이 깨진다.
# 주 5일만 보정해도 주말 효과의 대부분은 걷힌다.
# ─────────────────────────────────────────────────────────────
영업일수_연 = 261          # 365일 중 주말을 뺀 평균 영업일


def 영업일수(d0, d1):
    """d0 다음 날부터 d1 까지의 월~금 일수. 사건이 일어날 수 있었던 날의 수다."""
    rng = pd.date_range(pd.to_datetime(d0) + pd.Timedelta(days=1), pd.to_datetime(d1))
    return int((rng.dayofweek < 5).sum())


def series():
    """인접 스냅샷 쌍을 모두 대조해 시계열 표를 만든다.

    스냅샷이 쌓일수록 이 표가 길어진다. 하루치만 보고 단정하지 말라는 뜻으로
    구간별 값을 전부 남긴다."""
    files = snapshots()
    if len(files) < 2:
        print("[시계열] 스냅샷이 2개 이상이어야 한다.")
        return None

    rows = []
    for f0, f1 in zip(files[:-1], files[1:]):
        a = pd.read_csv(f0, dtype=str).drop_duplicates(subset=[KEY]).set_index(KEY)
        b = pd.read_csv(f1, dtype=str).drop_duplicates(subset=[KEY]).set_index(KEY)
        d0, d1 = os.path.basename(f0)[6:14], os.path.basename(f1)[6:14]
        공통 = a.index.intersection(b.index)
        이탈 = len(a.index.difference(b.index))
        신규 = len(b.index.difference(a.index))
        갱신 = int((a.loc[공통, "issueenddate"].values
                    != b.loc[공통, "issueenddate"].values).sum())
        days = (pd.to_datetime(d1) - pd.to_datetime(d0)).days
        bd = 영업일수(d0, d1)
        rows.append({
            "시작": d0, "끝": d1, "달력일": days, "영업일": bd,
            "시작_유효": len(a), "끝_유효": len(b),
            "이탈": 이탈, "신규": 신규, "갱신": 갱신,
            "연환산_달력": round(이탈 / len(a) / days * 365 * 100, 2) if days else None,
            "연환산_영업": round(이탈 / len(a) / bd * 영업일수_연 * 100, 2) if bd else None,
        })

    t = pd.DataFrame(rows)
    t.insert(0, "구분", "구간")

    # 누적(첫 스냅샷 -> 각 스냅샷)을 끝 날짜마다 한 줄씩 더한다.
    # 보고서는 특정 끝 날짜의 누적 줄에 고정해 인용한다(30번 감사의 스냅샷_컷).
    # 그래야 매일 스냅샷이 쌓여도 문서의 숫자가 따라 흔들리지 않는다.
    a = pd.read_csv(files[0], dtype=str).drop_duplicates(subset=[KEY]).set_index(KEY)
    d0 = os.path.basename(files[0])[6:14]
    누적 = []
    for f1 in files[1:]:
        b = pd.read_csv(f1, dtype=str).drop_duplicates(subset=[KEY]).set_index(KEY)
        d1 = os.path.basename(f1)[6:14]
        days, bd = (pd.to_datetime(d1) - pd.to_datetime(d0)).days, 영업일수(d0, d1)
        이탈 = len(a.index.difference(b.index))
        공통 = a.index.intersection(b.index)
        누적.append({
            "구분": "누적", "시작": d0, "끝": d1, "달력일": days, "영업일": bd,
            "시작_유효": len(a), "끝_유효": len(b), "이탈": 이탈,
            "신규": len(b.index.difference(a.index)),
            "갱신": int((a.loc[공통, "issueenddate"].values
                         != b.loc[공통, "issueenddate"].values).sum()),
            "연환산_달력": round(이탈 / len(a) / days * 365 * 100, 2) if days else None,
            "연환산_영업": round(이탈 / len(a) / bd * 영업일수_연 * 100, 2) if bd else None,
        })
    t = pd.concat([t, pd.DataFrame(누적)], ignore_index=True)

    os.makedirs("output", exist_ok=True)
    t.to_csv("output/스냅샷_시계열.csv", index=False, encoding="utf-8-sig")
    print("=" * 78)
    print(" 스냅샷 시계열  (구간 = 인접한 두 스냅샷 / 누적 = 첫 스냅샷에서 각 스냅샷까지)")
    print("=" * 78)
    print(t.to_string(index=False))
    print("\n[저장] output/스냅샷_시계열.csv")
    if len(t) > 1:
        구간 = t[t["구분"] == "구간"]
        print("\n 구간별 연환산(영업일) {:.2f} ~ {:.2f}%  중앙값 {:.2f}%".format(
            구간["연환산_영업"].min(), 구간["연환산_영업"].max(),
            구간["연환산_영업"].median()))
        print(" 하루치 값은 흔들린다. 구간 수가 적을 때는 누적 줄을 보는 편이 낫다.")
    return t


# ─────────────────────────────────────────────────────────────
# 비교
# ─────────────────────────────────────────────────────────────
def snapshots():
    return sorted(glob.glob(f"{SNAP_DIR}/haccp_*.csv"))


def diff(t0=None, t1=None):
    """두 스냅샷을 대조해 이탈·신규·갱신 사건을 뽑는다."""
    files = snapshots()
    if len(files) < 2:
        print("[비교] 스냅샷이 2개 이상이어야 한다. 먼저 collect 를 실행할 것.")
        print(f"  현재: {[os.path.basename(f) for f in files]}")
        return None

    p0 = f"{SNAP_DIR}/haccp_{t0}.csv" if t0 else files[0]
    p1 = f"{SNAP_DIR}/haccp_{t1}.csv" if t1 else files[-1]
    a = pd.read_csv(p0, dtype=str).drop_duplicates(subset=[KEY]).set_index(KEY)
    b = pd.read_csv(p1, dtype=str).drop_duplicates(subset=[KEY]).set_index(KEY)
    d0 = os.path.basename(p0)[6:14]
    d1 = os.path.basename(p1)[6:14]

    사라짐 = a.loc[a.index.difference(b.index)].copy()      # ★ 관측된 이탈
    신규 = b.loc[b.index.difference(a.index)].copy()
    공통 = a.index.intersection(b.index)
    갱신 = 공통[a.loc[공통, "issueenddate"].values != b.loc[공통, "issueenddate"].values]

    days = (pd.to_datetime(d1) - pd.to_datetime(d0)).days
    print("=" * 70)
    print(f" 스냅샷 대조  {d0} → {d1}  ({days}일)")
    print("=" * 70)
    print(f"  1차 유효 인증        {len(a):,}건")
    print(f"  2차 유효 인증        {len(b):,}건")
    print(f"  ★ 사라진 인증(이탈)  {len(사라짐):,}건")
    print(f"    신규 인증          {len(신규):,}건")
    print(f"    만료일 연장(갱신)  {len(갱신):,}건")
    영업일 = 영업일수(d0, d1)
    if days > 0 and len(a):
        print(f"  → 관측 이탈률 연환산(달력일) {len(사라짐)/len(a)/days*365*100:.2f}%")
        if 영업일 > 0:
            print(f"  → 관측 이탈률 연환산(영업일) "
                  f"{len(사라짐)/len(a)/영업일*영업일수_연*100:.2f}%")
        print(f"    관측 {days}일 중 영업일 {영업일}일")
        print("    (03의 뺄셈 추정 6.16%와 대조하면 방법 자체의 검증이 된다)")

    os.makedirs("output", exist_ok=True)
    out = f"output/이탈사건_{d0}_{d1}.csv"
    사라짐.reset_index().to_csv(out, index=False, encoding="utf-8-sig")
    print(f"\n[저장] 개별 이탈 사건 명단 → {out}")

    if len(사라짐):
        print("\n 이탈 인증의 구성")
        for col in ["productGb", "companykindNm", "area1"]:
            if col in 사라짐:
                print(f"\n  [{col}]")
                print(사라짐[col].value_counts().head(8).to_string())
    return 사라짐


# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    seed_from_raw()
    cmd = sys.argv[1] if len(sys.argv) > 1 else "diff"
    if cmd == "collect":
        collect()
    elif cmd == "series":
        series()
    else:
        diff(*sys.argv[2:4])
