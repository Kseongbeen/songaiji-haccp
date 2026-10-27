# -*- coding: utf-8 -*-
"""
02. HACCP 제품별 인증정보 수집기 (중단·재개 가능)

배경: 전체 약 140만 건 / 페이지당 100건 상한 → 약 14,000회 호출이 필요하다.
      응답이 호출당 10초 내외로 느리고 동시 요청 시 실패율이 높아,
      낮은 동시성으로 나눠 받고 페이지 단위로 저장해 언제든 재개할 수 있게 한다.
      (개발계정 트래픽 한도 10,000건/일 → 이틀에 걸쳐 수집)

산출: data/prd_shards/p{페이지}.csv  →  data/haccp_제품_raw.csv 로 병합
"""

import os
import sys
import glob
import json
import xml.etree.ElementTree as ET
import time
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

SERVICE_KEY = os.environ.get(
    "HACCP_API_KEY",
    "",
)
BASE_URL = "https://apis.data.go.kr/B553748/CertPrdListService/getCertPrdListService"
SHARD_DIR = "data/prd_shards"
PAGE_SIZE = 100
WORKERS = 4          # 동시성을 높이면 타임아웃 급증. 4가 안정적
DAILY_LIMIT = 9500   # 개발계정 일 한도(10,000) 여유분


def done_pages():
    """이미 저장된 페이지 번호 집합.

    빈 파일은 완료로 세지 않는다. 예전 판이 items 가 비어도 CSV를 써 버려서
    빈 shard 가 '완료'로 남고 영영 다시 받지 않는 문제가 있었다.
    빈 파일은 지워서 다음 실행이 다시 받게 한다."""
    done = set()
    for f in glob.glob(f"{SHARD_DIR}/p*.csv"):
        if os.path.getsize(f) < 20:
            try:
                os.remove(f)
            except OSError:
                pass
            continue
        done.add(int(os.path.basename(f)[1:-4]))
    return done


def parse(text):
    """이 API는 returnType=json 을 줘도 가끔 XML로 답한다.
    둘 다 받아서 (totalCount, [레코드...]) 로 통일해 돌려준다."""
    t = text.lstrip()
    if t.startswith("{"):
        body = json.loads(t).get("body") or {}
        items = [row.get("item", row) for row in (body.get("items") or [])]
        return int(body.get("totalCount", 0)), items
    if t.startswith("<"):
        root = ET.fromstring(t)
        tc = root.findtext(".//totalCount") or "0"
        items = []
        for it in root.iter("item"):
            items.append({c.tag: (c.text or "") for c in it})
        return int(tc), items
    raise ValueError("JSON도 XML도 아닌 응답: " + t[:120])


def total_count(session):
    r = session.get(
        f"{BASE_URL}?serviceKey={SERVICE_KEY}&pageNo=1&numOfRows=1&returnType=json", timeout=180)
    return parse(r.text)[0]


def fetch(session, page, retries=3):
    """한 페이지를 받아 shard로 저장. 성공 시 True."""
    for _ in range(retries):
        try:
            r = session.get(
                f"{BASE_URL}?serviceKey={SERVICE_KEY}&pageNo={page}&numOfRows={PAGE_SIZE}&returnType=json",
                timeout=180,
            )
            _, items = parse(r.text)
            if items:
                pd.DataFrame(items).to_csv(
                    f"{SHARD_DIR}/p{page}.csv", index=False, encoding="utf-8-sig")
                return True
        except Exception:
            time.sleep(3)
    return False


def collect():
    os.makedirs(SHARD_DIR, exist_ok=True)
    session = requests.Session()
    total = total_count(session)
    last_page = total // PAGE_SIZE + 1
    todo = sorted(set(range(1, last_page + 1)) - done_pages())[:DAILY_LIMIT]
    print(f"[제품수집] 전체 {total:,}건 / 남은 페이지 {len(todo):,} (완료 {len(done_pages()):,})", flush=True)

    ok = 0
    with ThreadPoolExecutor(WORKERS) as ex:
        for i, r in enumerate(ex.map(lambda p: fetch(session, p), todo), 1):
            ok += r
            if i % 200 == 0:
                print(f"  진행 {i:,}/{len(todo):,} 성공 {ok:,}", flush=True)
    print(f"[제품수집] 종료. 이번 실행 성공 {ok:,}페이지, 누적 {len(done_pages()):,}페이지", flush=True)


def merge():
    """shard 병합 + 인증번호별 제품 수 집계.

    주의 — shard 마다 컬럼 순서가 다를 수 있다. API가 JSON으로 답할 때와
    XML로 답할 때 필드 순서가 다르기 때문이다. 컬럼 이름은 같으므로
    pd.concat 이 이름으로 정렬해 붙인다. 순서 차이는 문제가 되지 않는다."""
    files = [f for f in glob.glob(f"{SHARD_DIR}/p*.csv") if os.path.getsize(f) >= 20]
    print(f"[병합] shard {len(files):,}개 읽는 중...", flush=True)
    df = pd.concat((pd.read_csv(f, dtype=str) for f in files), ignore_index=True)
    df = df.drop_duplicates(subset=["prdlstReportNo"])
    df.to_csv("data/haccp_제품_raw.csv", index=False, encoding="utf-8-sig")

    agg = df.groupby("appointno").agg(
        제품수=("prdlstNm", "size"),
        제품유형수=("businessitemNm", "nunique"),
    ).reset_index().rename(columns={"appointno": "인증번호"})
    agg.to_csv("output/인증번호별_제품수.csv", index=False, encoding="utf-8-sig")
    print(f"[병합] 제품 {len(df):,}건 / 인증번호 {len(agg):,}개")
    return agg


if __name__ == "__main__":
    # 병합(merge)은 shard 수천 개를 한 DataFrame 으로 올려 메모리를 많이 쓴다.
    # 수집 중에 같이 돌리면 시스템 메모리 부족으로 죽을 수 있어 기본값에서 뺐다.
    #   python 02_제품정보_수집.py          수집만 (매일 이걸 돌린다)
    #   python 02_제품정보_수집.py merge    병합만 (수집이 끝난 뒤 한 번)
    #   python 02_제품정보_수집.py all      둘 다
    모드 = sys.argv[1] if len(sys.argv) > 1 else "collect"
    if 모드 in ("collect", "all"):
        collect()
    if 모드 in ("merge", "all"):
        merge()
