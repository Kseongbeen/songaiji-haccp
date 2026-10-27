# -*- coding: utf-8 -*-
"""
36. 스마트HACCP에 농장은 들어가 있는가

【검사하는 주장】
"축산은 스마트HACCP이 안 되는 것 같다." 사용자의 감이다.
제도상으로는 축산업 등록 농가도 신청 대상이다. 그러면 실제로는 몇 곳이 들어가 있나.

【쓰는 데이터 — 둘 다 주최기관(한국식품안전관리인증원)이 직접 낸다】
  ① 축산물 스마트HACCP 인증업체 정보   B553748/SmartCertLiveStockListService/getLiveStockList
     스마트HACCP 등록을 마친 축산물 업체. 업태·품목·중요관리점(CCP)이 붙어 있다.
  ② 스마트HACCP 기술지원 업체 정보      B553748/CounselSmartListService/getCounselSmartListService
     등록 전에 기술지원을 요청한 업체. 들어오는 파이프라인이다.

【비교 기준】
업태별 스마트HACCP 등록 수를 같은 업태의 HACCP 인증 업소 수(01번 스냅샷)로 나눈다.
'인증은 받았는데 스마트로 넘어간 비율'이다.

산출: data/smart_축산물.csv, data/smart_기술지원.csv, output/스마트_업태별.csv
"""

import io
import os
import sys
import time
import requests
import pandas as pd

OUT = "output"
SERVICE_KEY = os.environ.get(
    "HACCP_API_KEY",
    "",
)
API = {
    "축산물": "https://apis.data.go.kr/B553748/SmartCertLiveStockListService/getLiveStockList",
    "기술지원": "https://apis.data.go.kr/B553748/CounselSmartListService/getCounselSmartListService",
}
PAGE = 100


def 받기(이름):
    """전 페이지를 받는다. 이 API는 가끔 XML로 답하므로 JSON 실패 시 재시도한다."""
    url = API[이름]
    rows, page, total = [], 1, None
    while True:
        for 시도 in range(5):
            try:
                r = requests.get("{}?serviceKey={}&pageNo={}&numOfRows={}&returnType=json".format(
                    url, SERVICE_KEY, page, PAGE), timeout=120)
                j = r.json()
                break
            except Exception as e:
                print("   재시도 {} ({})".format(시도 + 1, str(e)[:60]), flush=True)
                time.sleep(3)
        else:
            sys.exit("[중단] {} {}페이지를 받지 못했다".format(이름, page))
        body = j["body"]
        total = int(body["totalCount"])
        items = body.get("items") or []
        if isinstance(items, dict):
            items = items.get("item", [])
            items = items if isinstance(items, list) else [items]
        for it in items:
            rows.append(it.get("item", it))
        print("[{}] {}페이지 누적 {:,} / {:,}".format(이름, page, len(rows), total), flush=True)
        if len(rows) >= total or not items:
            break
        page += 1
    d = pd.DataFrame(rows)
    p = "data/smart_{}.csv".format(이름)
    d.to_csv(p, index=False, encoding="utf-8-sig")
    print("[저장] {}  {:,}건".format(p, len(d)))
    return d


def 분석(sm, cs):
    print()
    print("=" * 78)
    print(" ① 축산물 스마트HACCP 등록 업체 — 업태별")
    print("=" * 78)
    sm = sm.drop_duplicates("appointno")
    t = sm["businesstype"].value_counts().rename("스마트등록").to_frame()

    h = pd.read_csv("data/haccp_인증업체_raw.csv", dtype=str)
    h = h[h["productGb"] == "축산물"].drop_duplicates("appointno")
    base = h["businesstypeNm"].value_counts()
    t["HACCP인증"] = t.index.map(base).fillna(0).astype(int)
    for 업태 in ["가축사육업"]:
        if 업태 not in t.index:
            t.loc[업태] = [0, int(base.get(업태, 0))]
    t["스마트전환%"] = (t["스마트등록"] / t["HACCP인증"].where(t["HACCP인증"] > 0) * 100).round(2)
    t = t.sort_values("스마트등록", ascending=False)
    print(t.to_string())
    t.to_csv(OUT + "/스마트_업태별.csv", encoding="utf-8-sig")

    농장 = int(t.loc["가축사육업", "스마트등록"]) if "가축사육업" in t.index else 0
    농장인증 = int(base.get("가축사육업", 0))
    print()
    print(" 가축사육업(농장): HACCP 인증 {:,}건 중 스마트HACCP 등록 {:,}건".format(농장인증, 농장))

    print()
    print("=" * 78)
    print(" ② 스마트HACCP 기술지원 요청 업체 — 업종별 (등록 전 파이프라인)")
    print("=" * 78)
    if len(cs):
        c = cs["indutyNm"].fillna("미상").value_counts()
        print(c.head(25).to_string())
        # 축산물 가공·유통은 업종명 핵심어로 묶고, 농장(가축사육업)은 따로 센다.
        가공 = c[c.index.str.contains("식육|축산물|유가공|알가공|식용란", na=False)
               & (c.index != "가축사육업")]
        print()
        print(" 축산물 가공·유통 {:,}건 / 농장 {:,}건 / 전체 {:,}건".format(
            int(가공.sum()), int(c.get("가축사육업", 0)), len(cs)))
        c.to_csv(OUT + "/스마트_기술지원_업종.csv", encoding="utf-8-sig")


def 중요관리점(sm):
    """스마트HACCP 등록 업체가 어떤 중요관리점(CCP)을 자동 기록하는가.

    API 응답은 563행인데 인증번호 기준으로는 298건이다. 한 업체가 연도·CCP별로
    여러 행을 가진다. 업체 수를 셀 때는 반드시 인증번호로 중복을 걷어낸다.

    CCP 표기는 '완제품 냉장보관', '완제품냉장보관공정'처럼 제각각이라 핵심어로 묶는다.
    """
    sm = sm.drop_duplicates("appointno")
    c = sm["ccp"].dropna().str.split(",").explode().str.strip()
    c = c[c != ""]
    규칙 = [("금속", "금속검출"), ("냉장", "냉장·냉동 보관"), ("냉동", "냉장·냉동 보관"),
          ("동결", "냉장·냉동 보관"), ("살균", "가열·살균"), ("가열", "가열·살균"),
          ("열처리", "가열·살균"), ("자숙", "가열·살균"), ("훈연", "가열·살균"),
          ("이물", "이물 검출"), ("X-ray", "이물 검출"), ("엑스레이", "이물 검출")]

    def 군(x):
        for k, v in 규칙:
            if k in x:
                return v
        return "기타"

    g = c.map(군).value_counts()
    기계 = g.drop(["기타"], errors="ignore").sum()
    print()
    print("=" * 78)
    print(" ③ 스마트HACCP 등록 업체의 중요관리점 — 무엇을 자동 기록하는가")
    print("=" * 78)
    print(" 등록 {:,}건 / 중요관리점 표기 {:,}개".format(len(sm), len(c)))
    print(g.to_string())
    print(" 온도·검출기·가열처럼 기계가 기록하는 항목 {:.1f}%".format(기계 / len(c) * 100))
    pd.DataFrame([{
        "등록": len(sm), "CCP표기": len(c),
        "냉장냉동": int(g.get("냉장·냉동 보관", 0)), "금속검출": int(g.get("금속검출", 0)),
        "가열살균": int(g.get("가열·살균", 0)), "이물검출": int(g.get("이물 검출", 0)),
        "기타": int(g.get("기타", 0)),
        "기계기록형%": round(기계 / len(c) * 100, 1),
    }]).to_csv(OUT + "/스마트_CCP유형.csv", index=False, encoding="utf-8-sig")
    print("[저장] output/스마트_CCP유형.csv")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "all"
    if mode in ("all", "collect"):
        sm = 받기("축산물")
        cs = 받기("기술지원")
    else:
        sm = pd.read_csv("data/smart_축산물.csv", dtype=str)
        cs = pd.read_csv("data/smart_기술지원.csv", dtype=str)
    분석(sm, cs)
    중요관리점(sm)
