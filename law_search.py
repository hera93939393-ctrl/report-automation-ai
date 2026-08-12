# -*- coding: utf-8 -*-
"""
법령 검색 (F5) — 국가법령정보센터 Open API
- 검색어(키워드)만 외부로 전송한다. 문서 원본이나 개인정보는 전송하지 않는다.
- 사용 전 OC(API 이용자 ID) 발급 필요: https://open.law.go.kr 회원가입 후 Open API 신청
  발급받은 OC는 환경변수 LAW_API_OC 로 설정해서 쓴다 (코드에 직접 적지 않음).
"""
import os
import requests
import xml.etree.ElementTree as ET

LAW_SEARCH_URL = "http://www.law.go.kr/DRF/lawSearch.do"
OC_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "law_api_oc.txt")


def _read_oc_from_file() -> str:
    if os.path.exists(OC_FILE):
        with open(OC_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return ""


OC = os.environ.get("LAW_API_OC") or _read_oc_from_file()


def search_law(keyword: str, display: int = 10):
    """법령명에 키워드가 포함된 법령 목록을 검색해 반환한다.
    반환: [{"법령명": ..., "법령ID": ..., "공포일자": ..., "소관부처": ..., "상세링크": ...}, ...]
    """
    if not OC:
        raise ValueError(
            "LAW_API_OC 환경변수가 설정되지 않았습니다. "
            "https://open.law.go.kr 에서 OC를 발급받아 설정하세요."
        )

    params = {
        "OC": OC,
        "target": "law",
        "type": "XML",
        "query": keyword,
        "display": display,
    }
    res = requests.get(LAW_SEARCH_URL, params=params, timeout=30)
    res.raise_for_status()
    root = ET.fromstring(res.content)

    results = []
    for law in root.findall("law"):
        results.append({
            "법령명": (law.findtext("법령명한글") or "").strip(),
            "법령ID": (law.findtext("법령ID") or "").strip(),
            "공포일자": (law.findtext("공포일자") or "").strip(),
            "소관부처": (law.findtext("소관부처명") or "").strip(),
            "상세링크": (law.findtext("법령상세링크") or "").strip(),
        })
    return results


if __name__ == "__main__":
    if not OC:
        print("LAW_API_OC 환경변수를 설정한 뒤 다시 실행하세요.")
        print('예: $env:LAW_API_OC = "발급받은OC값"  (PowerShell)')
    else:
        # 이 API는 법령 "제목"에서 찾기 때문에, 문장이 아니라 짧은 단어 하나로 검색해야 잘 찾아짐
        results = search_law("청소년")
        for r in results:
            print(r)
