import io
import math
import re
import ipaddress
import socket
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from statistics import mean
from typing import Dict, List, Tuple
from urllib.parse import urljoin, urlparse, quote, unquote

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# =========================================================
# 기본 설정
# =========================================================
BASE_URL = "https://temp.garak.co.kr"
RESULT_PATH = "/price/resultAuctionList.do"

MARKETS = {
    "가락": "1",
    "강서": "3",
}

GRADE_ORDER = {
    "특": 0,
    "상": 1,
    "보통": 2,
    "중": 2,
    "하": 3,
    "등외": 4,
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Linux; Android 16; Mobile) "
        "AppleWebKit/537.36 Chrome/154 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

st.set_page_config(
    page_title="가락·강서 경매조회",
    page_icon="🥕",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
      .block-container {
        padding-top: 1.1rem;
        padding-bottom: 4rem;
        max-width: 1100px;
      }
      h1 { font-size: 1.55rem !important; }
      h2 { font-size: 1.25rem !important; }
      h3 { font-size: 1.08rem !important; }
      div[data-testid="stMetric"] {
        border: 1px solid rgba(120,120,120,.20);
        border-radius: 12px;
        padding: .6rem .75rem;
        background: rgba(250,250,250,.65);
      }
      .market-badge {
        display:inline-block;
        padding:.18rem .55rem;
        margin-right:.35rem;
        border-radius:999px;
        font-weight:700;
        font-size:.78rem;
        background:#eef2f7;
      }
      .company-card {
        border:1px solid rgba(120,120,120,.22);
        border-radius:14px;
        padding:.75rem .9rem;
        margin:.3rem 0 .7rem 0;
      }
      .price-legend {
        font-size:.88rem;
        opacity:.88;
        margin-bottom:.4rem;
      }
      @media (max-width: 640px) {
        .block-container { padding-left: .8rem; padding-right: .8rem; }
        h1 { font-size: 1.35rem !important; }
        div[data-testid="stMetricValue"] { font-size: 1.15rem; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# HTTP / HTML 파서
# =========================================================
def make_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.45,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST"]),
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=12, pool_maxsize=12)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    s.headers.update(HEADERS)
    return s


def result_url(market_code: str) -> str:
    return f"{BASE_URL}{RESULT_PATH}?market_cd={market_code}&menu_flag=1"


def parse_form_fields(html: str) -> Dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    form = soup.find("form", id="aForm") or soup.find("form", attrs={"name": "aForm"}) or soup.find("form")
    fields: Dict[str, str] = {}
    if not form:
        return fields

    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        typ = (inp.get("type") or "text").lower()
        if typ in {"submit", "button", "image", "reset", "file"}:
            continue
        if typ in {"checkbox", "radio"} and not inp.has_attr("checked"):
            continue
        fields[name] = inp.get("value", "")

    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name:
            continue
        opt = sel.find("option", selected=True) or sel.find("option")
        fields[name] = opt.get("value", "") if opt else ""

    return fields


def clean_company_name(text: str, market_name: str) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    text = re.sub(rf"^{re.escape(market_name)}(?:시장)?\s*[-:·]\s*", "", text)
    return text.strip()


def parse_company_options(html: str, market_name: str) -> List[Dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    sel = soup.find("select", attrs={"name": "s_bubin"})
    if not sel:
        return []

    raw = []
    prefixed = []

    for opt in sel.find_all("option"):
        code = (opt.get("value") or "").strip()
        text = re.sub(r"\s+", " ", opt.get_text(" ", strip=True))
        if not code or code == "-1":
            continue
        if not re.fullmatch(r"\d+", code):
            continue
        if any(x in text for x in ("법인명", "전체", "선택")):
            continue

        item = {
            "code": code,
            "raw_name": text,
            "name": clean_company_name(text, market_name),
            "market": market_name,
        }
        raw.append(item)

        if text.startswith(f"{market_name}-") or text.startswith(f"{market_name}시장-"):
            prefixed.append(item)

    # 페이지가 공용 법인목록이면 접두사가 있는 해당 시장 법인만,
    # 시장별 목록이면 접두사가 없을 수 있으므로 전체 유효옵션을 사용.
    return prefixed if prefixed else raw


def make_payload(
    base_fields: Dict[str, str],
    market_code: str,
    date_dot: str,
    item: str,
    company_code: str,
    page: int,
) -> Dict[str, str]:
    d = dict(base_fields)
    d.update(
        {
            "market_cd": market_code,
            "menu_flag": "1",
            "currentPage": str(page),
            "s_date": date_dot,
            "e_date": date_dot,
            "s_bubin": company_code,
            "s_pummok": item,
            "s_pummok1": item,
            "s_sangi": "",
            "s_sangi2": "",
            "s_danwi": "",
            "s_qty": "",
            "s_gubun": "",
            "R010680": "10",
            "R010690": "10",
            "R010700": "10",
            "R010710": "",
            "market": "garak" if market_code == "1" else "gangseo",
        }
    )
    return d


def parse_auction_rows(html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "html.parser")
    out = []

    for tr in soup.find_all("tr"):
        cells = [
            re.sub(r"\s+", " ", td.get_text(" ", strip=True)).strip()
            for td in tr.find_all(["th", "td"])
        ]
        if len(cells) != 6:
            continue

        no, item, unit, grade, price_text, origin = cells
        if not re.fullmatch(r"\d+", no):
            continue
        if not re.fullmatch(r"\d{1,3}(?:,\d{3})*|\d+", price_text):
            continue

        out.append(
            {
                "번호": int(no),
                "품목(품종)": item,
                "단위": unit,
                "등급": grade,
                "경락가": int(price_text.replace(",", "")),
                "출하지": origin,
            }
        )
    return out


@st.cache_data(ttl=600, show_spinner=False)
def get_market_context(market_name: str) -> Tuple[Dict[str, str], List[Dict[str, str]]]:
    market_code = MARKETS[market_name]
    with make_session() as s:
        r = s.get(result_url(market_code), timeout=25)
        r.raise_for_status()
        html = r.text
    return parse_form_fields(html), parse_company_options(html, market_name)


@st.cache_data(ttl=300, show_spinner=False)
def probe_company(date_yyyymmdd: str, item: str, market_name: str, company_code: str) -> int:
    market_code = MARKETS[market_name]
    date_dot = f"{date_yyyymmdd[:4]}.{date_yyyymmdd[4:6]}.{date_yyyymmdd[6:8]}"
    base_fields, _ = get_market_context(market_name)

    with make_session() as s:
        payload = make_payload(base_fields, market_code, date_dot, item, company_code, 1)
        r = s.post(result_url(market_code), data=payload, timeout=30)
        r.raise_for_status()
        rows = parse_auction_rows(r.text)
    return len(rows)


@st.cache_data(ttl=60, show_spinner=False)
def discover_company_status(
    date_yyyymmdd: str, item: str, markets: Tuple[str, ...]
) -> List[Dict[str, str]]:
    """
    선택한 시장의 법인을 전부 보여주되,
    해당 날짜/품목의 경매결과가 있는 법인만 selectable=True 로 표시한다.
    """
    candidates: List[Dict[str, str]] = []

    for market_name in markets:
        _, opts = get_market_context(market_name)
        candidates.extend(opts)

    status_rows = []
    workers = min(10, max(1, len(candidates)))

    with ThreadPoolExecutor(max_workers=workers) as ex:
        future_map = {
            ex.submit(
                probe_company,
                date_yyyymmdd,
                item,
                c["market"],
                c["code"],
            ): c
            for c in candidates
        }

        for fut in as_completed(future_map):
            c = future_map[fut]
            try:
                first_page_count = fut.result()
            except Exception:
                first_page_count = 0

            status_rows.append(
                {
                    **c,
                    "first_page_count": first_page_count,
                    "available": first_page_count > 0,
                    "key": f'{c["market"]}|{c["code"]}',
                }
            )

    market_rank = {"가락": 0, "강서": 1}
    status_rows.sort(
        key=lambda x: (
            market_rank.get(x["market"], 99),
            0 if x["available"] else 1,
            x["name"],
        )
    )
    return status_rows


@st.cache_data(ttl=300, show_spinner=False)
def fetch_all_company_rows(
    date_yyyymmdd: str,
    item: str,
    market_name: str,
    company_code: str,
    max_pages: int = 100,
) -> pd.DataFrame:
    market_code = MARKETS[market_name]
    date_dot = f"{date_yyyymmdd[:4]}.{date_yyyymmdd[4:6]}.{date_yyyymmdd[6:8]}"
    base_fields, _ = get_market_context(market_name)

    all_rows = []
    previous_signature = None

    with make_session() as s:
        for page in range(1, max_pages + 1):
            payload = make_payload(
                base_fields,
                market_code,
                date_dot,
                item,
                company_code,
                page,
            )
            r = s.post(result_url(market_code), data=payload, timeout=35)
            r.raise_for_status()

            rows = parse_auction_rows(r.text)
            if not rows:
                break

            signature = tuple(
                (x["번호"], x["품목(품종)"], x["경락가"])
                for x in rows
            )
            if page > 1 and signature == previous_signature:
                break
            previous_signature = signature

            for row in rows:
                row["페이지"] = page
                row["시장"] = market_name
                row["법인코드"] = company_code
                all_rows.append(row)

    cols = ["시장", "법인코드", "페이지", "번호", "품목(품종)", "단위", "등급", "경락가", "출하지"]
    return pd.DataFrame(all_rows, columns=cols)



# =========================================================
# v16: 품목 미입력 시 오늘 진행된 경매 품목 탐색
# =========================================================
def _base_item_name_v16(raw_name: str) -> str:
    """
    공개 경매결과의 '품목(품종)' 표기에서 품목 선택용 이름을 만든다.
    예: '양배추(양배추일반)' -> '양배추'
    괄호가 없으면 원문을 그대로 사용한다.
    """
    text = re.sub(r"\s+", " ", str(raw_name or "")).strip()
    if not text:
        return ""
    base = re.sub(r"\s*[\(\（][^\)\）]*[\)\）]\s*$", "", text).strip()
    return base or text


@st.cache_data(ttl=120, show_spinner=False)
def discover_today_items_v16(
    date_yyyymmdd: str,
    markets: Tuple[str, ...],
) -> List[Dict[str, object]]:
    """
    품목을 비운 상태로 각 법인의 당일 경매결과를 조회해,
    실제로 결과가 올라온 품목만 모은다.

    반환:
      item       : 다시 검색할 품목명
      raw_names  : 실제 공개 품목(품종) 표기
      rows       : 확인된 거래행 수
      companies  : 결과가 확인된 법인 수
      markets    : 확인된 시장
    """
    candidates: List[Dict[str, str]] = []
    for market_name in markets:
        _, opts = get_market_context(market_name)
        candidates.extend(opts)

    # 품목별 집계
    agg: Dict[str, Dict[str, object]] = {}

    def fetch_candidate(c: Dict[str, str]):
        # 빈 품목으로 해당 법인의 오늘 결과 전체를 페이지 끝까지 조회
        df = fetch_all_company_rows(
            date_yyyymmdd,
            "",
            c["market"],
            c["code"],
            max_pages=100,
        )
        return c, df

    workers = min(8, max(1, len(candidates)))
    with ThreadPoolExecutor(max_workers=workers) as ex:
        future_map = {ex.submit(fetch_candidate, c): c for c in candidates}

        for fut in as_completed(future_map):
            c = future_map[fut]
            try:
                c2, df = fut.result()
            except Exception:
                continue

            if df is None or df.empty:
                continue

            company_label = f'{c2["market"]} · {c2["name"]}'

            for raw_name, g in df.groupby("품목(품종)", sort=False):
                item_name = _base_item_name_v16(raw_name)
                if not item_name:
                    continue

                slot = agg.setdefault(
                    item_name,
                    {
                        "item": item_name,
                        "raw_names": set(),
                        "rows": 0,
                        "company_labels": set(),
                        "market_names": set(),
                    },
                )
                slot["raw_names"].add(str(raw_name))
                slot["rows"] += int(len(g))
                slot["company_labels"].add(company_label)
                slot["market_names"].add(c2["market"])

    out: List[Dict[str, object]] = []
    for item_name, slot in agg.items():
        out.append(
            {
                "item": item_name,
                "raw_names": sorted(slot["raw_names"]),
                "rows": int(slot["rows"]),
                "companies": len(slot["company_labels"]),
                "company_labels": sorted(slot["company_labels"]),
                "markets": sorted(slot["market_names"]),
            }
        )

    # 거래행 많은 품목 우선, 같은 건수면 이름순
    out.sort(key=lambda x: (-int(x["rows"]), str(x["item"])))
    return out


def render_today_item_picker_v16(
    date_yyyymmdd: str,
    markets: Tuple[str, ...],
):
    st.divider()
    st.subheader("🥕 오늘 진행된 경매 품목")
    st.caption(
        "품목을 비워서 검색했기 때문에, 선택한 장일자에 실제 경매결과가 올라온 품목을 먼저 찾습니다. "
        "품목을 고르면 그 품목 기준으로 다시 법인 선택 화면으로 넘어갑니다."
    )

    reload_col, info_col = st.columns([1, 2.4])
    with reload_col:
        if st.button(
            "🔄 품목 새로고침",
            use_container_width=True,
            key=f"today_items_reload::{date_yyyymmdd}::{','.join(markets)}",
        ):
            discover_today_items_v16.clear()
            fetch_all_company_rows.clear()
            st.rerun()
    with info_col:
        st.caption(
            f"{date_yyyymmdd[:4]}-{date_yyyymmdd[4:6]}-{date_yyyymmdd[6:8]} · "
            + " + ".join(markets)
        )

    with st.spinner("오늘 진행된 품목을 법인별 경매결과에서 모으는 중..."):
        items = discover_today_items_v16(date_yyyymmdd, tuple(markets))

    if not items:
        st.warning(
            "현재 확인되는 경매 품목이 없습니다. 아직 경매결과가 올라오지 않았거나, "
            "공개 경매조회가 빈 품목 조회를 허용하지 않는 경우입니다."
        )
        return

    st.success(f"현재 확인된 품목 {len(items):,}개")

    option_labels = []
    by_label = {}
    for x in items:
        market_text = "/".join(x["markets"])
        label = (
            f'{x["item"]}  ·  {x["rows"]:,}건  ·  '
            f'{x["companies"]}법인  ·  {market_text}'
        )
        option_labels.append(label)
        by_label[label] = x

    selected_label = st.selectbox(
        "품목 선택",
        options=option_labels,
        index=None,
        placeholder="눌러서 오늘 경매된 품목을 선택하세요",
        key=f"today_item_select::{date_yyyymmdd}::{','.join(markets)}",
    )

    if selected_label:
        x = by_label[selected_label]
        with st.expander("이 품목이 확인된 법인/표기"):
            st.write("법인:", " · ".join(x["company_labels"]))
            st.write("공개 품목표기:", " / ".join(x["raw_names"]))

        if st.button(
            f'✅ {x["item"]} 선택하고 법인 다시 찾기',
            type="primary",
            use_container_width=True,
            key=f'choose_today_item::{date_yyyymmdd}::{x["item"]}',
        ):
            st.session_state["search"] = {
                "date": date_yyyymmdd,
                "item": str(x["item"]),
                "markets": tuple(markets),
            }
            st.session_state.pop("item_picker_request_v16", None)

            # 같은 날짜라도 이전 품목에서 선택한 법인이 새 품목으로 넘어오지 않도록 정리
            for k in list(st.session_state.keys()):
                if str(k).startswith("corp_selected::"):
                    st.session_state.pop(k, None)

            qp_set(
                date=date_yyyymmdd,
                item=str(x["item"]),
                markets=",".join(markets),
            )
            st.rerun()


# =========================================================
# 표시 / 계산
# =========================================================
def grade_rank(grade: str) -> Tuple[int, str]:
    text = str(grade)
    for keyword, rank in GRADE_ORDER.items():
        if keyword in text:
            return rank, text
    return 99, text


def actual_middle_price(prices: List[int]) -> int:
    """실제 경락값 중 중앙 순번의 값. 짝수 건도 실제 거래가격 하나를 선택."""
    s = sorted(int(x) for x in prices)
    return s[len(s) // 2]


def price_label(price: int, head: int, middle: int, tail: int) -> str:
    labels = []
    if price == head:
        labels.append("머리")
    if price == middle:
        labels.append("중간")
    if price == tail:
        labels.append("꼬리")
    return "·".join(labels)


def style_price_rows(df: pd.DataFrame):
    def row_style(row):
        tag = str(row.get("구간", ""))
        if "머리" in tag:
            return ["background-color: #fde8e8; font-weight: 700;"] * len(row)
        if "중간" in tag:
            return ["background-color: #fff4cc; font-weight: 700;"] * len(row)
        if "꼬리" in tag:
            return ["background-color: #e6f0ff; font-weight: 700;"] * len(row)
        return [""] * len(row)

    return (
        df.style
        .apply(row_style, axis=1)
        .format({"경락가": "{:,.0f}"})
    )


def render_company(company: Dict[str, str], date_yyyymmdd: str, item: str) -> pd.DataFrame:
    with st.spinner(f'{company["market"]} · {company["name"]} 전체 경매자료 불러오는 중...'):
        df = fetch_all_company_rows(
            date_yyyymmdd,
            item,
            company["market"],
            company["code"],
        )

    if df.empty:
        st.warning("경매자료가 없습니다. 새로고침 후 다시 검색해 주세요.")
        return df

    df = df.copy()
    df.insert(1, "법인", company["name"])

    st.markdown(
        f'<span class="market-badge">{company["market"]}</span>'
        f'<strong>{company["name"]}</strong>',
        unsafe_allow_html=True,
    )

    m1, m2, m3 = st.columns(3)
    m1.metric("경매건수", f"{len(df):,}건")
    m2.metric("품종표기", f'{df["품목(품종)"].nunique():,}개')
    m3.metric("가격범위", f'{df["경락가"].min():,} ~ {df["경락가"].max():,}')

    st.caption("품종 → 단위 → 등급 순으로 표시합니다. 공개 경매결과에 보이는 정보만 사용합니다.")
    st.markdown(
        '<div class="price-legend">🔴 머리(최고가) &nbsp; 🟡 중간(실제 경락 중 중앙값) &nbsp; 🔵 꼬리(최저가)</div>',
        unsafe_allow_html=True,
    )

    item_names = list(dict.fromkeys(df["품목(품종)"].tolist()))

    for idx, item_name in enumerate(item_names):
        item_df = df[df["품목(품종)"] == item_name].copy()

        with st.expander(
            f'{item_name}  ·  {len(item_df):,}건',
            expanded=(len(item_names) == 1 or idx == 0),
        ):
            units = list(dict.fromkeys(item_df["단위"].tolist()))
            for unit in units:
                unit_df = item_df[item_df["단위"] == unit].copy()
                st.markdown(f"#### {unit}")

                grades = sorted(
                    unit_df["등급"].dropna().unique().tolist(),
                    key=grade_rank,
                )

                for grade in grades:
                    g = unit_df[unit_df["등급"] == grade].copy()
                    g = g.sort_values(["경락가", "번호"], ascending=[False, False])

                    prices = g["경락가"].astype(int).tolist()
                    head = max(prices)
                    middle = actual_middle_price(prices)
                    tail = min(prices)
                    avg = round(mean(prices))

                    st.markdown(f"**{grade}** · {len(g):,}건")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("머리", f"{head:,}")
                    c2.metric("중간", f"{middle:,}")
                    c3.metric("꼬리", f"{tail:,}")
                    c4.metric("평균", f"{avg:,}")

                    view = g[["경락가", "출하지", "번호"]].copy()
                    view.insert(
                        0,
                        "구간",
                        [
                            price_label(int(p), head, middle, tail)
                            for p in view["경락가"]
                        ],
                    )

                    st.dataframe(
                        style_price_rows(view),
                        use_container_width=True,
                        hide_index=True,
                        height=min(430, 42 + 35 * min(len(view), 11)),
                    )

    render_integrated_spec_v11(company, date_yyyymmdd, item)

    with st.expander("이 법인 원자료 전체 보기"):
        raw_view = df[
            ["시장", "법인", "번호", "품목(품종)", "단위", "등급", "경락가", "출하지"]
        ].sort_values(["품목(품종)", "단위", "등급", "경락가"], ascending=[True, True, True, False])
        st.dataframe(
            raw_view.style.format({"경락가": "{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
        )

    return df



# =========================================================
# v14: 선택 법인 즉시 비교표
# =========================================================
def _cmp_unit_v14(v: str) -> str:
    t = re.sub(r"\s+", "", str(v or "")).lower()
    m = re.search(r"(\d+(?:\.\d+)?)(kg|g)", t, flags=re.I)
    if not m:
        return t or "-"
    n = float(m.group(1))
    ns = str(int(n)) if n.is_integer() else str(n).rstrip("0").rstrip(".")
    return f"{ns}{m.group(2).lower()}"


def _cmp_grade_v14(v: str) -> str:
    t = re.sub(r"\s+", "", str(v or ""))
    if not t or t in {"-", "없음", "미상", "무등급"} or "등외" in t or "기타" in t:
        return "세부규격 미확인"
    return t


def _comparison_rows_v14(frames: List[pd.DataFrame], search_item: str) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    all_df = pd.concat(frames, ignore_index=True)
    need = {"시장", "법인", "품목(품종)", "단위", "등급", "경락가"}
    if not need.issubset(set(all_df.columns)):
        return pd.DataFrame()
    all_df = all_df.copy()
    all_df["비교단위"] = all_df["단위"].map(_cmp_unit_v14)
    all_df["비교등급"] = all_df["등급"].map(_cmp_grade_v14)
    out = []
    for (u, gr), gg in all_df.groupby(["비교단위", "비교등급"], dropna=False, sort=False):
        corp_count = gg[["시장", "법인"]].drop_duplicates().shape[0]
        if corp_count < 2:
            continue
        for (mkt, corp), cg in gg.groupby(["시장", "법인"], sort=False):
            prices = []
            for x in cg["경락가"].tolist():
                try:
                    if pd.notna(x) and float(x) > 0:
                        prices.append(int(round(float(x))))
                except Exception:
                    pass
            if not prices:
                continue
            out.append({
                "검색품목": search_item,
                "비교단위": u,
                "비교등급": gr,
                "법인": f"{mkt} · {corp}",
                "품종표기": " / ".join(list(dict.fromkeys(cg["품목(품종)"].astype(str).tolist()))[:4]),
                "원등급": " / ".join(list(dict.fromkeys(cg["등급"].astype(str).tolist()))[:4]),
                "건수": len(prices),
                "머리": max(prices),
                "중간": actual_middle_price(prices),
                "꼬리": min(prices),
                "평균": round(mean(prices)),
            })
    return pd.DataFrame(out)


def _fallback_compare_v14(frames: List[pd.DataFrame], search_item: str) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()
    all_df = pd.concat(frames, ignore_index=True)
    out = []
    for (mkt, corp), g in all_df.groupby(["시장", "법인"], sort=False):
        prices = []
        for x in g["경락가"].tolist():
            try:
                if pd.notna(x) and float(x) > 0:
                    prices.append(int(round(float(x))))
            except Exception:
                pass
        if prices:
            out.append({
                "법인": f"{mkt} · {corp}",
                "검색품목": search_item,
                "품종표기": " / ".join(list(dict.fromkeys(g["품목(품종)"].astype(str).tolist()))[:5]),
                "건수": len(prices),
                "머리": max(prices),
                "중간": actual_middle_price(prices),
                "꼬리": min(prices),
                "평균": round(mean(prices)),
            })
    return pd.DataFrame(out)


def render_compare_v14(frames: List[pd.DataFrame], selected_labels: List[str], search_item: str):
    st.markdown("## 📊 선택 법인 비교")
    st.caption("v14 비교모드 · 법인 2곳 이상 선택 시 이 영역은 항상 표시됩니다.")
    if not frames:
        st.error("선택 법인의 비교용 경매자료를 불러오지 못했습니다.")
        return
    cdf = _comparison_rows_v14(frames, search_item)
    if cdf.empty:
        st.warning("공통 단위·등급 그룹이 없어 법인별 전체 요약으로 비교합니다.")
        fb = _fallback_compare_v14(frames, search_item)
        if not fb.empty:
            st.dataframe(
                fb.style.format({
                    "건수": "{:,.0f}", "머리": "{:,.0f}", "중간": "{:,.0f}",
                    "꼬리": "{:,.0f}", "평균": "{:,.0f}"
                }),
                use_container_width=True,
                hide_index=True,
            )
        return
    order = {x: idx for idx, x in enumerate(selected_labels)}
    cdf["순서"] = cdf["법인"].map(order).fillna(9999).astype(int)
    for (u, gr), g in cdf.groupby(["비교단위", "비교등급"], sort=False):
        g = g.sort_values(["순서", "중간"]).copy()
        base = None
        base_label = None
        for lab in selected_labels:
            r = g[g["법인"] == lab]
            if not r.empty:
                base = int(r.iloc[0]["중간"])
                base_label = lab
                break
        g["중간차이"] = g["중간"].astype(int) - (base if base is not None else 0)
        st.markdown(f"### {search_item} · {u} · {gr}")
        view = g[["법인", "품종표기", "원등급", "건수", "머리", "중간", "중간차이", "꼬리", "평균"]].copy()
        st.dataframe(
            view.style.format({
                "건수": "{:,.0f}", "머리": "{:,.0f}", "중간": "{:,.0f}",
                "중간차이": "{:+,.0f}", "꼬리": "{:,.0f}", "평균": "{:,.0f}"
            }),
            use_container_width=True,
            hide_index=True,
        )
        if base_label:
            st.caption(f"중간차이 기준: {base_label}")
        if gr == "세부규격 미확인":
            st.warning("⚠ 공개결과상 세부규격 미확인. 42/45망, 4내/4수 등이 섞였을 수 있어 참고가격 비교로 보세요.")


# =========================================================
# v15: 시세 약세 / 우리만 약세 진단
# - '시장 전체'라고 단정하지 않고 선택한 타 법인 기준으로 진단
# - 오늘 타 법인 중간가 vs 최근 3/5장 타 법인 일별 중간가 평균
# - 우리 오늘 중간가 vs 오늘 타 법인 중간가
# =========================================================
def _group_stats_v15(df: pd.DataFrame, label: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    work = df.copy()
    work["비교단위"] = work["단위"].map(_cmp_unit_v14)
    work["비교등급"] = work["등급"].map(_cmp_grade_v14)

    rows = []
    for (u, gr), g in work.groupby(["비교단위", "비교등급"], dropna=False, sort=False):
        prices = []
        for x in g["경락가"].tolist():
            try:
                if pd.notna(x) and float(x) > 0:
                    prices.append(int(round(float(x))))
            except Exception:
                pass
        if not prices:
            continue

        rows.append({
            "법인": label,
            "비교단위": u,
            "비교등급": gr,
            "품종표기": " / ".join(
                list(dict.fromkeys(g["품목(품종)"].astype(str).tolist()))[:5]
            ),
            "건수": len(prices),
            "중간": actual_middle_price(prices),
            "머리": max(prices),
            "꼬리": min(prices),
            "평균": round(mean(prices)),
        })
    return pd.DataFrame(rows)


def _selected_frame_map_v15(
    frames: List[pd.DataFrame],
) -> Dict[str, pd.DataFrame]:
    out: Dict[str, pd.DataFrame] = {}
    for df in frames:
        if df is None or df.empty:
            continue
        try:
            market_name = str(df.iloc[0]["시장"])
            corp_name = str(df.iloc[0]["법인"])
            label = f"{market_name} · {corp_name}"
            out[label] = df
        except Exception:
            continue
    return out


def _fetch_company_frame_v15(
    company: Dict[str, str],
    ymd: str,
    item: str,
) -> pd.DataFrame:
    df = fetch_all_company_rows(
        ymd,
        item,
        company["market"],
        company["code"],
    )
    if df.empty:
        return df
    df = df.copy()
    if "법인" not in df.columns:
        df.insert(1, "법인", company["name"])
    return df


def _peer_benchmark_for_group_v15(
    stats_by_label: Dict[str, pd.DataFrame],
    peer_labels: List[str],
    unit_key: str,
    grade_key: str,
) -> Tuple[int | None, int]:
    vals = []
    for label in peer_labels:
        sdf = stats_by_label.get(label)
        if sdf is None or sdf.empty:
            continue
        rows = sdf[
            (sdf["비교단위"] == unit_key) &
            (sdf["비교등급"] == grade_key)
        ]
        if rows.empty:
            continue
        vals.append(int(rows.iloc[0]["중간"]))

    if not vals:
        return None, 0

    # 법인별 중간가를 다시 실제 중앙순번 방식으로 대표값화
    return actual_middle_price(vals), len(vals)


def _diagnosis_label_v15(
    market_change_pct: float,
    our_gap_pct: float,
    threshold_pct: float,
) -> str:
    market_weak = market_change_pct <= -threshold_pct
    our_weak = our_gap_pct <= -threshold_pct

    if market_weak and our_weak:
        return "🔻 비교 시세도 약세 + 우리도 더 낮음"
    if market_weak:
        return "🔻 비교 법인 시세 자체가 약세"
    if our_weak:
        return "⚠️ 비교 법인은 버티는데 우리만 낮음"
    return "➡️ 뚜렷한 약세 신호 없음"


def build_weakness_diagnosis_v15(
    companies_by_label: Dict[str, Dict[str, str]],
    current_frames: List[pd.DataFrame],
    selected_labels: List[str],
    our_label: str,
    search_date: str,
    item: str,
    lookback_markets: int,
    threshold_pct: float,
    max_calendar_scan: int = 21,
) -> Tuple[pd.DataFrame, Dict[str, object]]:
    frame_map = _selected_frame_map_v15(current_frames)
    if our_label not in frame_map:
        return pd.DataFrame(), {"message": "우리 법인의 오늘 자료가 없습니다."}

    peer_labels = [x for x in selected_labels if x != our_label and x in companies_by_label]
    if not peer_labels:
        return pd.DataFrame(), {"message": "비교할 타 법인이 없습니다."}

    # 오늘 법인별 그룹 통계
    today_stats: Dict[str, pd.DataFrame] = {}
    for label in selected_labels:
        df = frame_map.get(label)
        if df is not None and not df.empty:
            today_stats[label] = _group_stats_v15(df, label)

    our_stats = today_stats.get(our_label, pd.DataFrame())
    if our_stats.empty:
        return pd.DataFrame(), {"message": "우리 법인의 오늘 비교 가능한 가격이 없습니다."}

    # 오늘 우리와 타 법인 모두 존재하는 그룹만 진단
    target_groups = []
    for _, r in our_stats.iterrows():
        u = r["비교단위"]
        gr = r["비교등급"]
        peer_today, peer_count = _peer_benchmark_for_group_v15(
            today_stats, peer_labels, u, gr
        )
        if peer_today is not None:
            target_groups.append((u, gr))

    if not target_groups:
        return pd.DataFrame(), {
            "message": "오늘 우리 법인과 타 법인에 공통인 단위·등급 그룹이 없습니다."
        }

    # 과거 일별 타 법인 대표 중간가 / 우리 중간가 저장
    history_peer: Dict[Tuple[str, str], List[Tuple[str, int, int]]] = {
        g: [] for g in target_groups
    }
    history_our: Dict[Tuple[str, str], List[Tuple[str, int]]] = {
        g: [] for g in target_groups
    }

    base_date = datetime.strptime(search_date, "%Y%m%d").date()
    scanned = 0

    for offset in range(1, max_calendar_scan + 1):
        # 모든 그룹이 필요한 장 수를 채웠으면 종료
        if all(len(history_peer[g]) >= lookback_markets for g in target_groups):
            break

        d = base_date - timedelta(days=offset)
        ymd = d.strftime("%Y%m%d")
        scanned += 1

        day_stats: Dict[str, pd.DataFrame] = {}

        # 우리 + 비교 법인 모두 조회. fetch_all_company_rows 자체가 캐시됨.
        for label in selected_labels:
            company = companies_by_label.get(label)
            if not company:
                continue
            try:
                hdf = _fetch_company_frame_v15(company, ymd, item)
            except Exception:
                continue
            if hdf.empty:
                continue
            day_stats[label] = _group_stats_v15(hdf, label)

        if not day_stats:
            continue

        for group in target_groups:
            u, gr = group

            if len(history_peer[group]) < lookback_markets:
                peer_value, peer_count = _peer_benchmark_for_group_v15(
                    day_stats, peer_labels, u, gr
                )
                if peer_value is not None:
                    history_peer[group].append((ymd, peer_value, peer_count))

            if len(history_our[group]) < lookback_markets:
                osdf = day_stats.get(our_label)
                if osdf is not None and not osdf.empty:
                    rr = osdf[
                        (osdf["비교단위"] == u) &
                        (osdf["비교등급"] == gr)
                    ]
                    if not rr.empty:
                        history_our[group].append((ymd, int(rr.iloc[0]["중간"])))

    rows = []

    for u, gr in target_groups:
        our_row = our_stats[
            (our_stats["비교단위"] == u) &
            (our_stats["비교등급"] == gr)
        ]
        if our_row.empty:
            continue

        our_today = int(our_row.iloc[0]["중간"])
        peer_today, peer_count_today = _peer_benchmark_for_group_v15(
            today_stats, peer_labels, u, gr
        )
        if peer_today is None or peer_today <= 0:
            continue

        ph = history_peer.get((u, gr), [])
        if not ph:
            rows.append({
                "단위": u,
                "등급": gr,
                "우리오늘": our_today,
                "타법인오늘": peer_today,
                f"타법인최근{lookback_markets}장": None,
                "시세변동%": None,
                "우리격차%": round((our_today / peer_today - 1) * 100, 1),
                "오늘비교법인수": peer_count_today,
                "과거확보장수": 0,
                "진단": "과거자료 부족",
            })
            continue

        peer_hist_values = [x[1] for x in ph[:lookback_markets]]
        peer_hist_avg = round(mean(peer_hist_values))
        market_change = (
            (peer_today / peer_hist_avg - 1) * 100
            if peer_hist_avg > 0 else 0.0
        )
        our_gap = (our_today / peer_today - 1) * 100

        label = _diagnosis_label_v15(
            market_change,
            our_gap,
            threshold_pct,
        )

        rows.append({
            "단위": u,
            "등급": gr,
            "우리오늘": our_today,
            "타법인오늘": peer_today,
            f"타법인최근{lookback_markets}장": peer_hist_avg,
            "시세변동%": round(market_change, 1),
            "우리격차%": round(our_gap, 1),
            "오늘비교법인수": peer_count_today,
            "과거확보장수": min(len(peer_hist_values), lookback_markets),
            "진단": label,
        })

    meta = {
        "scanned_calendar_days": scanned,
        "peer_labels": peer_labels,
        "lookback": lookback_markets,
        "threshold": threshold_pct,
    }
    return pd.DataFrame(rows), meta


def render_weakness_diagnosis_v15(
    companies_by_label: Dict[str, Dict[str, str]],
    current_frames: List[pd.DataFrame],
    selected_labels: List[str],
    search_date: str,
    item: str,
):
    st.markdown("## 📉 시세 약세 / 우리만 약세 진단")
    st.caption(
        "선택한 법인들만 기준으로 봅니다. `시장 전체`를 단정하는 기능은 아닙니다. "
        "오늘 타 법인 중간가와 최근 장들의 타 법인 일별 중간가를 비교합니다."
    )

    if len(selected_labels) < 2:
        st.info("법인을 2곳 이상 선택하면 진단할 수 있습니다.")
        return

    c1, c2, c3 = st.columns([1.4, 1, 1])
    with c1:
        our_label = st.selectbox(
            "우리 법인",
            options=selected_labels,
            index=0,
            key=f"diag_our::{search_date}::{item}",
        )
    with c2:
        lookback = st.selectbox(
            "최근 기준",
            options=[3, 5],
            index=1,
            format_func=lambda x: f"최근 {x}장",
            key=f"diag_lookback::{search_date}::{item}",
        )
    with c3:
        threshold = st.selectbox(
            "약세 기준",
            options=[3, 5, 7, 10],
            index=1,
            format_func=lambda x: f"{x}% 이상",
            key=f"diag_threshold::{search_date}::{item}",
        )

    diag_state_key = (
        f"diag_run::{search_date}::{item}::"
        + "::".join(selected_labels)
    )
    if st.button(
        "📉 약세 진단 실행",
        type="primary",
        use_container_width=True,
        key=f"diag_button::{search_date}::{item}",
    ):
        st.session_state[diag_state_key] = True

    if not st.session_state.get(diag_state_key, False):
        st.caption(
            "예: 타 법인 오늘 중간가가 최근 5장 평균보다 5% 이상 낮으면 "
            "`비교 시세 약세`, 우리 중간가가 타 법인보다 5% 이상 낮으면 "
            "`우리만 낮음` 신호로 봅니다."
        )
        return

    with st.spinner(
        f"최근 {lookback}장 비교자료 확인 중... "
        "휴장일은 건너뛰므로 처음 한 번은 시간이 조금 걸릴 수 있습니다."
    ):
        diag_df, meta = build_weakness_diagnosis_v15(
            companies_by_label,
            current_frames,
            selected_labels,
            our_label,
            search_date,
            item,
            lookback,
            float(threshold),
        )

    if diag_df.empty:
        st.warning(meta.get("message", "진단 가능한 공통 가격자료가 없습니다."))
        return

    hist_col = f"타법인최근{lookback}장"

    def _style_diag_v15(row):
        text = str(row.get("진단", ""))
        if "우리만 낮음" in text:
            return ["background-color: #fff0f0; font-weight: 700;"] * len(row)
        if "시세도 약세" in text:
            return ["background-color: #fff1e6; font-weight: 700;"] * len(row)
        if "시세 자체가 약세" in text:
            return ["background-color: #fff8df; font-weight: 700;"] * len(row)
        return [""] * len(row)

    show_cols = [
        "단위", "등급", "우리오늘", "타법인오늘", hist_col,
        "시세변동%", "우리격차%", "오늘비교법인수", "과거확보장수", "진단",
    ]

    st.dataframe(
        diag_df[show_cols].style
        .apply(_style_diag_v15, axis=1)
        .format({
            "우리오늘": "{:,.0f}",
            "타법인오늘": "{:,.0f}",
            hist_col: lambda x: "-" if pd.isna(x) else f"{x:,.0f}",
            "시세변동%": lambda x: "-" if pd.isna(x) else f"{x:+.1f}%",
            "우리격차%": lambda x: "-" if pd.isna(x) else f"{x:+.1f}%",
            "오늘비교법인수": "{:,.0f}",
            "과거확보장수": "{:,.0f}",
        }),
        use_container_width=True,
        hide_index=True,
    )

    peer_text = ", ".join(meta.get("peer_labels", []))
    st.caption(
        f"우리: {our_label} · 비교: {peer_text} · "
        f"약세 기준 {threshold}% · 휴장일은 자동 건너뜀"
    )

    if any(diag_df["등급"].astype(str) == "세부규격 미확인"):
        st.warning(
            "⚠ `세부규격 미확인` 그룹은 42/45망, 4내/4수 등 실제 규격이 섞였을 수 있어 "
            "진단도 참고용입니다."
        )


def build_excel_bytes(selected_frames: List[pd.DataFrame]) -> bytes:
    all_df = pd.concat(selected_frames, ignore_index=True)

    summary_rows = []
    group_cols = ["시장", "법인", "품목(품종)", "단위", "등급"]

    for keys, g in all_df.groupby(group_cols, sort=False):
        prices = g["경락가"].astype(int).tolist()
        summary_rows.append(
            {
                "시장": keys[0],
                "법인": keys[1],
                "품목(품종)": keys[2],
                "단위": keys[3],
                "등급": keys[4],
                "건수": len(prices),
                "머리": max(prices),
                "중간": actual_middle_price(prices),
                "꼬리": min(prices),
                "평균": round(mean(prices)),
            }
        )

    summary_df = pd.DataFrame(summary_rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "머리중간꼬리"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    sub_fill = PatternFill("solid", fgColor="D9EAF7")

    for c, name in enumerate(summary_df.columns, start=1):
        cell = ws.cell(1, c, name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(summary_df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(r_idx, c_idx, value)

    for col_name in ("건수", "머리", "중간", "꼬리", "평균"):
        if col_name in summary_df.columns:
            col_idx = summary_df.columns.get_loc(col_name) + 1
            for row in range(2, len(summary_df) + 2):
                ws.cell(row, col_idx).number_format = "#,##0"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 9, "B": 15, "C": 30, "D": 10, "E": 12,
        "F": 9, "G": 12, "H": 12, "I": 12, "J": 12
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    raw = wb.create_sheet("선택법인원자료")
    raw_cols = [
        "시장", "법인", "법인코드", "페이지", "번호",
        "품목(품종)", "단위", "등급", "경락가", "출하지",
    ]

    for c, name in enumerate(raw_cols, start=1):
        cell = raw.cell(1, c, name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(all_df[raw_cols].itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            raw.cell(r_idx, c_idx, value)
        raw.cell(r_idx, 9).number_format = "#,##0"

    raw.freeze_panes = "A2"
    raw.auto_filter.ref = raw.dimensions
    raw.column_dimensions["A"].width = 9
    raw.column_dimensions["B"].width = 15
    raw.column_dimensions["C"].width = 15
    raw.column_dimensions["D"].width = 9
    raw.column_dimensions["E"].width = 9
    raw.column_dimensions["F"].width = 30
    raw.column_dimensions["G"].width = 10
    raw.column_dimensions["H"].width = 12
    raw.column_dimensions["I"].width = 12
    raw.column_dimensions["J"].width = 28

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()



# =========================================================
# 공개 세부규격 흔적 검사기 v8
# - 품목 검색폼 자동 감지/제출
# - 검색결과의 품목 상세링크 추적
# - 품목별 현장규격 키워드 분리
# - 공개 GET/POST만 사용, 로그인/권한우회 없음
# =========================================================
TRACE_COMPANIES = {
    "가락 · 대아청과": {
        "urls": [
            "https://www.dagreen.co.kr/market_price_new/daily_market.asp",
            "https://www.dagreen.co.kr/",
        ]
    },
    "가락 · 동화청과": {
        "urls": ["https://www.donghwafp.com/"]
    },
    "가락 · 중앙청과": {
        "urls": ["https://www.ejoongang.co.kr/"]
    },
    "가락 · 한국청과": {
        "urls": ["https://www.hkck.co.kr/"]
    },
    "가락 · 서울청과": {
        "urls": ["http://www.sfvc.co.kr/"]
    },
    "가락 · 농협가락공판장": {
        "urls": ["https://newgp.nonghyup.com/"]
    },
    "강서 · 서부청과": {
        "urls": ["https://www.sbbot.com/mobile_web/m_itemList.do?gubn=vege"]
    },
    "강서 · 강서청과": {
        "urls": ["http://www.ksfresh.kr/"]
    },
    "강서 · 농협강서공판장": {
        "urls": ["https://newgp.nonghyup.com/"]
    },
}

DISCOVERY_WORDS = (
    "시세", "경매", "가격", "유통", "시장", "품목", "거래", "동향",
    "price", "auction", "market", "result", "trend", "item", "product"
)

SEARCH_WORDS = (
    "검색", "조회", "품목", "품종", "상품", "keyword", "search", "query",
    "item", "product", "pum", "jong", "name", "sch", "find"
)

DANGEROUS_WORDS = (
    "login", "logout", "member", "join", "write", "edit", "delete", "remove",
    "insert", "update", "upload", "mail", "send", "order", "reserve", "payment",
    "로그인", "회원", "가입", "글쓰기", "수정", "삭제", "주문", "예약", "결제"
)

GENERIC_FIELD_RE = re.compile(
    r'(?i)\b(?:size|spec|standard|grade|class|level|unit)'
    r'(?:[_-]?(?:cd|code|nm|name|no|id|seq))?\b'
)

PRICE_RE = re.compile(r'(?<!\d)(?:\d{1,3}(?:,\d{3})+|\d{4,7})\s*원?')

CABBAGE_RE = re.compile(
    r'(?<!\d)(?:36|38|40|42|43|45|47|48|50|52|55)'
    r'(?:\s*[-~]\s*(?:36|38|40|42|43|45|47|48|50|52|55))?\s*망'
)
FOUR_RE = re.compile(r'(?:4\s*수|4\s*개(?:들이)?|4\s*포기)')

SWEET_POTATO_RE = re.compile(
    r'(?:긴긴특|긴긴상|긴긴중|긴긴하|긴긴소|긴왕|긴특|긴상|긴중|긴하|긴소|'
    r'공특|공상|공중|공하|왕왕|특상|상중|파지|B품)',
    re.I,
)
POTATO_RE = re.compile(
    r'(?:왕왕|왕특|특대|대특|특상|상중|중하|파지|B품)',
    re.I,
)
SHIITAKE_RE = re.compile(
    r'(?:\bC\s*/?\s*T\b|\bCT\b|P\s*[-/]?\s*BOX|PP\s*대|상자)',
    re.I,
)

SEARCH_INPUT_RE = re.compile(
    r'(?i)(?:keyword|search|query|item|product|pum|pummok|pumjong|jong|'
    r'goods|name|sch|find|word|key)'
)


def _is_public_http_url(url: str) -> Tuple[bool, str]:
    """사설/로컬 주소는 서버측 요청 대상에서 제외."""
    try:
        p = urlparse(url.strip())
        if p.scheme not in ("http", "https"):
            return False, "http 또는 https 주소만 사용할 수 있습니다."
        if not p.hostname:
            return False, "도메인을 확인할 수 없습니다."

        host = p.hostname.lower()
        if host in {"localhost", "localhost.localdomain"} or host.endswith(".local"):
            return False, "로컬 주소는 사용할 수 없습니다."

        infos = socket.getaddrinfo(
            host,
            443 if p.scheme == "https" else 80,
            type=socket.SOCK_STREAM,
        )
        if not infos:
            return False, "도메인의 IP를 확인할 수 없습니다."

        for info in infos:
            ip_text = info[4][0]
            try:
                ip = ipaddress.ip_address(ip_text)
            except ValueError:
                continue
            if (
                ip.is_private or ip.is_loopback or ip.is_link_local
                or ip.is_reserved or ip.is_multicast or ip.is_unspecified
            ):
                return False, "공개 인터넷 주소가 아닌 대상은 검사하지 않습니다."
        return True, ""
    except socket.gaierror:
        return False, "도메인 DNS를 확인하지 못했습니다."
    except Exception as ex:
        return False, f"주소 확인 오류: {ex}"


def _same_host(a: str, b: str) -> bool:
    return (urlparse(a).hostname or "").lower() == (urlparse(b).hostname or "").lower()


def _safe_public_request(
    session: requests.Session,
    method: str,
    url: str,
    params=None,
    data=None,
    timeout: int = 16,
):
    """공개 URL만 요청. 리다이렉트 대상도 매번 재검증."""
    current = url.strip()
    req_method = method.upper()
    req_params = params
    req_data = data

    for _ in range(6):
        ok, reason = _is_public_http_url(current)
        if not ok:
            raise ValueError(reason)

        r = session.request(
            req_method,
            current,
            params=req_params,
            data=req_data,
            timeout=timeout,
            allow_redirects=False,
        )

        if r.status_code in (301, 302, 303, 307, 308):
            loc = r.headers.get("Location")
            if not loc:
                r.raise_for_status()
                return r
            current = urljoin(current, loc)
            # 검색 POST 뒤 일반적인 302/303은 GET으로 따라감
            if r.status_code in (301, 302, 303):
                req_method = "GET"
                req_params = None
                req_data = None
            continue

        r.raise_for_status()
        return r

    raise RuntimeError("리다이렉트가 너무 많습니다.")


def _safe_public_get(session: requests.Session, url: str, timeout: int = 16):
    return _safe_public_request(session, "GET", url, timeout=timeout)


def _decode_public_response(r: requests.Response) -> str:
    enc = (r.encoding or "").lower()
    if enc and enc not in {"iso-8859-1", "ascii"}:
        return r.text
    try:
        return r.content.decode(r.apparent_encoding or "utf-8", errors="replace")
    except Exception:
        return r.content.decode("utf-8", errors="replace")


def _clean_html_text(raw: str) -> str:
    try:
        return re.sub(
            r"\s+",
            " ",
            BeautifulSoup(raw, "html.parser").get_text(" ", strip=True),
        ).strip()
    except Exception:
        return re.sub(r"\s+", " ", raw).strip()


def _item_windows(text: str, item: str, radius: int = 2200) -> List[Tuple[int, int]]:
    if not item:
        return [(0, len(text))]
    positions = [m.start() for m in re.finditer(re.escape(item), text, flags=re.I)]
    if not positions:
        return []

    ranges = [(max(0, p-radius), min(len(text), p+len(item)+radius)) for p in positions[:120]]
    ranges.sort()
    merged: List[Tuple[int, int]] = []
    for a, b in ranges:
        if not merged or a > merged[-1][1]:
            merged.append((a, b))
        else:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
    return merged


def _patterns_for_item(item: str):
    text = (item or "").replace(" ", "")
    patterns = []
    if "양배추" in text or text == "배추":
        patterns += [("망규격", CABBAGE_RE), ("4수/4개", FOUR_RE)]
    if "고구마" in text:
        patterns += [("고구마 세부등급", SWEET_POTATO_RE)]
    if "감자" in text:
        patterns += [("감자 세부등급", POTATO_RE)]
    if "표고" in text:
        patterns += [("표고 포장규격", SHIITAKE_RE)]
    # 필드+숫자는 모든 품목에서 후보로 보되 실제 품목 주변일 때만
    patterns += [(
        "필드+숫자",
        re.compile(
            r'(?i)(?:size|spec|standard|grade|class|level)'
            r'[^<>{}\n]{0,100}(?:36|38|40|42|43|45|47|48|50|52|55)'
        )
    )]
    return patterns


def _add_hit(
    hits: List[Dict[str, str]],
    source: str,
    item: str,
    kind: str,
    match: str,
    full_text: str,
    start: int,
    end: int,
):
    left = max(0, start - 260)
    right = min(len(full_text), end + 260)
    context = _clean_html_text(full_text[left:right])

    prices = []
    for pm in PRICE_RE.finditer(context):
        raw = pm.group(0).strip()
        digits = re.sub(r"\D", "", raw)
        if digits:
            n = int(digits)
            if 100 <= n <= 10000000:
                prices.append(raw)
    prices = list(dict.fromkeys(prices))[:10]

    hits.append({
        "품목": item,
        "종류": kind,
        "일치": match,
        "가격후보": " / ".join(prices),
        "출처": source,
        "주변내용": context[:1500],
    })


def _scan_item_context(text: str, source: str, item: str):
    ranges = _item_windows(text, item)
    if item and not ranges:
        return [], False

    hits: List[Dict[str, str]] = []
    patterns = _patterns_for_item(item)

    for a, b in ranges:
        segment = text[a:b]
        for kind, pat in patterns:
            for m in pat.finditer(segment):
                _add_hit(
                    hits, source, item, kind, m.group(0),
                    text, a + m.start(), a + m.end()
                )

        # 품목 주변 숫자코드 후보
        for m in re.finditer(r"\[[1-9]\]", segment):
            _add_hit(
                hits, source, item, "숫자코드 후보", m.group(0),
                text, a + m.start(), a + m.end()
            )

    # DOM 행/리스트 단위 보강: 품목과 규격이 같은 행에 있으면 직접 잡음
    try:
        soup = BeautifulSoup(text, "html.parser")
        for tag in soup.find_all(["tr", "li", "article", "p"]):
            row_text = tag.get_text(" ", strip=True)
            if item and item.lower() not in row_text.lower():
                continue
            for kind, pat in patterns:
                for m in pat.finditer(row_text):
                    fake = row_text
                    _add_hit(hits, source, item, kind, m.group(0), fake, m.start(), m.end())
    except Exception:
        pass

    uniq = []
    seen = set()
    for row in hits:
        key = (row["종류"], row["일치"], row["출처"], row["주변내용"])
        if key not in seen:
            seen.add(key)
            uniq.append(row)
    return uniq, True


def _schema_candidates(text: str, source: str):
    rows = []
    for m in GENERIC_FIELD_RE.finditer(text):
        left = max(0, m.start()-120)
        right = min(len(text), m.end()+120)
        rows.append({
            "종류": "필드명",
            "일치": m.group(0),
            "출처": source,
            "주변내용": _clean_html_text(text[left:right])[:800],
        })
        if len(rows) >= 100:
            break
    return rows


def _discover_links(html_text: str, base_url: str, item: str = "", limit: int = 16):
    """같은 도메인의 시세/경매/품목 관련 링크 후보."""
    soup = BeautifulSoup(html_text, "html.parser")
    host = (urlparse(base_url).hostname or "").lower()
    scored = []

    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        if not href or href.startswith(("javascript:", "#", "mailto:", "tel:")):
            continue
        u = urljoin(base_url, href)
        if (urlparse(u).hostname or "").lower() != host:
            continue
        label = (a.get_text(" ", strip=True) + " " + u).lower()
        score = 0
        if item and item.lower() in label:
            score += 20
        score += sum(2 for w in DISCOVERY_WORDS if w.lower() in label)
        if "view" in label or "detail" in label:
            score += 1
        if score:
            scored.append((score, u))

    scored.sort(key=lambda x: (-x[0], len(x[1])))
    out = []
    seen = set()
    for _, u in scored:
        if u not in seen:
            seen.add(u)
            out.append(u)
        if len(out) >= limit:
            break
    return out


def _form_is_safe_search(form, base_url: str) -> bool:
    method = (form.get("method") or "get").lower()
    if method not in ("get", "post"):
        return False

    action = urljoin(base_url, form.get("action") or base_url)
    if not _same_host(action, base_url):
        return False

    desc = (
        action + " " + form.get_text(" ", strip=True) + " "
        + " ".join(str(v) for tag in form.find_all(["input", "select"])
                   for v in [tag.get("name", ""), tag.get("id", ""),
                             tag.get("placeholder", ""), tag.get("type", "")])
    ).lower()

    if any(w in desc for w in DANGEROUS_WORDS):
        return False
    if form.find("input", attrs={"type": re.compile(r"password|file", re.I)}):
        return False

    return any(w.lower() in desc for w in SEARCH_WORDS)


def _choose_select_value(select_tag):
    options = select_tag.find_all("option")
    if not options:
        return ""

    # 제목/품목/전체 검색 계열을 우선
    for opt in options:
        txt = opt.get_text(" ", strip=True).lower()
        val = opt.get("value", "")
        if val and any(k in txt for k in ("품목", "제목", "전체", "상품", "title", "item", "name")):
            return val

    selected = select_tag.find("option", selected=True)
    if selected is not None:
        return selected.get("value", "")

    for opt in options:
        val = opt.get("value", "")
        if val:
            return val
    return ""


def _build_search_submission(form, base_url: str, item: str):
    if not _form_is_safe_search(form, base_url):
        return None

    action = urljoin(base_url, form.get("action") or base_url)
    method = (form.get("method") or "get").upper()
    payload = {}
    candidate_names = []

    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        typ = (inp.get("type") or "text").lower()
        if typ in {"submit", "button", "image", "reset", "file", "password"}:
            continue
        if typ in {"checkbox", "radio"} and not inp.has_attr("checked"):
            continue
        value = inp.get("value", "")
        payload[name] = value

        desc = " ".join([
            name, inp.get("id", ""), inp.get("placeholder", ""),
            inp.get("title", ""), inp.get("aria-label", "")
        ]).lower()
        if typ in {"text", "search", ""} and (
            SEARCH_INPUT_RE.search(desc)
            or any(w.lower() in desc for w in SEARCH_WORDS)
        ):
            candidate_names.append((name, desc))

    for sel in form.find_all("select"):
        name = sel.get("name")
        if name:
            payload[name] = _choose_select_value(sel)

    if not candidate_names:
        # 검색폼으로 판정됐는데 이름이 애매한 text input이 하나뿐이면 사용
        texts = []
        for inp in form.find_all("input"):
            typ = (inp.get("type") or "text").lower()
            name = inp.get("name")
            if name and typ in {"text", "search", ""}:
                texts.append(name)
        if len(texts) == 1:
            candidate_names = [(texts[0], texts[0])]

    if not candidate_names:
        return None

    # 가장 검색스러운 입력칸 하나에 품목 입력
    scored = []
    for name, desc in candidate_names:
        score = sum(1 for w in SEARCH_WORDS if w.lower() in desc)
        if re.search(r'(?i)keyword|search|query|pum|item|product|name', desc):
            score += 3
        scored.append((score, name))
    scored.sort(reverse=True)
    payload[scored[0][1]] = item

    return method, action, payload


def _submit_search_forms(session, html_text: str, base_url: str, item: str, limit: int = 4):
    """공개 검색폼만 실제 품목명으로 조회."""
    soup = BeautifulSoup(html_text, "html.parser")
    results = []
    signatures = set()

    for form in soup.find_all("form"):
        sub = _build_search_submission(form, base_url, item)
        if not sub:
            continue
        method, action, payload = sub
        signature = (method, action, tuple(sorted(payload.items())))
        if signature in signatures:
            continue
        signatures.add(signature)

        try:
            if method == "GET":
                r = _safe_public_request(session, "GET", action, params=payload, timeout=18)
            else:
                r = _safe_public_request(session, "POST", action, data=payload, timeout=18)
            text = _decode_public_response(r)
            results.append({
                "url": r.url or action,
                "html": text,
                "method": method,
                "action": action,
                "payload": payload,
            })
        except Exception:
            continue

        if len(results) >= limit:
            break

    return results


@st.cache_data(ttl=300, show_spinner=False)
def deep_scan_public_site(start_url: str, item: str, deep: bool = True):
    """
    1) 시작페이지
    2) 시세/경매 링크
    3) 안전한 검색폼에 품목 입력
    4) 검색결과에서 품목명 상세링크 추적
    """
    session = make_session()

    queue = [(start_url, 0, "시작")]
    seen_urls = set()
    hits = []
    schema = []
    pages = []
    errors = []
    form_queries = []
    detail_pages = 0
    total_bytes = 0
    item_found_any = False

    max_pages = 24 if deep else 10
    max_depth = 2 if deep else 1

    while queue and len(pages) < max_pages:
        url, depth, via = queue.pop(0)
        if url in seen_urls:
            continue
        seen_urls.add(url)

        try:
            r = _safe_public_get(session, url, timeout=18)
            text = _decode_public_response(r)
            final_url = r.url or url
            total_bytes += len(r.content)
            pages.append({"url": final_url, "경로": via, "깊이": depth})

            phits, item_found = _scan_item_context(text, final_url, item)
            if item_found:
                item_found_any = True
                hits.extend(phits)
            schema.extend(_schema_candidates(text, final_url))

            # 공개 검색폼 자동 제출
            if deep and depth <= 1:
                submitted = _submit_search_forms(session, text, final_url, item, limit=4)
                for sub in submitted:
                    form_queries.append({
                        "method": sub["method"],
                        "action": sub["action"],
                        "url": sub["url"],
                    })
                    stext = sub["html"]
                    surl = sub["url"]
                    shits, sfound = _scan_item_context(stext, surl, item)
                    if sfound:
                        item_found_any = True
                        hits.extend(shits)
                    schema.extend(_schema_candidates(stext, surl))

                    if surl not in seen_urls:
                        pages.append({"url": surl, "경로": "검색폼", "깊이": depth + 1})
                        seen_urls.add(surl)
                        total_bytes += len(stext.encode("utf-8", errors="ignore"))

                    # 검색결과에서 품목명이 붙은 링크를 우선 상세 추적
                    for link in _discover_links(stext, surl, item=item, limit=10):
                        if link not in seen_urls and len(queue) + len(pages) < max_pages + 10:
                            queue.insert(0, (link, min(depth+1, max_depth), "검색결과 상세"))
                            detail_pages += 1

            # 일반 시세/경매 관련 링크 탐색
            if depth < max_depth:
                for link in _discover_links(text, final_url, item=item, limit=14):
                    if link not in seen_urls:
                        queue.append((link, depth + 1, "관련링크"))

        except Exception as ex:
            errors.append(f"{url} :: {ex}")

    # dedupe
    uniq_hits = []
    seen = set()
    for row in hits:
        key = (row["종류"], row["일치"], row["출처"], row["주변내용"])
        if key not in seen:
            seen.add(key)
            uniq_hits.append(row)

    uniq_schema = []
    seen_schema = set()
    for row in schema:
        key = (row["일치"].lower(), row["출처"], row["주변내용"])
        if key not in seen_schema:
            seen_schema.add(key)
            uniq_schema.append(row)

    # 같은 사이트에서 너무 많은 일반 필드명은 잘라냄
    uniq_schema = uniq_schema[:200]

    return {
        "item_found": item_found_any,
        "hits": uniq_hits,
        "schema": uniq_schema,
        "pages": pages,
        "errors": errors,
        "form_queries": form_queries,
        "detail_pages": detail_pages,
        "bytes": total_bytes,
    }


def _merge_results(parts):
    merged = {
        "item_found": False,
        "hits": [],
        "schema": [],
        "pages": [],
        "errors": [],
        "form_queries": [],
        "detail_pages": 0,
        "bytes": 0,
    }
    for r in parts:
        merged["item_found"] = merged["item_found"] or r["item_found"]
        for key in ("hits", "schema", "pages", "errors", "form_queries"):
            merged[key].extend(r[key])
        merged["detail_pages"] += r["detail_pages"]
        merged["bytes"] += r["bytes"]

    # pages/hits dedupe
    pseen = set()
    merged["pages"] = [
        p for p in merged["pages"]
        if not (p["url"] in pseen or pseen.add(p["url"]))
    ]
    hseen = set()
    merged["hits"] = [
        h for h in merged["hits"]
        if not (
            (h["종류"], h["일치"], h["출처"], h["주변내용"]) in hseen
            or hseen.add((h["종류"], h["일치"], h["출처"], h["주변내용"]))
        )
    ]
    qseen = set()
    merged["form_queries"] = [
        q for q in merged["form_queries"]
        if not (
            (q["method"], q["action"], q["url"]) in qseen
            or qseen.add((q["method"], q["action"], q["url"]))
        )
    ]
    return merged


def _render_scan_result(label: str, result: Dict[str, object], item: str):
    hits = pd.DataFrame(result["hits"])
    schema = pd.DataFrame(result["schema"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("직접단서", f"{len(hits):,}")
    c2.metric("페이지", f'{len(result["pages"]):,}')
    c3.metric("검색폼", f'{len(result["form_queries"]):,}')
    c4.metric("오류", f'{len(result["errors"]):,}')

    if not result["item_found"]:
        st.warning(
            f"검사한 공개 응답에서 **{item}** 자체를 찾지 못했습니다. "
            "다른 품목의 규격은 결과에 섞지 않았습니다."
        )
    elif hits.empty:
        st.info(
            f"**{item}**은 찾았지만, 그 주변에서 현재 정의한 세부규격 직접단서는 못 찾았습니다."
        )
    else:
        st.subheader("🎯 품목과 연결된 직접 단서")
        cols = ["품목", "종류", "일치", "가격후보", "출처", "주변내용"]
        st.dataframe(hits[cols], use_container_width=True, hide_index=True)

    if result["form_queries"]:
        with st.expander(f"실제로 조회한 공개 검색폼 {len(result['form_queries'])}건"):
            st.dataframe(pd.DataFrame(result["form_queries"]), use_container_width=True, hide_index=True)

    if result["pages"]:
        with st.expander(f"검사한 공개 페이지 {len(result['pages'])}개"):
            st.dataframe(pd.DataFrame(result["pages"]), use_container_width=True, hide_index=True)

    if not schema.empty:
        with st.expander(f"사이트 구조/필드 후보 {len(schema)}건 (직접규격 아님)"):
            st.dataframe(schema, use_container_width=True, hide_index=True)

    if result["errors"]:
        with st.expander(f"접속 실패/건너뜀 {len(result['errors'])}건"):
            st.code("\n".join(result["errors"][:40]))

    lines = [
        f"법인: {label}",
        f"품목: {item}",
        f"직접단서: {len(result['hits'])}",
        f"검사페이지: {len(result['pages'])}",
        f"검색폼: {len(result['form_queries'])}",
        "",
    ]
    for h in result["hits"]:
        lines += [
            f'[{h["종류"]}] {h["일치"]} / 가격후보={h["가격후보"]}',
            h["출처"],
            h["주변내용"],
            "",
        ]

    st.download_button(
        f"{label} 결과 TXT 저장",
        data="\n".join(lines).encode("utf-8-sig"),
        file_name=f"세부규격_{re.sub(r'[^0-9A-Za-z가-힣]+','_',label)}_{item}.txt",
        mime="text/plain",
        use_container_width=True,
        key=f"v8download::{label}::{item}",
    )


def render_trace_scanner():
    st.title("🔍 세부규격 정밀찾기")
    st.caption(
        "단순 홈페이지 검색이 아니라 **공개 검색폼에 품목을 실제 입력해 조회하고, "
        "그 결과의 품목 상세페이지까지 따라가며** 세부규격 흔적을 찾습니다."
    )

    mode = st.radio(
        "검사 방식",
        ["법인별 정밀 자동검사", "직접 주소 정밀검사"],
        horizontal=True,
    )

    item = st.text_input(
        "찾을 품목",
        value="양배추",
        placeholder="예: 양배추, 고구마, 감자, 표고",
    ).strip()

    precision = st.toggle(
        "정밀 모드",
        value=True,
        help="켜면 공개 검색폼 제출 + 관련 상세페이지를 최대 2단계까지 확인합니다.",
    )

    if mode == "법인별 정밀 자동검사":
        defaults = ["가락 · 대아청과", "가락 · 동화청과", "가락 · 중앙청과"]
        selected = st.multiselect(
            "검사할 법인",
            list(TRACE_COMPANIES.keys()),
            default=defaults,
        )

        st.caption(
            "공개 페이지/공개 검색폼만 사용합니다. 로그인, 권한 우회, 비공개 주소 추측은 하지 않습니다."
        )

        if st.button("선택 법인 정밀검사", type="primary", use_container_width=True):
            if not item:
                st.error("품목을 입력해 주세요.")
                return
            if not selected:
                st.error("법인을 하나 이상 선택해 주세요.")
                return

            results = []
            progress = st.progress(0.0)
            status = st.empty()

            for idx, label in enumerate(selected, start=1):
                status.caption(f"{label} · {item} 검색폼/상세페이지 검사 중...")
                parts = []
                for start_url in TRACE_COMPANIES[label]["urls"]:
                    try:
                        parts.append(deep_scan_public_site(start_url, item, deep=precision))
                    except Exception as ex:
                        parts.append({
                            "item_found": False, "hits": [], "schema": [], "pages": [],
                            "errors": [f"{start_url} :: {ex}"], "form_queries": [],
                            "detail_pages": 0, "bytes": 0,
                        })
                results.append((label, _merge_results(parts)))
                progress.progress(idx / len(selected))

            status.empty()
            progress.empty()

            summary = []
            for label, r in results:
                summary.append({
                    "법인": label,
                    "접속": "✅" if r["pages"] else "❌",
                    "품목발견": "✅" if r["item_found"] else "—",
                    "직접단서": len(r["hits"]),
                    "검색폼": len(r["form_queries"]),
                    "검사페이지": len(r["pages"]),
                    "오류": len(r["errors"]),
                })

            st.subheader("법인별 정밀 결과")
            st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

            for label, r in results:
                with st.expander(
                    f'{label} · 직접단서 {len(r["hits"])}건 · 검색폼 {len(r["form_queries"])}건',
                    expanded=(len(r["hits"]) > 0),
                ):
                    _render_scan_result(label, r, item)

    else:
        url = st.text_input("공개 페이지 주소", placeholder="https://...").strip()
        if st.button("주소 정밀검사", type="primary", use_container_width=True):
            if not item:
                st.error("품목을 입력해 주세요.")
                return
            if not url:
                st.error("주소를 입력해 주세요.")
                return

            with st.spinner("공개 검색폼과 상세페이지까지 검사 중..."):
                try:
                    result = deep_scan_public_site(url, item, deep=precision)
                except Exception as ex:
                    st.error("이 공개 사이트를 검사하지 못했습니다.")
                    st.code(str(ex))
                    return
            _render_scan_result("직접 주소", result, item)

    st.info(
        "품목별로 찾는 단서를 분리했습니다. 예: 양배추=42/45망·4수, "
        "고구마=긴특·긴상·공중·공하 등, 표고=C/T·P-BOX·PP대. "
        "다른 품목의 규격은 결과에 섞지 않습니다."
    )




# =========================================================
# 농림축산식품부 도매시장 원천데이터 - "등외 해체" 엔진
# =========================================================
MAFRA_API_HOST = "http://211.237.50.150:7080/openapi"

MAFRA_SERVICES = {
    "settlement": "Grid_20240625000000000653_1",  # 정산 가격 정보: STD/SIZENAME/LVNAME
    "live": "Grid_20240625000000000654_1",     # 도매시장 실시간 경락 정보
    "raw": "Grid_20240625000000000655_1",      # 도매시장 원천데이터 정산 가격
    "market": "Grid_20240625000000000661_1",   # 도매시장 코드
    "corp": "Grid_20240626000000000662_1",     # 법인 코드
    "grade": "Grid_20240626000000000663_1",    # 등급 코드
    "unit": "Grid_20240626000000000664_1",     # 단위 코드
    "pack": "Grid_20240626000000000665_1",     # 포장 코드
    "size": "Grid_20240626000000000666_1",     # 크기 코드
    "item": "Grid_20240626000000000668_1",     # 품목 코드
}

SEOUL_WHOLESALE_MARKETS = {
    "가락 · 서울가락": "110001",
    "강서 · 서울강서": "110008",
}


def _mafra_key_path(api_key: str) -> str:
    # 이미 URL-encoded 된 키를 붙여 넣어도 한 번 풀었다가 안전하게 재인코딩
    key = unquote((api_key or "").strip())
    return quote(key, safe="")


def _xml_rows_and_meta(text: str):
    try:
        root = ET.fromstring(text)
    except Exception as e:
        raise RuntimeError(f"API XML 해석 실패: {e}")

    result = root.find("result")
    if result is not None:
        code = (result.findtext("code") or "").strip()
        message = (result.findtext("message") or "").strip()
        if code and code != "INFO-000":
            raise RuntimeError(f"{code}: {message or 'API 오류'}")

    total_text = (root.findtext("totalCnt") or "0").strip()
    try:
        total = int(float(total_text))
    except Exception:
        total = 0

    rows = []
    for row in root.findall("row"):
        d = {}
        for child in list(row):
            d[child.tag] = (child.text or "").strip()
        rows.append(d)
    return rows, total


def _mafra_request_xml(
    api_key: str,
    service: str,
    start: int,
    end: int,
    params: Dict[str, str] | None = None,
    timeout: int = 25,
):
    key_path = _mafra_key_path(api_key)
    url = f"{MAFRA_API_HOST}/{key_path}/xml/{service}/{start}/{end}"

    session = make_session()
    try:
        r = session.get(url, params=params or {}, timeout=timeout)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(
            "농식품 공공데이터 API 접속 실패. "
            "Streamlit Cloud가 211.237.50.150:7080 연결을 막는 경우가 있습니다. "
            f"원인: {e}"
        )

    return _xml_rows_and_meta(_decode_public_response(r))


@st.cache_data(ttl=3600, show_spinner=False)
def mafra_code_map(api_key: str, service_key: str) -> Dict[str, str]:
    service = MAFRA_SERVICES[service_key]
    first_rows, total = _mafra_request_xml(api_key, service, 1, 1000)

    rows = list(first_rows)
    start = 1001
    while start <= total and start <= 10000:
        end = min(start + 999, total)
        part, _ = _mafra_request_xml(api_key, service, start, end)
        rows.extend(part)
        start = end + 1

    result = {}
    for row in rows:
        code = str(row.get("CODEID", "")).strip()
        name = str(row.get("CODENAME", "")).strip()
        if code:
            result[code] = name
    return result


@st.cache_data(ttl=300, show_spinner=False)
def mafra_fetch_raw_company(
    api_key: str,
    sale_date: str,
    market_code: str,
    corp_code: str,
    max_rows: int = 6000,
) -> Tuple[List[Dict[str, str]], int]:
    service = MAFRA_SERVICES["raw"]
    params = {
        "SALEDATE": sale_date,
        "WHSALCD": market_code,
        "CMPCD": corp_code,
    }

    rows = []
    start = 1
    total = 0
    while start <= max_rows:
        end = min(start + 999, max_rows)
        part, total = _mafra_request_xml(api_key, service, start, end, params=params, timeout=35)
        rows.extend(part)

        if not part or len(rows) >= total:
            break
        start = end + 1

    return rows, total


@st.cache_data(ttl=90, show_spinner=False)
def mafra_fetch_live_company(
    api_key: str,
    sale_date: str,
    market_code: str,
    corp_code: str,
    max_rows: int = 5000,
) -> Tuple[List[Dict[str, str]], int]:
    service = MAFRA_SERVICES["live"]
    params = {
        "SALEDATE": sale_date,
        "WHSALCD": market_code,
        "CMPCD": corp_code,
    }

    rows = []
    start = 1
    total = 0
    while start <= max_rows:
        end = min(start + 999, max_rows)
        part, total = _mafra_request_xml(api_key, service, start, end, params=params, timeout=30)
        rows.extend(part)

        if not part or len(rows) >= total:
            break
        start = end + 1

    return rows, total


def _num(v, default=0):
    try:
        if v is None or str(v).strip() == "":
            return default
        return float(str(v).replace(",", ""))
    except Exception:
        return default


def _int_if_whole(v):
    n = _num(v, 0)
    if float(n).is_integer():
        return int(n)
    return n


def _code_name(code_map: Dict[str, str], code: str, unknown_prefix: str = "코드") -> str:
    c = str(code or "").strip()
    if not c:
        return "-"
    return code_map.get(c, f"{unknown_prefix}:{c}")


def _raw_to_dataframe(
    rows: List[Dict[str, str]],
    market_name: str,
    corp_name: str,
    grade_map: Dict[str, str],
    unit_map: Dict[str, str],
    pack_map: Dict[str, str],
    size_map: Dict[str, str],
) -> pd.DataFrame:
    data = []

    for r in rows:
        unit_name = _code_name(unit_map, r.get("DANCD", ""), "단위")
        pack_name = _code_name(pack_map, r.get("POJCD", ""), "포장")
        size_name = _code_name(size_map, r.get("SIZECD", ""), "크기")
        grade_name = _code_name(grade_map, r.get("LVCD", ""), "등급")

        danq = str(r.get("DANQ", "")).strip()
        parts = []
        if danq and danq not in {"0", "0.0"}:
            parts.append(f"{danq}{'' if unit_name == '-' else unit_name}")
        elif unit_name != "-":
            parts.append(unit_name)
        if pack_name != "-":
            parts.append(pack_name)
        if size_name != "-":
            parts.append(size_name)

        data.append({
            "장일자": r.get("SALEDATE", ""),
            "시장": market_name,
            "법인": corp_name,
            "원표": r.get("SEQ", ""),
            "경매순서": r.get("NO1", ""),
            "품목": r.get("PUMNAME", ""),
            "품종": r.get("GOODNAME", ""),
            "단량": _num(r.get("DANQ", ""), 0),
            "단위": unit_name,
            "포장": pack_name,
            "크기/규격": size_name,
            "등급": grade_name,
            "규격표시": " ".join(parts).strip() or "-",
            "물량": _num(r.get("QTY", ""), 0),
            "경락가": _num(r.get("COST", ""), 0),
            "산지": r.get("SANNAME", ""),
            "낙찰시간": r.get("SBIDTIME", ""),
            "매매방법코드": r.get("MMCD", ""),
            "크기코드": r.get("SIZECD", ""),
            "등급코드": r.get("LVCD", ""),
            "포장코드": r.get("POJCD", ""),
            "단위코드": r.get("DANCD", ""),
        })

    return pd.DataFrame(data)


def _live_to_dataframe(rows: List[Dict[str, str]]) -> pd.DataFrame:
    data = []
    for r in rows:
        data.append({
            "장일자": r.get("SALEDATE", ""),
            "시장": r.get("WHSALNAME", ""),
            "법인": r.get("CMPNAME", ""),
            "부류": r.get("LARGENAME", ""),
            "품목": r.get("MIDNAME", ""),
            "품종": r.get("SMALLNAME", ""),
            "공개규격": r.get("STD", ""),
            "물량": _num(r.get("QTY", ""), 0),
            "경락가": _num(r.get("COST", ""), 0),
            "산지": r.get("SANNAME", ""),
            "낙찰시간": r.get("SBIDTIME", ""),
        })
    return pd.DataFrame(data)


def _item_filter(df: pd.DataFrame, keyword: str, cols: List[str]) -> pd.DataFrame:
    if df.empty or not keyword:
        return df
    kw = keyword.strip()
    mask = pd.Series(False, index=df.index)
    for c in cols:
        if c in df.columns:
            mask = mask | df[c].astype(str).str.contains(re.escape(kw), case=False, na=False)
    return df[mask].copy()


def _is_nonstandard_grade(name: str) -> bool:
    text = str(name or "").strip()
    if text in {"특", "상", "보통"}:
        return False
    return True


def _summary_from_decoded(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    rows = []
    group_cols = ["시장", "법인", "품목", "품종", "규격표시", "등급"]

    for keys, g in df.groupby(group_cols, dropna=False, sort=False):
        prices = [int(round(x)) for x in g["경락가"].tolist() if _num(x, 0) > 0]
        if not prices:
            continue
        rows.append({
            "시장": keys[0],
            "법인": keys[1],
            "품목": keys[2],
            "품종": keys[3],
            "규격": keys[4],
            "등급": keys[5],
            "건수": len(g),
            "물량": round(g["물량"].sum(), 2),
            "머리": max(prices),
            "중간": actual_middle_price(prices),
            "꼬리": min(prices),
            "평균": round(mean(prices)),
        })

    return pd.DataFrame(rows)


def _get_secret_api_key() -> str:
    try:
        return str(st.secrets.get("MAFRA_API_KEY", "") or "")
    except Exception:
        return ""



@st.cache_data(ttl=300, show_spinner=False)
def mafra_fetch_settlement_market(
    api_key: str,
    sale_date: str,
    market_code: str,
    max_rows: int = 5000,
) -> Tuple[List[Dict[str, str]], int]:
    service = MAFRA_SERVICES["settlement"]
    params = {"SALEDATE": sale_date, "WHSALCD": market_code}
    rows = []
    start = 1
    total = 0
    while start <= max_rows:
        end = min(start + 999, max_rows)
        part, total = _mafra_request_xml(
            api_key, service, start, end, params=params, timeout=35
        )
        rows.extend(part)
        if not part or len(rows) >= total:
            break
        start = end + 1
    return rows, total


def _normalize_company_name_v11(name: str) -> str:
    text = re.sub(r"\s+", "", str(name or ""))
    for token in ("(주)", "㈜", "주식회사", "농협경제지주", "농협중앙회"):
        text = text.replace(token, "")
    return re.sub(r"[()（）·\-_]", "", text).lower()


def _company_score_v11(public_name: str, official_name: str, market_name: str) -> int:
    a = _normalize_company_name_v11(public_name)
    b = _normalize_company_name_v11(official_name)
    if not a or not b:
        return 0
    if a == b:
        return 100
    if a in b or b in a:
        return 85

    for key in ("서울청과", "중앙청과", "동화청과", "한국청과", "대아청과", "서부청과", "강서청과"):
        k = key.lower()
        if k in a and k in b:
            return 95

    if "농협" in a and "농협" in b:
        if market_name == "가락" and "가락" in b:
            return 95
        if market_name == "강서" and "강서" in b:
            return 95
        return 40
    return 0


def _settlement_df_v11(rows: List[Dict[str, str]]) -> pd.DataFrame:
    data = []
    for r in rows:
        data.append({
            "장일자": r.get("SALEDATE", ""),
            "시장코드": r.get("WHSALCD", ""),
            "시장": r.get("WHSALNAME", ""),
            "법인코드": r.get("CMPCD", ""),
            "법인": r.get("CMPNAME", ""),
            "부류": r.get("LARGENAME", ""),
            "품목": r.get("MIDNAME", ""),
            "품종": r.get("SMALLNAME", ""),
            "단량": _num(r.get("DANQ", ""), 0),
            "공개규격": r.get("STD", ""),
            "크기/규격": r.get("SIZENAME", ""),
            "등급": r.get("LVNAME", ""),
            "산지": r.get("SANNAME", ""),
            "총물량": _num(r.get("TOTQTY", ""), 0),
            "최저": _num(r.get("MINAMT", ""), 0),
            "최고": _num(r.get("MAXAMT", ""), 0),
            "평균": _num(r.get("AVGAMT", ""), 0),
        })
    return pd.DataFrame(data)


def _official_market_code_v11(market_name: str) -> str:
    return {"가락": "110001", "강서": "110008"}.get(market_name, "")


@st.cache_data(ttl=300, show_spinner=False)
def integrated_spec_v11(
    api_key: str,
    sale_date: str,
    market_name: str,
    public_company_name: str,
    item_keyword: str,
) -> Dict[str, object]:
    market_code = _official_market_code_v11(market_name)
    if not market_code:
        return {"summary": pd.DataFrame(), "raw": pd.DataFrame(), "message": "지원하지 않는 시장입니다."}

    # 1) official settlement summary: directly exposes SIZENAME/LVNAME/STD
    rows, total = mafra_fetch_settlement_market(api_key, sale_date, market_code)
    sdf = _settlement_df_v11(rows)
    if sdf.empty:
        return {"summary": pd.DataFrame(), "raw": pd.DataFrame(), "message": "공식 정산자료가 아직 없습니다."}

    sdf = _item_filter(sdf, item_keyword, ["품목", "품종", "부류"])
    if sdf.empty:
        return {"summary": pd.DataFrame(), "raw": pd.DataFrame(), "message": "공식 정산자료에서 해당 품목을 찾지 못했습니다."}

    scores = {
        n: _company_score_v11(public_company_name, n, market_name)
        for n in sdf["법인"].dropna().astype(str).unique()
    }
    best = max(scores.values()) if scores else 0
    if best <= 0:
        return {
            "summary": pd.DataFrame(),
            "raw": pd.DataFrame(),
            "message": "선택 법인과 공식 법인명을 연결하지 못했습니다.",
            "official_names": list(scores.keys()),
        }

    matched_names = [n for n, score in scores.items() if score == best]
    sdf = sdf[sdf["법인"].isin(matched_names)].copy()
    corp_codes = [x for x in sdf["법인코드"].dropna().astype(str).unique() if x]

    # 2) raw auction rows: actual COST/QTY + SIZECD/LVCD/POJCD
    raw_frames = []
    if corp_codes:
        grade_map = mafra_code_map(api_key, "grade")
        unit_map = mafra_code_map(api_key, "unit")
        pack_map = mafra_code_map(api_key, "pack")
        size_map = mafra_code_map(api_key, "size")

        for corp_code in corp_codes[:3]:
            raw_rows, _ = mafra_fetch_raw_company(
                api_key, sale_date, market_code, corp_code
            )
            if not raw_rows:
                continue
            rdf = _raw_to_dataframe(
                raw_rows,
                market_name,
                matched_names[0],
                grade_map,
                unit_map,
                pack_map,
                size_map,
            )
            rdf = _item_filter(rdf, item_keyword, ["품목", "품종"])
            if not rdf.empty:
                raw_frames.append(rdf)

    raw = pd.concat(raw_frames, ignore_index=True) if raw_frames else pd.DataFrame()
    return {
        "summary": sdf,
        "raw": raw,
        "message": "",
        "matched_names": matched_names,
        "settlement_total": total,
    }


def render_integrated_spec_v11(
    company: Dict[str, str],
    date_yyyymmdd: str,
    item: str,
):
    if not bool(st.session_state.get("main_spec_enrich_enabled", False)):
        return

    st.markdown("### 🧩 확인된 세부규격")
    api_key = str(st.session_state.get("main_mafra_api_key", "") or "").strip()
    if not api_key:
        st.info("위 `세부규격 자동보강 설정`에 공식 API 키를 넣으면 여기에 규격이 붙습니다.")
        return

    try:
        with st.spinner("공식 원천데이터에서 같은 장일자·법인·품목 규격 확인 중..."):
            result = integrated_spec_v11(
                api_key, date_yyyymmdd, company["market"], company["name"], item
            )
    except Exception as e:
        st.warning("공식 원천데이터 보강에 실패했습니다.")
        st.code(str(e))
        return

    raw = result.get("raw", pd.DataFrame())
    summary = result.get("summary", pd.DataFrame())
    if raw.empty and summary.empty:
        st.info(result.get("message", "") or "확인된 규격자료가 없습니다.")
        return

    names = result.get("matched_names", [])
    if names:
        st.caption(
            "공식 법인명: " + ", ".join(names) +
            " · 같은 장일자/시장/법인/품목 기준. 공사 모바일 행과 임의로 1:1 짝짓지는 않습니다."
        )

    if not raw.empty:
        st.success(f"실제 원천 거래 {len(raw):,}건의 포장·크기·등급을 확인했습니다.")
        summary_raw = _summary_from_decoded(raw)
        if not summary_raw.empty:
            st.dataframe(
                summary_raw.style.format({
                    "물량": "{:,.2f}",
                    "머리": "{:,.0f}",
                    "중간": "{:,.0f}",
                    "꼬리": "{:,.0f}",
                    "평균": "{:,.0f}",
                }),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("세부규격 확인된 실제 원천 거래행"):
            cols = [
                "시장", "법인", "품목", "품종", "규격표시", "등급",
                "물량", "경락가", "산지", "낙찰시간",
                "크기코드", "등급코드", "포장코드",
            ]
            st.dataframe(
                raw[cols].sort_values(
                    ["품종", "규격표시", "등급", "경락가"],
                    ascending=[True, True, True, False],
                ).style.format({"물량": "{:,.2f}", "경락가": "{:,.0f}"}),
                use_container_width=True,
                hide_index=True,
            )

    if not summary.empty:
        with st.expander("공식 정산 규격 요약", expanded=raw.empty):
            cols = [
                "법인", "품목", "품종", "단량", "공개규격",
                "크기/규격", "등급", "산지", "총물량", "최저", "최고", "평균",
            ]
            st.dataframe(
                summary[cols].sort_values(
                    ["품종", "공개규격", "크기/규격", "등급", "최고"],
                    ascending=[True, True, True, True, False],
                ).style.format({
                    "총물량": "{:,.2f}",
                    "최저": "{:,.0f}",
                    "최고": "{:,.0f}",
                    "평균": "{:,.0f}",
                }),
                use_container_width=True,
                hide_index=True,
            )


def render_grade_decoder():
    st.title("🧩 등외 해체")
    st.caption(
        "농림축산식품부 **도매시장 원천데이터**의 단위·포장·크기·등급 코드를 따로 해석합니다. "
        "공사 화면의 `등외`만 보는 대신 실제 원천의 `크기/규격`을 분리해서 비교합니다."
    )

    source_mode = st.radio(
        "데이터",
        ["원천 정산 · 등급+크기 해체", "실시간 경락 · 공개규격 확인"],
        horizontal=True,
    )

    saved_key = _get_secret_api_key()
    api_key = st.text_input(
        "농림축산식품 공공데이터 API 키",
        value=saved_key,
        type="password",
        placeholder="data.mafra.go.kr에서 발급받은 키",
        help="키는 채팅에 보내지 말고 이 입력칸 또는 Streamlit secrets의 MAFRA_API_KEY에 넣으세요.",
    ).strip()

    if not api_key:
        st.info(
            "이 기능은 공식 원천 API 키가 필요합니다. "
            "농림축산식품 공공데이터포털에서 `도매시장 원천데이터 정산 가격` OpenAPI를 신청한 뒤 키를 입력하세요."
        )
        st.code(
            'Streamlit secrets 예시\n\nMAFRA_API_KEY = "여기에_발급키"',
            language="toml",
        )
        return

    col1, col2 = st.columns([1, 1.35])
    with col1:
        query_date = st.date_input("장일자", value=date.today(), key="mafra_decoder_date")
    with col2:
        keyword = st.text_input(
            "품목",
            value="양배추",
            placeholder="예: 양배추, 고구마, 감자, 표고, 오이",
            key="mafra_decoder_item",
        ).strip()

    market_labels = st.multiselect(
        "시장",
        list(SEOUL_WHOLESALE_MARKETS.keys()),
        default=list(SEOUL_WHOLESALE_MARKETS.keys()),
        key="mafra_decoder_markets",
    )

    if not market_labels:
        st.warning("가락 또는 강서를 하나 이상 선택해 주세요.")
        return

    try:
        with st.spinner("공식 코드표 확인 중..."):
            corp_map = mafra_code_map(api_key, "corp")
    except Exception as e:
        st.error("농림축산식품 공공데이터 API 키 또는 연결을 확인해 주세요.")
        st.code(str(e))
        return

    market_codes = [SEOUL_WHOLESALE_MARKETS[x] for x in market_labels]

    corps = []
    for market_label, market_code in zip(market_labels, market_codes):
        market_short = market_label.split(" · ")[0]
        for code, name in corp_map.items():
            # 법인코드는 공식 샘플처럼 시장코드+2자리 형태가 일반적.
            if code.startswith(market_code):
                corps.append({
                    "label": f"{market_short} · {name}",
                    "market_label": market_short,
                    "market_code": market_code,
                    "corp_code": code,
                    "corp_name": name,
                })

    # 혹시 코드 prefix가 다르게 오는 경우를 대비해 전체 법인 검색도 가능하게
    if not corps:
        st.warning(
            "공식 법인코드표에서 선택 시장과 연결되는 법인을 자동 식별하지 못했습니다. "
            "코드표 구조가 바뀌었을 수 있습니다."
        )
        return

    corp_options = [x["label"] for x in corps]
    selected_corp_labels = st.multiselect(
        "법인",
        corp_options,
        default=corp_options,
        key="mafra_decoder_corps",
        help="처음에는 전체 법인을 두고 비교해도 됩니다. 데이터가 많으면 필요한 법인만 선택하세요.",
    )

    selected_corps = [x for x in corps if x["label"] in selected_corp_labels]

    if not selected_corps:
        st.warning("법인을 하나 이상 선택해 주세요.")
        return

    if source_mode.startswith("원천 정산"):
        grade_scope = st.radio(
            "등급 범위",
            ["특/상/보통 외만", "전체 등급"],
            horizontal=True,
            help="`특/상/보통 외만`은 4등·5등·등외·없음·기타·미상 등을 모두 잡습니다.",
        )

        run = st.button("원천 규격 해체", type="primary", use_container_width=True)
        if not run:
            st.caption(
                "정산 원천자료는 실시간 화면보다 늦게 확정될 수 있습니다. "
                "대신 단위·포장·크기·등급이 별도 코드로 있어 가장 정확하게 해체할 수 있습니다."
            )
            return

        try:
            with st.spinner("단위·포장·크기·등급 코드표 불러오는 중..."):
                grade_map = mafra_code_map(api_key, "grade")
                unit_map = mafra_code_map(api_key, "unit")
                pack_map = mafra_code_map(api_key, "pack")
                size_map = mafra_code_map(api_key, "size")
        except Exception as e:
            st.error("공식 코드표를 불러오지 못했습니다.")
            st.code(str(e))
            return

        sale_date = query_date.strftime("%Y%m%d")
        frames = []
        truncated = []

        progress = st.progress(0.0)
        status = st.empty()

        def fetch_one(c):
            rows, total = mafra_fetch_raw_company(
                api_key, sale_date, c["market_code"], c["corp_code"]
            )
            frame = _raw_to_dataframe(
                rows,
                c["market_label"],
                c["corp_name"],
                grade_map,
                unit_map,
                pack_map,
                size_map,
            )
            return c, frame, total, len(rows)

        with ThreadPoolExecutor(max_workers=min(4, len(selected_corps))) as ex:
            futures = {ex.submit(fetch_one, c): c for c in selected_corps}
            done = 0
            for fut in as_completed(futures):
                c = futures[fut]
                try:
                    c2, frame, total, fetched = fut.result()
                    if total > fetched:
                        truncated.append(f'{c2["label"]}: {fetched:,}/{total:,}건')
                    frames.append(frame)
                except Exception as e:
                    st.warning(f'{c["label"]} 조회 실패: {e}')
                done += 1
                progress.progress(done / len(selected_corps))
                status.caption(f"{done}/{len(selected_corps)} 법인 조회")

        progress.empty()
        status.empty()

        if not frames:
            st.warning("가져온 원천 정산자료가 없습니다.")
            return

        df = pd.concat(frames, ignore_index=True)
        df = _item_filter(df, keyword, ["품목", "품종"])

        if grade_scope == "특/상/보통 외만" and not df.empty:
            df = df[df["등급"].map(_is_nonstandard_grade)].copy()

        if df.empty:
            st.info(
                "선택한 장일자·품목·법인에서 조건에 맞는 원천 거래를 찾지 못했습니다. "
                "당일 자료가 아직 정산 전이면 `실시간 경락 · 공개규격 확인`을 사용해 보세요."
            )
            return

        st.success(f"원천 거래 {len(df):,}건에서 규격/등급을 분리했습니다.")

        if truncated:
            st.warning("일부 법인은 안전 상한까지만 불러왔습니다: " + " · ".join(truncated))

        summary = _summary_from_decoded(df)

        if not summary.empty:
            st.subheader("규격별 경쟁가")
            st.caption("같은 `등외`라도 포장·크기/규격이 다르면 별도 행으로 분리합니다.")
            st.dataframe(
                summary.style.format({
                    "물량": "{:,.2f}",
                    "머리": "{:,.0f}",
                    "중간": "{:,.0f}",
                    "꼬리": "{:,.0f}",
                    "평균": "{:,.0f}",
                }),
                use_container_width=True,
                hide_index=True,
            )

        st.subheader("원천 거래행")
        view_cols = [
            "시장", "법인", "품목", "품종", "규격표시", "등급",
            "물량", "경락가", "산지", "낙찰시간",
            "크기코드", "등급코드", "포장코드",
        ]
        st.dataframe(
            df[view_cols].sort_values(
                ["법인", "품목", "품종", "규격표시", "등급", "경락가"],
                ascending=[True, True, True, True, True, False],
            ).style.format({"물량": "{:,.2f}", "경락가": "{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
        )

        csv_data = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "등외 해체 원천 CSV 저장",
            data=csv_data,
            file_name=f"등외해체_{sale_date}_{keyword or '전체'}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    else:
        run = st.button("실시간 공개규격 조회", type="primary", use_container_width=True)
        if not run:
            st.caption(
                "실시간 API는 법인·품목·산지·경락가·물량과 `STD(규격)`를 제공합니다. "
                "정산 원천의 SIZECD/LVCD처럼 완전히 분리되지는 않지만 경매 직후 확인용으로 빠릅니다."
            )
            return

        sale_date = query_date.strftime("%Y%m%d")
        frames = []

        progress = st.progress(0.0)
        status = st.empty()

        def fetch_live_one(c):
            rows, total = mafra_fetch_live_company(
                api_key, sale_date, c["market_code"], c["corp_code"]
            )
            return c, _live_to_dataframe(rows), total

        with ThreadPoolExecutor(max_workers=min(4, len(selected_corps))) as ex:
            futures = {ex.submit(fetch_live_one, c): c for c in selected_corps}
            done = 0
            for fut in as_completed(futures):
                c = futures[fut]
                try:
                    _, frame, _ = fut.result()
                    frames.append(frame)
                except Exception as e:
                    st.warning(f'{c["label"]} 실시간 조회 실패: {e}')
                done += 1
                progress.progress(done / len(selected_corps))
                status.caption(f"{done}/{len(selected_corps)} 법인 조회")

        progress.empty()
        status.empty()

        if not frames:
            st.warning("실시간 경락자료가 없습니다.")
            return

        df = pd.concat(frames, ignore_index=True)
        df = _item_filter(df, keyword, ["품목", "품종", "부류"])

        if df.empty:
            st.info("선택한 조건의 실시간 거래가 아직 없습니다.")
            return

        st.success(f"실시간 거래 {len(df):,}건")
        st.dataframe(
            df.sort_values(["법인", "공개규격", "경락가"], ascending=[True, True, False])
            .style.format({"물량": "{:,.0f}", "경락가": "{:,.0f}"}),
            use_container_width=True,
            hide_index=True,
        )

        # 규격별 머리/중간/꼬리
        rows = []
        for keys, g in df.groupby(["시장", "법인", "품목", "품종", "공개규격"], dropna=False):
            prices = [int(round(x)) for x in g["경락가"].tolist() if _num(x, 0) > 0]
            if not prices:
                continue
            rows.append({
                "시장": keys[0],
                "법인": keys[1],
                "품목": keys[2],
                "품종": keys[3],
                "규격": keys[4],
                "건수": len(g),
                "머리": max(prices),
                "중간": actual_middle_price(prices),
                "꼬리": min(prices),
                "평균": round(mean(prices)),
            })

        if rows:
            st.subheader("실시간 규격별 가격")
            sdf = pd.DataFrame(rows)
            st.dataframe(
                sdf.style.format({
                    "머리": "{:,.0f}", "중간": "{:,.0f}",
                    "꼬리": "{:,.0f}", "평균": "{:,.0f}",
                }),
                use_container_width=True,
                hide_index=True,
            )

        st.info(
            "실시간 `STD`가 `10kg 상자`까지만 내려오면 그 시점 공개 실시간 자료에는 크기값이 없는 것입니다. "
            "정산 후에는 원천자료의 `SIZECD`와 공식 크기코드표로 다시 확인할 수 있습니다."
        )


# =========================================================
# URL 상태
# =========================================================
def qp_get(name: str, default: str = "") -> str:
    try:
        value = st.query_params.get(name, default)
        if isinstance(value, list):
            return value[0] if value else default
        return str(value)
    except Exception:
        return default


def qp_set(**kwargs):
    try:
        for k, v in kwargs.items():
            st.query_params[k] = str(v)
    except Exception:
        pass


def parse_default_date(raw: str) -> date:
    try:
        return datetime.strptime(raw, "%Y%m%d").date()
    except Exception:
        return date.today()


# =========================================================
# UI
# =========================================================
mode = st.radio(
    "기능",
    ["📊 경매조회", "🧩 등외 해체", "🔍 세부규격 찾기"],
    horizontal=True,
    label_visibility="collapsed",
)

if mode == "🧩 등외 해체":
    render_grade_decoder()
    st.stop()

if mode == "🔍 세부규격 찾기":
    render_trace_scanner()
    st.stop()

st.title("가락·강서 경매조회")
st.caption("✅ APP VERSION: v16-TODAY-ITEM-PICKER")
st.caption("품목 입력 시 바로 법인검색 · 품목을 비우면 오늘 진행 품목 선택 후 법인검색")

default_item = qp_get("item", "")
default_date = parse_default_date(qp_get("date", ""))

market_param = qp_get("markets", "가락,강서")
default_markets = [x for x in market_param.split(",") if x in MARKETS]
if not default_markets:
    default_markets = ["가락", "강서"]

c1, c2 = st.columns([1, 1.35])
with c1:
    selected_date = st.date_input("장일자", value=default_date)
    st.caption("예: 10월 1일 장 = 9월 30일 밤부터 10월 1일 새벽까지의 장")
with c2:
    item_text = st.text_input(
        "품목",
        value=default_item,
        placeholder="예: 표고, 당근, 쪽파",
    ).strip()

market_scope = st.multiselect(
    "시장",
    options=["가락", "강서"],
    default=default_markets,
    help="둘 다 선택하면 가락·강서에서 경매결과가 있는 법인을 함께 찾습니다.",
)

with st.expander("🧩 세부규격 자동보강 설정", expanded=False):
    secret_key = _get_secret_api_key()
    current_key = str(st.session_state.get("main_mafra_api_key", "") or secret_key)

    enrich_toggle = st.toggle(
        "경매결과 아래에 공식 세부규격 붙이기",
        value=bool(st.session_state.get("main_spec_enrich_enabled", False)),
        key="main_spec_enrich_widget",
    )
    api_key_input = st.text_input(
        "농림축산식품 공공데이터 API 키",
        value=current_key,
        type="password",
        placeholder="발급받은 키",
        help="키는 채팅에 보내지 말고 여기 입력하세요.",
        key="main_mafra_api_key_widget",
    ).strip()

    st.session_state["main_spec_enrich_enabled"] = enrich_toggle
    st.session_state["main_mafra_api_key"] = api_key_input

    if enrich_toggle and api_key_input:
        st.success("자동보강 ON · 경매결과 아래에 확인된 규격을 붙입니다.")
    elif enrich_toggle:
        st.warning("자동보강 ON · API 키를 입력해야 실제 보강이 시작됩니다.")

search_clicked = st.button("경매 법인 찾기 / 오늘 품목 보기", type="primary", use_container_width=True)

if search_clicked:
    if not market_scope:
        st.error("가락 또는 강서를 하나 이상 선택해 주세요.")
        st.stop()

    ymd = selected_date.strftime("%Y%m%d")

    if not item_text:
        # v16: 품목 미입력 -> 당일 실제 진행 품목 선택 화면
        st.session_state.pop("search", None)
        st.session_state["item_picker_request_v16"] = {
            "date": ymd,
            "markets": tuple(market_scope),
        }
        qp_set(date=ymd, item="", markets=",".join(market_scope))
    else:
        st.session_state.pop("item_picker_request_v16", None)
        st.session_state["search"] = {
            "date": ymd,
            "item": item_text,
            "markets": tuple(market_scope),
        }
        qp_set(date=ymd, item=item_text, markets=",".join(market_scope))

# 품목을 비워 검색한 경우: 오늘 진행 품목 선택 -> 선택 후 다시 법인 선택
if "item_picker_request_v16" in st.session_state:
    picker = st.session_state["item_picker_request_v16"]
    render_today_item_picker_v16(
        picker["date"],
        tuple(picker["markets"]),
    )
    st.stop()

# URL에 검색조건이 있으면 새로고침 후에도 같은 검색화면 복원
if "search" not in st.session_state and default_item:
    st.session_state["search"] = {
        "date": default_date.strftime("%Y%m%d"),
        "item": default_item,
        "markets": tuple(default_markets),
    }

if "search" not in st.session_state:
    st.info("품목을 입력해 바로 검색하거나, 품목을 비운 채 **경매 법인 찾기 / 오늘 품목 보기**를 누르세요.")
    st.stop()

search = st.session_state["search"]

st.divider()

title_col, reload_col = st.columns([4.6, 1.4], vertical_alignment="center")
with title_col:
    st.subheader("법인 선택")
with reload_col:
    reload_clicked = st.button(
        "🔄 재로딩",
        use_container_width=True,
        help="아직 경매결과가 없던 법인을 다시 확인합니다.",
    )

if reload_clicked:
    # 법인 활성/비활성 상태와 선택 법인의 최신 경매행을 다시 확인
    probe_company.clear()
    discover_company_status.clear()
    fetch_all_company_rows.clear()
    st.rerun()

st.caption(
    f'{search["date"][:4]}-{search["date"][4:6]}-{search["date"][6:8]} · '
    f'{search["item"]} · {" + ".join(search["markets"])}'
)
st.caption("✅ 결과 있음 · ⏳ 아직 없음 · 우측 **재로딩**으로 다시 확인")

try:
    with st.spinner("법인별 경매결과 확인 중..."):
        companies = discover_company_status(
            search["date"],
            search["item"],
            tuple(search["markets"]),
        )
except Exception as e:
    st.error("가락시장 경매 사이트에 연결하지 못했습니다.")
    st.code(str(e))
    st.caption(
        "Streamlit Cloud 서버에서 temp.garak.co.kr 접속이 차단되는 경우가 있습니다. "
        "그 경우 수집 서버만 별도로 국내에 두는 구조로 바꾸면 됩니다."
    )
    st.stop()

if not companies:
    st.warning("선택한 시장의 법인 목록을 가져오지 못했습니다.")
    st.stop()

# 시장별 요약
for market_name in search["markets"]:
    market_companies = [c for c in companies if c["market"] == market_name]
    active_count = sum(1 for c in market_companies if c["available"])
    if market_companies:
        st.markdown(
            f'<span class="market-badge">{market_name}</span>'
            f' 활성 {active_count} / 전체 {len(market_companies)}',
            unsafe_allow_html=True,
        )

# 법인 상태 버튼:
# - ✅ = 오늘 해당 품목 경매결과가 올라옴 + 누를 수 있음
# - ⏳ = 아직 결과 없음 + 누를 수 없음
# - 여러 법인을 동시에 선택 가능
selection_state_key = (
    f'corp_selected::{search["date"]}::{search["item"]}::'
    + ",".join(search["markets"])
)
if selection_state_key not in st.session_state:
    st.session_state[selection_state_key] = []

selected_keys = set(st.session_state[selection_state_key])
company_by_key = {c["key"]: c for c in companies}

# 재로딩 후 더 이상 목록에 없는 키 제거
selected_keys = {k for k in selected_keys if k in company_by_key and company_by_key[k]["available"]}

for market_name in search["markets"]:
    market_companies = [c for c in companies if c["market"] == market_name]
    if not market_companies:
        continue

    st.markdown(f"##### {market_name}시장")
    cols = st.columns(2)

    for idx, c in enumerate(market_companies):
        is_selected = c["key"] in selected_keys

        if c["available"]:
            # 결과가 올라온 법인은 전부 ✅ 상태로 보임.
            # 선택된 법인은 버튼 색으로 구분.
            label = f'✅ {c["name"]}'
            btn_type = "primary" if is_selected else "secondary"
            help_text = (
                "선택됨 · 다시 누르면 해제"
                if is_selected
                else "경매결과 있음 · 눌러서 선택"
            )
        else:
            label = f'⏳ {c["name"]}'
            btn_type = "secondary"
            help_text = "아직 경매결과 없음 · 재로딩 후 다시 확인"

        with cols[idx % 2]:
            clicked = st.button(
                label,
                key=f'corp_btn::{search["date"]}::{search["item"]}::{c["key"]}',
                disabled=not c["available"],
                type=btn_type,
                use_container_width=True,
                help=help_text,
            )

        if clicked and c["available"]:
            if c["key"] in selected_keys:
                selected_keys.remove(c["key"])
            else:
                selected_keys.add(c["key"])
            st.session_state[selection_state_key] = list(selected_keys)
            st.rerun()

st.session_state[selection_state_key] = list(selected_keys)

selected_labels = []
company_by_label = {}
for c in companies:
    label = f'{c["market"]} · {c["name"]}'
    company_by_label[label] = c
    if c["key"] in selected_keys and c["available"]:
        selected_labels.append(label)

if selected_labels:
    st.caption("선택: " + " · ".join(selected_labels))

if not any(c["available"] for c in companies):
    st.info("아직 선택 가능한 법인이 없습니다. 경매결과가 올라온 뒤 우측 **재로딩**을 눌러 주세요.")
    st.stop()

if not selected_labels:
    st.info("✅ 법인을 하나 이상 누르면 해당 법인의 품종·등급별 가격을 보여줍니다.")
    st.stop()

label_to_company = company_by_label

# v15: 선택 문구 바로 아래에 비교표 + 약세 진단
if len(selected_labels) >= 2:
    st.markdown("### 🔎 v15 비교/진단 준비")
    compare_frames: List[pd.DataFrame] = []
    with st.spinner("선택한 법인들의 오늘 경매자료를 비교 중..."):
        for label in selected_labels:
            c = label_to_company[label]
            df_cmp = fetch_all_company_rows(
                search["date"], search["item"], c["market"], c["code"]
            )
            if not df_cmp.empty:
                df_cmp = df_cmp.copy()
                df_cmp.insert(1, "법인", c["name"])
                compare_frames.append(df_cmp)

    render_compare_v14(compare_frames, selected_labels, search["item"])
    render_weakness_diagnosis_v15(
        label_to_company,
        compare_frames,
        selected_labels,
        search["date"],
        search["item"],
    )

st.divider()

selected_frames: List[pd.DataFrame] = []

if len(selected_labels) == 1:
    company = label_to_company[selected_labels[0]]
    frame = render_company(company, search["date"], search["item"])
    if not frame.empty:
        selected_frames.append(frame)
else:
    tabs = st.tabs(selected_labels)
    for tab, label in zip(tabs, selected_labels):
        with tab:
            company = label_to_company[label]
            frame = render_company(company, search["date"], search["item"])
            if not frame.empty:
                selected_frames.append(frame)

if selected_frames:
    st.divider()
    try:
        excel_bytes = build_excel_bytes(selected_frames)
        filename = f'경매비교_{search["date"]}_{search["item"]}.xlsx'
        st.download_button(
            "선택한 법인 엑셀 다운로드",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.caption(f"엑셀 생성 중 오류: {e}")
