import io
import math
import re
import ipaddress
import socket
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from statistics import mean
from typing import Dict, List, Tuple
from urllib.parse import urljoin, urlparse

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
# 공개 세부규격 흔적 검사기 (모바일 1차 조사)
# =========================================================
TRACE_KEYWORDS = [
    "42망", "45망", "40망", "38망", "36망",
    "4수", "4개들이", "4포기", "망치수", "치수", "세부규격", "규격",
    "size", "spec", "standard", "grade", "class", "level", "unit",
    "sizecd", "size_cd", "speccd", "spec_cd",
    "gradecd", "grade_cd", "classcd", "class_cd",
    "standardcd", "standard_cd",
]

TRACE_SIZE_RE = re.compile(
    r'(?<!\d)(?:36|38|40|42|43|45|47|48|50|52|55)'
    r'(?:\s*[-~]\s*(?:36|38|40|42|43|45|47|48|50|52|55))?\s*망'
)
TRACE_FOUR_RE = re.compile(r'(?:4\s*수|4\s*개(?:들이)?|4\s*포기)')
TRACE_FIELD_RE = re.compile(
    r'(?i)\b(?:size|spec|standard|grade|class|level|unit)'
    r'(?:[_-]?(?:cd|code|nm|name|no|id|seq))?\b'
)
TRACE_FIELD_NUMBER_RE = re.compile(
    r'(?i)(?:size|spec|standard|grade|class|level)'
    r'[^<>{}\n]{0,70}(?:36|38|40|42|43|45|47|48|50|52|55)'
)


def _is_public_http_url(url: str) -> Tuple[bool, str]:
    """서버측 URL 요청이 사설/로컬 주소로 가지 않게 막는다."""
    try:
        p = urlparse(url.strip())
        if p.scheme not in ("http", "https"):
            return False, "http 또는 https 주소만 사용할 수 있습니다."
        if not p.hostname:
            return False, "도메인을 확인할 수 없습니다."

        host = p.hostname.lower()
        if host in {"localhost", "localhost.localdomain"} or host.endswith(".local"):
            return False, "로컬 주소는 사용할 수 없습니다."

        # DNS 결과 중 하나라도 사설/로컬 계열이면 차단.
        infos = socket.getaddrinfo(host, 443 if p.scheme == "https" else 80, type=socket.SOCK_STREAM)
        if not infos:
            return False, "도메인의 IP를 확인할 수 없습니다."

        for info in infos:
            ip_text = info[4][0]
            try:
                ip = ipaddress.ip_address(ip_text)
            except ValueError:
                continue
            if (
                ip.is_private
                or ip.is_loopback
                or ip.is_link_local
                or ip.is_reserved
                or ip.is_multicast
                or ip.is_unspecified
            ):
                return False, "공개 인터넷 주소가 아닌 대상은 검사하지 않습니다."
        return True, ""
    except socket.gaierror:
        return False, "도메인 DNS를 확인하지 못했습니다."
    except Exception as e:
        return False, f"주소 확인 오류: {e}"


def _safe_public_get(session: requests.Session, url: str, timeout: int = 18):
    """리다이렉트도 한 단계씩 검증하면서 공개 URL만 GET."""
    current = url.strip()
    for _ in range(6):
        ok, reason = _is_public_http_url(current)
        if not ok:
            raise ValueError(reason)

        r = session.get(current, timeout=timeout, allow_redirects=False)
        if r.status_code in (301, 302, 303, 307, 308):
            location = r.headers.get("Location")
            if not location:
                r.raise_for_status()
                return r
            current = urljoin(current, location)
            continue

        r.raise_for_status()
        return r

    raise RuntimeError("리다이렉트가 너무 많습니다.")


def _decode_public_response(r: requests.Response) -> str:
    enc = (r.encoding or "").lower()
    if enc and enc not in {"iso-8859-1", "ascii"}:
        return r.text
    try:
        return r.content.decode(r.apparent_encoding or "utf-8", errors="replace")
    except Exception:
        return r.content.decode("utf-8", errors="replace")


def _trace_add(hits: List[Dict[str, str]], source: str, kind: str, match: str, text: str, start: int, end: int):
    left = max(0, start - 110)
    right = min(len(text), end + 110)
    context = re.sub(r"\s+", " ", text[left:right]).strip()
    hits.append({
        "종류": kind,
        "일치": match,
        "출처": source,
        "주변내용": context,
    })


def _trace_text(text: str, source: str, item: str = "") -> List[Dict[str, str]]:
    hits: List[Dict[str, str]] = []

    for kind, pattern in (
        ("망규격", TRACE_SIZE_RE),
        ("4수/4개", TRACE_FOUR_RE),
        ("필드명", TRACE_FIELD_RE),
        ("필드+숫자", TRACE_FIELD_NUMBER_RE),
    ):
        for m in pattern.finditer(text):
            _trace_add(hits, source, kind, m.group(0), text, m.start(), m.end())

    low = text.lower()
    for kw in TRACE_KEYWORDS:
        k = kw.lower()
        start = 0
        while True:
            pos = low.find(k, start)
            if pos < 0:
                break
            _trace_add(hits, source, "키워드", kw, text, pos, pos + len(kw))
            start = pos + len(k)

    # 품목명 근처에서 [1]~[9] 같은 내부 코드 후보 탐색.
    if item:
        item_positions = [m.start() for m in re.finditer(re.escape(item), text, flags=re.I)]
        bracket_re = re.compile(r"\[[1-9]\]")
        for ipos in item_positions[:100]:
            window_start = max(0, ipos - 350)
            window_end = min(len(text), ipos + 700)
            window = text[window_start:window_end]
            for m in bracket_re.finditer(window):
                real_s = window_start + m.start()
                real_e = window_start + m.end()
                _trace_add(hits, source, "숫자코드 후보", m.group(0), text, real_s, real_e)

    # 중복 제거
    uniq: List[Dict[str, str]] = []
    seen = set()
    for row in hits:
        key = (row["종류"], row["일치"], row["출처"], row["주변내용"])
        if key not in seen:
            seen.add(key)
            uniq.append(row)
    return uniq


@st.cache_data(ttl=300, show_spinner=False)
def scan_public_trace(url: str, item: str) -> Dict[str, object]:
    s = make_session()
    r = _safe_public_get(s, url)
    html_text = _decode_public_response(r)
    final_url = r.url or url

    hits = _trace_text(html_text, final_url, item=item)
    soup = BeautifulSoup(html_text, "html.parser")

    structured_rows: List[Dict[str, str]] = []
    for tag in soup.find_all(["input", "select", "option", "form", "meta"]):
        attrs = {}
        for k, v in tag.attrs.items():
            attrs[k] = " ".join(v) if isinstance(v, list) else str(v)
        attr_text = " ".join(f"{k}={v}" for k, v in attrs.items())
        visible = tag.get_text(" ", strip=True)
        probe = f"{attr_text} {visible}"

        interesting = (
            TRACE_FIELD_RE.search(probe)
            or TRACE_SIZE_RE.search(probe)
            or TRACE_FOUR_RE.search(probe)
            or any(k.lower() in probe.lower() for k in TRACE_KEYWORDS)
        )
        if interesting:
            structured_rows.append({
                "태그": tag.name,
                "name": attrs.get("name", ""),
                "id": attrs.get("id", ""),
                "value": attrs.get("value", ""),
                "text": visible[:250],
                "속성": attr_text[:600],
            })

    # 동일 도메인에서 페이지가 직접 불러오는 공개 JS만 추가 검사.
    script_urls = []
    base_host = (urlparse(final_url).hostname or "").lower()
    for script in soup.find_all("script", src=True):
        js_url = urljoin(final_url, script.get("src", ""))
        if (urlparse(js_url).hostname or "").lower() == base_host:
            script_urls.append(js_url)

    js_scanned = 0
    for js_url in list(dict.fromkeys(script_urls))[:12]:
        try:
            jr = _safe_public_get(s, js_url, timeout=12)
            js_text = _decode_public_response(jr)
            hits.extend(_trace_text(js_text, jr.url or js_url, item=item))
            js_scanned += 1
        except Exception:
            continue

    # JS 추가 후 다시 중복 제거
    uniq = []
    seen = set()
    for row in hits:
        key = (row["종류"], row["일치"], row["출처"], row["주변내용"])
        if key not in seen:
            seen.add(key)
            uniq.append(row)

    return {
        "final_url": final_url,
        "status_code": r.status_code,
        "html_bytes": len(r.content),
        "js_scanned": js_scanned,
        "hits": uniq,
        "structured": structured_rows,
        "raw_html": html_text,
    }


def render_trace_scanner():
    st.title("🔍 세부규격 찾기")
    st.caption("공개 페이지와 그 페이지가 불러오는 공개 JS에서 42·45·4수·size/spec/grade 같은 흔적을 찾습니다.")

    presets = {
        "서부청과 모바일 채소": "https://www.sbbot.com/mobile_web/m_itemList.do?gubn=vege",
        "대아청과 시황 예시": "https://dagreen.co.kr/market_price_new/daily_market_view.asp?board_seq=23969&keyword=&opt=&page=2138",
        "직접 주소 입력": "",
    }

    choice = st.selectbox("사이트", list(presets.keys()))
    default_url = presets[choice]
    item = st.text_input("찾을 품목", value="양배추", placeholder="예: 양배추, 고구마, 표고").strip()
    url = st.text_input(
        "공개 페이지 주소",
        value=default_url,
        placeholder="https://...",
        disabled=(choice != "직접 주소 입력"),
    ).strip()

    st.caption("화면에 안 보이는 값도 HTML·폼 필드·공개 JS에 남아 있으면 잡힐 수 있습니다. 로그인/권한 우회는 하지 않습니다.")

    if st.button("흔적 찾기", type="primary", use_container_width=True):
        if not url:
            st.error("페이지 주소를 넣어 주세요.")
            return

        with st.spinner("공개 소스에서 규격 흔적 찾는 중..."):
            try:
                result = scan_public_trace(url, item)
            except Exception as e:
                st.error("이 공개 페이지를 가져오지 못했습니다.")
                st.code(str(e))
                st.caption("사이트가 서버 접속을 막거나 Streamlit Cloud에서 해당 도메인 접속이 안 되는 경우 PC 정밀검사를 사용하면 됩니다.")
                return

        hits = pd.DataFrame(result["hits"])
        structured = pd.DataFrame(result["structured"])

        c1, c2, c3 = st.columns(3)
        c1.metric("흔적", f"{len(hits):,}건")
        c2.metric("HTML", f'{int(result["html_bytes"]):,} B')
        c3.metric("검사 JS", f'{int(result["js_scanned"]):,}개')

        if hits.empty:
            st.warning("이 공개 응답에서는 42·45·4수 또는 규격 필드 흔적을 찾지 못했습니다.")
        else:
            direct = hits[hits["종류"].isin(["망규격", "4수/4개", "필드+숫자"])]
            candidates = hits[hits["종류"].isin(["숫자코드 후보", "필드명"])]
            others = hits[~hits.index.isin(direct.index) & ~hits.index.isin(candidates.index)]

            if not direct.empty:
                st.subheader("🎯 직접 단서")
                st.dataframe(direct, use_container_width=True, hide_index=True)
            if not candidates.empty:
                st.subheader("🧩 코드/필드 후보")
                st.dataframe(candidates, use_container_width=True, hide_index=True)
            with st.expander(f"전체 흔적 {len(hits):,}건"):
                st.dataframe(hits, use_container_width=True, hide_index=True)

            report_lines = [
                "공개 세부규격 흔적 검사",
                f'URL: {result["final_url"]}',
                f'품목: {item}',
                "",
            ]
            for row in result["hits"]:
                report_lines.append(f'[{row["종류"]}] {row["일치"]}')
                report_lines.append(row["출처"])
                report_lines.append(row["주변내용"])
                report_lines.append("")
            report_text = "\n".join(report_lines)

            st.download_button(
                "검사결과 TXT 저장",
                data=report_text.encode("utf-8-sig"),
                file_name=f"세부규격흔적_{item or '품목'}.txt",
                mime="text/plain",
                use_container_width=True,
            )

        if not structured.empty:
            with st.expander(f"폼/숨은 필드 후보 {len(structured):,}건"):
                st.dataframe(structured, use_container_width=True, hide_index=True)

        st.info("`[1]~[6]` 같은 숫자코드는 42/45망으로 자동 해석하지 않습니다. 같은 응답에서 실제 규격명과 연결되는 증거가 있어야 매핑합니다.")


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
    ["📊 경매조회", "🔍 세부규격 찾기"],
    horizontal=True,
    label_visibility="collapsed",
)

if mode == "🔍 세부규격 찾기":
    render_trace_scanner()
    st.stop()

st.title("가락·강서 경매조회")
st.caption("장일자 + 품목 검색 → 법인 상태 확인 → 원하는 법인 선택 → 머리·중간·꼬리 확인")

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

search_clicked = st.button("경매 법인 찾기", type="primary", use_container_width=True)

if search_clicked:
    if not item_text:
        st.error("품목을 입력해 주세요.")
        st.stop()
    if not market_scope:
        st.error("가락 또는 강서를 하나 이상 선택해 주세요.")
        st.stop()

    ymd = selected_date.strftime("%Y%m%d")
    st.session_state["search"] = {
        "date": ymd,
        "item": item_text,
        "markets": tuple(market_scope),
    }
    qp_set(date=ymd, item=item_text, markets=",".join(market_scope))

# URL에 검색조건이 있으면 새로고침 후에도 같은 검색화면 복원
if "search" not in st.session_state and default_item:
    st.session_state["search"] = {
        "date": default_date.strftime("%Y%m%d"),
        "item": default_item,
        "markets": tuple(default_markets),
    }

if "search" not in st.session_state:
    st.info("장일자와 품목을 입력한 뒤 **경매 법인 찾기**를 누르세요.")
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
