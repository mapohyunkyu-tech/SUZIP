import io
import math
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from statistics import mean
from typing import Dict, List, Tuple

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

    st.caption("품종 → 단위 → 등급(특 → 상 → 보통 → 기타) 순으로 표시합니다.")
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
st.title("가락·강서 경매조회")
st.caption("품목 검색 → 전체 법인 표시 → 경매결과가 올라온 법인만 선택 가능 → 재로딩으로 갱신")

default_item = qp_get("item", "")
default_date = parse_default_date(qp_get("date", ""))

market_param = qp_get("markets", "가락,강서")
default_markets = [x for x in market_param.split(",") if x in MARKETS]
if not default_markets:
    default_markets = ["가락", "강서"]

c1, c2 = st.columns([1, 1.35])
with c1:
    selected_date = st.date_input("조회일", value=default_date)
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
    st.info("날짜와 품목을 입력한 뒤 **경매 법인 찾기**를 누르세요.")
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
st.caption("✅ = 오늘 경매결과 올라옴 · ⏳ = 아직 결과 없음. 우측 **재로딩**으로 상태를 다시 확인합니다.")

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
