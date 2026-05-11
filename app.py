import time
import zipfile
import io
import xml.etree.ElementTree as ET
import datetime

import streamlit as st
import requests
import pandas as pd
from io import BytesIO

API_KEY = st.secrets.get("DART_API_KEY", "f0b490beadb3e0200407e2f237b58bca1ae74ac4")
BASE_URL = "https://opendart.fss.or.kr/api"

REPORT_TYPES = {
    "연간 (사업보고서)": "11011",
    "반기 (반기보고서)": "11012",
    "1분기 (분기보고서)": "11013",
    "3분기 (분기보고서)": "11014",
}

FS_LABELS = {
    "BS":  "재무상태표",
    "IS":  "손익계산서",
    "CIS": "포괄손익계산서",
    "CF":  "현금흐름표",
    "SCE": "자본변동표",
}

THROTTLE_SEC = 0.2

# IFRS XBRL 표준 명명 규칙에서 현금유출을 양수로 저장하는 패턴
# 특정 회사·계정명이 아닌 IFRS full taxonomy 원칙 기반
_CF_OUTFLOW_PATTERNS = [
    "PurchaseOf",         # 자산 취득 (유형자산, 무형자산, 금융자산 등)
    "Repayments",         # 차입금·사채 상환
    "DividendsPaid",      # 배당금 지급
    "InterestPaid",       # 이자 지급
    "IncomeTaxesPaid",    # 법인세 납부
    "PaymentsFor",        # 리스부채 상환 등 (IFRS 16)
    "PaymentsTo",         # 직접법 영업CF (공급자/종업원 지급액)
    "AcquisitionOf",      # 자기주식 취득 등 dart_ 커스텀 취득 요소
]

# 위 패턴에 해당하지만 실제로는 이미 서명된(signed) 행을 보호하는 제외 패턴
_CF_OUTFLOW_EXCLUDE = [
    "CashFlowsFromUsedIn",   # 소계행 — DART가 이미 부호 처리
    "ProceedsFrom",           # 처분·상환 수취액(유입)
]


def _is_cf_outflow(account_id: str) -> bool:
    """
    IFRS XBRL 표준에서 현금유출을 양수로 저장하는 요소인지 판정.
    - ifrs-full_ 및 dart_ 커스텀 요소 모두 포함
    - 이미 서명된 소계행·유입 요소는 제외
    """
    if not account_id or "표준계정코드 미사용" in account_id:
        return False
    if any(exc in account_id for exc in _CF_OUTFLOW_EXCLUDE):
        return False
    if any(p in account_id for p in _CF_OUTFLOW_PATTERNS):
        return True
    # CashFlowsUsedIn 단독 (사업결합 순현금유출 등) — FromUsedIn 소계는 위에서 제외됨
    if "CashFlowsUsedIn" in account_id:
        return True
    return False


# 계정명 중복 레이블 매핑: account_id 키워드 → 표시 prefix
# 삼성 CIS 사례 외에도 IFRS taxonomy 공통 패턴을 등록
_DUPLICATE_LABEL_RULES: list[tuple[str, str]] = [
    ("WillNotBeReclassifiedToProfitOrLoss", "[재분류X]"),
    ("WillBeReclassifiedToProfitOrLoss",    "[재분류O]"),
    # 향후 추가 가능: ("SomeOtherPattern", "[Label]")
]


def _make_display_nm(account_nm: str, account_id: str, is_dup: bool) -> str:
    """
    중복 계정명(is_dup=True)이면 account_id 기반 prefix를 붙여 구분.
    - 알려진 패턴은 의미 있는 한글 label 사용
    - 미등록 패턴은 account_id 말미를 fallback으로 사용
    """
    if not is_dup:
        return account_nm
    for keyword, label in _DUPLICATE_LABEL_RULES:
        if keyword in account_id:
            return f"{label} {account_nm}"
    # fallback: account_id의 마지막 토큰(언더스코어 분리)
    suffix = account_id.split("_")[-1][:20] if account_id else "?"
    return f"{account_nm} [{suffix}]"


# ── 데이터 로직 ───────────────────────────────────────────────────────────────

@st.cache_data(ttl=86400, show_spinner=False)
def load_corp_codes() -> pd.DataFrame:
    r = requests.get(f"{BASE_URL}/corpCode.xml", params={"crtfc_key": API_KEY}, timeout=30)
    r.raise_for_status()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    root = ET.fromstring(z.read("CORPCODE.xml"))
    rows = [
        {
            "corp_code":  item.findtext("corp_code", "").strip(),
            "corp_name":  item.findtext("corp_name", "").strip(),
            "stock_code": item.findtext("stock_code", "").strip(),
            "modify_date":item.findtext("modify_date", "").strip(),
        }
        for item in root.findall("list")
    ]
    return pd.DataFrame(rows)


def search_company(query: str, corp_df: pd.DataFrame) -> pd.DataFrame:
    query = query.strip()
    if not query:
        return pd.DataFrame()
    if query.isdigit() and len(query) == 6:
        result = corp_df[corp_df["stock_code"] == query]
    else:
        result = corp_df[corp_df["corp_name"].str.contains(query, na=False)]
    listed   = result[result["stock_code"] != ""].copy()
    unlisted = result[result["stock_code"] == ""].copy()
    return pd.concat([listed, unlisted], ignore_index=True)


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_financials(corp_code: str, year: int, reprt_code: str) -> pd.DataFrame:
    for fs_div in ("CFS", "OFS"):
        time.sleep(THROTTLE_SEC)
        r = requests.get(
            f"{BASE_URL}/fnlttSinglAcntAll.json",
            params={"crtfc_key": API_KEY, "corp_code": corp_code,
                    "bsns_year": str(year), "reprt_code": reprt_code, "fs_div": fs_div},
            timeout=15,
        )
        data = r.json()
        if data.get("status") == "000" and data.get("list"):
            df = pd.DataFrame(data["list"])
            df["year"] = year
            df["fs_div_used"] = fs_div
            return df
    return pd.DataFrame()


def parse_amount(val):
    try:
        cleaned = str(val).replace(",", "").strip()
        return int(cleaned) if cleaned not in ("", "-", "None") else None
    except (ValueError, TypeError):
        return None


def build_excel(corp_code: str, corp_name: str, stock_code: str,
                years: list[int], reprt_code: str) -> bytes:
    all_frames = []
    progress = st.progress(0, text="데이터 수집 중...")

    for i, year in enumerate(years):
        progress.progress(i / len(years), text=f"{year}년 재무제표 수집 중...")
        df = fetch_financials(corp_code, year, reprt_code)
        if not df.empty:
            all_frames.append(df)

    progress.progress(1.0, text="수집 완료")

    if not all_frames:
        st.error("재무 데이터를 가져오지 못했습니다. 연도나 보고서 종류를 확인해주세요.")
        return b""

    raw = pd.concat(all_frames, ignore_index=True)

    fs_div_used = raw["fs_div_used"].iloc[0] if "fs_div_used" in raw.columns else "?"
    fs_label = "연결재무제표" if fs_div_used == "CFS" else "별도재무제표"

    for col in ["thstrm_amount", "frmtrm_amount", "bfefrmtrm_amount"]:
        if col in raw.columns:
            raw[col] = raw[col].apply(parse_amount)

    # Bug #2 수정: CF 세부 라인 부호 복원
    if "account_id" in raw.columns:
        cf_mask = raw["sj_div"] == "CF"
        outflow_mask = raw["account_id"].apply(
            lambda aid: _is_cf_outflow(str(aid) if aid else "")
        )
        for col in ["thstrm_amount", "frmtrm_amount", "bfefrmtrm_amount"]:
            if col in raw.columns:
                negate = cf_mask & outflow_mask & raw[col].notna() & (raw[col] > 0)
                raw.loc[negate, col] = -raw.loc[negate, col]

    # Bug #1 수정: 재무제표·연도 내 중복 account_nm을 자동 감지 후 display_nm 생성
    if "account_id" in raw.columns:
        # 중복 여부: 동일 sj_div + year 안에서 account_nm이 2회 이상 등장하는 경우
        dup_keys = set(
            raw.groupby(["sj_div", "year", "account_nm"])
               .size()
               .loc[lambda s: s > 1]
               .reset_index()[["sj_div", "account_nm"]]
               .apply(tuple, axis=1)
        )
        raw["display_nm"] = raw.apply(
            lambda r: _make_display_nm(
                r.get("account_nm", ""),
                str(r.get("account_id", "")),
                is_dup=(r.get("sj_div"), r.get("account_nm")) in dup_keys,
            ),
            axis=1,
        )
    else:
        raw["display_nm"] = raw["account_nm"]

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        meta = pd.DataFrame([
            {"항목": "기업명",           "내용": corp_name},
            {"항목": "종목코드",         "내용": stock_code or "-"},
            {"항목": "DART corp_code",   "내용": corp_code},
            {"항목": "재무제표 종류",    "내용": fs_label},
            {"항목": "수집 연도",        "내용": f"{years[0]}~{years[-1]}"},
            {"항목": "수집일",           "내용": pd.Timestamp.now().strftime("%Y-%m-%d")},
        ])
        meta.to_excel(writer, sheet_name="정보", index=False)

        cols_keep = [c for c in ["year", "sj_div", "sj_nm", "account_id",
                                  "account_nm", "thstrm_amount", "fs_div_used"]
                     if c in raw.columns]
        raw[cols_keep].rename(columns={
            "year": "연도", "sj_div": "재무제표구분코드", "sj_nm": "재무제표명",
            "account_id": "계정ID", "account_nm": "계정명",
            "thstrm_amount": "금액(원)", "fs_div_used": "연결/별도",
        }).to_excel(writer, sheet_name="원본데이터", index=False)

        for fs_code, fs_name in FS_LABELS.items():
            subset = raw[raw["sj_div"] == fs_code]
            if subset.empty:
                continue
            pivot = (
                subset.groupby(["display_nm", "year"])["thstrm_amount"]
                .first().unstack("year").reset_index()
            )
            pivot.columns.name = None
            pivot = pivot.rename(columns={"display_nm": "계정명"})
            year_cols = sorted([c for c in pivot.columns if isinstance(c, int)])
            pivot = pivot[["계정명"] + year_cols]
            pivot.to_excel(writer, sheet_name=fs_name[:31], index=False)
            ws = writer.sheets[fs_name[:31]]
            ws.column_dimensions["A"].width = 45
            for col_idx in range(2, len(year_cols) + 2):
                ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = 18

    return output.getvalue()


# ── CSS ──────────────────────────────────────────────────────────────────────

def inject_css():
    st.markdown("""
    <style>
    /* 상단 Streamlit 툴바/햄버거 메뉴 숨김 */
    #MainMenu, header, footer { visibility: hidden; }

    /* 전체 여백 */
    .block-container { padding: 2.5rem 2rem 4rem; max-width: 780px; }

    /* 헤더 영역 */
    .app-header {
        padding: 2rem 0 1.5rem;
        border-bottom: 1px solid #E2E8F0;
        margin-bottom: 2rem;
    }
    .app-header h1 {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0 0 0.25rem;
        letter-spacing: -0.02em;
    }
    .app-header p {
        font-size: 0.875rem;
        color: #64748B;
        margin: 0;
    }

    /* 섹션 카드 */
    .card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.25rem;
    }
    .card-title {
        font-size: 0.7rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #94A3B8;
        margin-bottom: 1rem;
    }

    /* 기업 정보 배지 */
    .corp-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        background: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        margin-bottom: 1rem;
        width: 100%;
    }
    .corp-name { font-size: 1rem; font-weight: 700; color: #1E40AF; }
    .corp-meta { font-size: 0.8rem; color: #3B82F6; }
    .corp-tag  {
        font-size: 0.7rem; font-weight: 600;
        background: #DBEAFE; color: #1D4ED8;
        padding: 0.15rem 0.5rem; border-radius: 4px;
    }

    /* 다운로드 버튼 재스타일 */
    [data-testid="stDownloadButton"] button {
        background: #2563EB !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 0.6rem 1.2rem !important;
        font-size: 0.9rem !important;
        transition: background 0.15s !important;
    }
    [data-testid="stDownloadButton"] button:hover {
        background: #1D4ED8 !important;
    }

    /* 제출 버튼 */
    [data-testid="stFormSubmitButton"] button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }

    /* input / selectbox 테두리 */
    [data-baseweb="input"] > div,
    [data-baseweb="select"] > div:first-child {
        border-radius: 8px !important;
    }

    /* 구분선 */
    hr { border: none; border-top: 1px solid #E2E8F0; margin: 1.5rem 0; }
    </style>
    """, unsafe_allow_html=True)


# ── UI ───────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="DART 재무제표", page_icon="📋", layout="centered")
inject_css()

st.markdown("""
<div class="app-header">
  <h1>📋 DART 재무제표 수집기</h1>
  <p>국내 상장사의 재무제표를 DART에서 자동으로 가져와 연도별 엑셀로 정리합니다.</p>
</div>
""", unsafe_allow_html=True)

# 기업코드 목록 로드
with st.spinner("DART 기업 목록 초기화 중..."):
    try:
        corp_df = load_corp_codes()
    except Exception as e:
        st.error(f"DART 기업 목록 로드 실패: {e}")
        st.stop()

# 검색 폼
st.markdown('<div class="card"><div class="card-title">기업 검색</div>', unsafe_allow_html=True)
with st.form("main_form"):
    query = st.text_input(
        "회사명 또는 종목코드",
        placeholder="예: 삼성전자 / 005930",
        label_visibility="collapsed",
    )

    col1, col2, col3 = st.columns([2, 2, 3])
    today = datetime.date.today()
    default_year = today.year - 1

    with col1:
        end_year = st.number_input("기준 연도", min_value=2010,
                                   max_value=today.year, value=default_year, step=1)
    with col2:
        num_years = st.number_input("수집 연수", min_value=1, max_value=10, value=5, step=1)
    with col3:
        report_type = st.selectbox("보고서 종류", list(REPORT_TYPES.keys()))

    submitted = st.form_submit_button("검색 및 수집", use_container_width=True, type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# 결과
if submitted and query.strip():
    results = search_company(query.strip(), corp_df)

    if results.empty:
        st.error("검색 결과가 없습니다. 회사명을 다시 확인해주세요.")
        st.stop()

    if len(results) > 1:
        options = [
            f"{row['corp_name']}  ·  {row['stock_code'] or '비상장'}"
            for _, row in results.head(10).iterrows()
        ]
        choice_idx = st.selectbox("검색된 기업 선택", range(len(options)),
                                   format_func=lambda i: options[i])
        selected = results.iloc[choice_idx]
    else:
        selected = results.iloc[0]

    corp_code  = selected["corp_code"]
    corp_name  = selected["corp_name"]
    stock_code = selected["stock_code"]
    years      = list(range(int(end_year) - int(num_years) + 1, int(end_year) + 1))
    reprt_code = REPORT_TYPES[report_type]
    year_range = f"{years[0]} – {years[-1]}"

    st.markdown(f"""
    <div class="corp-badge">
      <span class="corp-name">{corp_name}</span>
      <span class="corp-tag">{stock_code or "비상장"}</span>
      <span class="corp-meta">· {year_range} · {report_type.split()[0]}</span>
    </div>
    """, unsafe_allow_html=True)

    excel_bytes = build_excel(corp_code, corp_name, stock_code, years, reprt_code)

    if excel_bytes:
        fs_div_used = "CFS"  # build_excel 내부에서 결정되므로 표시는 생략
        c1, c2, c3 = st.columns(3)
        c1.metric("수집 연도", year_range)
        c2.metric("연수", f"{num_years}개년")
        c3.metric("보고서", report_type.split()[0])

        st.markdown("<hr>", unsafe_allow_html=True)
        st.download_button(
            label="엑셀 다운로드  (.xlsx)",
            data=excel_bytes,
            file_name=f"{corp_name}_재무제표_{years[0]}-{years[-1]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

elif submitted:
    st.warning("회사명 또는 종목코드를 입력해주세요.")
