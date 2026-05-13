import os
import time
import zipfile
import io
import json
import xml.etree.ElementTree as ET
import datetime

import streamlit as st
import requests
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

API_KEY = (
    os.environ.get("DART_API_KEY")
    or st.secrets.get("DART_API_KEY", "f0b490beadb3e0200407e2f237b58bca1ae74ac4")
)
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
    # SCE(자본변동표)는 2D 매트릭스 구조로 온전한 재현 불가 → L2 참조
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
# 순서 중요: 먼저 매칭된 규칙이 적용됨
_DUPLICATE_LABEL_RULES: list[tuple[str, str]] = [
    # CIS 재분류 구분
    ("WillNotBeReclassifiedToProfitOrLoss", "[재분류X]"),
    ("WillBeReclassifiedToProfitOrLoss",    "[재분류O]"),
    # 유동/비유동 구분 (BS에서 가장 흔한 중복 원인)
    ("Shortterm",    "[단기]"),
    ("Current",      "[유동]"),
    ("Longterm",     "[장기]"),
    ("Noncurrent",   "[비유동]"),
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
    # 비표준 커스텀 항목은 (기타) 표기
    if not account_id or "표준계정코드 미사용" in account_id:
        return f"{account_nm} (기타)"
    # 그 외 미등록 패턴은 account_id 말미로 구분
    suffix = account_id.split("_")[-1][:20]
    return f"{account_nm} [{suffix}]"


# ── Excel 스타일 상수 (디자인 사양 고정) ─────────────────────────────────────

# 배경색
_F_BLACK    = PatternFill("solid", fgColor="000000")  # 제목 배경
_F_NAVY     = PatternFill("solid", fgColor="1F4E78")  # 1차 헤더
_F_LBLUE    = PatternFill("solid", fgColor="D9E1F2")  # 2차 헤더 (결산일)
_F_YELLOW   = PatternFill("solid", fgColor="FFF2CC")  # 대분류 헤더
_F_GREEN    = PatternFill("solid", fgColor="E2EFDA")  # 합계/총계 (연한 초록)
_F_DK_GREEN = PatternFill("solid", fgColor="C6E0B4")  # 최종 결과 강조 (진한 초록)
_F_WHITE    = PatternFill("solid", fgColor="FFFFFF")  # 기본 배경

# 테두리
_S_THIN   = Side(style="thin",   color="000000")
_S_THICK  = Side(style="thick",  color="000000")
_S_NONE   = Side(style=None)
_S_DOUBLE = Side(style="double", color="000000")

# 글꼴
_FT_TITLE    = Font(name="맑은 고딕", bold=True,   color="FFFFFF", size=14)
_FT_H1       = Font(name="맑은 고딕", bold=True,   color="FFFFFF", size=11)
_FT_H2       = Font(name="맑은 고딕", bold=True,   color="000000", size=11)
_FT_BOLD     = Font(name="맑은 고딕", bold=True,   color="000000", size=11)
_FT_NORM     = Font(name="맑은 고딕",               color="000000", size=11)
_FT_ITALIC   = Font(name="맑은 고딕", italic=True,  color="000000", size=11)
_FT_UNIT     = Font(name="맑은 고딕", italic=True,  color="000000", size=9)
_FT_GREY     = Font(name="맑은 고딕",               color="BFBFBF", size=11)  # 빈값
_FT_GREY_NRM = Font(name="맑은 고딕",               color="7F7F7F", size=11)  # 보조 텍스트
_FT_GREY_BLD = Font(name="맑은 고딕", bold=True,    color="7F7F7F", size=11)  # 카테고리 라벨
_FT_GREY_ITA = Font(name="맑은 고딕", italic=True,  color="7F7F7F", size=11)  # % 보조지표
_FT_BLUE     = Font(name="맑은 고딕",               color="0000FF", size=11)  # 수동입력 강조

# 정렬
_AL_C  = Alignment(horizontal="center", vertical="center")
_AL_L  = Alignment(horizontal="left",   vertical="center")
_AL_L1 = Alignment(horizontal="left",   vertical="center", indent=1)
_AL_L2 = Alignment(horizontal="left",   vertical="center", indent=2)
_AL_L3 = Alignment(horizontal="left",   vertical="center", indent=3)
_AL_R  = Alignment(horizontal="right",  vertical="center")

# 숫자 서식: 양수=콤마 / 음수=괄호 / 0=하이픈
_NUM_FMT = '#,##0;(#,##0);-'

# ── 하이라이트 레벨 (IFRS element name 기반) ─────────────────────────────────
# account_id의 마지막 _ 이후 element name으로 판정 (한글 계정명 의존 X)

_HL3_ELS = {    # Level 3: 합계/총계 → 초록 배경 + 굵게
    # BS
    "Assets", "EquityAndLiabilities", "LiabilitiesAndEquity",
    # IS / CIS (당기순이익, 포괄이익은 해당 재무제표에서만 최상위 합계)
    "ProfitLoss",
    "ComprehensiveIncome",
    # CF
    "CashAndCashEquivalentsAtEndOfPeriodCf",
}
_HL2_ELS = {    # Level 2: 중분류/소계 → 흰색 + 굵게, indent 1
    # BS는 _hl_level 내에서 sj_div=="BS" 분기로 처리 → 여기선 제외
    # IS는 _hl_level 내에서 sj_div=="IS" 분기로 처리 → 여기선 제외
    # CIS
    "OtherComprehensiveIncome",
    # CF — 3대 활동 소계만 볼드, 나머지(환율변동·증감·기초) Level 0
    "CashFlowsFromUsedInOperatingActivities",
    "CashFlowsFromUsedInInvestingActivities",
    "CashFlowsFromUsedInFinancingActivities",
}

# CF에서는 ProfitLoss·ComprehensiveIncome이 간접법 출발 세부항목
# (IS/CIS에서는 최상위 합계지만 CF에서는 Level 0이어야 함)
_CF_NOT_TOTAL = {"ProfitLoss", "ComprehensiveIncome"}


def _hl_level(account_id: str, has_values: bool, sj_div: str = "") -> int:
    """
    0 = 세부항목  (흰색, 보통, indent 3)
    1 = 대분류    (노랑, 굵게, 값 없음)  ― BS에서는 미사용
    2 = 중분류/소계 (흰색, 굵게, indent 1) ― BS에서는 미사용
    3 = 합계/총계  (초록, 굵게, indent 0)

    sj_div: 재무제표 구분 코드 (BS/IS/CIS/CF/SCE).
    """
    # ── BS ───────────────────────────────────────────────────────────────────
    if sj_div == "BS":
        if not account_id or "표준계정코드 미사용" in account_id:
            return 0
        el = account_id.rsplit("_", 1)[-1]
        if el in _HL3_ELS:
            return 3
        if el in {"Liabilities", "Equity"}:      # 부채총계, 자본총계 → 노란색
            return 1
        if el in {"CurrentAssets", "NoncurrentAssets"}:  # 유동/비유동자산 → 볼드
            return 2
        return 0

    # ── IS ───────────────────────────────────────────────────────────────────
    if sj_div == "IS":
        if not has_values:
            return 1  # 구조 헤더(당기순이익의 귀속 등) → indent 0
        if not account_id or "표준계정코드 미사용" in account_id:
            return 0
        el = account_id.rsplit("_", 1)[-1]
        if el in {"Revenue", "RevenueFromContractsWithCustomers"}:
            return 3  # 매출액 → 초록
        if el in {"ProfitLossFromOperatingActivities", "OperatingIncomeLoss",
                  "ProfitLoss"}:
            return 1  # 영업이익·당기순이익 → 노란색
        return 0

    # ── CIS ──────────────────────────────────────────────────────────────────
    if sj_div == "CIS":
        if not has_values:
            return 1  # 구조 헤더
        if not account_id or "표준계정코드 미사용" in account_id:
            return 0
        el = account_id.rsplit("_", 1)[-1]
        if el == "ComprehensiveIncome":
            return 3  # 총포괄이익 → 초록
        if el == "ProfitLoss":
            return 1  # 당기순이익 → 노란색 (IS와 통일)
        if el == "OtherComprehensiveIncome":
            return 2  # 기타포괄손익 → 볼드
        return 0

    # ── SCE ──────────────────────────────────────────────────────────────────
    if sj_div == "SCE":
        if not has_values:
            return 1  # 구조 헤더
        if not account_id or "표준계정코드 미사용" in account_id:
            return 0
        el = account_id.rsplit("_", 1)[-1]
        # 기말잔액 계열 → 초록
        if "EndOfPeriod" in el or "ClosingBalance" in el:
            return 3
        # 기초잔액 계열 → 노란색
        if "BeginningOfPeriod" in el or "OpeningBalance" in el:
            return 1
        return 0

    # ── CF ───────────────────────────────────────────────────────────────────
    if not has_values:
        return 1  # 값 없는 행 = 대분류 헤더
    if not account_id or "표준계정코드 미사용" in account_id:
        return 0
    el = account_id.rsplit("_", 1)[-1]
    if sj_div == "CF" and el in _CF_NOT_TOTAL:
        return 0
    if el in _HL3_ELS:
        return 3
    if el in _HL2_ELS:
        return 2
    return 0


def _write_fs_sheet(ws, pivot: pd.DataFrame, id_map: dict,
                    corp_name: str, fs_name: str,
                    years: list[int], period_map: dict,
                    fs_code: str = "") -> dict:
    """재무제표 시트 작성 — 디자인 사양 준수.
    Returns: {display_nm: excel_row}  ← Cash Flow Overview 참조 수식 생성용
    """
    year_cols = [c for c in pivot.columns if isinstance(c, int)]
    n_yr = len(year_cols)
    CB   = 2   # B열 = column index 2
    LAST = CB + n_yr  # 마지막 데이터 열

    def _w(row, col, value=None, fill=_F_WHITE, font=_FT_NORM,
           align=_AL_L, num_fmt=None, border=None):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill = fill; cell.font = font; cell.alignment = align
        if num_fmt:
            cell.number_format = num_fmt
        if border:
            cell.border = border
        return cell

    # ── 1행: 상단 여백 ────────────────────────────────────────────────────
    for c in range(CB, LAST + 1):
        _w(1, c)
    ws.row_dimensions[1].height = 8

    # ── 2행: 메인 제목 ────────────────────────────────────────────────────
    _w(2, CB, f"{corp_name} - {fs_name}", _F_BLACK, _FT_TITLE, _AL_L)
    for c in range(CB + 1, LAST + 1):
        _w(2, c, fill=_F_BLACK)
    ws.merge_cells(start_row=2, start_column=CB, end_row=2, end_column=LAST)
    ws.row_dimensions[2].height = 30

    # ── 3행: 단위 ─────────────────────────────────────────────────────────
    _w(3, CB, "(단위: 원)", _F_WHITE, _FT_UNIT, _AL_L)
    for c in range(CB + 1, LAST + 1):
        _w(3, c)
    ws.row_dimensions[3].height = 16

    # ── 4행: 빈 줄 ────────────────────────────────────────────────────────
    for c in range(CB, LAST + 1):
        _w(4, c)
    ws.row_dimensions[4].height = 8

    # ── 5행: 1차 헤더 (FY 연도) ───────────────────────────────────────────
    _w(5, CB, "과  목", _F_NAVY, _FT_H1, _AL_C)
    for j, yr in enumerate(year_cols):
        _w(5, CB + 1 + j, f"FY{yr}", _F_NAVY, _FT_H1, _AL_C)
    ws.row_dimensions[5].height = 22

    # ── 6행: 2차 헤더 (기간 상세) ─────────────────────────────────────────
    _w(6, CB, "기간", _F_LBLUE, _FT_H2, _AL_C)
    for j, yr in enumerate(year_cols):
        _w(6, CB + 1 + j, period_map.get(yr, ""), _F_LBLUE, _FT_H2, _AL_C)
    ws.row_dimensions[6].height = 18

    # ── 7행: 빈 줄 ────────────────────────────────────────────────────────
    for c in range(CB, LAST + 1):
        _w(7, c)
    ws.row_dimensions[7].height = 6

    # ── 8행~: 데이터 (1패스: 셀 작성 + 레벨 기록) ────────────────────────
    DATA_START = 8
    row_levels: list[tuple[int, int]] = []  # (excel_row, level)
    nm_to_row: dict[str, int] = {}          # display_nm → excel_row (CFO 참조용)

    for i, (_, row_s) in enumerate(pivot.iterrows()):
        er = DATA_START + i
        display_nm = row_s["계정명"]
        nm_to_row[display_nm] = er
        account_id = id_map.get(display_nm, "")
        vals       = [row_s.get(yr) for yr in year_cols]
        has_vals   = any(v is not None and not pd.isna(v) for v in vals)
        level      = _hl_level(account_id, has_vals, fs_code)
        row_levels.append((er, level))

        # 행 스타일
        # Level 3(총계): indent 0 — 가장 돌출되게
        # Level 2(소계): indent 1
        # Level 1(헤더): indent 0
        # Level 0(세부): BS는 indent 1, 그 외 indent 3
        if level == 3:
            fill, font, nm_align = _F_GREEN,  _FT_BOLD, _AL_L
        elif level == 2:
            # BS: 부채총계/자본총계는 총계와 같은 indent 0 (볼드만 구분)
            # IS/CIS/CF: 소계는 indent 1
            sub_align = _AL_L if fs_code == "BS" else _AL_L1
            fill, font, nm_align = _F_WHITE,  _FT_BOLD, sub_align
        elif level == 1:
            fill, font, nm_align = _F_YELLOW, _FT_BOLD, _AL_L
        else:
            # BS: indent 1 / 그 외(IS·CIS·CF·SCE): indent 2
            detail_align = _AL_L1 if fs_code == "BS" else _AL_L2
            fill, font, nm_align = _F_WHITE,  _FT_NORM, detail_align

        _w(er, CB, display_nm, fill, font, nm_align)

        for j, yr in enumerate(year_cols):
            val  = row_s.get(yr)
            empty = val is None or pd.isna(val)
            fnt  = _FT_GREY if empty else font
            cell = _w(er, CB + 1 + j, None if empty else val,
                      fill, fnt, _AL_R, _NUM_FMT)

        ws.row_dimensions[er].height = 17

    # ── 2패스: 테두리 ─────────────────────────────────────────────────────
    # 마지막 총계 행 → double bottom
    last_total = max((r for r, lv in row_levels if lv == 3), default=None)

    for er, level in row_levels:
        if level in (2, 3):
            bot = _S_DOUBLE if er == last_total else _S_THIN
            bdr = Border(top=_S_THIN, bottom=bot)
            for c in range(CB, LAST + 1):
                ws.cell(row=er, column=c).border = bdr

    # ── 열 너비 / 틀 고정 ─────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions[get_column_letter(CB)].width = 40
    for j in range(n_yr):
        ws.column_dimensions[get_column_letter(CB + 1 + j)].width = 17

    ws.freeze_panes = f"{get_column_letter(CB + 1)}8"
    return nm_to_row


# ── Cash Flow Overview 분석 시트 ──────────────────────────────────────────────

def _cfo_extract(raw: pd.DataFrame, sj_div: str, patterns: list,
                 year_cols: list,
                 nm_fallbacks: list | None = None) -> tuple[dict, dict]:
    """raw에서 특정 XBRL 패턴의 계정을 연도별로 추출.
    Returns: (values_dict, display_nm_dict)
      values_dict    : {yr: value_or_None}
      display_nm_dict: {yr: display_nm_or_None}  ← FS 시트 참조 수식 생성용

    Fallback 순서:
      1) 당해 연도 행의 thstrm_amount (account_id 패턴 → account_nm 패턴)
      2) 다음 연도 행의 frmtrm_amount (동일 패턴 순서) ← 전기 비교컬럼 역이용
    """
    sub = raw[raw["sj_div"] == sj_div]
    result:   dict = {yr: None for yr in year_cols}
    found_dnm: dict = {yr: None for yr in year_cols}

    def _search(rows: pd.DataFrame, amount_col: str):
        """rows에서 패턴 매칭 후 (value, display_nm) 반환, 없으면 (None, None)."""
        # 1차: account_id 패턴
        for pat in patterns:
            m = rows[rows["account_id"].str.contains(pat, na=False, regex=False)]
            if not m.empty:
                v = m.iloc[0][amount_col]
                if v is not None and not pd.isna(v):
                    dnm = m.iloc[0]["display_nm"] if "display_nm" in m.columns else None
                    return v, dnm
        # 2차: account_nm 한국어 fallback
        if nm_fallbacks and "account_nm" in rows.columns:
            for nm_pat in nm_fallbacks:
                m = rows[rows["account_nm"].str.contains(nm_pat, na=False, regex=False)]
                m = m[~m["account_nm"].str.contains("합계|총계", na=False)]
                if not m.empty:
                    v = m.iloc[0][amount_col]
                    if v is not None and not pd.isna(v):
                        dnm = m.iloc[0]["display_nm"] if "display_nm" in m.columns else None
                        return v, dnm
        return None, None

    for i, yr in enumerate(year_cols):
        # 1차: 당해 연도 thstrm
        result[yr], found_dnm[yr] = _search(sub[sub["year"] == yr], "thstrm_amount")
        # 2차: 다음 연도 frmtrm (전기 비교컬럼)
        if result[yr] is None and i + 1 < len(year_cols):
            nxt = year_cols[i + 1]
            v, dnm = _search(sub[sub["year"] == nxt], "frmtrm_amount")
            if v is not None:
                result[yr], found_dnm[yr] = v, dnm

    return result, found_dnm


def _cfo_add(a: dict, b: dict, year_cols: list) -> dict:
    """두 연도별 dict를 합산 (None 처리 포함)."""
    out = {}
    for yr in year_cols:
        va, vb = a.get(yr), b.get(yr)
        if va is None and vb is None:
            out[yr] = None
        else:
            out[yr] = (va or 0) + (vb or 0)
    return out


def _write_cfo_sheet(ws, raw: pd.DataFrame, corp_name: str,
                     year_cols: list, period_map: dict,
                     fs_row_maps: dict | None = None,
                     fs_sheet_names: dict | None = None) -> None:
    """Cash Flow Overview 분석 시트 작성.
    fs_row_maps   : {fs_code: {display_nm: excel_row}}  ← FS 시트 참조 수식용
    fs_sheet_names: {fs_code: sheet_name_string}
    """

    PCT_FMT = "0.0%;(0.0%);-"
    CB  = 2   # B열 (label)
    NC  = len(year_cols)
    LAST = CB + NC  # 마지막 데이터 열

    def col(j):  # j=0 → C열, j=1 → D열 ...
        return get_column_letter(CB + 1 + j)

    def _w(row, c, val=None, fill=_F_WHITE, font=_FT_NORM,
           align=_AL_L, fmt=None, border=None):
        cell = ws.cell(row=row, column=c, value=val)
        cell.fill, cell.font, cell.alignment = fill, font, align
        if fmt:   cell.number_format = fmt
        if border: cell.border = border
        return cell

    def _row(r, label, fill=_F_WHITE, font=_FT_NORM,
             vals=None, val_font=None, val_fill=None, fmt=_NUM_FMT,
             border_top=None, border_bot=None):
        """행 전체 작성 헬퍼."""
        _w(r, CB, label, fill, font, _AL_L)
        vf = val_font or font
        vfl = val_fill or fill
        top_bdr = Border(top=border_top) if border_top else None
        bot_bdr = Border(bottom=border_bot) if border_bot else None
        for j in range(NC):
            v = vals[j] if vals else None
            bdr = None
            if border_top and border_bot:
                bdr = Border(top=border_top, bottom=border_bot)
            elif border_top:
                bdr = Border(top=border_top)
            elif border_bot:
                bdr = Border(bottom=border_bot)
            c = CB + 1 + j
            if v is None:
                _w(r, c, None, vfl, _FT_GREY, _AL_R, fmt, bdr)
            else:
                _w(r, c, v, vfl, vf, _AL_R, fmt, bdr)
        # label 셀 테두리
        if border_top or border_bot:
            lbdr = Border(
                top=border_top if border_top else _S_NONE,
                bottom=border_bot if border_bot else _S_NONE,
            )
            ws.cell(row=r, column=CB).border = lbdr
        ws.row_dimensions[r].height = 17

    def _formula_row(r, label, formula_tpl, fill=_F_WHITE, font=_FT_NORM,
                     fmt=_NUM_FMT, border_top=None, border_bot=None):
        """수식 행 작성 — formula_tpl에서 {c} 를 열 문자로 치환."""
        _w(r, CB, label, fill, font, _AL_L)
        for j in range(NC):
            c_ltr = col(j)
            formula = formula_tpl.format(c=c_ltr)
            bdr = None
            if border_top and border_bot:
                bdr = Border(top=border_top, bottom=border_bot)
            elif border_top:
                bdr = Border(top=border_top)
            elif border_bot:
                bdr = Border(bottom=border_bot)
            _w(r, CB + 1 + j, formula, fill, font, _AL_R, fmt, bdr)
        if border_top or border_bot:
            lbdr = Border(
                top=border_top if border_top else _S_NONE,
                bottom=border_bot if border_bot else _S_NONE,
            )
            ws.cell(row=r, column=CB).border = lbdr
        ws.row_dimensions[r].height = 17

    def _blank(r):
        for c in range(CB, LAST + 1):
            _w(r, c, None, _F_WHITE, _FT_NORM)
        ws.row_dimensions[r].height = 8

    # ── 데이터 추출 ──────────────────────────────────────────────────────────
    rev,  rev_dnm  = _cfo_extract(raw, "IS", ["Revenue", "RevenueFromContracts"], year_cols)
    ebit, ebit_dnm = _cfo_extract(raw, "IS",
                                  ["ProfitLossFromOperatingActivities", "OperatingIncomeLoss"],
                                  year_cols)
    cogs, _        = _cfo_extract(raw, "IS", ["CostOfSales", "CostOfGoodsSold"], year_cols)
    ar,   ar_dnm   = _cfo_extract(raw, "BS",
                                  ["TradeAndOtherCurrentReceivables", "TradeReceivables",
                                   "TradeAndOtherReceivables", "CurrentTradeReceivables",
                                   "AccountsAndNotesReceivableCurrent"],
                                  year_cols, nm_fallbacks=["매출채권"])
    inv,  inv_dnm  = _cfo_extract(raw, "BS",
                                  ["Inventories", "CurrentInventories"],
                                  year_cols, nm_fallbacks=["재고자산"])
    ap,   ap_dnm   = _cfo_extract(raw, "BS",
                                  ["TradeAndOtherCurrentPayables", "TradeAndOtherPayables",
                                   "TradePayables", "CurrentTradePayables"],
                                  year_cols, nm_fallbacks=["매입채무"])
    ppe,  _        = _cfo_extract(raw, "CF",
                                  ["PurchaseOfPropertyPlantAndEquipment"], year_cols)
    intan, _       = _cfo_extract(raw, "CF",
                                  ["PurchaseOfIntangibleAssets"], year_cols)
    capex = _cfo_add(ppe, intan, year_cols)

    def yr_vals(d: dict) -> list:
        return [d.get(yr) for yr in year_cols]

    # ── FS 시트 참조 수식 헬퍼 ────────────────────────────────────────────────
    def _ref_vals_simple(values: dict, found_dnm: dict, fs_code: str) -> list:
        """FS 시트 셀 참조 수식. 없으면 값 직접 기입."""
        row_map  = (fs_row_maps   or {}).get(fs_code, {})
        sheet_nm = (fs_sheet_names or {}).get(fs_code, "")
        dnm = next((v for v in found_dnm.values() if v is not None), None)
        row = row_map.get(dnm) if (dnm and row_map) else None
        result = []
        for j, yr in enumerate(year_cols):
            if row and sheet_nm:
                c = get_column_letter(CB + 1 + j)
                result.append(f"=IF('{sheet_nm}'!{c}{row}=\"\",\"\",'{sheet_nm}'!{c}{row})")
            else:
                result.append(values.get(yr))
        return result

    def _nwc_mv_formulas(bal_row: int, outflow: bool) -> list:
        """NWC Movement: 같은 시트 내 잔액 행을 참조하는 수식 리스트.
        outflow=True  → 증가가 현금유출 (매출채권·재고): -(cur-prv)
        outflow=False → 증가가 현금유입 (매입채무):       cur-prv
        """
        result = []
        for j in range(len(year_cols)):
            if j == 0:
                result.append(None)
            else:
                c  = get_column_letter(CB + 1 + j)
                pc = get_column_letter(CB + j)
                if outflow:
                    f = (f"=IF(OR({c}{bal_row}=\"\",{pc}{bal_row}=\"\"),\"\","
                         f"-({c}{bal_row}-{pc}{bal_row}))")
                else:
                    f = (f"=IF(OR({c}{bal_row}=\"\",{pc}{bal_row}=\"\"),\"\","
                         f"{c}{bal_row}-{pc}{bal_row})")
                result.append(f)
        return result

    # ── 행 번호 정의 ─────────────────────────────────────────────────────────
    R = {
        "title": 2, "unit": 3, "blank1": 4,
        "hdr1": 5, "hdr2": 6, "blank2": 7,
        "rev": 8, "growth": 9,
        "ebit": 10, "da": 11, "ebitda": 12, "ebitda_pct": 13,
        "blank3": 14,
        "nwc_hdr": 15,
        "ar_chg": 16, "inv_chg": 17, "ap_chg": 18, "nwc_tot": 19,
        "blank4": 20,
        "ebitda_nwc": 21, "conv_pct": 22,
        "blank5": 23,
        "capex_hdr": 24,
        "capex_tot": 25, "capex_maint": 26, "capex_exp": 27,
        "capex_pct": 28,
        "blank6": 29,
        "fcf": 30, "cash_conv": 31,
        "blank7": 32,
        "ccc_hdr": 33,
        "nwc_bal_cat": 34,
        "ar_bal": 35, "inv_bal": 36, "ap_bal": 37,
        "blank8": 38,
        "ccc_cat": 39,
        "dso": 40, "dio": 41, "dpo": 42,
        "blank9": 43,
        "ccc": 44,
    }

    # ── 타이틀 ───────────────────────────────────────────────────────────────
    title_cell = ws.cell(row=R["title"], column=CB,
                         value=f"{corp_name} - Cash Flow Overview")
    title_cell.fill, title_cell.font, title_cell.alignment = (
        _F_BLACK, _FT_TITLE, _AL_L)
    for c in range(CB + 1, LAST + 1):
        _w(R["title"], c, fill=_F_BLACK, font=_FT_H1)
    ws.row_dimensions[R["title"]].height = 22

    _w(R["unit"], CB, "(단위: 원)", _F_WHITE, _FT_UNIT, _AL_L)
    ws.row_dimensions[R["unit"]].height = 14
    _blank(R["blank1"])

    # ── 헤더 행 ──────────────────────────────────────────────────────────────
    _w(R["hdr1"], CB, "과 목", _F_NAVY, _FT_H1, _AL_C)
    for j, yr in enumerate(year_cols):
        _w(R["hdr1"], CB + 1 + j, f"FY{str(yr)[-2:]}", _F_NAVY, _FT_H1, _AL_C)
    ws.row_dimensions[R["hdr1"]].height = 18

    _w(R["hdr2"], CB, "기간", _F_LBLUE, _FT_H2, _AL_C)
    for j, yr in enumerate(year_cols):
        period = period_map.get(yr, "")
        _w(R["hdr2"], CB + 1 + j, period, _F_LBLUE, _FT_H2, _AL_C)
    ws.row_dimensions[R["hdr2"]].height = 18
    _blank(R["blank2"])

    # ── Revenue ──────────────────────────────────────────────────────────────
    _row(R["rev"], "Revenue", font=_FT_BOLD,
         vals=_ref_vals_simple(rev, rev_dnm, "IS"))
    # Growth % — 첫 열은 이전 연도 없으므로 공백, 이후 열은 수식
    _w(R["growth"], CB, "  Growth %", _F_WHITE, _FT_GREY_ITA, _AL_L)
    _w(R["growth"], CB + 1, None, _F_WHITE, _FT_GREY_ITA, _AL_R, PCT_FMT)
    for j in range(1, NC):
        c_ltr = col(j)
        prev_c = col(j - 1)
        formula = f"=IF({prev_c}{R['rev']}<>0,{c_ltr}{R['rev']}/{prev_c}{R['rev']}-1,\"\")"
        _w(R["growth"], CB + 1 + j, formula, _F_WHITE, _FT_GREY_ITA, _AL_R, PCT_FMT)
    ws.row_dimensions[R["growth"]].height = 17

    # ── EBIT / D&A / EBITDA ──────────────────────────────────────────────────
    _row(R["ebit"], "Operating Profit (EBIT)", font=_FT_BOLD,
         vals=_ref_vals_simple(ebit, ebit_dnm, "IS"), border_bot=_S_THIN)
    # D&A — 수동 입력 (파란색 빈 셀)
    _w(R["da"], CB, "  (+) D&A", _F_WHITE, _FT_ITALIC, _AL_L)
    for j in range(NC):
        _w(R["da"], CB + 1 + j, None, _F_WHITE, _FT_BLUE, _AL_R, _NUM_FMT)
    ws.row_dimensions[R["da"]].height = 17

    _formula_row(R["ebitda"], "EBITDA",
                 f"=IF({{c}}{R['ebit']}=\"\",\"\",{{c}}{R['ebit']}+IF({{c}}{R['da']}=\"\",0,{{c}}{R['da']}))",
                 font=_FT_BOLD, border_top=_S_THIN)
    _formula_row(R["ebitda_pct"], "  EBITDA %",
                 f"=IF({{c}}{R['rev']}<>0,{{c}}{R['ebitda']}/{{c}}{R['rev']},\"\")",
                 font=_FT_GREY_ITA, fmt=PCT_FMT)
    _blank(R["blank3"])

    # ── NWC Movement ─────────────────────────────────────────────────────────
    _w(R["nwc_hdr"], CB, "NWC Movement", _F_YELLOW, _FT_BOLD, _AL_L)
    for c in range(CB + 1, LAST + 1):
        _w(R["nwc_hdr"], c, fill=_F_YELLOW)
    ws.row_dimensions[R["nwc_hdr"]].height = 17

    # NWC Movement: NWC Balance 행(하단 CCC 섹션)을 참조하는 인시트 수식
    _row(R["ar_chg"],  "  매출채권 증감 (증가시 -)", font=_FT_NORM,
         vals=_nwc_mv_formulas(R["ar_bal"],  outflow=True))
    _row(R["inv_chg"], "  재고자산 증감 (증가시 -)", font=_FT_NORM,
         vals=_nwc_mv_formulas(R["inv_bal"], outflow=True))
    _row(R["ap_chg"],  "  매입채무 증감 (증가시 +)", font=_FT_NORM,
         vals=_nwc_mv_formulas(R["ap_bal"],  outflow=False), border_bot=_S_THIN)
    _formula_row(R["nwc_tot"], "Total NWC Movement",
                 f"=IF(AND({{c}}{R['ar_chg']}=\"\",{{c}}{R['inv_chg']}=\"\",{{c}}{R['ap_chg']}=\"\"),\"\","
                 f"IF({{c}}{R['ar_chg']}=\"\",0,{{c}}{R['ar_chg']})"
                 f"+IF({{c}}{R['inv_chg']}=\"\",0,{{c}}{R['inv_chg']})"
                 f"+IF({{c}}{R['ap_chg']}=\"\",0,{{c}}{R['ap_chg']}))",
                 font=_FT_BOLD, border_top=_S_THIN)
    _blank(R["blank4"])

    # ── EBITDA after NWC ─────────────────────────────────────────────────────
    _formula_row(R["ebitda_nwc"], "EBITDA after NWC Movement",
                 f"=IF({{c}}{R['ebitda']}=\"\",\"\",{{c}}{R['ebitda']}+IF({{c}}{R['nwc_tot']}=\"\",0,{{c}}{R['nwc_tot']}))",
                 fill=_F_GREEN, font=_FT_BOLD,
                 border_top=_S_THIN, border_bot=_S_THIN)
    _formula_row(R["conv_pct"], "  Conversion % (post NWC / EBITDA)",
                 f"=IF({{c}}{R['ebitda']}<>0,{{c}}{R['ebitda_nwc']}/{{c}}{R['ebitda']},\"\")",
                 font=_FT_GREY_ITA, fmt=PCT_FMT)
    _blank(R["blank5"])

    # ── CAPEX ────────────────────────────────────────────────────────────────
    _w(R["capex_hdr"], CB, "CAPEX", _F_YELLOW, _FT_BOLD, _AL_L)
    for c in range(CB + 1, LAST + 1):
        _w(R["capex_hdr"], c, fill=_F_YELLOW)
    ws.row_dimensions[R["capex_hdr"]].height = 17

    _row(R["capex_tot"], "  Total CAPEX", font=_FT_BOLD, vals=yr_vals(capex))
    # Maintenance / Expansion — 수동 입력 (파란색)
    _w(R["capex_maint"], CB, "    ㄴ Maintenance CAPEX", _F_WHITE, _FT_GREY_NRM, _AL_L)
    _w(R["capex_exp"],   CB, "    ㄴ Expansion CAPEX",   _F_WHITE, _FT_GREY_NRM, _AL_L)
    for j in range(NC):
        _w(R["capex_maint"], CB + 1 + j, None, _F_WHITE, _FT_BLUE, _AL_R, _NUM_FMT)
        _w(R["capex_exp"],   CB + 1 + j, None, _F_WHITE, _FT_BLUE, _AL_R, _NUM_FMT)
    ws.row_dimensions[R["capex_maint"]].height = 17
    ws.row_dimensions[R["capex_exp"]].height   = 17

    _formula_row(R["capex_pct"], "  CAPEX % of Revenue",
                 f"=IF({{c}}{R['rev']}<>0,IF({{c}}{R['capex_tot']}=\"\",\"\",{{c}}{R['capex_tot']}/{{c}}{R['rev']}),\"\")",
                 font=_FT_GREY_ITA, fmt=PCT_FMT, border_bot=_S_THIN)
    _blank(R["blank6"])

    # ── Free Cash Flow ────────────────────────────────────────────────────────
    _formula_row(R["fcf"], "Free Cash Flow",
                 f"=IF({{c}}{R['ebitda_nwc']}=\"\",\"\",{{c}}{R['ebitda_nwc']}-IF({{c}}{R['capex_tot']}=\"\",0,{{c}}{R['capex_tot']}))",
                 fill=_F_DK_GREEN, font=_FT_BOLD,
                 border_top=_S_THICK, border_bot=_S_DOUBLE)
    _formula_row(R["cash_conv"], "  Cash Conversion % (FCF / EBITDA)",
                 f"=IF({{c}}{R['ebitda']}<>0,{{c}}{R['fcf']}/{{c}}{R['ebitda']},\"\")",
                 font=_FT_GREY_ITA, fmt=PCT_FMT)
    _blank(R["blank7"])

    # ── NWC Balance & CCC ────────────────────────────────────────────────────
    _w(R["ccc_hdr"], CB, "NWC Balance & Cash Conversion Cycle (CCC)",
       _F_YELLOW, _FT_BOLD, _AL_L)
    for c in range(CB + 1, LAST + 1):
        _w(R["ccc_hdr"], c, fill=_F_YELLOW)
    ws.row_dimensions[R["ccc_hdr"]].height = 17

    _w(R["nwc_bal_cat"], CB, "[NWC Balance]", _F_WHITE, _FT_GREY_BLD, _AL_L)
    ws.row_dimensions[R["nwc_bal_cat"]].height = 17

    # NWC Balance: BS 시트 참조 수식 (가능 시) → NWC Movement가 이 셀들을 참조
    _row(R["ar_bal"],  "  매출채권", font=_FT_NORM,
         vals=_ref_vals_simple(ar,  ar_dnm,  "BS"))
    _row(R["inv_bal"], "  재고자산", font=_FT_NORM,
         vals=_ref_vals_simple(inv, inv_dnm, "BS"))
    _row(R["ap_bal"],  "  매입채무", font=_FT_NORM,
         vals=_ref_vals_simple(ap,  ap_dnm,  "BS"))
    _blank(R["blank8"])

    _w(R["ccc_cat"], CB, "[Cash Conversion Cycle (days)]", _F_WHITE, _FT_GREY_BLD, _AL_L)
    ws.row_dimensions[R["ccc_cat"]].height = 17

    # DSO / DIO / DPO 수식 (COGS 있으면 사용, 없으면 Revenue fallback)
    cogs_vals = yr_vals(cogs)
    has_cogs  = any(v is not None for v in cogs_vals)
    denom_row = R["rev"]   # fallback

    if has_cogs:
        # cogs를 별도 숨김 행에 기록 (수식 참조용) — 마지막 행 아래 숨김
        cogs_row = R["ccc"] + 2
        _row(cogs_row, "_cogs_helper", font=_FT_GREY, vals=cogs_vals)
        ws.row_dimensions[cogs_row].height = 0  # 숨김
        denom_row = cogs_row

    _formula_row(R["dso"], "  DSO (매출채권 회수일수)",
                 f"=IF({{c}}{R['rev']}<>0,{{c}}{R['ar_bal']}/{{c}}{R['rev']}*365,\"\")",
                 font=_FT_NORM, fmt="0.0")
    _formula_row(R["dio"], "  DIO (재고 회전일수)",
                 f"=IF({{c}}{denom_row}<>0,{{c}}{R['inv_bal']}/{{c}}{denom_row}*365,\"\")",
                 font=_FT_NORM, fmt="0.0")
    _formula_row(R["dpo"], "  DPO (매입채무 지급일수)",
                 f"=IF({{c}}{denom_row}<>0,{{c}}{R['ap_bal']}/{{c}}{denom_row}*365,\"\")",
                 font=_FT_NORM, fmt="0.0", border_bot=_S_THIN)
    _blank(R["blank9"])
    _formula_row(R["ccc"], "CCC (DSO + DIO - DPO)",
                 f"=IF({{c}}{R['dso']}=\"\",\"\",{{c}}{R['dso']}+{{c}}{R['dio']}-{{c}}{R['dpo']})",
                 fill=_F_GREEN, font=_FT_BOLD, fmt="0.0",
                 border_top=_S_THIN, border_bot=_S_DOUBLE)

    # ── 열 너비 / 틀 고정 ─────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions[get_column_letter(CB)].width = 44
    for j in range(NC):
        ws.column_dimensions[get_column_letter(CB + 1 + j)].width = 17
    ws.freeze_panes = f"{get_column_letter(CB + 1)}{R['hdr1']}"
    ws.sheet_view.showGridLines = False


# ── 데이터 로직 ───────────────────────────────────────────────────────────────

@st.cache_data(ttl=86400, show_spinner=False)
def load_corp_codes() -> pd.DataFrame:
    # 1순위: repo에 번들된 파일 (Streamlit Cloud에서 DART API 차단 대비)
    _here = os.path.dirname(os.path.abspath(__file__))
    _bundle = os.path.join(_here, "corp_codes.json")
    if os.path.exists(_bundle):
        with open(_bundle, encoding="utf-8") as f:
            return pd.DataFrame(json.load(f))
    # 2순위: DART API 직접 호출 (로컬 실행 또는 파일 없을 때)
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

    # Fix #3: 연도별 계정명 변경(리네임) 정규화
    # 동일 account_id가 연도마다 다른 account_nm을 가질 때 가장 최근 연도 기준으로 통일.
    # 예: '영업활동현금흐름표'(FY21) vs '영업활동현금흐름'(FY24) → 동일 IFRS 요소
    # 비표준 계정('-표준계정코드 미사용-')은 account_id가 없으므로 account_nm 유지.
    if "account_id" in raw.columns:
        non_std_mask = raw["account_id"].str.contains("표준계정코드 미사용", na=False)
        std_rows = raw[~non_std_mask]
        if not std_rows.empty:
            # account_id 별 최신 연도의 account_nm을 정규 명칭으로 채택
            canonical = (
                std_rows.sort_values("year", ascending=False)
                .groupby("account_id")["account_nm"]
                .first()
                .to_dict()
            )
            raw.loc[~non_std_mask, "account_nm"] = raw.loc[~non_std_mask, "account_id"].map(
                lambda aid: canonical.get(aid, "")
            )

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

    # display_nm → account_id 매핑 (하이라이트 레벨 판정용)
    id_map = {}
    if "account_id" in raw.columns:
        id_map = raw.groupby("display_nm")["account_id"].first().to_dict()

    # year → 기간 문자열 (헤더 서브행용)
    period_map = {}
    if "thstrm_nm" in raw.columns:
        period_map = raw.groupby("year")["thstrm_nm"].first().to_dict()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 정보 시트
        meta = pd.DataFrame([
            {"항목": "기업명",           "내용": corp_name},
            {"항목": "종목코드",         "내용": stock_code or "-"},
            {"항목": "DART corp_code",   "내용": corp_code},
            {"항목": "재무제표 종류",    "내용": fs_label},
            {"항목": "수집 연도",        "내용": f"{years[0]}~{years[-1]}"},
            {"항목": "수집일",           "내용": pd.Timestamp.now().strftime("%Y-%m-%d")},
        ])
        meta.to_excel(writer, sheet_name="정보", index=False)

        # 원본데이터 시트
        cols_keep = [c for c in ["year", "sj_div", "sj_nm", "account_id",
                                  "account_nm", "thstrm_amount", "fs_div_used"]
                     if c in raw.columns]
        raw[cols_keep].rename(columns={
            "year": "연도", "sj_div": "재무제표구분코드", "sj_nm": "재무제표명",
            "account_id": "계정ID", "account_nm": "계정명",
            "thstrm_amount": "금액(원)", "fs_div_used": "연결/별도",
        }).to_excel(writer, sheet_name="원본데이터", index=False)

        # 재무제표별 스타일 시트
        fs_row_maps: dict[str, dict] = {}   # {fs_code: {display_nm: excel_row}}
        for fs_code, fs_name in FS_LABELS.items():
            subset = raw[raw["sj_div"] == fs_code].copy()
            if subset.empty:
                continue

            if "ord" in subset.columns:
                subset["ord"] = pd.to_numeric(subset["ord"], errors="coerce")
                subset = subset.sort_values(["year", "ord"], na_position="last")

            # SCE: 자본총계(Equity) 열만 추출 → 난잡한 2D 매트릭스 대신 단일 컬럼 뷰
            if fs_code == "SCE" and "account_id" in subset.columns:
                sce_total_mask = subset["account_id"].str.contains(
                    r"(?:^|_)Equity$", regex=True, na=False
                )
                if sce_total_mask.any():
                    subset = subset[sce_total_mask]

            pivot = (
                subset.groupby(["display_nm", "year"], sort=False)["thstrm_amount"]
                .first().unstack("year").reset_index()
            )
            pivot.columns.name = None
            pivot = pivot.rename(columns={"display_nm": "계정명"})
            year_cols = sorted([c for c in pivot.columns if isinstance(c, int)])
            pivot = pivot[["계정명"] + year_cols]

            # ── frmtrm / bfefrmtrm 백필 ──────────────────────────────────────
            # DART는 한 번의 쿼리에 당기(thstrm)·전기(frmtrm)·전전기(bfefrmtrm)를
            # 모두 반환함. thstrm 단독 pivot에서 생긴 NaN을 이미 받아온 값으로 보완.
            # 추정·예측 없음 — API 응답에 이미 있는 숫자만 사용.
            bf: dict[tuple, float] = {}
            for col, offset in [("frmtrm_amount", 1), ("bfefrmtrm_amount", 2)]:
                if col not in subset.columns:
                    continue
                for _, r in subset.iterrows():
                    val = r.get(col)
                    if val is None or pd.isna(val):
                        continue
                    target = int(r["year"]) - offset
                    if target not in year_cols:
                        continue
                    key = (r["display_nm"], target)
                    if key not in bf:          # 인접 연도(frmtrm) 우선
                        bf[key] = val

            if bf:
                for idx, row_p in pivot.iterrows():
                    dnm = row_p["계정명"]
                    for yr in year_cols:
                        if pd.isna(row_p.get(yr)):
                            bv = bf.get((dnm, yr))
                            if bv is not None:
                                pivot.at[idx, yr] = bv

            # 전 기간 NaN 행 제거 ─────────────────────────────────────────────
            # · 표준 account_id 있는 행 → NaN이어도 유지 (구조 헤더: 당기순이익의 귀속 등)
            # · "-표준계정코드 미사용-" 행 → NaN이면 제거 (inapplicable placeholder)
            has_val = pivot[year_cols].notna().any(axis=1)
            is_nonstandard = pivot["계정명"].map(
                lambda nm: "표준계정코드 미사용" in id_map.get(nm, "")
            )
            pivot = pivot[has_val | ~is_nonstandard].reset_index(drop=True)

            sheet_name = fs_name[:31]
            writer.book.create_sheet(sheet_name)
            ws = writer.book[sheet_name]
            nm_to_row = _write_fs_sheet(ws, pivot, id_map, corp_name, fs_name,
                                        year_cols, period_map, fs_code)
            fs_row_maps[fs_code] = nm_to_row

        # ── Cash Flow Overview 분석 시트 ──────────────────────────────────
        all_year_cols = sorted(raw["year"].unique().tolist())
        ws_cfo = writer.book.create_sheet("Cash Flow Overview")
        _write_cfo_sheet(ws_cfo, raw, corp_name, all_year_cols, period_map,
                         fs_row_maps=fs_row_maps,
                         fs_sheet_names={k: v for k, v in FS_LABELS.items()})

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
