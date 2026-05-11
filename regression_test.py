"""FY25 삼성전자 회귀 테스트."""
import requests

API_KEY = "f0b490beadb3e0200407e2f237b58bca1ae74ac4"
CF_OUTFLOW_PATTERNS = [
    "PurchaseOf", "Repayments", "DividendsPaid", "InterestPaid",
    "IncomeTaxesPaid", "AcquisitionOfTreasury", "DecreaseInNoncontrolling",
]

EXPECTED_CF_NEGATIVES = [
    "이자의 지급", "법인세 납부액", "장기금융상품의 취득",
    "기타포괄손익-공정가치금융자산의 취득", "당기손익-공정가치금융자산의 취득",
    "관계기업 및 공동기업 투자의 취득", "유형자산의 취득", "무형자산의 취득",
    "사업결합으로 인한 순현금유출액", "사채 및 장기차입금의 상환",
    "배당금의 지급", "자기주식의 취득", "비지배지분의 증감",
]


def is_cf_outflow(aid: str) -> bool:
    if not aid or "표준계정코드 미사용" in aid:
        return False
    if any(p in aid for p in CF_OUTFLOW_PATTERNS):
        return True
    if "CashFlowsUsedIn" in aid and "CashFlowsFromUsedIn" not in aid:
        return True
    return False


def parse_amount(val):
    try:
        s = str(val).replace(",", "").strip()
        return int(s) if s not in ("", "-", "None") else None
    except (ValueError, TypeError):
        return None


def run():
    print("FY25 삼성전자 회귀 테스트 시작...")
    r = requests.get(
        "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json",
        params={"crtfc_key": API_KEY, "corp_code": "00126380",
                "bsns_year": "2025", "reprt_code": "11011", "fs_div": "CFS"},
        timeout=20,
    )
    items = r.json().get("list", [])
    cf_items = [x for x in items if x.get("sj_div") == "CF"]
    cis_items = [x for x in items if x.get("sj_div") == "CIS"]

    failures = []

    # ── Bug #2: CF 음수 13건 검증 ──────────────────────────────────────────
    print("\n[Bug #2] CF 음수 부호 검증 (13건)")
    cf_by_name = {}
    for x in cf_items:
        cf_by_name[x["account_nm"]] = x

    for nm in EXPECTED_CF_NEGATIVES:
        item = cf_by_name.get(nm)
        if item is None:
            failures.append(f"  MISSING: {nm}")
            print(f"  MISSING | {nm}")
            continue
        aid = item.get("account_id", "")
        raw_amt = parse_amount(item.get("thstrm_amount"))
        if raw_amt is None:
            failures.append(f"  NULL amount: {nm}")
            continue
        fixed = -raw_amt if (is_cf_outflow(aid) and raw_amt > 0) else raw_amt
        status = "PASS" if fixed < 0 else "FAIL"
        print(f"  {status} | {nm}: {fixed:,}")
        if status == "FAIL":
            failures.append(f"  CF 양수 오류: {nm} = {fixed:,}")

    # ── Bug #1: CIS 중복 계정명 2건 모두 존재 검증 ────────────────────────
    print("\n[Bug #1] CIS 중복 계정명 검증")
    dup_nm = "관계기업 및 공동기업의 기타포괄손익에 대한 지분"
    dup_items = [x for x in cis_items if x["account_nm"] == dup_nm]

    expected = {
        "WillNotBeReclassified": 92_804_000_000,
        "WillBeReclassified":   244_161_000_000,
    }
    found = {}
    for x in dup_items:
        aid = x.get("account_id", "")
        amt = parse_amount(x.get("thstrm_amount"))
        for key in expected:
            if key in aid:
                found[key] = amt

    for key, exp_amt in expected.items():
        got = found.get(key)
        label = "재분류X" if "Not" in key else "재분류O"
        if got == exp_amt:
            print(f"  PASS | [{label}] {dup_nm}: {got:,}")
        else:
            msg = f"  FAIL | [{label}] {dup_nm}: expected {exp_amt:,}, got {got}"
            print(msg)
            failures.append(msg)

    # ── 최종 결과 ─────────────────────────────────────────────────────────
    print("\n" + "=" * 50)
    if not failures:
        print("ALL TESTS PASSED")
    else:
        print(f"FAILED ({len(failures)} errors):")
        for f in failures:
            print(f)
        raise SystemExit(1)


if __name__ == "__main__":
    run()
