"""간단한 소송비용계산서
설치: python -m pip install streamlit python-docx
실행: python -m streamlit run app.py
"""
from decimal import Decimal, ROUND_FLOOR
from io import BytesIO
import re
import streamlit as st
from fractions import Fraction

# 2020.12.28. 개정 변호사보수의 소송비용 산입에 관한 규칙 별표
BRACKETS = [
    (3_000_000, 0, 300_000, "0"),
    (20_000_000, 3_000_000, 300_000, ".10"),
    (50_000_000, 20_000_000, 2_000_000, ".08"),
    (100_000_000, 50_000_000, 4_400_000, ".06"),
    (150_000_000, 100_000_000, 7_400_000, ".04"),
    (200_000_000, 150_000_000, 9_400_000, ".02"),
    (500_000_000, 200_000_000, 10_400_000, ".01"),
    (None, 500_000_000, 13_400_000, ".005"),
]

STAGES = {
    # 민사
    "가합": 1, "가단": 1, "가소": 1, "나": 2, "다": 3,
    # 행정
    "구합": 1, "구단": 1, "누": 2, "두": 3,
    # 가사
    "드합": 1, "드단": 1, "르": 2, "므": 3
}
DELIVERY_UNIT = 5_640
LAW_FIRM = '법무법인(유한)바른'
DELIVERY_ROUNDS = {
    '가합': 15, '가단': 15, '가소': 10, '나': 12, '다': 8,
    '구합': 15, '구단': 15, '누': 12, '두': 8,
    '드합': 15, '드단': 15, '르': 12, '므': 8
}


def won(value):
    return f"{value:,}원"


def parse_fraction(text):
    text = text.strip()
    if not text:
        return Fraction(1, 1)
    try:
        if '/' in text:
            num, den = map(int, text.split('/'))
            return Fraction(num, den)
        else:
            return Fraction(int(text), 1)
    except ValueError:
        raise ValueError("부담비율은 '1/3', '1/1'과 같은 분수 형식으로 입력해주세요.")


def fee_limit(soga):
    if soga <= 0:
        return 0
    for upper, lower, base, rate in BRACKETS:
        if upper is None or soga <= upper:
            value = Decimal(base) + Decimal(soga - lower) * Decimal(rate)
            return int(value.to_integral_value(rounding=ROUND_FLOOR))
    return 0


def parse_cases(text):
    cases = []
    used = set()
    code_pattern = "|".join(STAGES.keys())
    regex = rf"(.*?)\s*(\d{{4}})\s*({code_pattern})\s*(\d+)"
    for line in re.split(r"[\n,;]+", text):
        if not line.strip():
            continue
        match = re.fullmatch(regex, line.strip())
        if not match:
            raise ValueError("사건번호 형식을 확인해주세요. (예: 서울중앙지방법원 2023가합12345, 서울고등법원 2025나12345)")
        court, year, code, number = match.groups()
        stage = STAGES[code]
        if stage in used:
            raise ValueError("같은 심급은 한 건씩 입력해주세요.")
        used.add(stage)
        cases.append({"stage": stage, "court": court.strip(), "number": year + code + number, "code": code})
    if not cases:
        raise ValueError("사건번호를 입력해주세요.")
    return sorted(cases, key=lambda case: case["stage"])


def optional_money(text):
    text = text.strip().replace(",", "").removesuffix("원").strip()
    if not text:
        return None
    if not re.fullmatch(r"[0-9]+", text):
        raise ValueError("금액은 0 이상의 숫자로 입력해주세요. 쉼표는 사용할 수 있습니다.")
    return int(text)


def filing_estimate(soga, case, electronic=True, party_count=1):
    if soga <= 0 or party_count < 1:
        return 0, 0
    if soga < 10_000_000:
        numerator = soga * 50
    elif soga < 100_000_000:
        numerator = soga * 45 + 50_000_000
    elif soga < 1_000_000_000:
        numerator = soga * 40 + 550_000_000
    else:
        numerator = soga * 35 + 5_550_000_000
    
    base = max(1_000, (numerator // 1_000_000) * 100)
    multiplier = {1: Decimal("1.0"), 2: Decimal("1.5"), 3: Decimal("2.0")}[case['stage']]
    stamp = int(Decimal(base) * multiplier)
    if electronic:
        stamp = stamp * 9 // 10
    stamp = max(1_000 if not electronic else 900, (stamp // 100) * 100)
    
    rounds = DELIVERY_ROUNDS.get(case['code'], 15)
    delivery = party_count * rounds * DELIVERY_UNIT
    return stamp, delivery


def fee_formula(soga, is_reduced=False):
    if soga <= 0:
        return "0원"
    for upper, lower, base, rate in BRACKETS:
        if upper is None or soga <= upper:
            if rate == '0':
                formula = won(base)
            else:
                percent = format(Decimal(rate) * 100, 'f').rstrip('0').rstrip('.')
                formula = f"{base:,}원 + ({soga:,}원 − {lower:,}원) × {percent}%"
            if is_reduced:
                formula = f"({formula}) × 1/2 (제5조 보수 감액)"
            return formula


def make_model(cases, soga_dict, actual_fees, filing_costs, options):
    is_reduced = options.get('is_reduced', False)
    stage_settings = options.get('stage_settings', {})
    rows = []
    filing_rows = []
    
    main_total = 0
    borne_main_total = 0
    
    for case in cases:
        c_num = case["number"]
        stage = case["stage"]
        soga = soga_dict.get(c_num, 0)
        if soga <= 0:
            raise ValueError(f"{stage}심 소가를 올바르게 입력해주세요.")
            
        limit = fee_limit(soga)
        if is_reduced:
            limit = limit // 2
            
        actual = actual_fees.get(c_num)
        if actual is not None and actual < 0:
            raise ValueError(f"{stage}심 실제 보수는 음수일 수 없습니다.")
            
        calc_fee = limit if actual is None else min(limit, actual)
        
        # 인지대 및 송달료
        f_entry = filing_costs.get(c_num, {})
        include = f_entry.get('include', False)
        stamp = f_entry.get('stamp', 0) if include else 0
        delivery = f_entry.get('delivery', 0) if include else 0
        filing_rows.append(dict(stage=stage, number=c_num, stamp=stamp, delivery=delivery, include=include))

        # 추가 수기 금액 및 부담비율(분수) 적용
        settings = stage_settings.get(c_num, {'fraction': Fraction(1, 1), 'manual_add': 0})
        fraction = settings['fraction']
        manual_add = settings['manual_add']
        
        stage_sum = calc_fee + stamp + delivery + manual_add
        stage_borne = int(Decimal(stage_sum) * Decimal(fraction.numerator) / Decimal(fraction.denominator))
        
        rows.append({
            "stage": stage,
            "number": c_num,
            "soga": soga,
            "limit": limit,
            "actual": actual,
            "calc_fee": calc_fee,
            "manual_add": manual_add,
            "fraction": fraction,
            "stage_sum": stage_sum,
            "stage_borne": stage_borne
        })
        
        main_total += stage_sum
        borne_main_total += stage_borne

    # 신청사건 비용
    resp_count = options.get('respondent_count', 1)
    app_stamp = 900 if options.get('electronic', True) else 1_000
    app_delivery = (resp_count + 1) * 3 * DELIVERY_UNIT
    application_total = app_stamp + app_delivery

    final_total = borne_main_total + application_total
    provisional = any(r["actual"] is None for r in rows)
    is_all_won = all(r['fraction'] == Fraction(1, 1) for r in rows)

    # 상대방 소송비용 상계 계산 (Expander 옵션)
    use_offset = options.get('use_offset', False)
    opp_rows = []
    opp_total = 0
    offset_final = final_total

    if use_offset:
        opp_settings = options.get('opp_settings', {})
        for case in cases:
            c_num = case["number"]
            stage = case["stage"]
            soga = soga_dict.get(c_num, 0)
            settings = stage_settings.get(c_num, {'fraction': Fraction(1, 1)})
            # 신청인 부담비율 = 1 - 피신청인 부담비율
            app_burden_fraction = max(Fraction(0, 1), Fraction(1, 1) - settings['fraction'])
            
            o_entry = opp_settings.get(c_num, {})
            o_limit = fee_limit(soga)
            if is_reduced:
                o_limit = o_limit // 2
            o_actual =
