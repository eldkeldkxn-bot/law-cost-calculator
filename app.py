"""간단한 소송비용계산서
설치: python -m pip install streamlit python-docx
실행: python -m streamlit run app.py
"""
from decimal import Decimal, ROUND_FLOOR
from io import BytesIO
import html
import re
import streamlit as st
import streamlit.components.v1 as components
from fractions import Fraction

# 2020.12.28. 개정 변호사보수 규칙 별표
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

COURT_CARD_CSS = """
.court-doc-card {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 32px 36px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
    margin-top: 12px;
    margin-bottom: 24px;
    color: #0f172a;
    font-family: 'Pretendard', 'Malgun Gothic', '맑은 고딕', serif;
}
.court-doc-title {
    text-align: center;
    font-size: 22px;
    font-weight: 800;
    letter-spacing: 0.25em;
    margin-bottom: 22px;
    color: #0f172a;
}
.court-meta-box {
    background: #f8fafc;
    border-left: 4px solid #334155;
    padding: 12px 16px;
    font-size: 14px;
    line-height: 1.6;
    margin-bottom: 16px;
    border-radius: 0 6px 6px 0;
}
.court-burden-list {
    margin-top: 6px;
    padding-top: 6px;
    border-top: 1px dashed #cbd5e1;
    color: #334155;
    font-size: 13.5px;
}
.court-table-topbar {
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    margin-bottom: 6px;
    font-size: 13.5px;
}
.court-table-topbar .sec-label {
    font-weight: 700;
    font-size: 14.5px;
}
.court-table-topbar .unit-label {
    color: #64748b;
    font-size: 12.5px;
}
.court-cost-table {
    width: 100% !important;
    border-collapse: collapse !important;
    border-top: 2px solid #0f172a !important;
    border-bottom: 2px solid #0f172a !important;
    font-size: 13.5px !important;
    table-layout: fixed !important;
    background: #ffffff !important;
}
.court-cost-table th,
.court-cost-table td {
    border: 1px solid #cbd5e1 !important;
    padding: 9px 12px !important;
    vertical-align: middle !important;
    word-break: keep-all !important;
    line-height: 1.45 !important;
}
.court-cost-table thead th {
    background-color: #f1f5f9 !important;
    color: #0f172a !important;
    font-weight: 700 !important;
    text-align: center !important;
    border-bottom: 1.5px solid #64748b !important;
}
.cell-stage {
    background-color: #fcfcfd !important;
    font-weight: 700 !important;
    text-align: center !important;
    color: #0f172a !important;
}
.cell-center {
    text-align: center !important;
}
.cell-right {
    text-align: right !important;
    font-variant-numeric: tabular-nums !important;
    font-weight: 500 !important;
    white-space: nowrap !important;
}
.cell-left {
    text-align: left !important;
    color: #334155 !important;
    font-size: 13px !important;
}
.row-subtotal td {
    background-color: #f8fafc !important;
    font-weight: 600 !important;
}
.row-main-subtotal td {
    background-color: #f1f5f9 !important;
    font-weight: 700 !important;
    color: #0f172a !important;
    border-top: 1.5px solid #94a3b8 !important;
    border-bottom: 1.5px solid #94a3b8 !important;
}
.row-grand-total td {
    background-color: #eff6ff !important;
    font-weight: 700 !important;
    color: #1e3a8a !important;
    font-size: 14.5px !important;
    border-top: 2px solid #2563eb !important;
}
.court-reimburse-box {
    margin-top: 14px;
    padding: 12px 16px;
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    font-weight: 700;
    font-size: 14.5px;
    color: #0f172a;
    line-height: 1.6;
}
"""


def won(value):
    return f"{value:,}원"


def trunc_1(value):
    """1원 미만 버림 (소수점 이하만 버리고 1원 단위까지 보존)"""
    return int(Decimal(value).to_integral_value(rounding=ROUND_FLOOR))


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
        raise ValueError("부담비율은 '1/3', '9/10', '1/1'과 같은 분수 형식으로 입력해주세요.")


def format_frac(frac):
    """분수를 법원 서식 비율(백분율 또는 기약분수)로 변환"""
    if frac == Fraction(1, 1):
        return "100%"
    if frac == Fraction(0, 1):
        return "0%"
    if 100 % frac.denominator == 0:
        return f"{int((frac.numerator / frac.denominator) * 100)}%"
    return f"{frac.numerator}/{frac.denominator}"


def fee_limit(soga):
    if soga <= 0:
        return 0
    for upper, lower, base, rate in BRACKETS:
        if upper is None or soga <= upper:
            value = Decimal(base) + Decimal(soga - lower) * Decimal(rate)
            return trunc_1(value)
    return 0


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


def parse_cases(text):
    if not text.strip():
        return [{"stage": 1, "court": "", "number": "", "code": "가단"}]

    cases = []
    used = set()
    code_pattern = "|".join(STAGES.keys())
    regex = rf"(.*?)\s*(\d{{4}})\s*({code_pattern})\s*(\d+)"
    for line in re.split(r"[\n,;]+", text):
        if not line.strip():
            continue
        match = re.fullmatch(regex, line.strip())
        if not match:
            raise ValueError("사건번호 형식을 확인해주세요. (예: 서울중앙지방법원 2023가합12345)")
        court, year, code, number = match.groups()
        stage = STAGES[code]
        if stage in used:
            raise ValueError("같은 심급은 한 건씩 입력해주세요.")
        used.add(stage)
        cases.append({"stage": stage, "court": court.strip(), "number": year + code + number, "code": code})
    if not cases:
        return [{"stage": 1, "court": "", "number": "", "code": "가단"}]
    return sorted(cases, key=lambda case: case["stage"])


def optional_money(text):
    text = text.strip().replace(",", "").removesuffix("원").strip()
    if not text:
        return None
    if not re.fullmatch(r"[0-9]+", text):
        raise ValueError("금액은 0 이상의 숫자로 입력해주세요.")
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


def get_applicant_labels(options):
    """공동신청인 수 및 이름 설정에 따른 열 제목 리스트 생성"""
    app_count = options.get('applicant_count', 1)
    raw_names = options.get('applicant_names', '').strip()
    names = [n.strip() for n in raw_names.split(',') if n.strip()]
    if app_count == 1:
        return [f"신청인 ({names[0]})" if names else "신청인"]
    labels = []
    for i in range(app_count):
        if i < len(names):
            labels.append(f"신청인{i+1} {names[i]}")
        else:
            labels.append(f"신청인{i+1}")
    return labels


def parse_weights(weight_str, app_count):
    """공동신청인 안분 비율 파싱 (예: '1:1' 또는 '920:520')"""
    if app_count <= 1 or not weight_str.strip():
        return [Fraction(1, app_count) for _ in range(app_count)]
    parts = [p.strip().replace(",", "") for p in re.split(r"[:,/]+|\s+", weight_str.strip()) if p.strip()]
    if len(parts) != app_count:
        raise ValueError(f"안분 비율은 신청인 수({app_count}명)에 맞춰 {app_count}개의 숫자로 입력해주세요. (예: 1:1 또는 920:520)")
    try:
        nums = [Decimal(p) for p in parts]
    except Exception:
        raise ValueError("안분 비율은 숫자로 입력해주세요.")
    total_w = sum(nums)
    if total_w <= 0:
        raise ValueError("안분 비율의 합은 0보다 커야 합니다.")
    int_nums = [int(n * 10000) for n in nums]
    int_total = sum(int_nums)
    return [Fraction(x, int_total) for x in int_nums]


def make_model(cases, soga_dict, actual_fees, filing_costs, options):
    is_reduced = options.get('is_reduced', False)
    exclude_app_cost = options.get('exclude_app_cost', False)
    stage_settings = options.get('stage_settings', {})
    app_count = options.get('applicant_count', 1)
    app_labels = get_applicant_labels(options)
    weights = options.get('app_weights', [Fraction(1, app_count)] * app_count)

    rows = []
    filing_rows = []
    main_total = 0
    borne_main_total = 0
    per_borne_main_totals = [0] * app_count

    for case in cases:
        c_num = case["number"]
        c_key = c_num if c_num else f"stage_{case['stage']}"
        stage = case["stage"]
        soga = soga_dict.get(c_key, 0)
        if soga <= 0:
            raise ValueError(f"{stage}심 소가를 올바르게 입력해주세요.")

        limit = fee_limit(soga)
        if is_reduced:
            limit = trunc_1(Decimal(limit) / Decimal(2))

        actual = actual_fees.get(c_key)
        if actual is not None and actual < 0:
            raise ValueError(f"{stage}심 실제 보수는 음수일 수 없습니다.")

        calc_fee = limit if actual is None else min(limit, actual)

        f_entry = filing_costs.get(c_key, {})
        include = f_entry.get('include', False)
        stamp = f_entry.get('stamp', 0) if include else 0
        delivery = f_entry.get('delivery', 0) if include else 0
        filing_rows.append(dict(stage=stage, number=c_num, stamp=stamp, delivery=delivery, include=include))

        settings = stage_settings.get(c_key, {
            'fraction': Fraction(1, 1),
            'fractions': [Fraction(1, 1)] * app_count,
            'manual_name': '기타수기비용',
            'manual_add': 0
        })
        fractions = settings.get('fractions')
        if not fractions or len(fractions) != app_count:
            fractions = [settings.get('fraction', Fraction(1, 1))] * app_count
        fraction = fractions[0]
        manual_name = settings['manual_name']
        manual_add = settings['manual_add']

        per_limits = [trunc_1(Decimal(limit) * Decimal(w.numerator) / Decimal(w.denominator)) for w in weights]
        per_calc_fees = [trunc_1(Decimal(calc_fee) * Decimal(w.numerator) / Decimal(w.denominator)) for w in weights]
        per_stamps = [trunc_1(Decimal(stamp) * Decimal(w.numerator) / Decimal(w.denominator)) for w in weights]
        per_deliveries = [trunc_1(Decimal(delivery) * Decimal(w.numerator) / Decimal(w.denominator)) for w in weights]
        per_manuals = [trunc_1(Decimal(manual_add) * Decimal(w.numerator) / Decimal(w.denominator)) for w in weights]

        per_stage_sums = [
            per_calc_fees[i] + per_stamps[i] + per_deliveries[i] + per_manuals[i]
            for i in range(app_count)
        ]
        per_stage_bornes = [
            trunc_1(Decimal(per_stage_sums[i]) * Decimal(fractions[i].numerator) / Decimal(fractions[i].denominator))
            for i in range(app_count)
        ]

        if app_count == 1:
            stage_sum = calc_fee + stamp + delivery + manual_add
            stage_borne = trunc_1(Decimal(stage_sum) * Decimal(fraction.numerator) / Decimal(fraction.denominator))
        else:
            stage_sum = sum(per_stage_sums)
            stage_borne = sum(per_stage_bornes)

        rows.append({
            "stage": stage,
            "number": c_num,
            "court": case["court"],
            "soga": soga,
            "limit": limit,
            "per_limits": per_limits,
            "actual": actual,
            "calc_fee": calc_fee,
            "per_calc_fees": per_calc_fees,
            "stamp": stamp,
            "per_stamps": per_stamps,
            "delivery": delivery,
            "per_deliveries": per_deliveries,
            "manual_name": manual_name,
            "manual_add": manual_add,
            "per_manuals": per_manuals,
            "fraction": fraction,
            "fractions": fractions,
            "stage_sum": stage_sum,
            "stage_borne": stage_borne,
            "per_stage_sums": per_stage_sums,
            "per_stage_bornes": per_stage_bornes
        })

        main_total += stage_sum
        borne_main_total += stage_borne
        for i in range(app_count):
            per_borne_main_totals[i] += per_stage_bornes[i]

    resp_count = options.get('respondent_count', 1)
    if exclude_app_cost:
        app_stamp = 0
        app_delivery = 0
        application_total = 0
        per_app_stamps = [0] * app_count
        per_app_deliveries = [0] * app_count
        per_application_totals = [0] * app_count
    else:
        app_stamp = 900 if options.get('electronic', True) else 1_000
        app_delivery = (resp_count + app_count) * 3 * DELIVERY_UNIT
        application_total = app_stamp + app_delivery

        per_app_stamps = [trunc_1(Decimal(app_stamp) / Decimal(app_count))] * app_count
        per_app_deliveries = [trunc_1(Decimal(app_delivery) / Decimal(app_count))] * app_count
        per_application_totals = [per_app_stamps[i] + per_app_deliveries[i] for i in range(app_count)]

    per_final_totals = [per_borne_main_totals[i] + per_application_totals[i] for i in range(app_count)]
    final_total = borne_main_total + application_total if app_count == 1 else sum(per_final_totals)
    provisional = any(r["actual"] is None for r in rows)
    is_all_won = all(all(f == Fraction(1, 1) for f in r['fractions']) for r in rows)

    use_offset = options.get('use_offset', False)
    hide_opp_in_doc = options.get('hide_opp_in_doc', False)
    apply_offset_in_doc = use_offset and not hide_opp_in_doc

    opp_rows = []
    opp_total = 0
    per_opp_totals = [0] * app_count
    opp_settings = options.get('opp_settings', {})

    for case in cases:
        c_num = case["number"]
        c_key = c_num if c_num else f"stage_{case['stage']}"
        stage = case["stage"]
        soga = soga_dict.get(c_key, 0)
        settings = stage_settings.get(c_key, {'fraction': Fraction(1, 1), 'fractions': [Fraction(1, 1)] * app_count})
        fractions = settings.get('fractions')
        if not fractions or len(fractions) != app_count:
            fractions = [settings.get('fraction', Fraction(1, 1))] * app_count

        app_burden_fractions = [max(Fraction(0, 1), Fraction(1, 1) - f) for f in fractions]
        app_burden_fraction = app_burden_fractions[0]

        o_entry = opp_settings.get(c_key, {})
        o_limit = fee_limit(soga)
        if is_reduced:
            o_limit = trunc_1(Decimal(o_limit) / Decimal(2))
        o_actual = o_entry.get('fee')
        o_calc_fee = o_limit if o_actual is None else min(o_limit, o_actual)

        o_stamp = o_entry.get('stamp', 0) if o_entry.get('include', False) else 0
        o_delivery = o_entry.get('delivery', 0) if o_entry.get('include', False) else 0
        o_manual = o_entry.get('manual', 0)

        o_sum = o_calc_fee + o_stamp + o_delivery + o_manual
        per_o_sums = [trunc_1(Decimal(o_sum) * Decimal(w.numerator) / Decimal(w.denominator)) for w in weights]
        per_o_bornes = [
            trunc_1(Decimal(per_o_sums[i]) * Decimal(app_burden_fractions[i].numerator) / Decimal(app_burden_fractions[i].denominator))
            for i in range(app_count)
        ]

        if app_count == 1:
            o_borne = trunc_1(Decimal(o_sum) * Decimal(app_burden_fraction.numerator) / Decimal(app_burden_fraction.denominator))
        else:
            o_borne = sum(per_o_bornes)

        opp_rows.append({
            "stage": stage,
            "number": c_num,
            "calc_fee": o_calc_fee,
            "actual": o_actual,
            "limit": o_limit,
            "stamp": o_stamp,
            "delivery": o_delivery,
            "manual": o_manual,
            "opp_sum": o_sum,
            "app_fraction": app_burden_fraction,
            "app_fractions": app_burden_fractions,
            "per_opp_bornes": per_o_bornes,
            "opp_borne": o_borne
        })
        if use_offset:
            opp_total += o_borne
            for i in range(app_count):
                per_opp_totals[i] += per_o_bornes[i]

    per_offset_finals = [
        (per_final_totals[i] - per_opp_totals[i]) if use_offset else per_final_totals[i]
        for i in range(app_count)
    ]
    offset_final = final_total - opp_total if (use_offset and app_count == 1) else (sum(per_offset_finals) if use_offset else final_total)

    return dict(
        rows=rows, filing_rows=filing_rows, main_total=main_total,
        borne_main_total=borne_main_total, per_borne_main_totals=per_borne_main_totals,
        application_total=application_total, per_application_totals=per_application_totals,
        app_stamp=app_stamp, app_delivery=app_delivery,
        per_app_stamps=per_app_stamps, per_app_deliveries=per_app_deliveries,
        total=final_total, per_totals=per_final_totals,
        provisional=provisional, respondent_count=resp_count,
        applicant_count=app_count, applicant_labels=app_labels,
        app_weights=weights,
        is_all_won=is_all_won, exclude_app_cost=exclude_app_cost,
        use_offset=use_offset, hide_opp_in_doc=hide_opp_in_doc,
        apply_offset_in_doc=apply_offset_in_doc,
        opp_rows=opp_rows, opp_total=opp_total, per_opp_totals=per_opp_totals,
        offset_final=offset_final, per_offset_finals=per_offset_finals
    )


def application_sections(cases, model, options):
    refs = [f"{c['court']} {c['number']}".strip() for c in cases]
    refs = [r for r in refs if r]
    references = ', '.join(refs) if refs else '[법원 및 사건번호]'
    name = options['case_name'].strip() or '[사건명]'
    lawyer = LAW_FIRM
    date_text = (f"이 판결은 {options['final_date'].strip()} 확정되었습니다."
                 if options['final_date'].strip() else '')
    is_all_won = model['is_all_won']
    apply_offset = model.get('apply_offset_in_doc', False)
    app_count = model.get('applicant_count', 1)
    app_labels = model.get('applicant_labels', ['신청인'])

    if app_count == 1:
        claim_amount = model['offset_final'] if apply_offset and model['offset_final'] > 0 else model['total']
        prayer = (f'위 당사자 사이의 {references} {name} 사건의 판결에 의하여 피신청인이 신청인에게 '
                  f'상환해야 할 소송비용액은 금 {won(claim_amount)}임을 확정한다.\n라는 결정을 구합니다.')
    else:
        claims = model['per_offset_finals'] if apply_offset else model['per_totals']
        each_claims = ", ".join([f"{lbl}에게 금 {won(max(0, claims[i]))}" for i, lbl in enumerate(app_labels)])
        prayer = (f'위 당사자 사이의 {references} {name} 사건의 판결에 의하여 피신청인이 '
                  f'{each_claims}을(를) 각 상환해야 할 소송비용액임을 확정한다.\n라는 결정을 구합니다.')

    app_word = "신청인들" if app_count > 1 else "신청인"
    if is_all_won:
        opening = f'{app_word}({options["role"]})을 상대로 피신청인이 제기한 {references} {name} 사건에서 최종적으로 전부승소하였습니다. ' if options['role'] == '피고' else f'{app_word}({options["role"]})이 피신청인을 상대로 제기한 {references} {name} 사건에서 최종적으로 전부승소하였습니다. '
        action = '응소' if options['role'] == '피고' else '제기·수행'
        burden_text = '위 사건의 소송비용은 피신청인이 전액 부담하도록 정하여졌으므로'
    else:
        opening = f'{app_word}({options["role"]})을 상대로 피신청인이 제기한 {references} {name} 사건에서 최종적으로 일부승소 판결이 선고되었습니다. ' if options['role'] == '피고' else f'{app_word}({options["role"]})이 피신청인을 상대로 제기한 {references} {name} 사건에서 최종적으로 일부승소 판결이 선고되었습니다. '
        action = '응소' if options['role'] == '피고' else '제기·수행'
        burden_text = '위 사건들의 각 판결에서 정해진 피신청인의 소송비용 부담비율을 적용하여 산정하였으므로'

    offset_mention = " 또한, 당사자 쌍방의 부담액에 관하여 민사소송법 제114조에 따라 대등액에서 상계하고 남은 잔액을 청구합니다." if apply_offset else ""

    cause = (
        f'1. {opening}{date_text}\n\n'
        f'2. {app_word}은 위 사건을 {action}하기 위하여 {lawyer}을 소송대리인으로 선임하였으며, 상환을 구하는 비용은 별지 소송비용액계산서와 같습니다.\n\n'
        f'3. {burden_text},{offset_mention} {app_word}은 별지와 같은 금원의 상환을 구하기 위하여 이 사건 신청에 이르렀습니다.'
    )
    return prayer, cause


def generate_table_data(cases, model, options):
    """표 데이터 생성 (공동신청인 N명 열 분리, 피신청인 표시 여부 및 신청비용 제외 옵션 호환)"""
    table_rows = []
    show_opp = model.get('apply_offset_in_doc', False)
    is_all_won = model.get('is_all_won', True)
    exclude_app_cost = model.get('exclude_app_cost', False)
    app_count = model.get('applicant_count', 1)
    weights = model.get('app_weights', [Fraction(1, app_count)] * app_count)
    is_equal_w = all(w == Fraction(1, app_count) for w in weights)

    for idx, row in enumerate(model['rows']):
        stage_str = f"{row['stage']}심"
        stage_rows = []
        opp_row = model['opp_rows'][idx] if idx < len(model.get('opp_rows', [])) else {}
        formula_str = fee_formula(row['soga'], options.get('is_reduced', False))

        # 1. 변호사보수
        if app_count > 1:
            w_desc = f"1/{app_count} 균등안분" if is_equal_w else "청구비율 안분"
            if row['actual'] is not None and row['actual'] < row['limit']:
                app_note = f"합산 규칙 제3조 한도 {won(row['limit'])} 이내 실지출액 {won(row['actual'])} ({w_desc})"
            elif row['actual'] is not None:
                app_note = f"규칙 제3조 합산 한도: {formula_str} = {won(row['limit'])} (합산 실지출: {won(row['actual'])}, {w_desc})"
            else:
                app_note = f"규칙 제3조 합산 한도: {formula_str} = {won(row['limit'])} ({w_desc})"
        else:
            if row['actual'] is not None and row['actual'] < row['limit']:
                app_note = f"규칙 제3조 산정: {formula_str} = {won(row['limit'])}이나, 실제 지급한 변호사보수를 산입함"
            elif row['actual'] is not None:
                app_note = f"변호사보수의 소송비용산입에 관한 규칙 제3조에 의거 산정하면\n{formula_str} = {won(row['limit'])} (실제 지급액: {won(row['actual'])})"
            else:
                app_note = f"변호사보수의 소송비용산입에 관한 규칙 제3조에 의거 산정하면\n{formula_str} = {won(row['limit'])}"

        opp_val_str = '-'
        opp_note = ''
        if show_opp and opp_row:
            opp_val_str = won(opp_row['calc_fee'])
            if opp_row['actual'] is not None and opp_row['actual'] < opp_row['limit']:
                opp_note = f"규칙 제3조 산정: {formula_str} = {won(opp_row['limit'])}이나, 실제 지급한 변호사보수를 산입함"
            elif opp_row['actual'] is not None:
                opp_note = f"규칙 제3조 산정: {formula_str} = {won(opp_row['limit'])} (실제 지급액: {won(opp_row['actual'])})"
            else:
                opp_note = f"규칙 제3조 산정: {formula_str} = {won(opp_row['limit'])}"

        if show_opp:
            note_lines = [f"[신청인] {app_note}"]
            if opp_note:
                note_lines.append(f"[피신청인] {opp_note}")
        else:
            note_lines = [app_note]

        stage_rows.append({
            '비목': '변호사보수',
            'app_val': won(row['calc_fee']),
            'app_vals': [won(v) for v in row['per_calc_fees']],
            'opp_val': opp_val_str,
            '비고': "\n".join(note_lines),
            'row_type': 'item'
        })

        # 2. 인지대
        has_app_stamp = row['stamp'] > 0
        has_opp_stamp = show_opp and opp_row.get('stamp', 0) > 0
        if has_app_stamp or has_opp_stamp:
            stamp_note = f"합산 {won(row['stamp'])} 안분" if (has_app_stamp and app_count > 1) else ""
            stage_rows.append({
                '비목': '인지대',
                'app_val': won(row['stamp']) if has_app_stamp else '-',
                'app_vals': ([won(v) for v in row['per_stamps']]) if has_app_stamp else (['-'] * app_count),
                'opp_val': won(opp_row['stamp']) if has_opp_stamp else '-',
                '비고': stamp_note,
                'row_type': 'item'
            })

        # 3. 송달료
        has_app_del = row['delivery'] > 0
        has_opp_del = show_opp and opp_row.get('delivery', 0) > 0
        if has_app_del or has_opp_del:
            del_note = f"실제 지출비용 산입 (합산 {won(row['delivery'])} 안분)" if (has_app_del and app_count > 1) else "실제 지출비용 산입"
            stage_rows.append({
                '비목': '송달료',
                'app_val': won(row['delivery']) if has_app_del else '-',
                'app_vals': ([won(v) for v in row['per_deliveries']]) if has_app_del else (['-'] * app_count),
                'opp_val': won(opp_row['delivery']) if has_opp_del else '-',
                '비고': del_note,
                'row_type': 'item'
            })

        # 4. 기타수기비용
        has_app_man = row['manual_add'] > 0
        has_opp_man = show_opp and opp_row.get('manual', 0) > 0
        if has_app_man or has_opp_man:
            m_name = row['manual_name'] if has_app_man else '기타 비용'
            man_note = f"합산 {won(row['manual_add'])} 안분" if (has_app_man and app_count > 1) else ""
            stage_rows.append({
                '비목': m_name,
                'app_val': won(row['manual_add']) if has_app_man else '-',
                'app_vals': ([won(v) for v in row['per_manuals']]) if has_app_man else (['-'] * app_count),
                'opp_val': won(opp_row['manual']) if has_opp_man else '-',
                '비고': man_note,
                'row_type': 'item'
            })

        bimok_count = len(stage_rows)

        # 5. 소계 (합산 계산식 표시)
        if app_count == 1:
            sum_parts = [won(row['calc_fee'])]
            if row['stamp'] > 0:
                sum_parts.append(won(row['stamp']))
            if row['delivery'] > 0:
                sum_parts.append(won(row['delivery']))
            if row['manual_add'] > 0:
                sum_parts.append(won(row['manual_add']))
            subtotal_note = " + ".join(sum_parts) if len(sum_parts) > 1 else ""
        else:
            subtotal_note = f"신청인 합산 소계: {won(row['stage_sum'])}"

        stage_rows.append({
            '비목': '소계',
            'app_val': won(row['stage_sum']),
            'app_vals': [won(v) for v in row['per_stage_sums']],
            'opp_val': won(opp_row['opp_sum']) if (show_opp and opp_row) else '-',
            '비고': subtotal_note,
            'row_type': 'subtotal',
            'is_span_row': True
        })

        if not is_all_won or show_opp:
            app_ratio_strs = [format_frac(f) for f in row['fractions']]
            if opp_row:
                opp_f_list = opp_row.get('app_fractions', [opp_row.get('app_fraction', Fraction(0, 1))])
                opp_f_unique = list(dict.fromkeys([format_frac(f) for f in opp_f_list]))
                opp_ratio_str = ", ".join(opp_f_unique)
            else:
                opp_ratio_str = '-'

            stage_rows.append({
                '비목': '상대방부담비율',
                'app_val': app_ratio_strs[0],
                'app_vals': app_ratio_strs,
                'opp_val': opp_ratio_str,
                '비고': '당사자별 부담비율 적용' if app_count > 1 else '',
                'row_type': 'subtotal',
                'is_span_row': True
            })

            if app_count == 1:
                borne_note = f"{won(row['stage_sum'])} × {app_ratio_strs[0]} = {won(row['stage_borne'])} (1원 미만 버림)"
            else:
                b_lines = []
                for idx_a, lbl in enumerate(model['applicant_labels']):
                    b_lines.append(f"[{lbl}] {won(row['per_stage_sums'][idx_a])} × {app_ratio_strs[idx_a]} = {won(row['per_stage_bornes'][idx_a])}")
                borne_note = "\n".join(b_lines)

            stage_rows.append({
                '비목': '상대방부담비용',
                'app_val': won(row['stage_borne']),
                'app_vals': [won(v) for v in row['per_stage_bornes']],
                'opp_val': won(opp_row['opp_borne']) if (show_opp and opp_row) else '-',
                '비고': borne_note,
                'row_type': 'subtotal',
                'is_span_row': True
            })

        for i, sr in enumerate(stage_rows):
            is_span = sr.get('is_span_row', False)
            table_rows.append({
                '심급': stage_str if i == 0 else (sr['비목'] if is_span else ''),
                '비목': '' if is_span else sr['비목'],
                'app_val': sr['app_val'],
                'app_vals': sr['app_vals'],
                'opp_val': sr['opp_val'],
                '비고': sr['비고'],
                'merge_len': bimok_count if i == 0 else 0,
                'is_span_row': is_span,
                'row_type': sr.get('row_type', 'item')
            })

    # 본안비용소계 행 (신청비용이 포함될 때만 중간 소계로 표시)
    if app_count == 1:
        app_stage_expr = " + ".join([f"{r['stage']}심 {won(r['stage_borne'])}" for r in model['rows']])
    else:
        app_stage_expr = f"신청인 {app_count}명 합산 본안비용 {won(model['borne_main_total'])}"
    opp_stage_expr = " + ".join([f"{r['stage']}심 {won(r['opp_borne'])}" for r in model['opp_rows']])
    prefix_app = "[신청인] " if show_opp else ""
    note_sub = prefix_app + app_stage_expr
    if show_opp:
        note_sub += "\n[피신청인] " + opp_stage_expr

    if not exclude_app_cost:
        table_rows.append({
            '심급': '본안비용소계',
            '비목': '',
            'app_val': won(model['borne_main_total']),
            'app_vals': [won(v) for v in model['per_borne_main_totals']],
            'opp_val': won(model['opp_total']) if show_opp else '-',
            '비고': note_sub,
            'merge_len': 0,
            'is_span_row': True,
            'row_type': 'main_subtotal'
        })

        # 소송비용확정신청 행
        stamp_app_note = f"총 {won(model['app_stamp'])} × 1/{app_count} 안분" if app_count > 1 else ""
        del_app_note = f"당사자 총 {app_count + model['respondent_count']}명 기준 총 {won(model['app_delivery'])} × 1/{app_count} 안분" if app_count > 1 else f"피신청인 {model['respondent_count']}명 기준 (잔액은 확정 후 환급)"
        app_bimok_rows = [
            {
                '비목': '인지대',
                'app_val': won(model['app_stamp']),
                'app_vals': [won(v) for v in model['per_app_stamps']],
                'opp_val': '-',
                '비고': stamp_app_note
            },
            {
                '비목': '송달료',
                'app_val': won(model['app_delivery']),
                'app_vals': [won(v) for v in model['per_app_deliveries']],
                'opp_val': '-',
                '비고': del_app_note
            },
        ]
        for i, ar in enumerate(app_bimok_rows):
            table_rows.append({
                '심급': '소송비용\n확정신청' if i == 0 else '',
                '비목': ar['비목'],
                'app_val': ar['app_val'],
                'app_vals': ar['app_vals'],
                'opp_val': ar['opp_val'],
                '비고': ar['비고'],
                'merge_len': len(app_bimok_rows) if i == 0 else 0,
                'is_span_row': False,
                'row_type': 'item'
            })

        # 신청비용소계
        table_rows.append({
            '심급': '신청비용소계',
            '비목': '',
            'app_val': won(model['application_total']),
            'app_vals': [won(v) for v in model['per_application_totals']],
            'opp_val': '-',
            '비고': f"{won(model['per_app_stamps'][0])} + {won(model['per_app_deliveries'][0])}" + (" (1인당)" if app_count > 1 else ""),
            'merge_len': 0,
            'is_span_row': True,
            'row_type': 'subtotal'
        })

    # 최종 합계 행
    if app_count == 1:
        if show_opp:
            if exclude_app_cost:
                total_note = f"[신청인] 본안비용 합계 ({app_stage_expr})"
            else:
                total_note = f"[신청인] 본안비용 {won(model['borne_main_total'])} + 신청비용 {won(model['application_total'])}"
            total_note += f"\n[피신청인] 본안비용 {won(model['opp_total'])}"
            if model['offset_final'] >= 0:
                total_note += f"\n▶ 대등액 상계 후 최종 상환액: {won(model['offset_final'])}"
            else:
                total_note += f"\n▶ 대등액 상계 결과 신청인 초과부담: {won(abs(model['offset_final']))}"
        else:
            if exclude_app_cost:
                total_note = f"본안비용 합계 ({app_stage_expr})"
            else:
                total_note = f"본안비용 {won(model['borne_main_total'])} + 신청비용 {won(model['application_total'])}"
    else:
        if show_opp:
            note_parts = []
            for i_a, lbl in enumerate(model['applicant_labels']):
                note_parts.append(
                    f"[{lbl}] 청구액 {won(model['per_totals'][i_a])} − 피신청인 상계액 {won(model['per_opp_totals'][i_a])} = 잔액 {won(model['per_offset_finals'][i_a])}"
                )
            total_note = "\n".join(note_parts)
        else:
            total_note = f"신청인 {app_count}명 전체 합산 총액: {won(model['total'])}"

    table_rows.append({
        '심급': '합계',
        '비목': '',
        'app_val': won(model['total']),
        'app_vals': [won(v) for v in model['per_totals']],
        'opp_val': won(model['opp_total']) if show_opp else '-',
        '비고': total_note,
        'merge_len': 0,
        'is_span_row': True,
        'row_type': 'grand_total'
    })

    return table_rows


def render_html_table(cases, model, options):
    """법원 결정문 스타일의 HTML 표 생성"""
    table_data = generate_table_data(cases, model, options)
    show_opp = model.get('apply_offset_in_doc', False)
    app_count = model.get('applicant_count', 1)
    app_labels = model.get('applicant_labels', ['신청인'])

    refs = [f"{c['court']} {c['number']}".strip() for c in cases if (c['court'] or c['number'])]
    ref_paren = f" ({', '.join(refs)})" if refs else ""

    soga_parts = [f"제{r['stage']}심 소가 : {won(r['soga'])}" for r in model['rows']]
    soga_line = ", ".join(soga_parts)

    burden_lines = []
    for r in model['rows']:
        stage_label = {1: "소송비용", 2: "항소비용", 3: "상고비용"}.get(r['stage'], f"{r['stage']}심 비용")
        if app_count == 1:
            if r['fraction'] == Fraction(1, 1):
                burden_lines.append(f"▶ {stage_label}은 피신청인({'원고' if options['role']=='피고' else '피고'})이 부담")
            else:
                burden_lines.append(f"▶ {stage_label} 중 {format_frac(r['fraction'])}은 피신청인이 부담")
        else:
            b_parts = [f"{app_labels[i]} 부분 중 {format_frac(r['fractions'][i])}" for i in range(app_count)]
            burden_lines.append(f"▶ {stage_label} : {', '.join(b_parts)}은 피신청인이 부담")
    burden_html = "<br>".join(burden_lines)

    num_val_cols = app_count + (1 if show_opp else 0)
    if num_val_cols == 1:
        colgroup_html = (
            '<colgroup>'
            '<col style="width: 12%;">'
            '<col style="width: 16%;">'
            '<col style="width: 18%;">'
            '<col style="width: 54%;">'
            '</colgroup>'
        )
        thead_html = (
            '<thead>'
            '<tr><th>심급</th><th>비목</th><th>비용액</th><th>비고</th></tr>'
            '</thead>'
        )
    else:
        stage_w = 10
        bimok_w = 13
        val_w = max(10, min(15, int(42 / num_val_cols)))
        note_w = 100 - stage_w - bimok_w - (val_w * num_val_cols)
        cols = [f'<col style="width: {stage_w}%;">', f'<col style="width: {bimok_w}%;">']
        for _ in range(num_val_cols):
            cols.append(f'<col style="width: {val_w}%;">')
        cols.append(f'<col style="width: {note_w}%;">')
        colgroup_html = f"<colgroup>{''.join(cols)}</colgroup>"

        sub_ths = [f"<th>{html.escape(lbl)}</th>" for lbl in app_labels]
        if show_opp:
            sub_ths.append("<th>피신청인</th>")
        thead_html = (
            '<thead>'
            f'<tr><th rowspan="2">심급</th><th rowspan="2">비목</th><th colspan="{num_val_cols}">비용액</th><th rowspan="2">비고</th></tr>'
            f'<tr>{"".join(sub_ths)}</tr>'
            '</thead>'
        )

    tbody_rows = []
    for row in table_data:
        r_type = row.get('row_type', 'item')
        tr_class = {
            'item': '',
            'subtotal': ' class="row-subtotal"',
            'main_subtotal': ' class="row-main-subtotal"',
            'grand_total': ' class="row-grand-total"'
        }.get(r_type, '')

        tds = []
        if row.get('is_span_row'):
            label_html = html.escape(row['심급']).replace('\n', '<br>')
            tds.append(f'<td colspan="2" class="cell-center">{label_html}</td>')
        else:
            if row.get('merge_len', 0) > 0:
                stage_html = html.escape(row['심급']).replace('\n', '<br>')
                rowspan_attr = f' rowspan="{row["merge_len"]}"' if row['merge_len'] > 1 else ''
                tds.append(f'<td{rowspan_attr} class="cell-stage">{stage_html}</td>')
            bimok_html = html.escape(row['비목']).replace('\n', '<br>')
            tds.append(f'<td class="cell-center">{bimok_html}</td>')

        for app_v in row['app_vals']:
            app_v_str = str(app_v)
            app_cls = "cell-right" if app_v_str.endswith("원") else "cell-center"
            tds.append(f'<td class="{app_cls}">{html.escape(app_v_str)}</td>')

        if show_opp:
            opp_v = str(row['opp_val'])
            opp_cls = "cell-right" if opp_v.endswith("원") else "cell-center"
            tds.append(f'<td class="{opp_cls}">{html.escape(opp_v)}</td>')

        note_html = html.escape(row['비고']).replace('\n', '<br>')
        tds.append(f'<td class="cell-left">{note_html}</td>')

        tbody_rows.append(f'<tr{tr_class}>{"".join(tds)}</tr>')

    if app_count == 1:
        claim_amount = model['offset_final'] if show_opp and model['offset_final'] > 0 else model['total']
        reimburse_text = f"2. 피신청인이 신청인에게 상환하여야 할 소송비용액 금 {won(claim_amount)}"
        applicant_meta = f"<strong>신청인 :</strong> {html.escape(options['role'])}{html.escape(ref_paren)}"
    else:
        claims = model['per_offset_finals'] if show_opp else model['per_totals']
        each_str = ", ".join([f"{html.escape(lbl)}에게 금 {won(claims[i])}" for i, lbl in enumerate(app_labels)])
        reimburse_text = f"2. 피신청인이 상환하여야 할 소송비용액 : {each_str} (합계 금 {won(sum(claims))})"
        app_header_str = ", ".join([f"신청인({options['role']}){i+1} : {html.escape(lbl.replace(f'신청인{i+1}', '').strip() or f'신청인{i+1}')}" for i, lbl in enumerate(app_labels)])
        applicant_meta = f"<strong>{app_header_str}</strong>{html.escape(ref_paren)}"

    full_html = (
        '<div id="pdf-capture-target" class="court-doc-card">'
        '<div class="court-doc-title">소 송 비 용 계 산 서</div>'
        '<div class="court-meta-box">'
        f'<div>{applicant_meta}</div>'
        f'<div><strong>{html.escape(soga_line)}</strong></div>'
        f'<div class="court-burden-list">{burden_html}</div>'
        '</div>'
        '<div class="court-table-topbar">'
        '<span class="sec-label">1. 총 소송비용</span>'
        '<span class="unit-label">(단위 : 원, 1원 미만 버림)</span>'
        '</div>'
        f'<table class="court-cost-table">{colgroup_html}{thead_html}<tbody>{"".join(tbody_rows)}</tbody></table>'
        '<div class="court-reimburse-box">'
        f'{reimburse_text}'
        '</div>'
        '</div>'
    )
    return full_html


def render_text_calculation(cases, model, options):
    """줄글 양식에도 표 양식과 동일하게 상세 계산식이 출력되도록 개선"""
    show_opp = model.get('apply_offset_in_doc', False)
    exclude_app_cost = model.get('exclude_app_cost', False)
    app_count = model.get('applicant_count', 1)
    app_labels = model.get('applicant_labels', ['신청인'])
    before = ['1. 신청인의 지출비용 및 피신청인 부담액 (단위 : 원, 1원 미만 버림)']

    for index, row in enumerate(model['rows']):
        prefix = '가나다라마바사'[index]
        num_str = f" ({row['court']} {row['number']})".strip() if (row['court'] or row['number']) else ""
        stage_title = f"{prefix}. {row['stage']}심{num_str}"
        before.append(f"{stage_title} [소가: 금 {won(row['soga'])}]")

        formula_str = fee_formula(row['soga'], options.get('is_reduced', False))
        if row['actual'] is not None and row['actual'] < row['limit']:
            fee_calc_note = f"규칙 제3조 산정: {formula_str} = {won(row['limit'])}이나, 실제 지급한 변호사보수를 산입함"
        elif row['actual'] is not None:
            fee_calc_note = f"변호사보수의 소송비용산입에 관한 규칙 제3조 산정: {formula_str} = {won(row['limit'])} (실제 지급액: 금 {won(row['actual'])})"
        else:
            fee_calc_note = f"변호사보수의 소송비용산입에 관한 규칙 제3조 산정: {formula_str} = {won(row['limit'])}"

        if app_count == 1:
            before.append(f"   - 변호사보수: 금 {won(row['calc_fee'])}")
            before.append(f"     ({fee_calc_note})")

            sum_parts = [won(row['calc_fee'])]
            if row['stamp'] > 0:
                before.append(f"   - 인지대: 금 {won(row['stamp'])}")
                sum_parts.append(won(row['stamp']))
            if row['delivery'] > 0:
                before.append(f"   - 송달료: 금 {won(row['delivery'])} (실제 지출비용 산입)")
                sum_parts.append(won(row['delivery']))
            if row['manual_add'] > 0:
                before.append(f"   - {row['manual_name']}: 금 {won(row['manual_add'])}")
                sum_parts.append(won(row['manual_add']))

            sum_formula_str = f" ({' + '.join(sum_parts)})" if len(sum_parts) > 1 else ""
            before.append(f"   => {row['stage']}심 지출 소계: 금 {won(row['stage_sum'])}{sum_formula_str}")

            frac = row['fraction']
            ratio_str = format_frac(frac)
            if frac == Fraction(1, 1):
                before.append(f"   => 피신청인 부담비용 (부담비율 {ratio_str}): {won(row['stage_sum'])} × 100% = 금 {won(row['stage_borne'])}")
            else:
                before.append(f"   => 피신청인 부담비용 (부담비율 {ratio_str}): {won(row['stage_sum'])} × {frac.numerator}/{frac.denominator} = 금 {won(row['stage_borne'])} (1원 미만 버림)")
            before.append("")
        else:
            before.append(f"   - 합산 변호사보수 인정액: 금 {won(row['calc_fee'])} ({fee_calc_note})")
            for i, lbl in enumerate(app_labels):
                f_i = row['fractions'][i]
                before.append(
                    f"   - [{lbl}] 변호사보수 {won(row['per_calc_fees'][i])}, 인지대 {won(row['per_stamps'][i])}, 송달료 {won(row['per_deliveries'][i])}"
                    + (f", {row['manual_name']} {won(row['per_manuals'][i])}" if row['per_manuals'][i] > 0 else "")
                    + f" → 소계 {won(row['per_stage_sums'][i])} × 부담비율({f_i.numerator}/{f_i.denominator}) = 금 {won(row['per_stage_bornes'][i])}"
                )
            before.append("")

    after = ['2. 소송비용액 산정']
    if app_count == 1:
        stage_sum_expr = " + ".join([f"{r['stage']}심 {won(r['stage_borne'])}" for r in model['rows']])
        if exclude_app_cost:
            after.append(f"가. 본안 소송비용 중 피신청인 부담액 합계: 금 {won(model['borne_main_total'])} ({stage_sum_expr})")
            if not show_opp:
                after.append("나. 피신청인이 신청인에게 상환해야 할 소송비용액은,")
                after.append(f"    금 {won(model['total'])}   끝.")
        else:
            after.append(f"가. 본안 소송비용 중 피신청인 부담액 합계: 금 {won(model['borne_main_total'])} ({stage_sum_expr})")
            after.append(
                f"나. 소송비용액확정신청 비용 합계: 금 {won(model['application_total'])} "
                f"[인지대 {won(model['app_stamp'])} + 송달료 {won(model['app_delivery'])} (피신청인 {model['respondent_count']}명 기준)]"
            )
            after.append(
                f"다. 본안 및 신청비용 합계: 본안비용 {won(model['borne_main_total'])} + 신청비용 {won(model['application_total'])} = 금 {won(model['total'])}"
            )
            if not show_opp:
                after.append("라. 피신청인이 신청인에게 상환해야 할 소송비용액은,")
                after.append(f"    금 {won(model['total'])}   끝.")

        if show_opp:
            after.append("")
            after.append("3. 상대방 소송비용과의 대등액 상계 (민사소송법 제114조)")
            opp_stage_expr = " + ".join([f"{r['stage']}심 {won(r['opp_borne'])}" for r in model['opp_rows']])
            after.append(f"가. 상대방 본안 소송비용 중 신청인 부담액 합계: 금 {won(model['opp_total'])} ({opp_stage_expr})")
            if model['offset_final'] >= 0:
                after.append(f"나. 대등액 상계 후 피신청인이 상환해야 할 최종 잔액: 금 {won(model['total'])} − 금 {won(model['opp_total'])} = 금 {won(model['offset_final'])}   끝.")
            else:
                after.append(f"나. 대등액 상계 결과 (신청인의 초과 부담액 발생): 금 {won(model['total'])} − 금 {won(model['opp_total'])} = -금 {won(abs(model['offset_final']))}   끝.")
    else:
        claims = model['per_offset_finals'] if show_opp else model['per_totals']
        for i, lbl in enumerate(app_labels):
            if show_opp:
                after.append(f"   - [{lbl}] 본안 {won(model['per_borne_main_totals'][i])} + 신청비용 {won(model['per_application_totals'][i])} − 상대방 상계액 {won(model['per_opp_totals'][i])} = 금 {won(claims[i])}")
            else:
                after.append(f"   - [{lbl}] 본안 {won(model['per_borne_main_totals'][i])} + 신청비용 {won(model['per_application_totals'][i])} = 금 {won(claims[i])}")
        after.append(f"   => 피신청인이 상환해야 할 총 합계: 금 {won(sum(claims))}   끝.")

    return before, after


def render_pdf_download_button(html_card_content):
    """별지 계산서 HTML 그림 그대로 PDF로 다운로드하는 버튼 컴포넌트"""
    component_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{
            margin: 0;
            padding: 0;
            font-family: 'Pretendard', 'Malgun Gothic', '맑은 고딕', sans-serif;
            background: transparent;
        }}
        .pdf-btn {{
            width: 100%;
            background-color: #0f172a;
            color: #ffffff;
            border: 1px solid #0f172a;
            font-weight: 700;
            font-size: 15px;
            padding: 0.68rem 1rem;
            border-radius: 6px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            box-sizing: border-box;
            transition: background-color 0.15s ease;
        }}
        .pdf-btn:hover {{
            background-color: #334155;
        }}
        .pdf-btn:disabled {{
            opacity: 0.65;
            cursor: wait;
        }}
        #offscreen-wrapper {{
            position: fixed;
            left: -9999px;
            top: 0;
            width: 860px;
            background: #ffffff;
            padding: 10px;
        }}
        {COURT_CARD_CSS}
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    </head>
    <body>
        <button id="download-pdf-btn" class="pdf-btn" onclick="downloadCardAsPDF()">
            📄 별지 소송비용계산서 (그림 그대로) 다운로드 (.pdf)
        </button>
        <div id="offscreen-wrapper">
            {html_card_content}
        </div>
        <script>
        function downloadCardAsPDF() {{
            const btn = document.getElementById('download-pdf-btn');
            const element = document.getElementById('pdf-capture-target');
            const origText = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = '⏳ PDF 생성 중... 잠시만 기다려주세요';

            const opt = {{
                margin: [10, 10, 10, 10],
                filename: '별지_소송비용계산서.pdf',
                image: {{ type: 'jpeg', quality: 0.98 }},
                html2canvas: {{ scale: 2, useCORS: true, backgroundColor: '#ffffff' }},
                jsPDF: {{ unit: 'mm', format: 'a4', orientation: 'portrait' }}
            }};

            html2pdf().set(opt).from(element).save().then(() => {{
                btn.disabled = false;
                btn.innerHTML = origText;
            }}).catch((err) => {{
                console.error(err);
                btn.disabled = false;
                btn.innerHTML = origText;
            }});
        }}
        </script>
    </body>
    </html>
    """
    components.html(component_html, height=54)


def make_docx(cases, model, options, prayer, cause):
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL

    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.0)
    section.top_margin = section.bottom_margin = Cm(3.0)

    for style in doc.styles:
        for border in list(style.element.iter(qn('w:pBdr'))):
            border.getparent().remove(border)

    for name, font, size in [('Normal', '나눔명조', 11.5), ('Title', '나눔고딕', 13)]:
        style = doc.styles[name]
        style.font.name = font
        style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.size = Pt(size)
        style.paragraph_format.line_spacing = 1.25
        style.paragraph_format.space_after = Pt(6)
    doc.styles['Title'].font.bold = True

    def heading(text):
        p = doc.add_paragraph(text, 'Title')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True

    def paragraphs(lines):
        for line in lines:
            if line.strip():
                p = doc.add_paragraph(line)
                for run in p.runs:
                    run.font.name = '나눔명조'
                    run._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '나눔명조')

    heading('신 청 취 지')
    paragraphs(prayer.split('\n'))
    doc.add_paragraph()
    heading('신 청 원 인')
    paragraphs(cause.split('\n'))
    doc.add_page_break()

    # 별지 계산서 상단 메타정보
    doc.add_paragraph('별지')
    heading('소 송 비 용 계 산 서')

    refs_doc = [f"{c['court']} {c['number']}".strip() for c in cases if (c['court'] or c['number'])]
    ref_paren_doc = f" ({', '.join(refs_doc)})" if refs_doc else ""
    soga_parts_doc = [f"제{r['stage']}심 소가 : {won(r['soga'])}" for r in model['rows']]
    soga_line_doc = ", ".join(soga_parts_doc)

    app_count = model.get('applicant_count', 1)
    app_labels = model.get('applicant_labels', ['신청인'])
    if app_count == 1:
        doc.add_paragraph(f"신청인 : {options['role']}{ref_paren_doc}")
    else:
        app_header_str = ", ".join([f"신청인({options['role']}){i+1} : {lbl.replace(f'신청인{i+1}', '').strip() or f'신청인{i+1}'}" for i, lbl in enumerate(app_labels)])
        doc.add_paragraph(f"{app_header_str}{ref_paren_doc}")

    doc.add_paragraph(soga_line_doc)
    doc.add_paragraph("(단위 : 원, 1원 미만 버림)")
    doc.add_paragraph()

    use_table_format = options.get('output_format') == '표 양식 (법원 서식)'
    show_opp = model.get('apply_offset_in_doc', False)

    if use_table_format:
        table_data = generate_table_data(cases, model, options)
        num_val_cols = app_count + (1 if show_opp else 0)
        num_cols = 3 + num_val_cols
        hdr_rows_count = 2 if num_val_cols > 1 else 1

        table = doc.add_table(rows=len(table_data) + hdr_rows_count, cols=num_cols)
        table.style = 'Table Grid'
        table.autofit = False

        tblPr = table._tbl.tblPr
        cell_mar = parse_xml(
            f'<w:tblCellMar {nsdecls("w")}>'
            f'<w:top w:w="110" w:type="dxa"/>'
            f'<w:bottom w:w="110" w:type="dxa"/>'
            f'<w:left w:w="120" w:type="dxa"/>'
            f'<w:right w:w="120" w:type="dxa"/>'
            f'</w:tblCellMar>'
        )
        tblPr.append(cell_mar)

        total_cm = 17.0
        if num_val_cols == 1:
            col_widths = [Cm(2.2), Cm(2.8), Cm(3.0), Cm(9.0)]
            headers = ['심급', '비목', '비용액', '비고']
            for i, text in enumerate(headers):
                cell = table.rows[0].cells[i]
                cell.text = text
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>')
                cell._tc.get_or_add_tcPr().append(shd)
        else:
            stage_cm = 1.9
            bimok_cm = 2.4
            val_cm = max(2.1, min(2.7, 7.0 / num_val_cols))
            note_cm = total_cm - stage_cm - bimok_cm - (val_cm * num_val_cols)
            col_widths = [Cm(stage_cm), Cm(bimok_cm)] + [Cm(val_cm)] * num_val_cols + [Cm(note_cm)]

            r0 = table.rows[0].cells
            r1 = table.rows[1].cells
            r0[0].text = '심급'
            r0[1].text = '비목'
            r0[2].text = '비용액'
            r0[num_cols - 1].text = '비고'

            for idx_lbl, lbl in enumerate(app_labels):
                r1[2 + idx_lbl].text = lbl
            if show_opp:
                r1[2 + app_count].text = '피신청인'

            r0[0].merge(r1[0])
            r0[1].merge(r1[1])
            r0[2].merge(r0[2 + num_val_cols - 1])
            r0[num_cols - 1].merge(r1[num_cols - 1])

            for h_idx in range(hdr_rows_count):
                for c in table.rows[h_idx].cells:
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>')
                    c._tc.get_or_add_tcPr().append(shd)

        for r_idx, row in enumerate(table_data):
            t_row_idx = r_idx + hdr_rows_count
            cells = table.rows[t_row_idx].cells
            cells[0].text = row['심급']
            cells[1].text = row['비목']
            for a_i, a_val in enumerate(row['app_vals']):
                cells[2 + a_i].text = str(a_val)
            if show_opp:
                cells[2 + app_count].text = str(row['opp_val'])
            cells[num_cols - 1].text = row['비고']

        for r_idx, row in enumerate(table_data):
            t_row_idx = r_idx + hdr_rows_count
            m_len = row.get('merge_len', 0)
            if m_len > 1:
                start_cell = table.cell(t_row_idx, 0)
                end_cell = table.cell(t_row_idx + m_len - 1, 0)
                merged = start_cell.merge(end_cell)
                while len(merged.paragraphs) > 1:
                    p = merged.paragraphs[-1]._p
                    p.getparent().remove(p)

            if row.get('is_span_row'):
                start_cell = table.cell(t_row_idx, 0)
                end_cell = table.cell(t_row_idx, 1)
                merged = start_cell.merge(end_cell)
                while len(merged.paragraphs) > 1:
                    p = merged.paragraphs[-1]._p
                    p.getparent().remove(p)
                merged.paragraphs[0].text = row['심급']

                r_type = row.get('row_type', 'subtotal')
                fill_hex = "EFF6FF" if r_type == 'grand_total' else ("F1F5F9" if r_type == 'main_subtotal' else "F8FAFC")
                for c_i in range(num_cols):
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
                    table.cell(t_row_idx, c_i)._tc.get_or_add_tcPr().append(shd)

        amt_cols = list(range(2, 2 + num_val_cols))
        note_col = num_cols - 1

        for r_idx, row_obj in enumerate(table.rows):
            is_header = r_idx < hdr_rows_count
            data_row = table_data[r_idx - hdr_rows_count] if not is_header else None
            is_bold_row = is_header or (data_row and data_row.get('is_span_row'))

            for c_idx, cell in enumerate(row_obj.cells):
                cell.width = col_widths[c_idx]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15

                    for run in p.runs:
                        run.font.name = '나눔명조'
                        run._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '나눔명조')
                        run.font.size = Pt(9.0 if num_val_cols > 2 else 9.5)
                        if is_bold_row or c_idx == 0:
                            run.font.bold = True

                    if is_header:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        if c_idx in amt_cols:
                            val_text = p.text.strip()
                            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if val_text.endswith('원') else WD_ALIGN_PARAGRAPH.CENTER
                        elif c_idx == note_col:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        before, after = render_text_calculation(cases, model, options)
        paragraphs(before)
        doc.add_paragraph()
        paragraphs(after)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def main():
    st.set_page_config(page_title='소송비용 신청서 만들기', page_icon='⚖️', layout='wide')

    st.markdown(f"""
    <style>
    .block-container {{
        max-width: 1040px !important;
        padding-top: 2rem !important;
    }}
    [data-testid="stDownloadButton"] button {{
        background-color: #1e3a8a;
        color: #ffffff;
        border: 1px solid #1e3a8a;
        font-weight: 700;
        padding: 0.65rem 1rem;
        border-radius: 6px;
    }}
    [data-testid="stDownloadButton"] button:hover {{
        background-color: #172554;
        color: #ffffff;
    }}
    {COURT_CARD_CSS}
    </style>
    """, unsafe_allow_html=True)

    st.title('소송비용 신청서 만들기')
    st.caption('소송비용액확정 신청서 및 계산서 간편 생성기 (공동신청인 N명 분리·개별 부담비율 지원)')

    case_column, name_column = st.columns([3, 1])
    with case_column:
        raw = st.text_area('1·2·3심 법원·사건번호',
                           placeholder='서울중앙지방법원 2023가합12345, 서울고등법원 2025나12345', height=100)
    with name_column:
        case_name = st.text_input('사건명', placeholder='약정금')
    st.caption('진행한 심급만 쉼표나 줄바꿈으로 구분하세요. 행정, 가사 사건도 입력 가능합니다.')

    error, cases = None, []
    try:
        cases = parse_cases(raw)
    except ValueError as exc:
        error = str(exc)
        cases = [{"stage": 1, "court": "", "number": "", "code": "가단"}]

    role_col, app_cnt_col, app_name_col, weight_col = st.columns([1.1, 0.9, 1.8, 1.5])
    with role_col:
        role = st.radio('신청인은 본안에서', ['피고', '원고'], horizontal=True)
    with app_cnt_col:
        applicant_count = int(st.number_input('공동신청인 수(N명)', min_value=1, max_value=10, value=1, step=1,
                                              help="2명 이상 설정 시 전체 소가 기준 단일 법정 한도액을 안분하여 신청인1, 신청인2... 열로 각각 나누어 표시합니다."))
    with app_name_col:
        applicant_names = st.text_input('신청인 이름 (쉼표 구분, 선택)',
                                        placeholder='예: 이기옥, 최강둘',
                                        disabled=(applicant_count == 1))
    with weight_col:
        weight_str = st.text_input('비용 안분 비율 (균등=1:1 / 소가비율=920:520)',
                                   value='1:1',
                                   disabled=(applicant_count == 1),
                                   help="법정 한도액 및 공통 지출액을 신청인들 사이에 나눌 비율입니다.")

    try:
        app_weights = parse_weights(weight_str, applicant_count)
    except ValueError as exc:
        error = str(exc)
        app_weights = [Fraction(1, applicant_count)] * applicant_count

    app_labels = get_applicant_labels({'applicant_count': applicant_count, 'applicant_names': applicant_names})

    method_column, people_column = st.columns(2)
    with method_column:
        electronic = st.radio('제출 방식', ['전자소송', '종이소송'], horizontal=True) == '전자소송'
    with people_column:
        party_count = int(st.number_input('송달료 산정 기준 인원수', min_value=1,
                                          value=1 if electronic else 2, step=1))

    st.markdown("---")
    st.subheader("소송비용액확정 신청 비용 설정")
    resp_col1, resp_col2 = st.columns([1, 2])
    with resp_col1:
        respondent_count = int(st.number_input('상대방(피신청인) 수', min_value=1, value=1, step=1))
    with resp_col2:
        auto_app_stamp = 900 if electronic else 1_000
        total_parties = respondent_count + applicant_count
        auto_rounds = total_parties * 3
        auto_app_delivery = auto_rounds * DELIVERY_UNIT
        per_del_info = f" (1인당 안분: 인지 {won(auto_app_stamp // applicant_count)} / 송달료 {won(auto_app_delivery // applicant_count)})" if applicant_count > 1 else ""
        st.info(f"📌 **신청 비용 자동 계산 내역**\n\n- 인지대: **{won(auto_app_stamp)}** (전자 기준)\n- 송달료: **{won(auto_app_delivery)}**{per_del_info}")

    is_reduced = st.checkbox('무변론 판결, 자백간주, 이행권고결정에 따른 변호사보수 1/2 감액 일괄 적용 (규칙 제5조)')

    st.markdown("---")
    st.subheader("심급별 세부 비용 및 부담비율 설정")
    soga_dict = {}
    actual_fees = {}
    filing_costs = {}
    stage_settings = {}

    for case in cases:
        c_num = case['number']
        c_key = c_num if c_num else f"stage_{case['stage']}"
        stage = case['stage']
        court_info = f" ({case['court']} {c_num})".strip() if (case['court'] or c_num) else ""
        st.markdown(f"#### 🏛️ {stage}심{court_info}")

        soga_col, fee_col = st.columns(2)
        with soga_col:
            soga_val = int(st.number_input(f'{stage}심 소가(원)', min_value=0, value=50_000_000, step=1_000_000, key=f'soga_{c_key}'))
            soga_dict[c_key] = soga_val
            if soga_val > 0:
                lim_preview = fee_limit(soga_val)
                if is_reduced:
                    lim_preview = trunc_1(Decimal(lim_preview) / Decimal(2))
                st.caption(f"단일 법정 한도액(합산): **{won(lim_preview)}**")
        with fee_col:
            fee_val = st.text_input(f'{stage}심 실제 변호사보수(원)', key=f'fee_{c_key}',
                                    placeholder='생략 시 법정 한도액 적용 (다수 신청인인 경우 합산액)')
            try:
                actual_fees[c_key] = optional_money(fee_val)
            except ValueError as exc:
                error = str(exc)

            if applicant_count == 1:
                col1, col2, col3 = st.columns([1.2, 1.5, 1.5])
                with col1:
                    frac_str = st.text_input(f'{stage}심 피신청인 부담비율', value='1/1', key=f'frac_{c_key}', help="분수 입력")
                    try:
                        stage_fraction = parse_fraction(frac_str)
                    except ValueError as exc:
                        error = str(exc)
                        stage_fraction = Fraction(1, 1)
                with col2:
                    manual_name = st.text_input(f'{stage}심 수기 비용 항목명', value='기타 수기 비용', key=f'manual_name_{c_key}')
                with col3:
                    manual_add = int(st.number_input(f'{stage}심 기타 비용 금액(원)', min_value=0, value=0, step=10_000, key=f'manual_{c_key}', help="0원 입력시 계산서 적용 안됨"))
                stage_fractions = [stage_fraction]
            else:
                st.markdown(f"**신청인별 {stage}심 피신청인 부담비율** (예: `9/10`, `8/10`, `1/1`)")
                f_cols = st.columns(applicant_count)
                stage_fractions = []
                for idx_a in range(applicant_count):
                    with f_cols[idx_a]:
                        f_str = st.text_input(f'{app_labels[idx_a]} 부담비율', value='1/1', key=f'frac_{c_key}_{idx_a}')
                        try:
                            stage_fractions.append(parse_fraction(f_str))
                        except ValueError as exc:
                            error = str(exc)
                            stage_fractions.append(Fraction(1, 1))
                stage_fraction = stage_fractions[0]
                m_col1, m_col2 = st.columns(2)
                with m_col1:
                    manual_name = st.text_input(f'{stage}심 수기 비용 항목명', value='기타 수기 비용', key=f'manual_name_{c_key}')
                with m_col2:
                    manual_add = int(st.number_input(f'{stage}심 기타 비용 합산액(원)', min_value=0, value=0, step=10_000, key=f'manual_{c_key}', help="0원 입력시 계산서 적용 안됨"))

            stage_settings[c_key] = {
                'fraction': stage_fraction,
                'fractions': stage_fractions,
                'manual_name': manual_name,
                'manual_add': manual_add
            }

            include = st.checkbox(f'신청인이 납부한 {stage}심 인지대·송달료 포함',
                                  value=(role == '원고' and stage == 1),
                                  key=f'paid_{role}_{c_key}')
            stamp, delivery = (0, 0)
            if soga_val > 0:
                stamp, delivery = filing_estimate(soga_val, case, electronic, party_count)
            entry = dict(include=include, stamp=stamp, delivery=delivery)
            if include:
                st.caption(f"자동 계산: 인지대 {won(stamp)} / 송달료 {won(delivery)}")
                if st.checkbox(f'{stage}심 인지·송달료 실제 납부액으로 직접 수정', key=f'override_{c_key}'):
                    col_st, col_dl = st.columns(2)
                    with col_st:
                        entry['stamp'] = int(st.number_input(f'{stage}심 인지대 실부담액', min_value=0, value=stamp, key=f'stamp_{c_key}'))
                    with col_dl:
                        entry['delivery'] = int(st.number_input(f'{stage}심 송달료 실부담액', min_value=0, value=delivery, key=f'delivery_{c_key}'))
            filing_costs[c_key] = entry
            st.divider()

    st.caption('소송대리인: 법무법인(유한)바른')
    opp_settings = {}
    hide_opp_in_doc = False
    with st.expander('⚖️ 상대방 소송비용 상계(대등액 공제) 시뮬레이션 (선택)'):
        use_offset = st.checkbox('상대방 소송비용 상계 계산 활성화', value=False)
        if use_offset:
            hide_opp_in_doc = st.checkbox(
                '📄 신청서·계산서(표 및 줄글) 출력 시 피신청인(상대방) 내역 숨기기 (신청인 지출액만 단독 청구)',
                value=False,
                help="체크하면 상단에서 상계 시뮬레이션 금액은 확인하되, 실제 출력되는 신청서·표·줄글에서는 피신청인 칸과 상계 문구를 빼고 신청인 지출액 기준으로만 깔끔하게 생성합니다."
            )
            is_all_won = all(
                all(f == Fraction(1, 1) for f in stage_settings.get(c['number'] or f"stage_{c['stage']}", {}).get('fractions', [Fraction(1, 1)]))
                for c in cases
            )
            if is_all_won:
                st.warning("⚠️ **현재 피신청인 부담비율이 모두 '1/1(전부승소)'입니다.**\n\n신청인이 물어줘야 할 비용이 0%이므로 상계액이 계산되지 않습니다. 부담비율을 수정해주세요.")

            for case in cases:
                c_num = case['number']
                c_key = c_num if c_num else f"stage_{case['stage']}"
                stage = case['stage']
                soga_val = soga_dict.get(c_key, 0)
                court_info = f" ({case['court']} {c_num})".strip() if (case['court'] or c_num) else ""
                st.markdown(f"**🏛️ {stage}심 상대방 지출비용{court_info}**")

                col_of1, col_of2 = st.columns(2)
                with col_of1:
                    opp_fee_text = st.text_input(f'{stage}심 상대방 실제 변호사보수(원)', key=f'opp_fee_{c_key}', placeholder='생략 시 법정 한도액 적용')
                    try:
                        opp_fee_val = optional_money(opp_fee_text)
                    except ValueError as exc:
                        error = str(exc)
                        opp_fee_val = None
                with col_of2:
                    opp_manual_val = int(st.number_input(f'{stage}심 상대방 기타 비용(원)', min_value=0, value=0, step=10_000, key=f'opp_man_{c_key}', help="0원 입력시 계산서 적용 안됨"))

                opp_include = st.checkbox(f'상대방이 납부한 {stage}심 인지대·송달료 포함',
                                          value=(role == '피고' and stage == 1),
                                          key=f'opp_paid_{c_key}')
                o_stamp, o_delivery = 0, 0
                if soga_val > 0:
                    o_stamp, o_delivery = filing_estimate(soga_val, case, electronic, party_count)
                if opp_include:
                    st.caption(f"상대방 인지·송달료 추정치: 인지대 {won(o_stamp)} / 송달료 {won(o_delivery)}")
                    if st.checkbox(f'{stage}심 상대방 인지·송달료 직접 수정', key=f'opp_override_{c_key}'):
                        oc1, oc2 = st.columns(2)
                        with oc1:
                            o_stamp = int(st.number_input(f'{stage}심 상대방 인지대 실부담액', min_value=0, value=o_stamp, key=f'opp_st_{c_key}'))
                        with oc2:
                            o_delivery = int(st.number_input(f'{stage}심 상대방 송달료 실부담액', min_value=0, value=o_delivery, key=f'opp_dl_{c_key}'))

                opp_settings[c_key] = {
                    'fee': opp_fee_val,
                    'include': opp_include,
                    'stamp': o_stamp,
                    'delivery': o_delivery,
                    'manual': opp_manual_val
                }
                st.markdown("---")

    with st.expander('확정일 입력 (선택)'):
        final_date = st.text_input('확정일', placeholder='예: 2026. 7. 21.')
        st.caption('입력하지 않으면 확정일 문장은 문서에 넣지 않습니다.')

    # 확정일 바로 아래: 상대방 전달용 (확정신청 비용 제외) 옵션 추가
    exclude_app_cost = st.checkbox(
        '📨 상대방 전달용으로 만들기 (소송비용액확정 신청 인지대·송달료 제외)',
        value=False,
        help="법원에 확정신청서를 접수하기 전, 상대방에게 임의 변제 안내용으로 계산서를 전달할 때 체크하세요. 확정신청 인지대(900원)와 송달료가 제외된 본안 소송비용만 합산됩니다."
    )

    st.markdown("---")
    output_format = st.radio('📝 계산서 출력 양식', ['표 양식 (법원 서식)', '줄글 양식 (기본)'], horizontal=True)

    options = dict(electronic=electronic, party_count=party_count,
                   role=role, case_name=case_name, final_date=final_date,
                   is_reduced=is_reduced, respondent_count=respondent_count,
                   applicant_count=applicant_count, applicant_names=applicant_names,
                   app_weights=app_weights,
                   stage_settings=stage_settings, use_offset=use_offset,
                   hide_opp_in_doc=hide_opp_in_doc, opp_settings=opp_settings,
                   exclude_app_cost=exclude_app_cost, output_format=output_format)

    signature = repr((cases, raw, soga_dict, actual_fees, filing_costs, options))

    if st.button('신청서·계산서 만들기', type='primary', use_container_width=True):
        try:
            if error:
                raise ValueError(error)
            model = make_model(cases, soga_dict, actual_fees, filing_costs, options)
        except ValueError as exc:
            st.error(str(exc))
            st.session_state.pop('template_result', None)
        else:
            st.session_state['template_result'] = (signature, model)

    saved = st.session_state.get('template_result')
    if not saved or saved[0] != signature:
        return

    model = saved[1]
    st.success("✅ 계산서 작성이 완료되었습니다.")

    if model.get('applicant_count', 1) > 1:
        m_cols = st.columns(model['applicant_count'] + 1)
        claims = model['per_offset_finals'] if model.get('use_offset') else model['per_totals']
        for i, lbl in enumerate(model['applicant_labels']):
            with m_cols[i]:
                delta_s = f"상계 공제: -{won(model['per_opp_totals'][i])}" if model.get('use_offset') else None
                st.metric(f"{lbl} 청구액", won(claims[i]), delta=delta_s)
        with m_cols[model['applicant_count']]:
            st.metric("전체 합산 총액", won(sum(claims)))
    elif model.get('use_offset'):
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric('상계 전 신청인 단독 청구액', won(model['total']))
        with col_m2:
            delta_str = f"-{won(model['opp_total'])}" if model['opp_total'] > 0 else "0원"
            label_suffix = " (참고용·문서 미반영)" if model.get('hide_opp_in_doc') else ""
            st.metric(f'상계 후 예상 상환액{label_suffix}', won(model['offset_final']), delta=f"상대방 채권 공제: {delta_str}")
    else:
        metric_label = '최종 상환액 합계 (신청비용 제외·상대방 전달용)' if model.get('exclude_app_cost') else '최종 상환액 합계'
        st.metric(metric_label, won(model['total']))

    prayer, cause = application_sections(cases, model, options)
    prayer = st.text_area('신청취지 편집', value=prayer, height=140)
    cause = st.text_area('신청원인 편집', value=cause, height=260)

    st.subheader('별지 소송비용액계산서')

    html_table = render_html_table(cases, model, options)

    if output_format == '표 양식 (법원 서식)':
        st.markdown(html_table, unsafe_allow_html=True)
    else:
        before, after = render_text_calculation(cases, model, options)
        for line in before:
            st.write(line)
        for line in after:
            st.write(line)

    # 1. 별지 계산서 그림 그대로 PDF 다운로드 버튼 (.docx 다운로드 바로 위)
    render_pdf_download_button(html_table)

    # 2. 기존 Word (.docx) 다운로드 버튼
    args = (cases, model, options, prayer, cause)
    try:
        document = make_docx(*args)
    except ImportError:
        st.info('Word 저장: 터미널에서 python -m pip install python-docx를 실행하세요.')
    else:
        st.download_button('📥 신청서·계산서 초안 다운로드 (.docx)', document,
                           file_name='소송비용액확정_신청서_계산서.docx',
                           mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                           use_container_width=True)


if __name__ == '__main__':
    main()
