"""간단한 소송비용계산서
설치: python -m pip install streamlit python-docx
실행: python -m streamlit run app.py
"""
from decimal import Decimal, ROUND_FLOOR
from io import BytesIO
import re
import streamlit as st
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

def won(value):
    return f"{value:,}원"

def trunc_10(value):
    """10원 미만 절사 (국고금 관리법 준용)"""
    return (int(value) // 10) * 10

def parse_fraction(text):
    text = text.strip()
    if not text:
        return Fraction(1, 1)
    try:
        value = Fraction(text)
    except (ValueError, ZeroDivisionError):
        raise ValueError("부담비율은 0~1 사이의 분수로 입력해주세요. 분모는 0일 수 없습니다. (예: 1/3, 1/1)")
    if not 0 <= value <= 1:
        raise ValueError("부담비율은 0 이상 1 이하만 가능합니다. (예: 0/1, 1/2, 1/1)")
    return value


def fee_limit(soga):
    if soga <= 0:
        return 0
    for upper, lower, base, rate in BRACKETS:
        if upper is None or soga <= upper:
            value = Decimal(base) + Decimal(soga - lower) * Decimal(rate)
            return int(value.to_integral_value(rounding=ROUND_FLOOR))
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
    # 미입력 시 1심 기본 모델 자동 반환
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

def make_model(cases, soga_dict, actual_fees, filing_costs, options):
    stage_settings = options.get('stage_settings', {})
    
    rows = []
    filing_rows = []
    main_total = 0
    borne_main_total = 0
    
    for case in cases:
        c_num = case["number"]
        c_key = c_num if c_num else f"stage_{case['stage']}"
        stage = case["stage"]
        soga = soga_dict.get(c_key, 0)
        if soga <= 0:
            raise ValueError(f"{stage}심 소가를 올바르게 입력해주세요.")
            
        is_reduced = stage_settings.get(c_key, {}).get('is_reduced', False)
        limit = fee_limit(soga)
        if is_reduced:
            limit = limit // 2
            
        actual = actual_fees.get(c_key)
        if actual is not None and actual < 0:
            raise ValueError(f"{stage}심 실제 보수는 음수일 수 없습니다.")
            
        calc_fee = limit if actual is None else min(limit, actual)
        
        f_entry = filing_costs.get(c_key, {})
        include = f_entry.get('include', False)
        stamp = f_entry.get('stamp', 0) if include else 0
        delivery = f_entry.get('delivery', 0) if include else 0
        filing_rows.append(dict(stage=stage, number=c_num, stamp=stamp, delivery=delivery, include=include))

        settings = stage_settings.get(c_key, {'fraction': Fraction(1, 1), 'manual_name': '기타수기비용', 'manual_add': 0})
        fraction = settings['fraction']
        if not 0 <= fraction <= 1:
            raise ValueError('부담비율은 0 이상 1 이하만 가능합니다.')
        manual_name = settings['manual_name']
        manual_add = settings['manual_add']
        
        stage_sum = calc_fee + stamp + delivery + manual_add
        stage_borne = trunc_10(Decimal(stage_sum) * Decimal(fraction.numerator) / Decimal(fraction.denominator))
        
        rows.append({
            "stage": stage,
            "is_reduced": is_reduced,
            "number": c_num,
            "soga": soga,
            "limit": limit,
            "actual": actual,
            "calc_fee": calc_fee,
            "stamp": stamp,
            "delivery": delivery,
            "manual_name": manual_name,
            "manual_add": manual_add,
            "fraction": fraction,
            "stage_sum": stage_sum,
            "stage_borne": stage_borne
        })
        
        main_total += stage_sum
        borne_main_total += stage_borne

    resp_count = options.get('respondent_count', 1)
    app_stamp = 900 if options.get('electronic', True) else 1_000
    app_delivery = (resp_count + 1) * 3 * DELIVERY_UNIT
    application_total = app_stamp + app_delivery

    final_total = borne_main_total + application_total
    provisional = any(r["actual"] is None for r in rows)
    is_all_won = all(r['fraction'] == Fraction(1, 1) for r in rows)

    use_offset = options.get('use_offset', False)
    opp_rows = []
    opp_total = 0
    offset_final = final_total

    if use_offset:
        opp_settings = options.get('opp_settings', {})
        for case in cases:
            c_num = case["number"]
            c_key = c_num if c_num else f"stage_{case['stage']}"
            stage = case["stage"]
            soga = soga_dict.get(c_key, 0)
            settings = stage_settings.get(c_key, {'fraction': Fraction(1, 1)})
            app_burden_fraction = max(Fraction(0, 1), Fraction(1, 1) - settings['fraction'])
            
            o_entry = opp_settings.get(c_key, {})
            is_reduced = o_entry.get('is_reduced', False)
            o_limit = fee_limit(soga)
            if is_reduced:
                o_limit = o_limit // 2
            o_actual = o_entry.get('fee')
            o_calc_fee = o_limit if o_actual is None else min(o_limit, o_actual)
            
            o_stamp = o_entry.get('stamp', 0) if o_entry.get('include', False) else 0
            o_delivery = o_entry.get('delivery', 0) if o_entry.get('include', False) else 0
            o_manual = o_entry.get('manual', 0)
            
            o_sum = o_calc_fee + o_stamp + o_delivery + o_manual
            o_borne = trunc_10(Decimal(o_sum) * Decimal(app_burden_fraction.numerator) / Decimal(app_burden_fraction.denominator))
            
            opp_rows.append(dict(stage=stage, number=c_num, soga=soga,
                                 limit=o_limit, actual=o_actual, calc_fee=o_calc_fee,
                                 stamp=o_stamp, delivery=o_delivery, manual=o_manual,
                                 is_reduced=is_reduced, opp_sum=o_sum,
                                 app_fraction=app_burden_fraction, opp_borne=o_borne))
            opp_total += o_borne
        
        offset_final = final_total - opp_total

    return dict(rows=rows, filing_rows=filing_rows, main_total=main_total,
                borne_main_total=borne_main_total, application_total=application_total,
                app_stamp=app_stamp, app_delivery=app_delivery, total=final_total,
                provisional=provisional, respondent_count=resp_count, is_all_won=is_all_won,
                can_apply=offset_final > 0,
                use_offset=use_offset, opp_rows=opp_rows, opp_total=opp_total, offset_final=offset_final)


def application_sections(cases, model, options):
    refs = [f"{c['court']} {c['number']}".strip() for c in cases]
    refs = [r for r in refs if r]
    references = ', '.join(refs) if refs else '[법원 및 사건번호]'
    name = options['case_name'].strip() or '[사건명]'
    lawyer = LAW_FIRM
    date_text = (f"이 판결은 {options['final_date'].strip()} 확정되었습니다."
                 if options['final_date'].strip() else '')
    claim_amount = model['offset_final']
    if claim_amount <= 0:
        raise ValueError('신청인이 상환받을 잔액이 없어 신청서를 생성할 수 없습니다.')
    prayer = (f'위 당사자 사이의 {references} {name} 사건의 판결에 의하여 피신청인이 신청인에게 '
              f'상환해야 할 소송비용액은 금 {won(claim_amount)}임을 확정한다.\n라는 결정을 구합니다.')
    is_all_won = model['is_all_won']
    if is_all_won:
        opening = f'신청인({options["role"]})을 상대로 피신청인이 제기한 {references} {name} 사건에서 최종적으로 전부승소하였습니다. ' if options['role'] == '피고' else f'신청인({options["role"]})이 피신청인을 상대로 제기한 {references} {name} 사건에서 최종적으로 전부승소하였습니다. '
        action = '응소' if options['role'] == '피고' else '제기·수행'
        burden_text = '위 사건의 소송비용은 피신청인이 전액 부담하도록 정하여졌으므로'
    else:
        opening = f'신청인({options["role"]})을 상대로 피신청인이 제기한 {references} {name} 사건에서 최종적으로 일부승소 판결이 선고되었습니다. ' if options['role'] == '피고' else f'신청인({options["role"]})이 피신청인을 상대로 제기한 {references} {name} 사건에서 최종적으로 일부승소 판결이 선고되었습니다. '
        action = '응소' if options['role'] == '피고' else '제기·수행'
        burden_text = '위 사건들의 각 판결에서 정해진 피신청인의 소송비용 부담비율을 적용하여 산정하였으므로'

    offset_mention = " 또한, 당사자 쌍방의 부담액에 관하여 민사소송법 제112조에 따라 대등액에서 상계하고 남은 잔액을 청구합니다." if model.get('use_offset') else ""

    cause = (
        f'1. {opening}{date_text}\n\n'
        f'2. 신청인은 위 사건을 {action}하기 위하여 {lawyer}을 소송대리인으로 선임하였으며, 상환을 구하는 비용은 별지 소송비용액계산서와 같습니다.\n\n'
        f'3. {burden_text},{offset_mention} 신청인은 별지와 같은 금원의 상환을 구하기 위하여 이 사건 신청에 이르렀습니다.'
    )
    return prayer, cause


def generate_table_data(cases, model, options):
    table_rows = []
    
    for row in model['rows']:
        stage_str = f"{row['stage']}심"
        stage_rows = []
        
        formula_str = fee_formula(row['soga'], row['is_reduced'])
        if row['actual'] is not None and row['actual'] < row['limit']:
            note = f"규칙 제3조에 의거 산정하면 {formula_str} = {won(row['limit'])}이나, 실제 지급한 변호사보수를 소송비용으로 산입함"
        elif row['actual'] is not None:
            note = f"규칙 제3조에 의거 산정하면 {formula_str} = {won(row['limit'])}\n실제 지급한 변호사 보수 {won(row['actual'])}"
        else:
            note = f"규칙 제3조에 의거 산정: {formula_str} = {won(row['limit'])}"
            
        stage_rows.append({'비목': '변호사보수', '비용액': won(row['calc_fee']), '비고': note})
        
        if row['stamp']:
            stage_rows.append({'비목': '인지대', '비용액': won(row['stamp']), '비고': ''})
        if row['delivery']:
            stage_rows.append({'비목': '송달료', '비용액': won(row['delivery']), '비고': ''})
        if row['manual_add']:
            stage_rows.append({'비목': row['manual_name'], '비용액': won(row['manual_add']), '비고': ''})
            
        stage_rows.append({'비목': '소계', '비용액': won(row['stage_sum']), '비고': ''})
        
        frac = row['fraction']
        frac_str = f"{frac.numerator}/{frac.denominator}" if frac != Fraction(1, 1) else "100%"
        stage_rows.append({'비목': '상대방부담비율', '비용액': frac_str, '비고': ''})
        stage_rows.append({'비목': '상대방부담비용', '비용액': won(row['stage_borne']), '비고': ''})
        
        for i, sr in enumerate(stage_rows):
            table_rows.append({
                '심급': stage_str if i == 0 else '',
                '비목': sr['비목'],
                '비용액': sr['비용액'],
                '비고': sr['비고'],
                'merge_len': len(stage_rows) if i == 0 else 0
            })
            
    app_rows = [
        {'비목': '인지대', '비용액': won(model['app_stamp']), '비고': ''},
        {'비목': '송달료', '비용액': won(model['app_delivery']), '비고': f"피신청인 {model['respondent_count']}명 기준"},
        {'비목': '신청비용소계', '비용액': won(model['application_total']), '비고': ''}
    ]
    for i, ar in enumerate(app_rows):
        table_rows.append({
            '심급': '소송비용액확정신청' if i == 0 else '', 
            '비목': ar['비목'], 
            '비용액': ar['비용액'], 
            '비고': ar['비고'],
            'merge_len': len(app_rows) if i == 0 else 0
        })
        
    if model.get('use_offset'):
        for row in model['opp_rows']:
            frac = row['app_fraction']
            detail = [
                ('상대방 변호사보수', won(row['calc_fee']),
                 f"한도 {won(row['limit'])} / 지급·지급예정보수 "
                 + (won(row['actual']) if row['actual'] is not None else '미입력: 한도 적용')),
                ('상대방 인지대', won(row['stamp']), ''),
                ('상대방 송달료', won(row['delivery']), ''),
                ('상대방 기타 비용', won(row['manual']), ''),
                ('상대방 비용 소계', won(row['opp_sum']), ''),
                ('신청인 부담비율', f"{frac.numerator}/{frac.denominator}", ''),
                ('상계 공제액', won(row['opp_borne']), ''),
            ]
            for i, (item, amount, note) in enumerate(detail):
                table_rows.append({'심급': f"상대방 {row['stage']}심" if i == 0 else '',
                                   '비목': item, '비용액': amount, '비고': note,
                                   'merge_len': len(detail) if i == 0 else 0})
        table_rows.append({'심급': '상계 전 합계', '비목': '', '비용액': won(model['total']),
                           '비고': '申請人 본안 부담액 및 신청비용'.replace('申請人', '신청인'), 'merge_len': 1})
        table_rows.append({'심급': '공제액 합계', '비목': '', '비용액': won(model['opp_total']),
                           '비고': '상대방 비용 중 신청인 부담액', 'merge_len': 1})

    total_amount = model['offset_final']
    total_note = f"본안비용 {won(model['borne_main_total'])} + 신청비용 {won(model['application_total'])}"
    if model.get('use_offset'):
        total_note += f" − 상대방 비용 공제 {won(model['opp_total'])}"
        
    table_rows.append({
        '심급': '합계', '비목': '', '비용액': won(total_amount), '비고': total_note, 'merge_len': 1
    })
    
    return table_rows


def offset_comparison(model):
    """화면과 Word에 공통으로 사용하는 나란히 비교하는 상계표."""
    rows = []
    opponents = {r['stage']: r for r in model['opp_rows']}
    def add(stage, item, left='', right=''):
        rows.append({'심급': stage, '비목': item, '신청인 지출분': left, '상대방 지출분': right})
    for r in model['rows']:
        o = opponents[r['stage']]
        stage = f"{r['stage']}심"
        add(stage, '변호사보수 한도', won(r['limit']), won(o['limit']))
        add(stage, '실제 보수', won(r['actual']) if r['actual'] is not None else '미입력: 한도 적용',
            won(o['actual']) if o['actual'] is not None else '미입력: 한도 적용')
        add(stage, '산입 변호사보수', won(r['calc_fee']), won(o['calc_fee']))
        add(stage, '인지대', won(r['stamp']), won(o['stamp']))
        add(stage, '송달료', won(r['delivery']), won(o['delivery']))
        add(stage, '기타 비용', won(r['manual_add']), won(o['manual']))
        add(stage, '소계', won(r['stage_sum']), won(o['opp_sum']))
        f, g = r['fraction'], o['app_fraction']
        add(stage, '비용 부담비율', f"피신청인 {f.numerator}/{f.denominator}", f"신청인 {g.numerator}/{g.denominator}")
        add(stage, '비율 적용 금액', won(r['stage_borne']), won(o['opp_borne']))
        add(stage, '상계 후 차액', won(r['stage_borne'] - o['opp_borne']), '')
    net = model['borne_main_total'] - model['opp_total']
    add('합계', '본안 차액 합계', won(net))
    add('신청비용', '인지대', won(model['app_stamp']))
    add('신청비용', '송달료', won(model['app_delivery']))
    add('신청비용', '신청비용 소계', won(model['application_total']))
    add('최종', '최종 청구액' if model['offset_final'] > 0 else '상계 계산 결과', won(model['offset_final']))
    return rows


def render_text_calculation(cases, model, options):
    before = ['1. 신청인의 지출비용 및 피신청인 부담액']
    for index, row in enumerate(model['rows']):
        prefix = '가나다라마바사'[index]
        num_str = f" ({row['number']})" if row['number'] else ""
        stage_title = f"{prefix}. {row['stage']}심{num_str}"
        before.append(f"{stage_title} 소가: 금 {won(row['soga'])}")
        
        actual_str = f" (실제 지급액: 금 {won(row['actual'])})" if row['actual'] is not None else ""
        before.append(f"   - 변호사보수: 금 {won(row['calc_fee'])}{actual_str}")
        before.append(f"   - 최대인정보수 산식: {fee_formula(row['soga'], row['is_reduced'])}")

        if row['stamp']: before.append(f"   - 인지대: 금 {won(row['stamp'])}")
        if row['delivery']: before.append(f"   - 송달료: 금 {won(row['delivery'])}")
        if row['manual_add'] > 0: before.append(f"   - {row['manual_name']}: 금 {won(row['manual_add'])}")
            
        before.append(f"   => {row['stage']}심 지출 합계: 금 {won(row['stage_sum'])}")
        
        frac = row['fraction']
        if frac == Fraction(1, 1):
            before.append(f"   => 피신청인 부담비율(전부) 적용: 금 {won(row['stage_borne'])}")
        else:
            before.append(f"   => 피신청인 부담비율({frac.numerator}/{frac.denominator}) 적용: 금 {won(row['stage_borne'])}")
        before.append("")

    after = ['2. 소송비용액 산정']
    after.append(f"가. 본안 소송비용 중 피신청인 부담액 합계: 금 {won(model['borne_main_total'])}")
    after.append(f"나. 소송비용액확정신청 인지대 금 {won(model['app_stamp'])} 및 송달료 금 {won(model['app_delivery'])} (피신청인 {model['respondent_count']}명 기준)")
    after.append(f"다. 본안 및 신청비용 합계: {won(model['borne_main_total'])} + {won(model['application_total'])} = 금 {won(model['total'])}")

    if model.get('use_offset'):
        after.append("")
        after.append("3. 상대방 소송비용과의 대등액 상계 (민사소송법 제112조)")
        for row in model['opp_rows']:
            frac = row['app_fraction']
            after.append(f"- 상대방 {row['stage']}심: 변호사보수 {won(row['calc_fee'])}, 인지대 {won(row['stamp'])}, 송달료 {won(row['delivery'])}, 기타 {won(row['manual'])}")
            after.append(f"  소계 {won(row['opp_sum'])} × 신청인 부담비율 {frac.numerator}/{frac.denominator} = 공제액 {won(row['opp_borne'])}")
        after.append(f"가. 상대방 본안 소송비용 중 신청인 부담액 합계: 금 {won(model['opp_total'])}")
        if model['offset_final'] >= 0:
            after.append(f"나. 대등액 상계 후 피신청인이 상환해야 할 최종 잔액: 금 {won(model['total'])} − 금 {won(model['opp_total'])} = 금 {won(model['offset_final'])}    끝.")
        else:
            after.append(f"나. 대등액 상계 결과 (신청인의 초과 부담액 발생): 금 {won(abs(model['offset_final']))}    끝.")
    else:
        after.append("라. 피신청인이 신청인에게 상환해야 할 소송비용액은,")
        after.append(f"    금 {won(model['total'])}    끝.")
    
    return before, after


def make_docx(cases, model, options, prayer, cause):
    if model['offset_final'] <= 0:
        raise ValueError('신청인이 상환받을 잔액이 없어 DOCX를 생성할 수 없습니다.')
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin = section.bottom_margin = Cm(4)
    
    for style in doc.styles:
        for border in list(style.element.iter(qn('w:pBdr'))):
            border.getparent().remove(border)
            
    for name, font, size in [('Normal', '나눔명조', 12), ('Title', '나눔고딕', 13)]:
        style = doc.styles[name]
        style.font.name = font
        style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), font)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.font.size = Pt(size)
        style.paragraph_format.line_spacing = 1.25
        style.paragraph_format.space_after = Pt(8)
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
    
    doc.add_paragraph('별지')
    heading('소송비용액계산서')
    
    use_table_format = options.get('output_format') == '표 양식 (법원 서식)'
    
    if model.get('use_offset'):
        data = offset_comparison(model)
        headers = ['심급', '비목', '신청인 지출분', '상대방 지출분']
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        table.autofit = False
        widths = [Cm(1.8), Cm(4.2), Cm(5), Cm(5)]
        for col, width in zip(table.columns, widths):
            col.width = width
        for i, label in enumerate(headers):
            table.rows[0].cells[i].text = label
        repeat = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
        table.rows[0]._tr.get_or_add_trPr().append(repeat)
        for row in data:
            cells = table.add_row().cells
            for i, key in enumerate(headers):
                cells[i].text = row[key]
        for idx, row in enumerate(table.rows):
            row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            strong = idx == 0 or data[idx-1]['비목'] in ['소계', '상계 후 차액', '본안 차액 합계', '최종 청구액']
            for col, cell in enumerate(row.cells):
                cell.width = widths[col]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for para in cell.paragraphs:
                    para.paragraph_format.space_after = Pt(3)
                    para.paragraph_format.space_before = Pt(3)
                    para.paragraph_format.line_spacing = 1.1
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT if col >= 2 and idx else WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        run.font.size = Pt(10)
                        run.font.bold = strong
                if strong:
                    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F7"/>'))
        doc.add_paragraph('최종 청구액 = 각 심급 상계 후 차액 합계 + 소송비용액확정 신청비용')
    elif use_table_format:
        table_data = generate_table_data(cases, model, options)
        table = doc.add_table(rows=len(table_data)+1, cols=4)
        table.style = 'Table Grid'
        table.autofit = False
        
        tblPr = table._tbl.tblPr
        cell_mar = parse_xml(
            f'<w:tblCellMar {nsdecls("w")}>'
            f'<w:top w:w="100" w:type="dxa"/>'
            f'<w:bottom w:w="100" w:type="dxa"/>'
            f'<w:left w:w="140" w:type="dxa"/>'
            f'<w:right w:w="140" w:type="dxa"/>'
            f'</w:tblCellMar>'
        )
        tblPr.append(cell_mar)
        
        col_widths = [Cm(2.3), Cm(3.0), Cm(3.2), Cm(7.5)]

        headers = ['구분', '비목', '금액·비율', '비고']
        hdr_row = table.rows[0]
        for i, text in enumerate(headers):
            cell = hdr_row.cells[i]
            cell.text = text
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F4F7"/>')
            cell._tc.get_or_add_tcPr().append(shd)

        for r_idx, row in enumerate(table_data):
            cells = table.rows[r_idx+1].cells
            cells[0].text = row['심급']
            cells[1].text = row['비목']
            cells[2].text = str(row['비용액'])
            cells[3].text = row['비고']

        row_idx = 1
        for row in table_data:
            m_len = row.get('merge_len', 0)
            if m_len > 1:
                start_cell = table.cell(row_idx, 0)
                end_cell = table.cell(row_idx + m_len - 1, 0)
                merged = start_cell.merge(end_cell)
                while len(merged.paragraphs) > 1:
                    p = merged.paragraphs[-1]._p
                    p.getparent().remove(p)
                    
            if row['심급'] == '합계':
                start_cell = table.cell(row_idx, 0)
                end_cell = table.cell(row_idx, 1)
                merged = start_cell.merge(end_cell)
                while len(merged.paragraphs) > 1:
                    p = merged.paragraphs[-1]._p
                    p.getparent().remove(p)
                merged.paragraphs[0].text = '합계'
                
                for c in [merged, table.cell(row_idx, 2), table.cell(row_idx, 3)]:
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F9FAFB"/>')
                    c._tc.get_or_add_tcPr().append(shd)
                    
            row_idx += 1

        for r_idx, row_obj in enumerate(table.rows):
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
                        run.font.size = Pt(9.5)
                        if r_idx == 0 or (r_idx > 0 and table_data[r_idx-1]['비목'] in ['소계', '합계']):
                            run.font.bold = True
                    
                    if r_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        if c_idx == 2:
                            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        elif c_idx == 3:
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
    # 기본 중앙 정렬 레이아웃 복원
    st.set_page_config(page_title='소송비용 신청서 만들기', page_icon='⚖️')
    
    # 본문 폼은 아담하게 중앙 정렬 유지 + 표(Table)만 좌우로 확장하는 CSS Breakout 적용
    st.markdown("""
    <style>
    [data-testid="stDownloadButton"] button {
        background-color: #e0f2fe;
        color: #0369a1;
        border: 1px solid #bae6fd;
        font-weight: bold;
    }
    [data-testid="stDownloadButton"] button:hover {
        background-color: #bae6fd;
        color: #0c4a6e;
    }
    /* 표만 본문 폭을 넘어 시원하게 확장 (최대 1,000px) */
    div[data-testid="stMarkdownContainer"] table {
        width: 1000px !important;
        max-width: 90vw !important;
        margin-left: calc(-1 * (min(1000px, 90vw) - 100%) / 2) !important;
        border-collapse: collapse;
        font-size: 13.5px;
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        box-shadow: 0 2px 6px rgba(0,0,0,0.04);
        margin-top: 15px;
        margin-bottom: 25px;
    }
    div[data-testid="stMarkdownContainer"] th {
        background-color: #f8fafc;
        border-bottom: 2px solid #cbd5e1;
        padding: 9px 12px;
        font-weight: bold;
    }
    div[data-testid="stMarkdownContainer"] td {
        padding: 8px 12px;
        border-bottom: 1px solid #e2e8f0;
        vertical-align: middle;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title('소송비용 신청서 만들기')
    st.caption('소송비용액확정 신청서 및 계산서 간편 생성기')
    
    case_column, name_column = st.columns([3, 1])
    with case_column:
        raw = st.text_area('1·2·3심 법원·사건번호 (선택)',
                           placeholder='(예: 서울중앙지방법원 2023가합12345, 서울고등법원 2025나12345)', height=100)
    with name_column:
        case_name = st.text_input('사건명', placeholder='약정금')
    st.caption('진행한 심급만 쉼표나 줄바꿈으로 구분하세요. 미입력 시 1심 단독 사건으로 자동 처리됩니다.')
    
    error, cases = None, []
    try:
        cases = parse_cases(raw)
    except ValueError as exc:
        error = str(exc)
        cases = [{"stage": 1, "court": "", "number": "", "code": "가단"}]
        
    role = st.radio('신청인은 본안에서', ['피고', '원고'], horizontal=True)
    
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
        total_parties = respondent_count + 1
        auto_rounds = total_parties * 3
        auto_app_delivery = auto_rounds * DELIVERY_UNIT
        st.info(f"📌 **신청 비용 자동 계산 내역**\n\n- 인지대: **{won(auto_app_stamp)}** (전자 기준)\n- 송달료: **{won(auto_app_delivery)}**")

    st.caption('보수 감액 여부는 아래에서 해당 심급별로 선택하세요.')

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
        
        reduced = st.checkbox(f'{stage}심 신청인 변호사보수 1/2 감액 적용 (규칙 제5조)',
                              key=f'reduced_{c_key}',
                              help='피고의 전부자백·자백간주 판결, 무변론 판결, 이행권고결정 등 해당 심급에만 선택합니다.')
        soga_col, fee_col = st.columns(2)
        with soga_col:
            soga_val = int(st.number_input(f'{stage}심 소가(원)', min_value=0, value=50_000_000, step=1_000_000, key=f'soga_{c_key}'))
            soga_dict[c_key] = soga_val
        with fee_col:
            fee_val = st.text_input(f'{stage}심 실제 변호사보수(원)', key=f'fee_{c_key}',
                                    placeholder='생략 시 법정 한도액 적용')
            try:
                actual_fees[c_key] = optional_money(fee_val)
            except ValueError as exc:
                error = str(exc)

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
        
        stage_settings[c_key] = {'is_reduced': reduced, 'fraction': stage_fraction, 'manual_name': manual_name, 'manual_add': manual_add}

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
    with st.expander('⚖️ 상대방 소송비용 상계(대등액 공제) 시뮬레이션 (선택)'):
        use_offset = st.checkbox('상대방 소송비용 상계 계산 활성화', value=False)
        if use_offset:
            is_all_won = all(stage_settings.get(c['number'] or f"stage_{c['stage']}", {}).get('fraction', Fraction(1, 1)) == Fraction(1, 1) for c in cases)
            if is_all_won:
                st.warning("⚠️ **현재 피신청인 부담비율이 '1/1(전부승소)'입니다.**\n\n신청인이 물어줘야 할 비용이 0%이므로 상계액이 계산되지 않습니다. 부담비율을 수정해주세요.")
                
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
                    
                opp_reduced = st.checkbox(f'{stage}심 상대방 변호사보수 1/2 감액 적용 (규칙 제5조)', key=f'opp_reduced_{c_key}')
                opp_settings[c_key] = {
                    'is_reduced': opp_reduced,
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
        
    st.markdown("---")
    if use_offset:
        st.caption('상계 사용 시 화면과 DOCX 계산서는 신청인·상대방 비교표로 출력됩니다.')
    output_format = st.radio('📝 계산서 출력 양식', ['줄글 양식 (기본)', '표 양식 (법원 서식)'], horizontal=True)
        
    options = dict(electronic=electronic, party_count=party_count,
                   role=role, case_name=case_name, final_date=final_date,
                   respondent_count=respondent_count, 
                   stage_settings=stage_settings, use_offset=use_offset, opp_settings=opp_settings,
                   output_format=output_format)
    
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
    
    if model.get('use_offset'):
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric('상계 전 신청인 총 청구액', won(model['total']))
        with col_m2:
            delta_str = f"-{won(model['opp_total'])}" if model['opp_total'] > 0 else "0원"
            st.metric('상계 후 최종 상환액', won(model['offset_final']), delta=f"상대방 채권 공제: {delta_str}")
    else:
        st.metric('최종 상환액 합계', won(model['total']))
        
    if not model['can_apply']:
        st.warning('신청인이 상환받을 잔액이 없습니다. 신청서 및 DOCX 다운로드를 생성하지 않습니다.')
        if model['offset_final'] < 0:
            st.info(f"계산상 신청인의 초과 부담액: {won(-model['offset_final'])}")
        st.subheader('상계 계산 내역 확인')
        st.table(offset_comparison(model))
        return
    prayer, cause = application_sections(cases, model, options)
    prayer = st.text_area('신청취지 편집', value=prayer, height=180)
    cause = st.text_area('신청원인 편집', value=cause, height=300)
    
    st.subheader('별지 소송비용액계산서')
    
    if model.get('use_offset'):
        st.table(offset_comparison(model))
        st.caption('최종 청구액 = 각 심급 상계 후 차액 합계 + 소송비용액확정 신청비용')
    elif output_format == '표 양식 (법원 서식)':
        table_data = generate_table_data(cases, model, options)
        md_table = "| 구분 | 비목 | 금액·비율 | 비고 |\n|:---:|:---:|---:|:---|\n"
        for row in table_data:
            display_stage = row['심급']
            display_bimok = row['비목']
            if row['심급'] == '합계':
                display_bimok = ""
            formatted_row = [str(item).replace("\n", "<br>") for item in [display_stage, display_bimok, row['비용액'], row['비고']]]
            md_table += "| " + " | ".join(formatted_row) + " |\n"
        st.markdown(md_table, unsafe_allow_html=True)
    else:
        before, after = render_text_calculation(cases, model, options)
        for line in before:
            st.write(line)
        for line in after:
            st.write(line)
        
    args = (cases, model, options, prayer, cause)
    try:
        document = make_docx(*args)
    except ImportError:
        st.info('Word 저장: 터미널에서 python -m pip install python-docx를 실행하세요.')
    else:
        st.download_button('신청서·계산서 초안 다운로드', document,
                           file_name='소송비용액확정_신청서_계산서.docx',
                           mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                           use_container_width=True)

if __name__ == '__main__':
    main()
