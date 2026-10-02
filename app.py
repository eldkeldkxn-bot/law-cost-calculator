"""간단한 소송비용계산서
설치: python -m pip install streamlit python-docx
실행: python -m streamlit run app.py
"""
from decimal import Decimal, ROUND_FLOOR
from io import BytesIO
import re
import streamlit as st

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


def fee_limit(soga):
    if soga <= 0:
        raise ValueError("소가를 0원보다 크게 입력해주세요.")
    for upper, lower, base, rate in BRACKETS:
        if upper is None or soga <= upper:
            value = Decimal(base) + Decimal(soga - lower) * Decimal(rate)
            return int(value.to_integral_value(rounding=ROUND_FLOOR))


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


def calculate(soga, cases, actual_fees, extra, is_reduced=False):
    limit = fee_limit(soga)
    if is_reduced:
        limit = limit // 2
    if extra < 0:
        raise ValueError("추가 비용은 음수일 수 없습니다.")
    rows = []
    for case in cases:
        actual = actual_fees.get(case["number"])
        if actual is not None and actual < 0:
            raise ValueError("실제 보수는 음수일 수 없습니다.")
        rows.append({
            "심급": f"{case['stage']}심",
            "비목": "변호사보수",
            "최대인정보수": limit,
            "실제 보수": actual,
            "계산액": limit if actual is None else min(limit, actual)
        })
    provisional = any(row["실제 보수"] is None for row in rows)
    return rows, sum(row["계산액"] for row in rows) + extra, provisional


def filing_estimate(soga, case, electronic=True, party_count=1):
    if soga <= 0 or party_count < 1:
        raise ValueError('소가와 송달 대상 인원은 1 이상이어야 합니다.')
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


def make_model(soga, cases, actual_fees, options):
    is_reduced = options.get('is_reduced', False)
    rows, attorney_total, provisional = calculate(soga, cases, actual_fees, 0, is_reduced)
    direct = options['source'] == 'LawTop 합계액 입력'
    filing_rows = []
    
    if direct:
        main_total = options['lawtop_total']
        if main_total is None or main_total <= 0:
            raise ValueError('LawTop에서 확인한 전체 심급의 본안 합계액을 입력해주세요.')
        provisional = False
    else:
        for case, row in zip(cases, rows):
            estimate_stamp, estimate_delivery = filing_estimate(
                soga, case, options.get('electronic', True), options.get('party_count', 1))
            entry = options.get('filing_costs', {}).get(case['number'], {})
            include = entry.get('include', options['role'] == '원고' and case['stage'] == 1)
            stamp = entry.get('stamp', estimate_stamp) if include else 0
            delivery = entry.get('delivery', estimate_delivery) if include else 0
            if stamp < 0 or delivery < 0:
                raise ValueError('본안 인지대·송달료는 음수일 수 없습니다.')
            filing_rows.append(dict(stage=case['stage'], stamp=stamp, delivery=delivery, include=include))
        main_total = attorney_total + sum(r['stamp'] + r['delivery'] for r in filing_rows)

    # 일부승소 부담비율 적용 (본안비용 * 부담비율)
    burden_rate = Decimal(str(options.get('burden_rate', 100))) / Decimal("100")
    borne_main_total = int((Decimal(main_total) * burden_rate).to_integral_value(rounding=ROUND_FLOOR))

    # 신청사건 비용 (인지: 전자 900원 / 종이 1,000원, 송달료: 피신청인수 * 6회분)
    resp_count = options.get('respondent_count', 1)
    app_stamp = 900 if options.get('electronic', True) else 1_000
    app_delivery = resp_count * 6 * DELIVERY_UNIT
    application_total = app_stamp + app_delivery

    # 최종 상환액 = (본안 지출액 * 피신청인 부담비율) + 신청비용 전액
    final_total = borne_main_total + application_total

    return dict(rows=rows, main_total=main_total, borne_main_total=borne_main_total,
                application_total=application_total, app_stamp=app_stamp, app_delivery=app_delivery,
                total=final_total, provisional=provisional, direct=direct,
                filing_rows=filing_rows, burden_rate=options.get('burden_rate', 100))


def fee_formula(soga, is_reduced=False):
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


def application_sections(cases, model, options):
    references = ', '.join(' '.join(filter(None, [c['court'], c['number']])) for c in cases)
    name = options['case_name'].strip() or '[사건명]'
    lawyer = LAW_FIRM
    date_text = (f"이 판결은 {options['final_date'].strip()} 확정되었습니다."
                 if options['final_date'].strip() else '')
    amount = won(model['total'])
    burden_rate = model['burden_rate']
    
    prayer = (f'위 당사자 사이의 {references} {name} 사건의 판결에 의하여 피신청인이 신청인에게 '
              f'상환해야 할 소송비용액은 금 {amount}임을 확정한다.\n라는 결정을 구합니다.')

    # 승소 및 부담 문구 분기 (전부승소 vs 일부승소)
    if burden_rate == 100:
        if options['role'] == '피고':
            opening = (f'신청인(피고, 이하 신청인이라고 함)을 상대로 피신청인(원고, 이하 피신청인이라고 함)이 제기한 '
                       f'{references} {name} 사건에서 최종적으로 신청인에 대한 피신청인의 청구가 모두 기각되어 '
                       '신청인이 전부승소하였습니다. ')
            action = '응소'
        else:
            opening = (f'신청인(원고, 이하 신청인이라고 함)이 피신청인(피고, 이하 피신청인이라고 함)을 상대로 제기한 '
                       f'{references} {name} 사건에서 최종적으로 신청인의 청구가 모두 인용되어 신청인이 전부승소하였습니다. ')
            action = '제기·수행'
        burden_text = '위 사건의 소송비용은 피신청인이 전액 부담하도록 정하여졌으므로'
    else:
        if options['role'] == '피고':
            opening = (f'신청인(피고, 이하 신청인이라고 함)을 상대로 피신청인(원고, 이하 피신청인이라고 함)이 제기한 '
                       f'{references} {name} 사건에서 최종적으로 일부승소 판결이 선고되었습니다. ')
            action = '응소'
        else:
            opening = (f'신청인(원고, 이하 신청인이라고 함)이 피신청인(피고, 이하 피신청인이라고 함)을 상대로 제기한 '
                       f'{references} {name} 사건에서 최종적으로 일부승소 판결이 선고되었습니다. ')
            action = '제기·수행'
        burden_text = f'위 사건의 소송비용 중 피신청인의 부담비율은 {burden_rate}%로 정하여졌으므로'

    cause = (
        f'1. {opening}{date_text}\n\n'
        f'2. 신청인은 위 사건을 {action}하기 위하여 {lawyer}을 소송대리인으로 선임하였으며, '
        '상환을 구하는 비용은 별지 소송비용액계산서와 같습니다.\n\n'
        f'3. {burden_text}, '
        f'신청인은 별지와 같이 금 {amount}의 상환을 구하기 위하여 이 사건 신청에 이르렀습니다.'
    )
    return prayer, cause


def calculation_sections(soga, cases, model, options):
    labels = ', '.join(str(c['stage']) for c in cases)
    before = [f'1. 신청인의 지출비용', f'- {labels}심 소가: 금 {won(soga)}']
    if model['direct']:
        before.append('본안 소송비용 합계: 금 ' + won(model['main_total']) + ' (LawTop 입력값)')
    else:
        for index, row in enumerate(model['rows']):
            actual_str = f" (실제 지급액: 금 {won(row['실제 보수'])})" if row['실제 보수'] is not None else " (실제 지급액 미입력, 한도액 적용)"
            before.append(f"{'가나다'[index]}. {row['심급']} 변호사보수: 금 {won(row['계산액'])}{actual_str}")
            before.append(f"   [최대인정보수 산식: {fee_formula(soga, options.get('is_reduced', False))}]")
        for entry in model['filing_rows']:
            if entry['include']:
                before.append(f"{entry['stage']}심 인지대: 금 {won(entry['stamp'])} / 송달료: 금 {won(entry['delivery'])}")
        before.append('본안 소송비용 합계: 금 ' + won(model['main_total']))

    burden_rate = model['burden_rate']
    after = ['2. 소송비용액 산정']
    if burden_rate == 100:
        after.append('가. 신청인의 본안 소송비용: 금 ' + won(model['main_total']))
    else:
        after.append(f"가. 신청인의 본안 소송비용: 금 {won(model['main_total'])}")
        after.append(f"나. 피신청인 부담비율({burden_rate}%) 적용 본안비용: 금 {won(model['main_total'])} × {burden_rate}% = 금 {won(model['borne_main_total'])}")
    
    label_next = '나.' if burden_rate == 100 else '다.'
    label_final = '다.' if burden_rate == 100 else '라.'
    after.append(f"{label_next} 소송비용액확정신청에 따른 인지대 금 {won(model['app_stamp'])} 및 송달료 금 {won(model['app_delivery'])}")
    
    if burden_rate == 100:
        after.extend([
            f"{label_final} 소송비용은 피신청인의 부담이므로 피신청인이 신청인에게 상환해야 할 소송비용액은,",
            f"{won(model['main_total'])} + {won(model['application_total'])} = {won(model['total'])}    끝.",
        ])
    else:
        after.extend([
            f"{label_final} 피신청인이 신청인에게 상환해야 할 소송비용액은,",
            f"{won(model['borne_main_total'])} + {won(model['application_total'])} = {won(model['total'])}    끝.",
        ])
    notes = []
    return before, after, notes


def make_text(soga, cases, model, options, prayer, cause):
    before, after, notes = calculation_sections(soga, cases, model, options)
    return '\n'.join(['신 청 취 지', '', prayer, '', '신 청 원 인', '', cause,
                      '', '별지', '소송비용액계산서', '', *before, '', *after, '', *notes])


def make_docx(soga, cases, model, options, prayer, cause):
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    
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
    
    before, after, notes = calculation_sections(soga, cases, model, options)
    paragraphs(before)
    doc.add_paragraph()
    paragraphs(after)
    paragraphs(notes)
    
    output = BytesIO()
    doc.save(output)
    return output.getvalue()


def main():
    st.set_page_config(page_title='소송비용 신청서 만들기', page_icon='⚖️')
    st.title('소송비용 신청서 만들기')
    st.caption('소송비용액확정 신청서 및 계산서 간편 생성기')
    
    soga = int(st.number_input('소가(원)', min_value=0, value=50_000_000, step=100_000))
    case_column, name_column = st.columns([3, 1])
    with case_column:
        raw = st.text_area('1·2·3심 법원·사건번호',
                           placeholder='서울중앙지방법원 2023가합12345, 서울고등법원 2025나12345, 대법원 2026다12345', height=100)
    with name_column:
        case_name = st.text_input('사건명', placeholder='약정금')
    st.caption('진행한 심급만 쉼표나 줄바꿈으로 구분하세요. 행정, 가사 사건도 입력 가능합니다.')
    
    role = st.radio('신청인은 본안에서', ['피고', '원고'], horizontal=True)
    source = st.radio('본안 소송비용', ['소가로 계산', 'LawTop 합계액 입력'], horizontal=True)
    
    # 일부승소 부담비율 입력
    burden_rate = st.slider('피신청인 소송비용 부담비율(%)', min_value=1, max_value=100, value=100, step=1,
                            help='전부승소는 100%입니다. 판결 주문에 따라 부담 비율을 조절하세요 (예: 피고 70% 부담 시 70).')
    
    error, cases, lawtop_total = None, [], None
    try:
        cases = parse_cases(raw)
    except ValueError as exc:
        error = str(exc)
        
    electronic, party_count = True, 1
    if source == 'LawTop 합계액 입력':
        total_text = st.text_input('LawTop에서 확인한 본안 합계액(원)', placeholder='예: 33,669,820')
        st.caption('모든 심급의 본안 비용 합계를 입력하세요. 신청사건 비용은 별도로 계산되어 합산됩니다.')
        try:
            lawtop_total = optional_money(total_text)
        except ValueError:
            error = 'LawTop 합계액을 숫자로 입력해주세요. 쉼표는 사용할 수 있습니다.'
    else:
        method_column, people_column = st.columns(2)
        with method_column:
            electronic = st.radio('제출 방식', ['전자소송', '종이소송'], horizontal=True) == '전자소송'
        with people_column:
            party_count = int(st.number_input('본안 송달 대상 인원', min_value=1,
                                              value=1 if electronic else 2, step=1,
                                              help='전자소송은 상대방 수, 종이소송은 전체 당사자 수를 기본으로 확인해 입력하세요.'))
        st.caption('원고의 1심 인지대·송달료는 자동 포함됩니다. 항소·상고 비용은 아래에서 신청인이 납부한 심급만 선택하세요.')
        
    with st.expander('산정 추가 옵션 (무변론/자백간주 등)'):
        is_reduced = st.checkbox('무변론 판결, 자백간주, 이행권고결정에 따른 변호사보수 1/2 감액 적용 (규칙 제5조)')
        respondent_count = int(st.number_input('비용확정신청 송달 대상 피신청인 수', min_value=1, value=1, step=1,
                                               help='피신청인이 여러 명인 경우 신청 송달료(인원 × 6회분)가 증가합니다.'))
        st.info('💡 **변호사보수 부가세(VAT) 참고:** 신청인이 일반과세자 사업자로서 매입세액공제를 받는 경우에는 부가세를 제외한 공급가액을 실제 보수로 입력하세요.')

    actual_fees, filing_costs = {}, {}
    if source == '소가로 계산':
        with st.expander('심급별 실제 보수·납부 비용 확인 (선택)'):
            st.caption('실제 보수 미입력 시 법정 한도액을 사용합니다. 납부액·환급액이 다르면 실제 부담액으로 수정하세요.')
            for case in cases:
                st.markdown(f"**{case['stage']}심 ({case['number']})**")
                prefix = case['number']
                value = st.text_input('실제 변호사보수(원)', key='fee:' + prefix,
                                      placeholder='생략하면 법정 한도액 기준')
                try:
                    actual_fees[prefix] = optional_money(value)
                except ValueError as exc:
                    error = str(exc)
                include = st.checkbox('신청인이 납부한 이 심급의 인지대·송달료 포함',
                                      value=role == '원고' and case['stage'] == 1,
                                      key='paid:' + role + ':' + prefix)
                stamp, delivery = (0, 0)
                if soga > 0:
                    stamp, delivery = filing_estimate(soga, case, electronic, party_count)
                entry = dict(include=include, stamp=stamp, delivery=delivery)
                if include:
                    st.caption(f"자동 계산: 인지대 {won(stamp)} / 송달료 {won(delivery)}")
                    if st.checkbox('실제 납부액(또는 환급 후 잔액)으로 직접 수정', key='override:' + prefix):
                        entry['stamp'] = int(st.number_input('인지대 실제 부담액(원)', min_value=0,
                                                             value=stamp, key='stamp:' + prefix))
                        entry['delivery'] = int(st.number_input('송달료 실제 부담액(원)', min_value=0,
                                                                value=delivery, key='delivery:' + prefix))
                filing_costs[prefix] = entry

    st.caption('소송대리인: 법무법인(유한)바른')
    with st.expander('확정일 입력 (선택)'):
        final_date = st.text_input('확정일', placeholder='예: 2026. 7. 21.')
        st.caption('입력하지 않으면 확정일 문장은 문서에 넣지 않습니다.')
        
    options = dict(source=source, lawtop_total=lawtop_total, electronic=electronic, party_count=party_count,
                   filing_costs=filing_costs, role=role, case_name=case_name, final_date=final_date,
                   is_reduced=is_reduced, respondent_count=respondent_count, burden_rate=burden_rate)
    signature = repr((soga, cases, raw, actual_fees, options))
    
    if st.button('신청서·계산서 만들기', type='primary', use_container_width=True):
        try:
            if error:
                raise ValueError(error)
            if not case_name.strip():
                raise ValueError('사건번호 옆에 사건명을 입력해주세요.')
            model = make_model(soga, cases, actual_fees, options)
        except ValueError as exc:
            st.error(str(exc))
            st.session_state.pop('template_result', None)
        else:
            st.session_state['template_result'] = (signature, model)
            
    saved = st.session_state.get('template_result')
    if not saved or saved[0] != signature:
        return
        
    model = saved[1]
    st.metric('최종 상환액 합계', won(model['total']))
    if model['burden_rate'] == 100:
        st.caption(f"본안 비용 {won(model['main_total'])} + 비용확정신청 비용 {won(model['application_total'])}")
    else:
        st.caption(f"본안 지출액 {won(model['main_total'])} 중 {model['burden_rate']}% 부담액 {won(model['borne_main_total'])} + 비용확정신청 비용 {won(model['application_total'])}")
        
    if model['provisional']:
        st.caption('실제 보수를 입력하지 않은 심급은 법정 한도액이 반영되었습니다.')
        
    prayer, cause = application_sections(cases, model, options)
    prayer = st.text_area('신청취지 편집', value=prayer, height=180)
    cause = st.text_area('신청원인 편집', value=cause, height=300)
    
    st.subheader('별지 소송비용액계산서')
    before, after, notes = calculation_sections(soga, cases, model, options)
    for line in before:
        st.write(line)
    for line in after:
        st.write(line)
        
    args = (soga, cases, model, options, prayer, cause)
    st.download_button('신청서·계산서 TXT 저장', make_text(*args),
                       file_name='소송비용액확정_신청서_계산서.txt', mime='text/plain; charset=utf-8')
    try:
        document = make_docx(*args)
    except ImportError:
        st.info('Word 저장: 터미널에서 python -m pip install python-docx를 실행하세요.')
    else:
        st.download_button('신청서·계산서 DOCX 저장', document,
                           file_name='소송비용액확정_신청서_계산서.docx',
                           mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    st.markdown('[LawTop 소송비용 계산기](https://support.lawtop.co.kr/fee_renew/index1.asp)')


if __name__ == '__main__':
    main()