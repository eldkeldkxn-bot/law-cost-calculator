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
    rows = []
    filing_rows = []
    
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
        rows.append({
            "stage": stage,
            "number": c_num,
            "soga": soga,
            "limit": limit,
            "actual": actual,
            "calc_fee": calc_fee
        })
        
        # 인지대 및 송달료
        f_entry = filing_costs.get(c_num, {})
        include = f_entry.get('include', False)
        stamp = f_entry.get('stamp', 0) if include else 0
        delivery = f_entry.get('delivery', 0) if include else 0
        filing_rows.append(dict(stage=stage, number=c_num, stamp=stamp, delivery=delivery, include=include))

    attorney_total = sum(r["calc_fee"] for r in rows)
    filing_total = sum(r["stamp"] + r["delivery"] for r in filing_rows)
    main_total = attorney_total + filing_total

    # 피신청인 부담비율 적용 (본안비용 * 부담비율)
    burden_rate = Decimal(str(options.get('burden_rate', 100))) / Decimal("100")
    borne_main_total = int((Decimal(main_total) * burden_rate).to_integral_value(rounding=ROUND_FLOOR))

    # 신청사건 비용 (인지: 전자 900원 / 종이 1,000원, 송달료: 피신청인 수 * 6회분)
    resp_count = options.get('respondent_count', 1)
    app_stamp = 900 if options.get('electronic', True) else 1_000
    app_delivery = resp_count * 6 * DELIVERY_UNIT
    application_total = app_stamp + app_delivery

    final_total = borne_main_total + application_total
    provisional = any(r["actual"] is None for r in rows)

    return dict(rows=rows, filing_rows=filing_rows, main_total=main_total,
                borne_main_total=borne_main_total, application_total=application_total,
                app_stamp=app_stamp, app_delivery=app_delivery, total=final_total,
                provisional=provisional, burden_rate=options.get('burden_rate', 100),
                respondent_count=resp_count)


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


def calculation_sections(cases, model, options):
    before = ['1. 신청인의 지출비용']
    
    # 심급별 소가 및 변호사보수
    for index, row in enumerate(model['rows']):
        prefix = '가나다'[index]
        stage_title = f"{prefix}. {row['stage']}심 ({row['number']})"
        before.append(f"{stage_title} 소가: 금 {won(row['soga'])}")
        
        actual_str = f" (실제 지급액: 금 {won(row['actual'])})" if row['actual'] is not None else " (실제 지급액 미입력, 한도액 적용)"
        before.append(f"   - 변호사보수: 금 {won(row['calc_fee'])}{actual_str}")
        before.append(f"   - 최대인정보수 산식: {fee_formula(row['soga'], options.get('is_reduced', False))}")

    # 인지대 / 송달료 지출내역
    for entry in model['filing_rows']:
        if entry['include']:
            before.append(f"- {entry['stage']}심 인지대: 금 {won(entry['stamp'])} / 송달료: 금 {won(entry['delivery'])}")
            
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
    after.append(f"{label_next} 소송비용액확정신청에 따른 인지대 금 {won(model['app_stamp'])} 및 송달료 금 {won(model['app_delivery'])} (상대방 {model['respondent_count']}명 기준)")
    
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


def make_text(cases, model, options, prayer, cause):
    before, after, notes = calculation_sections(cases, model, options)
    return '\n'.join(['신 청 취 지', '', prayer, '', '신 청 원 인', '', cause,
                      '', '별지', '소송비용액계산서', '', *before, '', *after, '', *notes])


def make_docx(cases, model, options, prayer, cause):
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
    
    before, after, notes = calculation_sections(cases, model, options)
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
    
    case_column, name_column = st.columns([3, 1])
    with case_column:
        raw = st.text_area('1·2·3심 법원·사건번호',
                           placeholder='서울중앙지방법원 2023가합12345, 서울고등법원 2025나12345, 대법원 2026다12345', height=100)
    with name_column:
        case_name = st.text_input('사건명', placeholder='약정금')
    st.caption('진행한 심급만 쉼표나 줄바꿈으로 구분하세요. 행정, 가사 사건도 입력 가능합니다.')
    
    error, cases = None, []
    try:
        cases = parse_cases(raw)
    except ValueError as exc:
        error = str(exc)
        
    role = st.radio('신청인은 본안에서', ['피고', '원고'], horizontal=True)
    
    method_column, people_column = st.columns(2)
    with method_column:
        electronic = st.radio('제출 방식', ['전자소송', '종이소송'], horizontal=True) == '전자소송'
    with people_column:
        party_count = int(st.number_input('본안 송달 대상 인원수', min_value=1,
                                          value=1 if electronic else 2, step=1,
                                          help='본안 소송의 송달료 계산 기준 인원수입니다.'))

    # 소송비용확정 신청사건 상대방 수 및 송달료 자동계산
    st.markdown("---")
    st.subheader("소송비용액확정 신청 비용 설정")
    resp_col1, resp_col2 = st.columns([1, 2])
    with resp_col1:
        respondent_count = int(st.number_input('상대방(피신청인) 수', min_value=1, value=1, step=1,
                                               help='피신청인 수에 따라 신청사건 송달료(인원 × 6회분 × 5,640원)가 자동 산출됩니다.'))
    with resp_col2:
        auto_app_stamp = 900 if electronic else 1_000
        auto_app_delivery = respondent_count * 6 * DELIVERY_UNIT
        st.info(f"📌 **신청 비용 자동 계산 내역**\n\n- 인지대: **{won(auto_app_stamp)}** (전자 기준)\n- 송달료: **{won(auto_app_delivery)}** ({respondent_count}명 × 6회분 × 5,640원)")

    # 일부승소 부담비율 입력
    burden_rate = st.slider('피신청인 소송비용 부담비율(%)', min_value=1, max_value=100, value=100, step=1,
                            help='전부승소는 100%입니다. 판결 주문에 따라 부담 비율을 조절하세요 (예: 피고 70% 부담 시 70).')
    
    is_reduced = st.checkbox('무변론 판결, 자백간주, 이행권고결정에 따른 변호사보수 1/2 감액 적용 (규칙 제5조)')

    # 심급별 소가 및 실제 보수, 인지/송달료 입력
    st.markdown("---")
    st.subheader("심급별 소가 및 비용 설정")
    soga_dict = {}
    actual_fees = {}
    filing_costs = {}

    if cases:
        for case in cases:
            prefix = case['number']
            stage = case['stage']
            st.markdown(f"#### 🏛️ {stage}심 ({case['court']} {prefix})")
            
            soga_col, fee_col = st.columns(2)
            with soga_col:
                soga_val = int(st.number_input(f'{stage}심 소가(원)', min_value=0, value=50_000_000, step=1_000_000, key=f'soga_{prefix}'))
                soga_dict[prefix] = soga_val
            with fee_col:
                fee_val = st.text_input(f'{stage}심 실제 변호사보수(원)', key=f'fee_{prefix}',
                                        placeholder='생략 시 법정 한도액 적용')
                try:
                    actual_fees[prefix] = optional_money(fee_val)
                except ValueError as exc:
                    error = str(exc)

            # 인지대 송달료
            include = st.checkbox(f'신청인이 납부한 {stage}심 인지대·송달료 포함',
                                  value=role == '원고' and stage == 1,
                                  key=f'paid_{role}_{prefix}')
            stamp, delivery = (0, 0)
            if soga_val > 0:
                stamp, delivery = filing_estimate(soga_val, case, electronic, party_count)
            entry = dict(include=include, stamp=stamp, delivery=delivery)
            if include:
                st.caption(f"자동 계산: 인지대 {won(stamp)} / 송달료 {won(delivery)}")
                if st.checkbox(f'{stage}심 인지·송달료 실제 납부액으로 직접 수정', key=f'override_{prefix}'):
                    col_st, col_dl = st.columns(2)
                    with col_st:
                        entry['stamp'] = int(st.number_input(f'{stage}심 인지대 실제부담액', min_value=0, value=stamp, key=f'stamp_{prefix}'))
                    with col_dl:
                        entry['delivery'] = int(st.number_input(f'{stage}심 송달료 실제부담액', min_value=0, value=delivery, key=f'delivery_{prefix}'))
            filing_costs[prefix] = entry
            st.divider()
    else:
        st.info("상단에 법원 및 사건번호를 입력하면 심급별 소가 및 비용 입력창이 나타납니다.")

    st.caption('소송대리인: 법무법인(유한)바른')
    with st.expander('확정일 입력 (선택)'):
        final_date = st.text_input('확정일', placeholder='예: 2026. 7. 21.')
        st.caption('입력하지 않으면 확정일 문장은 문서에 넣지 않습니다.')
        
    options = dict(electronic=electronic, party_count=party_count,
                   role=role, case_name=case_name, final_date=final_date,
                   is_reduced=is_reduced, respondent_count=respondent_count, burden_rate=burden_rate)
    
    signature = repr((cases, raw, soga_dict, actual_fees, filing_costs, options))
    
    if st.button('신청서·계산서 만들기', type='primary', use_container_width=True):
        try:
            if error:
                raise ValueError(error)
            if not cases:
                raise ValueError('사건번호를 올바르게 입력해주세요.')
            if not case_name.strip():
                raise ValueError('사건명을 입력해주세요.')
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
    before, after, notes = calculation_sections(cases, model, options)
    for line in before:
        st.write(line)
    for line in after:
        st.write(line)
        
    args = (cases, model, options, prayer, cause)
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


if __name__ == '__main__':
    main()
