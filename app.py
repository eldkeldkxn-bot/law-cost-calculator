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
    output_format = st.radio('📝 계산서 출력 양식', ['줄글 양식 (기본)', '표 양식 (법원 서식)'], horizontal=True)
        
    options = dict(electronic=electronic, party_count=party_count,
                   role=role, case_name=case_name, final_date=final_date,
                   judgment_results=judgment_results, cost_order=cost_order, respondent_count=respondent_count, 
                   stage_settings=stage_settings, use_offset=use_offset, opp_settings=opp_settings,
                   output_format=output_format)
    
    signature = repr((cases, raw, soga_dict, actual_fees, filing_costs, options))
    
    if st.button('신청서·계산서 만들기', type='primary', use_container_width=True):
        try:
            if error:
                raise ValueError(error)
            if not cost_order.strip():
                raise ValueError('소송비용 부담 주문을 입력해주세요.')
            if any(not value.strip() for value in judgment_results.values()):
                raise ValueError('각 심급의 판결 결과를 입력해주세요.')
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
        st.table([{k: v for k, v in row.items() if k != 'merge_len'}
                  for row in generate_table_data(cases, model, options)])
        return
    prayer, cause = application_sections(cases, model, options)
    prayer = st.text_area('신청취지 편집', value=prayer, height=180)
    cause = st.text_area('신청원인 편집', value=cause, height=300)
    
    st.subheader('별지 소송비용액계산서')
    
    if output_format == '표 양식 (법원 서식)':
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
