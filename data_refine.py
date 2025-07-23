def process_all_files(file1, file2, file3, df_master):
    try:
        df_smartstore = pd.read_excel(file1)
        df_ecount_orig = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        df_ecount_orig['original_order'] = range(len(df_ecount_orig))
        
        # <<-- 최종 수정: 고도몰 실결제금액 처리 로직 전면 수정 -->>
        cols_to_numeric = ['상품별 품목금액', '총 배송 금액', '회 할인 금액', '쿠폰 할인 금액', '사용된 마일리지']
        for col in cols_to_numeric:
            df_godomall[col] = pd.to_numeric(df_godomall[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        df_godomall['수정될_금액_고도몰'] = (
            df_godomall['상품별 품목금액'] + df_godomall['총 배송 금액'] - df_godomall['회 할인 금액'] - 
            df_godomall['쿠폰 할인 금액'] - df_godomall['사용된 마일리지']
        )
        
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})[key_cols_smartstore + ['수정될_금액_스토어']].drop_duplicates(subset=key_cols_smartstore, keep='first')
        
        # <<-- 최종 수정: 고도몰 금액 보정을 위한 연결고리(Key) 변경 -->>
        key_cols_godomall = ['수취인 이름', '상품수량', '상품별 품목금액']
        godomall_prices_for_merge = df_godomall[key_cols_godomall + ['수정될_금액_고도몰']].rename(columns={'수취인 이름': '수령자명', '상품수량': '주문수량', '상품별 품목금액': '실결제금액'})
        godomall_prices_for_merge = godomall_prices_for_merge.drop_duplicates(subset=['수령자명', '주문수량', '실결제금액'], keep='first')
        
        # 데이터 병합 전, 키로 사용될 열들의 데이터 타입을 통일 (공백 제거 포함)
        df_final['수령자명'] = df_final['수령자명'].astype(str).str.strip()
        df_final['주문수량'] = pd.to_numeric(df_final['주문수량'], errors='coerce').fillna(0).astype(int)
        df_final['실결제금액'] = pd.to_numeric(df_final['실결제금액'], errors='coerce').fillna(0).astype(int)
        
        smartstore_prices['수령자명'] = smartstore_prices['수령자명'].astype(str).str.strip()
        smartstore_prices['주문수량'] = pd.to_numeric(smartstore_prices['주문수량'], errors='coerce').fillna(0).astype(int)
        
        godomall_prices_for_merge['수령자명'] = godomall_prices_for_merge['수령자명'].astype(str).str.strip()
        godomall_prices_for_merge['주문수량'] = pd.to_numeric(godomall_prices_for_merge['주문수량'], errors='coerce').fillna(0).astype(int)
        godomall_prices_for_merge['실결제금액'] = pd.to_numeric(godomall_prices_for_merge['실결제금액'], errors='coerce').fillna(0).astype(int)

        df_final = pd.merge(df_final, smartstore_prices, on=key_cols_smartstore, how='left')
        df_final = pd.merge(df_final, godomall_prices_for_merge, on=['수령자명', '주문수량', '실결제금액'], how='left')

        warnings = [f"- [금액보정 실패] **{row['쇼핑몰']}** / {row['수령자명']} / {row['SKU상품명']}" for _, row in df_final[(df_final['쇼핑몰'] == '스마트스토어') & (df_final['수정될_금액_스토어'].isna()) | (df_final['쇼핑몰'] == '고도몰5') & (df_final['수정될_금액_고도몰'].isna())].iterrows()]
        
        df_final['실결제금액'] = np.where(df_final['쇼핑몰'] == '고도몰5', df_final['수정될_금액_고도몰'].fillna(df_final['실결제금액']), df_final['실결제금액'])
        df_final['실결제금액'] = np.where(df_final['쇼핑몰'] == '스마트스토어', df_final['수정될_금액_스토어'].fillna(df_final['실결제금액']), df_final['실결제금액'])
        
        df_main_result = df_final[['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명', 'original_order']]
        
        homonym_warnings = []
        name_groups = df_main_result.groupby('수령자명')['original_order'].apply(list)
        for name, orders in name_groups.items():
            if len(orders) > 1 and (max(orders) - min(orders) + 1) != len(orders):
                homonym_warnings.append(f"- [동명이인 의심] **{name}** 님의 주문이 떨어져서 입력되었습니다. 확인이 필요합니다.")
        warnings.extend(homonym_warnings)

        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum().rename(columns={'주문수량': '개수'})
        df_packing_list = df_main_result.sort_values(by='original_order')[['SKU상품명', '주문수량', '수령자명', '쇼핑몰']].copy()
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        df_packing_list_final = df_packing_list.copy()
        df_packing_list_final['묶음번호'] = df_packing_list_final['묶음번호'].where(is_first_item, '')
        df_packing_list_final = df_packing_list_final[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]

        df_merged = pd.merge(df_main_result, df_master[['SKU코드', '과세여부', '입수량']], left_on='재고관리코드', right_on='SKU코드', how='left')
        
        unmastered = df_merged[df_merged['SKU코드'].isna()]
        for _, row in unmastered.iterrows():
            warnings.append(f"- [미등록 상품] **{row['재고관리코드']}** / {row['SKU상품명']}")

        client_map = {'쿠팡': '쿠팡 주식회사', '고도몰5': '고래미자사몰_현금영수증(고도몰)', '스마트스토어': '스토어팜'}
        
        df_ecount_upload = pd.DataFrame()
        
        df_ecount_upload['일자'] = datetime.now().strftime("%Y%m%d")
        df_ecount_upload['거래처명'] = df_merged['쇼핑몰'].map(client_map).fillna(df_merged['쇼핑몰'])
        df_ecount_upload['출하창고'] = '고래미'
        df_ecount_upload['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_ecount_upload['적요_전표'] = '오전/온라인'
        df_ecount_upload['품목코드'] = df_merged['재고관리코드']
        
        is_box_order = df_merged['SKU상품명'].str.contains("BOX", na=False)
        입수량 = pd.to_numeric(df_merged['입수량'], errors='coerce').fillna(1)
        base_quantity = np.where(is_box_order, df_merged['주문수량'] * 입수량, df_merged['주문수량'])
        is_3_pack = df_merged['SKU상품명'].str.contains("3개입|3개", na=False)
        final_quantity = np.where(is_3_pack, base_quantity * 3, base_quantity)
        df_ecount_upload['박스'] = np.where(is_box_order, df_merged['주문수량'], np.nan)
        df_ecount_upload['수량'] = final_quantity.astype(int)
        
        df_merged['실결제금액'] = pd.to_numeric(df_merged['실결제금액'], errors='coerce').fillna(0)
        공급가액 = np.where(df_merged['과세여부'] == '과세', df_merged['실결제금액'] / 1.1, df_merged['실결제금액'])
        df_ecount_upload['공급가액'] = 공급가액
        df_ecount_upload['부가세'] = df_merged['실결제금액'] - df_ecount_upload['공급가액']
        
        df_ecount_upload['쇼핑몰고객명'] = df_merged['수령자명']
        df_ecount_upload['original_order'] = df_merged['original_order']
        
        ecount_columns = [
            '일자', '순번', '거래처코드', '거래처명', '담당자', '출하창고', '거래유형', '통화', '환율', 
            '적요_전표', '미수금', '총합계', '연결전표', '품목코드', '품목명', '규격', '박스', '수량', 
            '단가', '외화금액', '공급가액', '부가세', '적요_품목', '생산전표생성', '시리얼/로트', 
            '관리항목', '쇼핑몰고객명', 'original_order'
        ]
        for col in ecount_columns:
            if col not in df_ecount_upload:
                df_ecount_upload[col] = ''
        
        for col in ['공급가액', '부가세']:
            df_ecount_upload[col] = df_ecount_upload[col].round().astype('Int64')
        
        # 정렬 직전 '거래유형'을 숫자 타입으로 강제 변환
        df_ecount_upload['거래유형'] = pd.to_numeric(df_ecount_upload['거래유형'])
        
        sort_order = ['고래미자사몰_현금영수증(고도몰)', '스토어팜', '쿠팡 주식회사']
        df_ecount_upload['거래처명_sort'] = pd.Categorical(df_ecount_upload['거래처명'], categories=sort_order, ordered=True)
        df_ecount_upload = df_ecount_upload.sort_values(
            by=['거래처명_sort', '거래유형', 'original_order'],
            ascending=[True, True, True]
        ).drop(columns=['거래처명_sort', 'original_order'])
        
        df_ecount_upload = df_ecount_upload[ecount_columns[:-1]]

        return df_main_result.drop(columns=['original_order']), df_quantity_summary, df_packing_list_final, df_ecount_upload, True, "모든 파일 처리가 성공적으로 완료되었습니다.", warnings

    except Exception as e:
        import traceback
        st.error(f"처리 중 심각한 오류가 발생했습니다: {e}")
        st.error(traceback.format_exc())
        return None, None, None, None, False, f"오류가 발생했습니다. 파일을 다시 확인하거나 관리자에게 문의하세요.", []
