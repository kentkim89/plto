import streamlit as st
import pandas as pd
import io
import numpy as np

# --------------------------------------------------------------------------
# 함수 정의
# --------------------------------------------------------------------------

def to_excel(df):
    """데이터프레임을 엑셀 파일 형식의 BytesIO 객체로 변환하는 함수"""
    output = io.BytesIO()
    # 인덱스는 저장하지 않음 (index=False)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    processed_data = output.getvalue()
    return processed_data

def process_files(file1, file2, file3):
    """세 개의 파일을 받아 세 종류의 결과물(최종본, 출고수량, 포장리스트)을 생성하는 함수"""
    try:
        # --- 1. (기존 로직) 데이터 읽기 및 금액 보정 ---
        df_smartstore = pd.read_excel(file1)
        df_ecount = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        df_final = df_ecount.copy()
        df_final = df_final.rename(columns={'금액': '실결제금액'})

        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})[key_cols_smartstore + ['수정될_금액_스토어']]
        smartstore_prices = smartstore_prices.drop_duplicates(subset=key_cols_smartstore, keep='first')

        godomall_prices = df_godomall.copy()
        last_col_name = godomall_prices.columns[-1]
        godomall_prices['수정될_금액_고도몰'] = pd.to_numeric(godomall_prices[last_col_name].astype(str).str.replace(',', ''), errors='coerce')
        key_cols_godomall_orig = ['수취인 이름', '상품수량', '상품별 품목금액']
        godomall_prices_for_merge = godomall_prices[key_cols_godomall_orig + ['수정될_금액_고도몰']]
        godomall_prices_for_merge = godomall_prices_for_merge.rename(columns={
            '수취인 이름': '수령자명', '상품수량': '주문수량', '상품별 품목금액': '실결제금액'
        })
        key_cols_godomall_merge = ['수령자명', '주문수량', '실결제금액']
        godomall_prices_for_merge = godomall_prices_for_merge.drop_duplicates(subset=key_cols_godomall_merge, keep='first')

        df_final = pd.merge(df_final, smartstore_prices, on=key_cols_smartstore, how='left')
        df_final = pd.merge(df_final, godomall_prices_for_merge, on=key_cols_godomall_merge, how='left')

        warnings = []
        unmatched_smartstore = df_final[(df_final['쇼핑몰'] == '스마트스토어') & (df_final['수정될_금액_스토어'].isna())]
        for _, row in unmatched_smartstore.iterrows():
            warnings.append(f"- [스마트스토어] 수령자명: **{row['수령자명']}**, 상품명: {row['SKU상품명']} (수량: {row['주문수량']})")
        
        unmatched_godomall = df_final[(df_final['쇼핑몰'] == '고도몰5') & (df_final['수정될_금액_고도몰'].isna())]
        for _, row in unmatched_godomall.iterrows():
            warnings.append(f"- [고도몰5] 수령자명: **{row['수령자명']}**, 상품명: {row['SKU상품명']} (수량: {row['주문수량']})")

        df_final['실결제금액'] = np.where(
            (df_final['쇼핑몰'] == '고도몰5') & (df_final['수정될_금액_고도몰'].notna()),
            df_final['수정될_금액_고도몰'],
            df_final['실결제금액']
        )
        df_final['실결제금액'] = np.where(
            (df_final['쇼핑몰'] == '스마트스토어') & (df_final['수정될_금액_스토어'].notna()),
            df_final['수정될_금액_스토어'],
            df_final['실결제금액']
        )
        
        final_columns = ['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명']
        df_main_result = df_final[final_columns]

        # --- 2. (추가 기능) 물류팀용 파일 2종 생성 ---

        # 2-1. 출고수량 요약 파일 생성
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum()
        df_quantity_summary = df_quantity_summary.rename(columns={'주문수량': '개수'})

        # 2-2. 포장 리스트 파일 생성
        # 필요한 컬럼만 복사하여 작업
        df_packing_list = df_main_result[['SKU상품명', '주문수량', '수령자명', '쇼핑몰']].copy()
        # '수령자명'으로 정렬하여 동일 수령인의 주문이 연달아 오도록 함
        df_packing_list = df_packing_list.sort_values(by='수령자명', kind='mergesort', ignore_index=True)
        # 각 수령인의 첫 번째 주문인지 확인
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        # 첫 번째 주문일 때만 누적 합계를 이용해 묶음번호 부여
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        # 첫 번째 주문이 아닌 경우, 묶음번호를 공란으로 처리
        df_packing_list['묶음번호'] = df_packing_list['묶음번호'].where(is_first_item, None)
        # 묶음번호를 정수형 문자열로 변환 (NaN 값은 빈 문자열로)
        df_packing_list['묶음번호'] = df_packing_list['묶음번호'].astype('Int64').astype(str).replace('<NA>', '')
        # 최종 컬럼 순서 정리
        df_packing_list = df_packing_list[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]

        return df_main_result, df_quantity_summary, df_packing_list, True, "데이터 처리가 성공적으로 완료되었습니다.", warnings

    except Exception as e:
        return None, None, None, False, f"오류가 발생했습니다: {e}. 업로드한 파일의 형식이나 컬럼명을 확인해주세요.", []


# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성
# --------------------------------------------------------------------------

st.set_page_config(page_title="주문 데이터 처리 자동화", layout="wide")
st.title("📑 주문 데이터 처리 및 파일 생성 자동화")
st.write("---")

# --- 파일 업로더 ---
st.header("1. 원본 엑셀 파일 3개 업로드")
with st.expander("파일 업로드 섹션 보기/숨기기", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        file1 = st.file_uploader("1️⃣ 스마트스토어 파일", type=['xlsx', 'xls'])
    with col2:
        file2 = st.file_uploader("2️⃣ 이카운트 등록용 파일 (기준)", type=['xlsx', 'xls'])
    with col3:
        file3 = st.file_uploader("3️⃣ 고도몰 확인용 파일", type=['xlsx', 'xls'])

st.write("---")

# --- 데이터 처리 및 결과 표시 ---
st.header("2. 처리 결과 확인 및 다운로드")
if st.button("🚀 모든 데이터 처리 및 파일 생성 실행"):
    if file1 and file2 and file3:
        with st.spinner('파일을 읽고 모든 데이터를 처리 중입니다...'):
            df_main, df_qty, df_pack, success, message, warnings = process_files(file1, file2, file3)
        
        if success:
            st.success(message)
            
            if warnings:
                st.warning("⚠️ 데이터 불일치 알림")
                st.info("아래 목록의 데이터는 수령자명 등의 정보가 파일 간에 일치하지 않아 금액 보정이 실패했을 수 있습니다. 원본 파일을 확인해주세요.")
                with st.expander("자세한 목록 보기..."):
                    for warning_message in warnings:
                        st.markdown(warning_message)

            # --- 결과물을 탭으로 보여주기 ---
            tab1, tab2, tab3 = st.tabs(["✅ 최종 금액 보정 리스트", "📦 물류팀용: 출고수량 요약", "📋 물류팀용: 포장 리스트"])

            with tab1:
                st.subheader("최종 주문 데이터 (금액 보정 완료)")
                st.dataframe(df_main)
                st.download_button(
                    label="📥 '최종 금액 보정 리스트' 엑셀 다운로드",
                    data=to_excel(df_main),
                    file_name="최종_실결제금액_보정완료.xlsx"
                )

            with tab2:
                st.subheader("상품별 총 출고수량")
                st.dataframe(df_qty)
                st.download_button(
                    label="📥 '출고수량 요약' 엑셀 다운로드",
                    data=to_excel(df_qty),
                    file_name="물류팀_전달용_출고수량.xlsx"
                )

            with tab3:
                st.subheader("수령자별 묶음 포장 리스트")
                st.dataframe(df_pack)
                st.download_button(
                    label="📥 '포장 리스트' 엑셀 다운로드",
                    data=to_excel(df_pack),
                    file_name="물류팀_전달용_포장리스트.xlsx"
                )
        else:
            st.error(message)
    else:
        st.warning("⚠️ 3개의 엑셀 파일을 모두 업로드해야 실행할 수 있습니다.")
