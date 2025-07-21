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
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='최종_병합_데이터', index=False)
    processed_data = output.getvalue()
    return processed_data

def process_files(file1, file2, file3):
    """세 개의 파일을 받아 데이터를 처리하고 최종 데이터프레임을 반환하는 함수"""
    try:
        # 1. 파일 읽기
        df_smartstore = pd.read_excel(file1)
        df_ecount = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        # --- 데이터 전처리 및 보정 로직 ---

        # 2. 기준 데이터프레임 준비 (이카운트 파일)
        df_final = df_ecount.copy()
        df_final = df_final.rename(columns={'금액': '실결제금액'})

        # 3. 보정할 금액 데이터 준비
        
        # 3-1. 스마트스토어 보정 데이터 준비
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})
        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = smartstore_prices[key_cols_smartstore + ['수정될_금액_스토어']]
        # <<-- 수정된 부분: 병합 전 스마트스토어 조회 데이터에서 중복 제거
        smartstore_prices = smartstore_prices.drop_duplicates(subset=key_cols_smartstore, keep='first')

        # 3-2. 고도몰5 보정 데이터 준비
        godomall_prices = df_godomall.copy()
        last_col_name = godomall_prices.columns[-1]
        godomall_prices['수정될_금액_고도몰'] = pd.to_numeric(godomall_prices[last_col_name].astype(str).str.replace(',', ''), errors='coerce')
        
        key_cols_godomall = ['수취인 이름', '상품수량', '상품별 품목금액']
        godomall_prices_for_merge = godomall_prices[key_cols_godomall + ['수정될_금액_고도몰']]
        godomall_prices_for_merge = godomall_prices_for_merge.rename(columns={
            '수취인 이름': '수령자명', '상품수량': '주문수량', '상품별 품목금액': '실결제금액'
        })
        # <<-- 수정된 부분: 병합 전 고도몰 조회 데이터에서 중복 제거
        key_cols_godomall_merge = ['수령자명', '주문수량', '실결제금액']
        godomall_prices_for_merge = godomall_prices_for_merge.drop_duplicates(subset=key_cols_godomall_merge, keep='first')


        # 4. 기준 데이터프레임에 보정용 금액 병합(Merge)
        df_final = pd.merge(df_final, smartstore_prices, on=key_cols_smartstore, how='left')
        df_final = pd.merge(df_final, godomall_prices_for_merge, on=key_cols_godomall_merge, how='left')

        # 5. 최종 '실결제금액' 결정
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
        
        # 6. 최종 결과물 정리
        final_columns = ['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명']
        df_result = df_final[final_columns]
        
        return df_result, True, "데이터 처리가 성공적으로 완료되었습니다."

    except Exception as e:
        return None, False, f"오류가 발생했습니다: {e}. 업로드한 파일의 형식이나 컬럼명을 확인해주세요."

# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성 (이하 코드는 이전과 동일)
# --------------------------------------------------------------------------

st.set_page_config(page_title="엑셀 금액 보정 및 병합", layout="wide")

st.title("📑 엑셀 데이터 병합 및 실결제금액 보정 프로그램")
st.write("---")
st.markdown("""
- **파일 1**: 스마트스토어 주문 건의 정확한 **실결제금액**이 포함된 파일
- **파일 2**: 전체 주문 내역의 기준이 되는 **이카운트 등록용** 파일
- **파일 3**: 고도몰 주문 건의 정확한 **실결제금액**을 계산하기 위한 파일
""")

st.header("1. 엑셀 파일 3개 업로드")
col1, col2, col3 = st.columns(3)
with col1:
    file1 = st.file_uploader("1️⃣ 스마트스토어 파일", type=['xlsx', 'xls'])
with col2:
    file2 = st.file_uploader("2️⃣ 이카운트 등록용 파일 (기준)", type=['xlsx', 'xls'])
with col3:
    file3 = st.file_uploader("3️⃣ 고도몰 확인용 파일", type=['xlsx', 'xls'])

st.write("---")

st.header("2. 데이터 처리 및 결과 다운로드")
if st.button("🚀 데이터 병합 및 금액 보정 실행"):
    if file1 and file2 and file3:
        with st.spinner('파일을 읽고 데이터를 처리 중입니다...'):
            df_result, success, message = process_files(file1, file2, file3)
        
        if success:
            st.success(message)
            st.subheader("✅ 처리 결과 미리보기 (상위 10건)")
            st.dataframe(df_result.head(10))

            st.subheader("📊 쇼핑몰별 주문 건수")
            st.bar_chart(df_result['쇼핑몰'].value_counts())
            
            excel_data = to_excel(df_result)
            st.download_button(
                label="📥 최종 엑셀 파일 다운로드",
                data=excel_data,
                file_name="최종_실결제금액_보정완료.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error(message)
    else:
        st.warning("⚠️ 3개의 엑셀 파일을 모두 업로드해주세요.")
