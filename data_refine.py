import streamlit as st
import pandas as pd
import io
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime

# --------------------------------------------------------------------------
# 함수 정의
# --------------------------------------------------------------------------

def to_excel_formatted(df, format_type=None):
    """데이터프레임을 서식이 적용된 엑셀 파일 형식의 BytesIO 객체로 변환하는 함수"""
    output = io.BytesIO()
    df.to_excel(output, index=False, sheet_name='Sheet1')
    
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    # 1. 셀 너비 자동 조절
    for column_cells in sheet.columns:
        # 헤더 길이를 기본으로 설정
        length = len(str(column_cells[0].value))
        # 데이터 내용 중 가장 긴 길이를 찾아 업데이트
        for cell in column_cells:
            if cell.value:
                # 줄바꿈 문자가 있는 경우 가장 긴 줄을 기준으로 계산
                cell_length = max(len(line) for line in str(cell.value).split('\n'))
                if cell_length > length:
                    length = cell_length
        # 여유를 주어 너비 설정
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = (length + 2) * 1.2

    # 2. 포장 리스트 고급 서식 적용
    if format_type == 'packing_list':
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        odd_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # 연한 회색
        
        # 전체 셀에 기본 테두리 적용
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border
        
        bundle_start_row = 2
        for row_num in range(2, sheet.max_row + 2):
            current_bundle_cell = sheet.cell(row=row_num, column=1)
            
            if current_bundle_cell.value or row_num == sheet.max_row + 1:
                if row_num > 2:
                    bundle_end_row = row_num - 1
                    prev_bundle_num_str = str(sheet.cell(row=bundle_start_row, column=1).value)

                    if prev_bundle_num_str.isdigit():
                        prev_bundle_num = int(prev_bundle_num_str)
                        if prev_bundle_num % 2 != 0: # 홀수 묶음번호에만 배경색 적용
                            for r in range(bundle_start_row, bundle_end_row + 1):
                                for c in range(1, sheet.max_column + 1):
                                    sheet.cell(row=r, column=c).fill = odd_fill
                    
                    if bundle_start_row != bundle_end_row:
                        sheet.merge_cells(start_row=bundle_start_row, start_column=1, end_row=bundle_end_row, end_column=1)
                        sheet.cell(row=bundle_start_row, column=1).alignment = Alignment(vertical='center', horizontal='center')
                
                bundle_start_row = row_num
    
    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)
    
    return final_output.getvalue()


def process_files(file1, file2, file3):
    """세 개의 파일을 받아 세 종류의 결과물(최종본, 출고수량, 포장리스트)을 생성하는 함수"""
    try:
        # 데이터 처리 로직 (이전과 동일)
        df_smartstore = pd.read_excel(file1)
        df_ecount = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        df_final = df_ecount.copy().rename(columns={'금액': '실결제금액'})
        df_final['original_order'] = range(len(df_final)) # 원본 순서 저장을 위한 인덱스 생성

        key_cols_smartstore = ['재고관리코드', '주문수량', '수령자명']
        smartstore_prices = df_smartstore.rename(columns={'실결제금액': '수정될_금액_스토어'})[key_cols_smartstore + ['수정될_금액_스토어']].drop_duplicates(subset=key_cols_smartstore, keep='first')
        
        godomall_prices = df_godomall.copy()
        last_col_name = godomall_prices.columns[-1]
        godomall_prices['수정될_금액_고도몰'] = pd.to_numeric(godomall_prices[last_col_name].astype(str).str.replace(',', ''), errors='coerce')
        key_cols_godomall_orig = ['수취인 이름', '상품수량', '상품별 품목금액']
        godomall_prices_for_merge = godomall_prices[key_cols_godomall_orig + ['수정될_금액_고도몰']].rename(columns={'수취인 이름': '수령자명', '상품수량': '주문수량', '상품별 품목금액': '실결제금액'}).drop_duplicates(subset=['수령자명', '주문수량', '실결제금액'], keep='first')
        
        df_final = pd.merge(df_final, smartstore_prices, on=key_cols_smartstore, how='left')
        df_final = pd.merge(df_final, godomall_prices_for_merge, on=['수령자명', '주문수량', '실결제금액'], how='left')

        warnings = [f"- [스마트스토어] 수령자명: **{row['수령자명']}**, 상품명: {row['SKU상품명']}" for _, row in df_final[(df_final['쇼핑몰'] == '스마트스토어') & (df_final['수정될_금액_스토어'].isna())].iterrows()]
        warnings.extend([f"- [고도몰5] 수령자명: **{row['수령자명']}**, 상품명: {row['SKU상품명']}" for _, row in df_final[(df_final['쇼핑몰'] == '고도몰5') & (df_final['수정될_금액_고도몰'].isna())].iterrows()])
        
        df_final['실결제금액'] = np.where(df_final['쇼핑몰'] == '고도몰5', df_final['수정될_금액_고도몰'].fillna(df_final['실결제금액']), df_final['실결제금액'])
        df_final['실결제금액'] = np.where(df_final['쇼핑몰'] == '스마트스토어', df_final['수정될_금액_스토어'].fillna(df_final['실결제금액']), df_final['실결제금액'])
        
        df_main_result = df_final.sort_values('original_order').drop(columns=['original_order'])[['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명']]
        
        # --- 물류팀용 파일 생성 로직 ---
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum().rename(columns={'주문수량': '개수'})
        
        # <<-- 수정된 부분: 인위적인 정렬(sort_values) 제거 -->>
        # 원본 데이터(df_main_result)의 순서를 그대로 사용하여 포장 리스트 생성
        df_packing_list = df_main_result[['SKU상품명', '주문수량', '수령자명', '쇼핑몰']].copy()
        
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        df_packing_list_final = df_packing_list.copy()
        df_packing_list_final['묶음번호'] = df_packing_list_final['묶음번호'].where(is_first_item, '')
        df_packing_list_final = df_packing_list_final[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]

        return df_main_result, df_quantity_summary, df_packing_list_final, True, "데이터 처리가 성공적으로 완료되었습니다.", warnings

    except Exception as e:
        return None, None, None, False, f"오류가 발생했습니다: {e}. 업로드한 파일의 형식이나 컬럼명을 확인해주세요.", []


# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성 (이전과 동일)
# --------------------------------------------------------------------------
st.set_page_config(page_title="주문 데이터 처리 자동화", layout="wide")
st.title("📑 주문 데이터 처리 및 파일 생성 자동화 (v.Final-2)")
st.write("---")

st.header("1. 원본 엑셀 파일 3개 업로드")
with st.expander("파일 업로드 섹션 보기/숨기기", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1: file1 = st.file_uploader("1️⃣ 스마트스토어 파일", type=['xlsx', 'xls'])
    with col2: file2 = st.file_uploader("2️⃣ 이카운트 등록용 파일 (기준)", type=['xlsx', 'xls'])
    with col3: file3 = st.file_uploader("3️⃣ 고도몰 확인용 파일", type=['xlsx', 'xls'])

st.write("---")
st.header("2. 처리 결과 확인 및 다운로드")
if st.button("🚀 모든 데이터 처리 및 파일 생성 실행"):
    if file1 and file2 and file3:
        with st.spinner('파일을 읽고 모든 데이터를 처리하며 엑셀 서식을 적용 중입니다...'):
            df_main, df_qty, df_pack, success, message, warnings = process_files(file1, file2, file3)
        
        if success:
            st.success(message)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            if warnings:
                st.warning("⚠️ 데이터 불일치 알림")
                with st.expander("금액 보정 실패 목록 보기..."):
                    st.info("아래 목록의 데이터는 수령자명 등의 정보가 파일 간에 일치하지 않아 금액 보정이 실패했을 수 있습니다.")
                    for warning_message in warnings: st.markdown(warning_message)

            tab1, tab2, tab3 = st.tabs(["✅ 최종 금액 보정 리스트", "📦 물류팀용: 출고수량 요약", "📋 물류팀용: 포장 리스트"])
            
            with tab1:
                st.subheader("최종 주문 데이터 (금액 보정 완료)")
                st.dataframe(df_main)
                st.download_button(
                    label="📥 '최종 금액 보정 리스트' 엑셀 다운로드",
                    data=to_excel_formatted(df_main),
                    file_name=f"최종_실결제금액_보정완료_{timestamp}.xlsx"
                )

            with tab2:
                st.subheader("상품별 총 출고수량")
                st.dataframe(df_qty)
                st.download_button(
                    label="📥 '출고수량 요약' 엑셀 다운로드",
                    data=to_excel_formatted(df_qty),
                    file_name=f"물류팀_전달용_출고수량_{timestamp}.xlsx"
                )

            with tab3:
                st.subheader("수령자별 묶음 포장 리스트")
                st.dataframe(df_pack)
                st.download_button(
                    label="📥 '포장 리스트' 엑셀 다운로드",
                    data=to_excel_formatted(df_pack, format_type='packing_list'),
                    file_name=f"물류팀_전달용_포장리스트_{timestamp}.xlsx"
                )
        else:
            st.error(message)
    else:
        st.warning("⚠️ 3개의 엑셀 파일을 모두 업로드해야 실행할 수 있습니다.")
