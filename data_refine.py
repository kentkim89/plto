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
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column].width = adjusted_width

    # 2. 포장 리스트 고급 서식 적용
    if format_type == 'packing_list':
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        odd_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border
        
        bundle_start_row = 2
        for row_num in range(2, sheet.max_row + 2):
            current_bundle_cell = sheet.cell(row=row_num, column=1)
            
            if (current_bundle_cell.value and str(current_bundle_cell.value).isdigit()) or row_num > sheet.max_row:
                if row_num > 2:
                    bundle_end_row = row_num - 1
                    prev_bundle_num_str = str(sheet.cell(row=bundle_start_row, column=1).value)
                    
                    if prev_bundle_num_str.isdigit():
                        prev_bundle_num = int(prev_bundle_num_str)
                        if prev_bundle_num % 2 != 0:
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


def process_all_files(file1, file2, file3, file_master):
    """4개의 파일을 받아 4종류의 최종 결과물을 생성하는 메인 함수"""
    try:
        # 1. 파일 읽기
        df_smartstore = pd.read_excel(file1)
        df_ecount_orig = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)
        # 상품 마스터 파일은 헤더가 복잡하므로 4줄을 건너뛰고 읽음
        df_master = pd.read_excel(file_master, skiprows=4)

        # 2. (기존 로직) 금액 보정하여 최종 주문 목록 생성
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        
        # ... (기존 금액 보정 로직은 생략 없이 그대로 유지) ...
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
        df_main_result = df_final[['재고관리코드', 'SKU상품명', '주문수량', '실결제금액', '쇼핑몰', '수령자명']]
        
        # 3. (기존 로직) 물류팀용 파일 2종 생성
        df_quantity_summary = df_main_result.groupby('SKU상품명', as_index=False)['주문수량'].sum().rename(columns={'주문수량': '개수'})
        df_packing_list = df_main_result[['SKU상품명', '주문수량', '수령자명', '쇼핑몰']].copy()
        is_first_item = df_packing_list['수령자명'] != df_packing_list['수령자명'].shift(1)
        df_packing_list['묶음번호'] = is_first_item.cumsum()
        df_packing_list_final = df_packing_list.copy()
        df_packing_list_final['묶음번호'] = df_packing_list_final['묶음번호'].where(is_first_item, '')
        df_packing_list_final = df_packing_list_final[['묶음번호', 'SKU상품명', '주문수량', '수령자명', '쇼핑몰']]

        # 4. (신규 로직) 이카운트 업로드용 파일 생성
        # 상품 마스터와 최종 주문 목록을 병합
        df_merged = pd.merge(df_main_result, df_master[['SKU코드', '과세여부', '입수량']], left_on='재고관리코드', right_on='SKU코드', how='left')
        
        # 상품 마스터에 없는 코드 경고 추가
        unmastered = df_merged[df_merged['SKU코드'].isna()]
        for _, row in unmastered.iterrows():
            warnings.append(f"- [미등록 상품] 상품코드: **{row['재고관리코드']}**가 상품 마스터에 없습니다.")

        # 거래처명 매핑 딕셔너리
        client_map = {
            '쿠팡': '쿠팡 주식회사',
            '고도몰5': '고래미자사몰_현금영수증(고도몰)',
            '스마트스토어': '스토어팜',
            '배민상회': '주식회사 우아한형제들(배민상회)',
            '이지웰': '주식회사 현대이지웰'
        }
        
        # 이카운트 양식에 맞춰 데이터 생성
        df_ecount_upload = pd.DataFrame()
        df_ecount_upload['일자'] = datetime.now().strftime("%Y%m%d")
        df_ecount_upload['순번'] = ''
        df_ecount_upload['거래처코드'] = ''
        df_ecount_upload['거래처명'] = df_merged['쇼핑몰'].map(client_map).fillna(df_merged['쇼핑몰']) # 매핑되지 않으면 원본 쇼핑몰 이름 사용
        df_ecount_upload['담당자'] = ''
        df_ecount_upload['출하창고'] = '고래미'
        df_ecount_upload['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_ecount_upload['통화'] = ''
        df_ecount_upload['환율'] = ''
        df_ecount_upload['적요'] = '오전/온라인'
        df_ecount_upload['미수금'] = ''
        df_ecount_upload['총합계'] = ''
        df_ecount_upload['연결전표'] = ''
        df_ecount_upload['품목코드'] = df_merged['재고관리코드']
        df_ecount_upload['품목명'] = ''
        df_ecount_upload['규격'] = ''
        
        is_box = df_merged['SKU상품명'].str.contains('BOX', na=False)
        df_ecount_upload['박스'] = np.where(is_box, df_merged['주문수량'], np.nan)
        
        입수량 = df_merged['입수량'].fillna(1) # 입수량이 NaN이면 1로 처리
        df_ecount_upload['수량'] = np.where(is_box, df_merged['주문수량'] * 입수량, df_merged['주문수량']).astype(int)
        
        df_ecount_upload['단가'] = ''
        df_ecount_upload['외화금액'] = ''

        공급가액 = np.where(df_merged['과세여부'] == '과세', df_merged['실결제금액'] / 1.1, df_merged['실결제금액'])
        df_ecount_upload['공급가액'] = 공급가액.round().astype(int)
        df_ecount_upload['부가세'] = (df_merged['실결제금액'] - df_ecount_upload['공급가액']).round().astype(int)
        
        df_ecount_upload['적요'] = ''
        df_ecount_upload['생산전표생성'] = ''
        df_ecount_upload['시리얼/로트'] = ''
        df_ecount_upload['관리항목'] = ''
        df_ecount_upload['쇼핑몰고객명'] = df_merged['수령자명']

        return df_main_result, df_quantity_summary, df_packing_list_final, df_ecount_upload, True, "모든 파일 처리가 성공적으로 완료되었습니다.", warnings

    except Exception as e:
        return None, None, None, None, False, f"오류가 발생했습니다: {e}. 업로드한 파일의 형식이나 컬럼명을 확인해주세요.", []


# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성
# --------------------------------------------------------------------------
st.set_page_config(page_title="주문-물류-회계 자동화", layout="wide")
st.title("📑 주문-물류-회계(ERP) 데이터 생성 자동화")
st.write("---")

st.header("1. 원본 엑셀 파일 4개 업로드")
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("1️⃣ 스마트스토어 (금액확인용)", type=['xlsx', 'xls'])
    file2 = st.file_uploader("2️⃣ 이카운트 다운로드 (주문목록)", type=['xlsx', 'xls'])
with col2:
    file3 = st.file_uploader("3️⃣ 고도몰 (금액확인용)", type=['xlsx', 'xls'])
    file_master = st.file_uploader("4️⃣ 상품 마스터 (플레이오토 코드)", type=['xlsx', 'xls'])

st.write("---")
st.header("2. 처리 결과 확인 및 다운로드")
if st.button("🚀 모든 데이터 처리 및 파일 생성 실행"):
    if file1 and file2 and file3 and file_master:
        with st.spinner('모든 파일을 읽고 데이터를 처리하며 엑셀 서식을 적용 중입니다...'):
            df_main, df_qty, df_pack, df_ecount, success, message, warnings = process_all_files(file1, file2, file3, file_master)
        
        if success:
            st.success(message)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            if warnings:
                st.warning("⚠️ 확인 필요 항목")
                with st.expander("자세한 목록 보기..."):
                    st.info("금액 보정 실패 또는 미등록 상품 등의 데이터입니다. 원본 파일을 확인해주세요.")
                    for warning_message in warnings: st.markdown(warning_message)
            
            tab_main, tab_qty, tab_pack, tab_erp = st.tabs(["✅ 최종 보정 리스트", "📦 출고수량 요약", "📋 포장 리스트", "🏢 이카운트 업로드용"])

            with tab_main:
                st.dataframe(df_main)
                st.download_button("📥 다운로드", to_excel_formatted(df_main), f"최종_실결제금액_보정완료_{timestamp}.xlsx")

            with tab_qty:
                st.dataframe(df_qty)
                st.download_button("📥 다운로드", to_excel_formatted(df_qty), f"물류팀_전달용_출고수량_{timestamp}.xlsx")

            with tab_pack:
                st.dataframe(df_pack)
                st.download_button("📥 다운로드", to_excel_formatted(df_pack, format_type='packing_list'), f"물류팀_전달용_포장리스트_{timestamp}.xlsx")
            
            with tab_erp:
                st.dataframe(df_ecount)
                st.download_button("📥 다운로드", to_excel_formatted(df_ecount), f"이카운트_업로드용_{timestamp}.xlsx")
        else:
            st.error(message)
    else:
        st.warning("⚠️ 4개의 엑셀 파일을 모두 업로드해야 실행할 수 있습니다.")
