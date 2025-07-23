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
    # NaN 값을 빈 문자열로 바꿔서 저장
    df_to_save = df.fillna('')
    df_to_save.to_excel(output, index=False, sheet_name='Sheet1')
    
    workbook = openpyxl.load_workbook(output)
    sheet = workbook.active

    # 1. 셀 너비 자동 조절
    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if cell.value:
                    # 셀 내용의 길이를 측정
                    cell_text = str(cell.value)
                    # 현재 셀의 길이가 기존 최대 길이보다 길면 업데이트
                    if len(cell_text) > max_length:
                        max_length = len(cell_text)
            except:
                pass
        # 계산된 최대 길이에 약간의 여유를 주어 너비 설정
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column].width = adjusted_width

    # 2. 포장 리스트 고급 서식 적용
    if format_type == 'packing_list':
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        odd_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        
        # 모든 셀에 기본 테두리 적용
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
            for cell in row:
                cell.border = thin_border
        
        bundle_start_row = 2
        for row_num in range(2, sheet.max_row + 2):
            current_bundle_cell = sheet.cell(row=row_num, column=1)
            
            # 묶음번호가 있거나 마지막 행에 도달하면 이전 그룹에 서식 적용
            if (current_bundle_cell.value) or (row_num > sheet.max_row):
                if row_num > 2:
                    bundle_end_row = row_num - 1
                    
                    # 이전 그룹의 묶음번호 (문자열로 처리하여 오류 방지)
                    prev_bundle_num_str = str(sheet.cell(row=bundle_start_row, column=1).value)
                    
                    # 묶음번호가 숫자일 경우에만 홀/짝 배경색 적용
                    if prev_bundle_num_str.isdigit():
                        prev_bundle_num = int(prev_bundle_num_str)
                        if prev_bundle_num % 2 != 0: # 홀수 묶음번호 그룹
                            for r in range(bundle_start_row, bundle_end_row + 1):
                                for c in range(1, sheet.max_column + 1):
                                    sheet.cell(row=r, column=c).fill = odd_fill
                    
                    # 셀 병합 (묶음이 2줄 이상일 경우)
                    if bundle_start_row < bundle_end_row:
                        sheet.merge_cells(start_row=bundle_start_row, start_column=1, end_row=bundle_end_row, end_column=1)
                        # 병합된 셀은 수직/수평 중앙 정렬
                        merged_cell = sheet.cell(row=bundle_start_row, column=1)
                        merged_cell.alignment = Alignment(vertical='center', horizontal='center')
                
                bundle_start_row = row_num
    
    final_output = io.BytesIO()
    workbook.save(final_output)
    final_output.seek(0)
    
    return final_output.getvalue()


def process_all_files(file1, file2, file3):
    """3개의 파일을 받아 4종류의 최종 결과물을 생성하는 메인 함수"""
    try:
        # <<-- 핵심 변경점: 상품 마스터 정보를 코드 내에 데이터로 직접 내장 -->>
        master_data_string = """SKU코드,SKU상품명,과세여부,입수량
G604E,[BOX] 가쓰오 슈토우 1kg,과세,6
S011E,[BOX] 고래미 가라아게파우더2kg,과세,
G510E,[BOX] 고래미 가쯔오참치다다끼1kg,면세,8
G028E,[BOX] 고래미 가쯔오참치다다끼2kg,면세,
G024E,[BOX] 고래미 간장대하장 [300g],과세,
G090E,[BOX] 고래미 간장새우비빔장 500g,과세,12
G035E,[BOX] 고래미 게가라아게500g,과세,
G686E,[BOX] 고래미 고노와다 500g,면세,12
G034E,[BOX] 고래미 깡새우가라아게1kg,과세,
G656E,[BOX] 고래미 날치알 스칼렛 300g,과세,20
G097E,[BOX] 고래미 마라해물모둠 920g x 10EA,과세,
G521E,[BOX] 고래미 모즈쿠900g (최신),과세,8
G037E,[BOX] 고래미 모찌리도후 파우더1kg,과세,
G971E,[BOX] 고래미 보타르가 150g + 미니 강판 세트,과세,6
G103E,[BOX] 고래미 불맛 쭈꾸미볶음 340g,과세,
G523E,[BOX] 고래미 불맛 쭈꾸미볶음 800g (400g x 2ea),과세,10
G033E,[BOX] 고래미 새우살가라아게1kg,과세,
G112E,[BOX] 고래미 숭어어란 슬라이스 20g x 10ea,과세,10
G952E,[BOX] 고래미 숭어어란 카라스미 1kg,과세,5
G137E,[BOX] 고래미 숭어어란 파우더 60g,과세,9
G610E,[BOX] 고래미 시메니싱 L 1kg,과세,6
G609E,[BOX] 고래미 시메니싱 M 1kg,과세,6
G646E,[BOX] 고래미 시메사바 L 150g,과세,30
G514E,[BOX] 고래미 시메사바 M 120g,과세,30
G649E,[BOX] 고래미 시메사바 S 100g,과세,30
G645E,[BOX] 고래미 시메사바 XS 70g,과세,30
G515E,[BOX] 고래미 시메사바720g,과세,
G516E,[BOX] 고래미 시메사바750g,과세,
G513E,[BOX] 고래미 시메사바S 700g x 8개입,과세,
G527E,[BOX] 고래미 양념꼬막장1kg,과세,6
G538E,[BOX] 고래미 어니언 타코와사비 1kg,과세,6
G528E,[BOX] 고래미 연어타다끼 1kg,과세,8
G059E,[BOX] 고래미 염장모즈쿠750g,면세,
G135E,[BOX] 고래미 유자폰즈소스 600g,과세,
G098E,[BOX] 고래미 육즙 가득 메로 구이 450g x 10EA,과세,
G099E,[BOX] 고래미 육즙 가득 메로 머리 구이 750g x 10EA,과세,
G100E,[BOX] 고래미 육즙 가득 연어 구이 750g x 10EA,과세,
D625E,[BOX] 고래미 육즙 가득 자숙 통오징어 M 500g x 10EA,과세,
G031E,[BOX] 고래미 이까게소가라아게1kg,과세,
G529E,[BOX] 고래미 자숙문어다리300g,면세,30
G023E,[BOX] 고래미 장어구이 [150g],과세,
G030E,[BOX] 고래미 참치알조림1kg,과세,
G787E,[BOX] 고래미 청어알 골드 160g (80g x 2ea) MSC 인증,과세,24
G532E,[BOX] 고래미 청어알 골드 300g MSC 인증,과세,20
G786E,[BOX] 고래미 청어알 그린 160g (80g x 2ea) MSC 인증,과세,24
G531E,[BOX] 고래미 청어알 그린 300g MSC 인증,과세,20
G788E,[BOX] 고래미 청어알 스칼렛 160g (80g x 2ea) MSC 인증,과세,24
G533E,[BOX] 고래미 청어알 스칼렛 300g MSC 인증,과세,20
D695E,[BOX] 고래미 카즈노코니싱 500g,과세,12
G094E,[BOX] 고래미 코모치 야리이까 750g,과세,
G032E,[BOX] 고래미 타코가라아게1kg,과세,
D945E,[BOX] 고래미 타코와사비 250g,과세,20
G539E,[BOX] 고래미 타코와사비 마일드 1kg,과세,6
G677E,[BOX] 고래미 타코와사비 프로 1kg,과세,6
D631E,[BOX] 고래미 프로펠러 클램 와사비 1kg,과세,6
D630E,[BOX] 고래미 프로펠러 클램 와사비 500g,과세,12
G565E,[BOX] 고래미 홍게살 크림 파스타 소스 1kg,과세,6
G608E,[BOX] 고래조개 와사비 1kg,과세,6
D603E,[BOX] 낙지연포탕(R) 600g x 12ea,과세,12
G605E,[BOX] 마구로 슈토우 1kg,과세,6
D615E,[BOX] 마라룽샤(R) 580g x 12ea,과세,12
D616E,[BOX] 마라새우(R) 580g x 12ea,과세,12
G522E,[BOX] 매콤칼칼 바지락술찜 500g (냉장),과세,10
D051E,[BOX] 모둠해물탕(R) 750g x 12ea,과세,12
D646E,[BOX] 모즈쿠 80g (냉장),과세,30
D055E,[BOX] 민어매운탕(R) 680g x 12ea,과세,12
D102E,[BOX] 설래담 간장새우비빔장 320g(80g x 4ea)(냉동),과세,8
D601E,[BOX] 설래담 간장새우비빔장 400g x 12개입,과세,
D013E,[BOX] 설래담 간장새우비빔장 80g X 30개입,과세,30
D008E,[BOX] 설래담 간장새우장 700g x 12개입,과세,
D739E,[BOX] 설래담 고노와다 320g(80g x 4ea)(냉동),면세,8
D602E,[BOX] 설래담 고노와다 400g X 12개입,면세,
D018E,[BOX] 설래담 고노와다 80g X 30개입,면세,30
D603E,[BOX] 설래담 낙지연포탕 600g x 12ea,과세,
D603E,[BOX] 설래담 낙지연포탕 680g,면세,12
D099E,[BOX] 설래담 더진한 붉은대게장 320g(80g x 4ea)(냉동),과세,8
D612E,[BOX] 설래담 더진한 붉은대게장 400g x 12개입,과세,
D014E,[BOX] 설래담 더진한 붉은대게장 80gX30개입,과세,30
D615E,[BOX] 설래담 마라룽샤 580g x 12ea,과세,
D010E,[BOX] 설래담 마라새우 580g,면세,
D616E,[BOX] 설래담 마라새우 580g x 12ea,과세,
D617E,[BOX] 설래담 매콤 칠게젓 400g x 12개입,과세,
D016E,[BOX] 설래담 매콤 칠게젓 80g x 30개입,과세,
D618E,[BOX] 설래담 모둠해물탕 750g x 12ea,과세,
D033E,[BOX] 설래담 민어 매운탕 800g,면세,
D619E,[BOX] 설래담 민어매운탕 680g x 12ea,과세,
D620E,[BOX] 설래담 바지락술찜 900g (450g x 2봉) x 12ea,과세,12
D621E,[BOX] 설래담 붉은대게딱지장 320g(80g x 4ea)(냉동),과세,8
D622E,[BOX] 설래담 붉은대게딱지장 400g x 12개입,과세,
D611E,[BOX] 설래담 순날치알골드 400g x 12개입,과세,12
D879E,[BOX] 설래담 양념 해물장 3종 480g,과세,
D624E,[BOX] 설래담 양념꽃게무침 650g,면세,12
D624E,[BOX] 설래담 양념꽃게무침 650g x 12ea,과세,
D085E,[BOX] 설래담 오독오독 조개살와사비 80g (냉동),과세,
D101E,[BOX] 설래담 타코와사비 320g(80g x 4ea)(냉동),과세,8
D001B,[BOX] 설래담 타코와사비 400g x 12개입,과세,
G039E,[BOX] 설래담 타코와사비 80gx30개입,과세,30
D632E,[BOX] 설래담 해물알탕 660g x 12ea,과세,
D632E,[BOX] 설래담 해물알탕 660g x 12ea,면세,12
G612E,[BOX] 시메사바 콤보 2kg,과세,5
D624E,[BOX] 양념꽃게무침(R) 650g x 12ea,과세,12
G529E,[BOX] 자숙문어다리 300g,과세,30
G629E,[BOX] 자숙문어다리 300g_TRITON,과세,
G508E,[BOX] 전복게우장 400g,과세,12
G642E,[BOX] 카네코메 연어알 500g,과세,
G132E,[BOX] 코모치 야리이까 알을 품은 한치 600g(300gx2ea),과세,15
G548E,[BOX] 크래비넌 매콤크랩 300g (30gx10ea),과세,16
G547E,[BOX] 크래비넌 오리지널 300g (30gx10ea),과세,16
G728E,[BOX] 키자미 와사비 250g,과세,20
D626E,[BOX] 타코와사비_1kg,과세,6
G540E,[BOX] 타코와사비_300g,과세,20
G541E,[BOX] 타코와사비_500g,과세,12
G643E,[BOX] 토와수산 연어알,과세,24
D052E,[BOX] 해물알탕(R) 660g x 12ea,과세,
G641E,[BOX] 혼마수산 연어알 500g,과세,
G654E,[BOX][봉초밥 키트 증정]  국내산 시메사바 230g,과세,
G651E,[BOX][봉초밥 키트 증정] 국내산 시메사바 140g,과세,
G652E,[BOX][봉초밥 키트 증정] 국내산 시메사바 170g,과세,
G653E,[BOX][봉초밥 키트 증정] 국내산 시메사바 200g,과세,
G655E,[BOX][봉초밥 키트 증정] 국내산 시메사바 260g,과세,
G658E,[BOX][봉초밥 키트 증정] 국내산 시메사바 290g,과세,
G501E,[BOX]가니미소1kg,과세,6
G502E,[BOX]가니미소400g,과세,12
G606E,[BOX]가쓰오 슈토우 300g,과세,20
G035E,[BOX]게가라아게500g_copy1,과세,
G503E,[BOX]고래미 가니미소 마일드 1kg,과세,6
G504E,[BOX]고래미 가니미소 마일드 400g,과세,12
D001E,[BOX]고래미 소라와사비 1kg,과세,6
G519E,[BOX]날치알[골드]300g_copy1,과세,20
G520E,[BOX]날치알[골드]800g_copy1,과세,6
G517E,[BOX]날치알[오렌지]300g,과세,20
G518E,[BOX]날치알[오렌지]800g_copy1,과세,6
G607E,[BOX]마구로 슈토우 300g,과세,20
G511E,[BOX]멍게고노와다300g,과세,20
G025E,[BOX]모즈쿠300g,과세,
G025E,[BOX]모즈쿠300g_copy1,과세,
G022E,[BOX]바지락술찜450g_copy1,과세,
G525E,[BOX]붉은대게딱지장1kg,과세,6
G524E,[BOX]붉은대게딱지장400g_copy1,과세,12
D013E,[BOX]설래담 간장새우비빔장80g_copy1,과세,30
D018E,[BOX]설래담 고노와다 80g,면세,30
D014E,[BOX]설래담 더진한붉은대게장80g_copy1,과세,30
D011E,[BOX]설래담 마라롱샤 580g,면세,
D016E,[BOX]설래담 매콤칠게젓80g_copy1,과세,
D015E,[BOX]설래담 붉은대게딱지장80g_copy1,과세,30
D017E,[BOX]설래담 순날치알(골드)80g_copy1,과세,30
D017E,[BOX]설래담 순날치알골드 80g x 30개입,과세,30
SET001,[BOX]설래담 울진 프리미엄 붉은대게 비빔밥 키트 162.8g,과세,
G039E,[BOX]설래담 타코와사비80g_copy1,과세,30
G967E,[BOX]소라와사비500g,과세,12
G059E,[BOX]염장모즈쿠750g,면세,
G535E,[BOX]큐브마요[오렌지]1kg_copy1,과세,
G534E,[BOX]큐브마요[오렌지]300_copy1,과세,
G537E,[BOX]큐브마요[와사비]1kg_copy1,과세,
G536E,[BOX]큐브마요[와사비]300g_copy1,과세,
T006E,[BOX]택미홈 굴짬뽕750g_copy1,면세,
T007E,[BOX]택미홈 나가사끼짬뽕640g_copy1,면세,
T002E,[BOX]택미홈 더진한붉은대게장60g_copy1,과세,
T004E,[BOX]택미홈 모둠해물탕750g_copy1,면세,
T003E,[BOX]택미홈 붉은대게딱지장80g,과세,
T005E,[BOX]택미홈 알탕600g_copy1,면세,
G010E,[BOX]호타테와사비1kg,과세,
G011E,[BOX]호타테와사비500g,과세,
G036E,{BOX} 고래미 장어구이 [500g],과세,
D619E,{BOX} 설래담 민어매운탕 680g x 12ea,면세,12
D015E,{BOX} 설래담 붉은대게딱지장 80gX30개입,과세,30
G651E,[봉초밥 키트 증정] 국내산 시메사바 140g,과세,
G652E,[봉초밥 키트 증정] 국내산 시메사바 170g,과세,
G653E,[봉초밥 키트 증정] 국내산 시메사바 200g,과세,
G654E,[봉초밥 키트 증정] 국내산 시메사바 230g,과세,
G655E,[봉초밥 키트 증정] 국내산 시메사바 260g,과세,
G658E,[봉초밥 키트 증정] 국내산 시메사바 290g,과세,
S103E,[추가구성] 미니락교 20g,과세,1
S105E,[추가구성] 봉초밥키트,과세,1
S106E,[추가구성] 사시미키트,과세,
G726E,[추가구성] 유자폰즈소스 30g,과세,
S102E,[추가구성] 초대리 50g,과세,1
G604E,가쓰오 슈토우 1kg,과세,6
G606E,가쓰오 슈토우 300g,과세,20
G501E,고래미 가니미소 [1kg],과세,6
G502E,고래미 가니미소 [400g],과세,12
G503E,고래미 가니미소 마일드 1kg,과세,6
G504E,고래미 가니미소 마일드 400g,과세,12
S011E,고래미 가라아게파우더 [2kg],과세,
G510E,고래미 가쯔오참치다다끼 [1kg],면세,
G028E,고래미 가쯔오참치다다끼 [2kg],면세,
G024E,고래미 간장대하장 [300g],과세,
G090E,고래미 간장새우비빔장 500g,과세,12
G035E,고래미 게가라아게500g,과세,
G686E,고래미 고노와다 500g,면세,12
G123E,고래미 고래진미 선물세트,과세,
G034E,고래미 깡새우가라아게 [1kg],과세,
G656E,고래미 날치알 스칼렛 300g,과세,20
G097E,고래미 마라해물모둠 920g x 10EA,과세,
G025E,고래미 모즈쿠 [300g],과세,
G037E,고래미 모찌리도후 파우더 [1kg],과세,
G125E,고래미 미식 선물세트,과세,
G971E,고래미 보타르가 150g + 미니 강판 세트,과세,6
G103E,고래미 불맛 쭈꾸미볶음 340g,과세,
G523E,고래미 불맛 쭈꾸미볶음 800g (400g x 2ea),과세,10
G525E,고래미 붉은대게딱지장 [1kg],과세,6
G524E,고래미 붉은대게딱지장 [400g],과세,12
G033E,고래미 새우살가라아게 1kg,과세,
G967E,고래미 소라와사비 [500g],과세,12
D001E,고래미 소라와사비 1kg,과세,6
G112E,고래미 숭어어란 슬라이스 20g,과세,10
G952E,고래미 숭어어란 카라스미 1kg,과세,5
G137E,고래미 숭어어란 파우더 60g,과세,9
G124E,고래미 스몰미식 선물세트,과세,
G610E,고래미 시메니싱 L 1kg,과세,6
G609E,고래미 시메니싱 M 1kg,과세,6
G646E,고래미 시메사바 L 150g,과세,30
G514E,고래미 시메사바 M 120g,과세,30
G649E,고래미 시메사바 S 100g,과세,30
G645E,고래미 시메사바 XS 70g,과세,30
O716E,고래미 시메사바 실속형 (오늘회리턴) 1200g (120g x 10ea),과세,
G515E,고래미 시메사바720g,과세,
G516E,고래미 시메사바750g,과세,
G513E,고래미 시메사바S [700g],과세,
G527E,고래미 양념꼬막장1kg,과세,6
G538E,고래미 어니언 타코와사비 1kg,과세,6
G138E,고래미 연어알 250g,과세,
G528E,고래미 연어타다끼 1kg,과세,8
G726E,고래미 유자폰즈소스 1.5kg (30g x 50ea),과세,
G135E,고래미 유자폰즈소스 600g,과세,
G098E,고래미 육즙 가득 메로 구이 450g x 10EA,과세,
G099E,고래미 육즙 가득 메로 머리 구이 750g x 10EA,과세,
G100E,고래미 육즙 가득 연어 구이 450g x 10EA,과세,
D625E,고래미 육즙 가득 자숙 통오징어 M 500g,과세,
G031E,고래미 이까게소가라아게 [1kg],과세,
G529E,고래미 자숙문어다리 [300g],면세,30
G023E,고래미 장어구이 [150g],과세,
G036E,고래미 장어구이 [500g],과세,
G030E,고래미 참치알조림 [1kg],과세,
G787E,고래미 청어알 골드 160g (80g x 2ea) MSC 인증,과세,24
G532E,고래미 청어알 골드 300g MSC 인증,과세,20
G786E,고래미 청어알 그린 160g (80g x 2ea) MSC 인증,과세,24
G531E,고래미 청어알 그린 300g MSC 인증,과세,20
G788E,고래미 청어알 스칼렛 160g (80g x 2ea) MSC 인증,과세,24
G533E,고래미 청어알 스칼렛 300g MSC 인증,과세,20
D695E,고래미 카즈노코니싱 500g,과세,12
G094E,고래미 코모치 야리이까 750g,과세,
O133E,고래미 코모치 야리이까 실속형 (오늘회리턴) 600g (300g x 2ea),과세,
G613E,고래미 쿠치코 10g,과세,1
G032E,고래미 타코가라아게 1kg,과세,
D945E,고래미 타코와사비 250g,과세,20
G539E,고래미 타코와사비 마일드 1kg,과세,6
G677E,고래미 타코와사비 프로 1kg,과세,6
D631E,고래미 프로펠러 클램 와사비 1kg,과세,6
D630E,고래미 프로펠러 클램 와사비 500g,과세,12
G010E,고래미 호타테와사비 [1kg],과세,
G011E,고래미 호타테와사비 [500g],과세,
G565E,고래미 홍게살 크림 파스타 소스 1kg,과세,6
G608E,고래조개 와사비 1kg,과세,6
D603E,낙지연포탕(R) 600g,과세,
G519E,날치알[골드]300g,과세,20
G520E,날치알[골드]800g_copy1,과세,6
G517E,날치알[오렌지]300g,과세,20
G518E,날치알[오렌지]800g,과세,6
G605E,마구로 슈토우 1kg,과세,6
G607E,마구로 슈토우 300g,과세,20
D615E,마라룽샤(R) 580g x 12ea,과세,
D616E,마라새우(R) 580g,과세,
G522E,매콤칼칼 바지락술찜 500g (냉장),과세,10
G511E,멍게고노와다300g,과세,20
D618E,모둠해물탕(R) 750g,과세,
D646E,모즈쿠 80g (냉장),과세,30
D647E,모즈쿠(냉장) 80g 1+1 ((2입)),과세,
G521E,모즈쿠900g_(최신),과세,8
E011E,미니강판,과세,
D055E,민어매운탕(R) 680g,과세,
G644E,북해도산 연어알,과세,
D102E,설래담 간장새우비빔장 320g(80g x 4ea)(냉동),과세,8
D601E,설래담 간장새우비빔장 400g (80g x 5개입),과세,
D013E,설래담 간장새우비빔장 80g,과세,30
D013E,설래담 간장새우비빔장 80g x 3개입,과세,
D008E,설래담 간장새우장 700g,과세,
D739E,설래담 고노와다 320g(80g x 4ea)(냉동),면세,8
D602E,설래담 고노와다 400g (80g x 5개입),면세,
D018E,설래담 고노와다 80g,면세,30
D018E,설래담 고노와다 80gX3개입,면세,
D603E,설래담 낙지연포탕 600g,과세,
D603E,설래담 낙지연포탕 600g x 3개입 (1.8kg),과세,
D603E,설래담 낙지연포탕 680g,면세,12
D099E,설래담 더진한 붉은대게장 320g(80g x 4ea)(냉동),과세,8
D612E,설래담 더진한 붉은대게장 400g (80g x 5개입),과세,
D014E,설래담 더진한 붉은대게장 80g,과세,30
D014E,설래담 더진한 붉은대게장 80gx3개입,과세,
D011E,설래담 마라롱샤 580g,면세,
D615E,설래담 마라룽샤 580g,과세,
D010E,설래담 마라새우 580g,면세,
D616E,설래담 마라새우 580g,과세,
D617E,설래담 매콤 칠게젓 400g (80g x 5개입),과세,
D016E,설래담 매콤 칠게젓 80g,과세,
D016E,설래담 매콤 칠게젓 80g x 3개입,과세,
D618E,설래담 모둠해물탕 750g,과세,
D033E,설래담 민어 매운탕 800g,면세,
D048E,설래담 민어매운탕 680g,과세,
D619E,설래담 민어매운탕 680g x 12ea,면세,12
D620E,설래담 바지락술찜 900g (450g x 2ea ) x 3개입 (2.7kg),과세,
D620E,설래담 바지락술찜 900g (450g x 2봉),과세,12
D621E,설래담 붉은대게딱지장 320g(80g x 4ea)(냉동),과세,8
D622E,설래담 붉은대게딱지장 400g (80g x 5개입),과세,
D015E,설래담 붉은대게딱지장 80g,과세,30
D015E,설래담 붉은대게딱지장 80g x 3개입,과세,
D611E,설래담 순날치알 400g (80gx5개입),과세,12
D017E,설래담 순날치알골드 80g,과세,30
D017E,설래담 순날치알골드 80g x 3개입,과세,
D879E,설래담 양념 해물장 3종 480g,과세,
D624E,설래담 양념꽃게무침 650g,면세,12
D624E,설래담 양념꽃게무침 650g,과세,
D968E,설래담 오독오독 조개살 와사비80g 2개 (냉동),과세,
D968E,설래담 오독오독 조개살 와사비80g 3개 (냉동),과세,
D968E,설래담 오독오독 조개살와사비 80g (냉동),과세,30
SET001,설래담 울진 프리미엄 붉은대게 비빔밥 키트 162.8g,과세,
D101E,설래담 타코와사비 320g(80g x 4ea)(냉동),과세,8
D001B,설래담 타코와사비 400g (80g x 5개입),과세,
G039E,설래담 타코와사비 80g,과세,30
G039E,설래담 타코와사비 80g x 3개입,과세,
D632E,설래담 해물알탕 660g,과세,
D632E,설래담 해물알탕 660g x 12ea,면세,12
G612E,시메사바 콤보 2kg,과세,5
D624E,양념꽃게무침(R) 650g,과세,
G059E,염장모즈쿠750g_copy1,면세,
G629E,자숙문어다리 300g_TRITON,과세,
G508E,전복게우장 400g,과세,12
E006E,조미 사각 유부 (T-100),과세,
G642E,카네코메 연어알 500g,과세,
G132E,코모치 야리이까 알을 품은 한치 600g(300gx2ea),과세,15
G535E,큐브마요[오렌지]1kg_copy1,과세,
G534E,큐브마요[오렌지]300g,과세,
G537E,큐브마요[와사비]1kg,과세,
G536E,큐브마요[와사비]300g_copy1,과세,
G548E,크래비넌 매콤크랩 300g (30g10ea),과세,16
C557E,크래비넌 매콤크랩 30g,과세,
G547E,크래비넌 오리지널 300g (30gx10ea),과세,16
C556E,크래비넌 오리지널 30g,과세,
G728E,키자미 와사비 250g,과세,20
D626E,타코와사비_1kg,과세,6
G540E,타코와사비_300g,과세,20
G541E,타코와사비_500g,과세,12
G003E,타코와사비300g,과세,
T006E,택미홈 굴짬뽕750g,면세,
T007E,택미홈 나가사끼짬뽕640g,면세,
T002E,"택미홈 더진한붉은대게장60g,3개입_copy1",과세,
T004E,택미홈 모둠해물탕 [750g],면세,
G022E,택미홈 바지락술찜 [450g],과세,
T003E,"택미홈 붉은대게딱지장80g,3개입",과세,
T005E,택미홈 해물알탕 [600g],면세,
G643E,토와수산 연어알 500g,과세,
P001E,특가판매 소라와사비 1kg,과세,
D052E,해물알탕(R) 660g,과세,
G641E,혼마수산 연어알 500g,과세,
"""
        # 문자열 데이터를 파일처럼 읽어서 데이터프레임으로 변환
        df_master = pd.read_csv(io.StringIO(master_data_string))
        
        # 1. 사용자 업로드 파일 읽기
        df_smartstore = pd.read_excel(file1)
        df_ecount_orig = pd.read_excel(file2)
        df_godomall = pd.read_excel(file3)

        # 2. (기존 로직) 금액 보정하여 최종 주문 목록 생성
        df_final = df_ecount_orig.copy().rename(columns={'금액': '실결제금액'})
        # ... (이하 모든 데이터 처리 로직은 이전과 동일하게 유지됩니다) ...
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
        df_merged = pd.merge(df_main_result, df_master[['SKU코드', '과세여부', '입수량']], left_on='재고관리코드', right_on='SKU코드', how='left')
        
        unmastered = df_merged[df_merged['SKU코드'].isna()]
        for _, row in unmastered.iterrows():
            warnings.append(f"- [미등록 상품] 상품코드: **{row['재고관리코드']}**가 상품 마스터에 없습니다.")

        client_map = {'쿠팡': '쿠팡 주식회사', '고도몰5': '고래미자사몰_현금영수증(고도몰)', '스마트스토어': '스토어팜', '배민상회': '주식회사 우아한형제들(배민상회)', '이지웰': '주식회사 현대이지웰'}
        
        df_ecount_upload = pd.DataFrame()
        df_ecount_upload['일자'] = datetime.now().strftime("%Y%m%d")
        df_ecount_upload['거래처명'] = df_merged['쇼핑몰'].map(client_map).fillna(df_merged['쇼핑몰'])
        df_ecount_upload['출하창고'] = '고래미'
        df_ecount_upload['거래유형'] = np.where(df_merged['과세여부'] == '면세', 12, 11)
        df_ecount_upload['적요'] = '오전/온라인'
        df_ecount_upload['품목코드'] = df_merged['재고관리코드']
        
        is_box = df_merged['SKU상품명'].str.contains('BOX', na=False)
        df_ecount_upload['박스'] = np.where(is_box, df_merged['주문수량'], "")
        
        입수량 = pd.to_numeric(df_merged['입수량'], errors='coerce').fillna(1)
        df_ecount_upload['수량'] = np.where(is_box, df_merged['주문수량'] * 입수량, df_merged['주문수량']).astype(int)
        
        df_merged['실결제금액'] = pd.to_numeric(df_merged['실결제금액'], errors='coerce').fillna(0)
        공급가액 = np.where(df_merged['과세여부'] == '과세', df_merged['실결제금액'] / 1.1, df_merged['실결제금액'])
        df_ecount_upload['공급가액'] = 공급가액
        df_ecount_upload['부가세'] = df_merged['실결제금액'] - df_ecount_upload['공급가액']
        
        df_ecount_upload['쇼핑몰고객명'] = df_merged['수령자명']
        
        ecount_columns = ['일자', '순번', '거래처코드', '거래처명', '담당자', '출하창고', '거래유형', '통화', '환율', '적요', '미수금', '총합계', '연결전표', '품목코드', '품목명', '규격', '박스', '수량', '단가', '외화금액', '공급가액', '부가세', '적요', '생산전표생성', '시리얼/로트', '관리항목', '쇼핑몰고객명']
        for col in ecount_columns:
            if col not in df_ecount_upload:
                df_ecount_upload[col] = ''
        
        # 금액 관련 컬럼을 정수형으로 변환
        for col in ['공급가액', '부가세']:
            df_ecount_upload[col] = df_ecount_upload[col].round().astype('Int64')

        df_ecount_upload = df_ecount_upload[ecount_columns]

        return df_main_result, df_quantity_summary, df_packing_list_final, df_ecount_upload, True, "모든 파일 처리가 성공적으로 완료되었습니다.", warnings

    except Exception as e:
        import traceback
        st.error(f"오류 발생: {e}")
        st.error(traceback.format_exc())
        return None, None, None, None, False, f"오류가 발생했습니다: {e}. 업로드한 파일 또는 내부 로직을 확인해주세요.", []


# --------------------------------------------------------------------------
# Streamlit 앱 UI 구성
# --------------------------------------------------------------------------
st.set_page_config(page_title="주문 처리 자동화 v.Ultimate", layout="wide")
st.title("📑 주문 처리 자동화 (v.Ultimate)")
st.info("💡 3개의 주문 관련 파일을 업로드하면, 금액 보정, 물류, ERP(이카운트)용 데이터가 한 번에 생성됩니다.")
st.write("---")

st.header("1. 원본 엑셀 파일 3개 업로드")
col1, col2, col3 = st.columns(3)
with col1:
    file1 = st.file_uploader("1️⃣ 스마트스토어 (금액확인용)", type=['xlsx', 'xls', 'csv'])
with col2:
    file2 = st.file_uploader("2️⃣ 이카운트 다운로드 (주문목록)", type=['xlsx', 'xls', 'csv'])
with col3:
    file3 = st.file_uploader("3️⃣ 고도몰 (금액확인용)", type=['xlsx', 'xls', 'csv'])

st.write("---")
st.header("2. 처리 결과 확인 및 다운로드")
if st.button("🚀 모든 데이터 처리 및 파일 생성 실행"):
    if file1 and file2 and file3:
        with st.spinner('모든 파일을 읽고 데이터를 처리하며 엑셀 서식을 적용 중입니다...'):
            df_main, df_qty, df_pack, df_ecount, success, message, warnings = process_all_files(file1, file2, file3)
        
        if success:
            st.success(message)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            if warnings:
                st.warning("⚠️ 확인 필요 항목")
                with st.expander("자세한 목록 보기..."):
                    st.info("금액 보정 실패 또는 미등록 상품 등의 데이터입니다. 원본 파일을 확인해주세요.")
                    for warning_message in warnings: st.markdown(warning_message)
            
            tab_erp, tab_pack, tab_qty, tab_main = st.tabs(["🏢 **이카운트 업로드용**", "📋 포장 리스트", "📦 출고수량 요약", "✅ 최종 보정 리스트"])
            
            with tab_erp:
                st.dataframe(df_ecount.astype(str))
                st.download_button("📥 다운로드", to_excel_formatted(df_ecount), f"이카운트_업로드용_{timestamp}.xlsx")

            with tab_pack:
                st.dataframe(df_pack)
                st.download_button("📥 다운로드", to_excel_formatted(df_pack, format_type='packing_list'), f"물류팀_전달용_포장리스트_{timestamp}.xlsx")

            with tab_qty:
                st.dataframe(df_qty)
                st.download_button("📥 다운로드", to_excel_formatted(df_qty), f"물류팀_전달용_출고수량_{timestamp}.xlsx")
            
            with tab_main:
                st.dataframe(df_main)
                st.download_button("📥 다운로드", to_excel_formatted(df_main), f"최종_실결제금액_보정완료_{timestamp}.xlsx")

        else:
            st.error(message)
    else:
        st.warning("⚠️ 3개의 엑셀 파일을 모두 업로드해야 실행할 수 있습니다.")
