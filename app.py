import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (숫자 포맷 완벽 적용)")
st.write("EDI Raw Data 업로드 시, 모든 수치 데이터를 '순수 숫자'로 변환하여 엑셀에 저장합니다.")

# 💡 깃허브의 고정 서식 파일 (ME코드 매핑 사전용)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 💡 센터별 고정 발주/배송코드 매핑
CENTER_CODE_MAP = {
    '오산센터': 81030907, # 텍스트가 아닌 순수 숫자로 저장
    '김해센터': 81030908
}

def clean_center_name(name):
    """센터명에서 '상온', '센타', '센터' 등을 정제하여 반환"""
    name = str(name).strip()
    name = re.sub(r'상온센타|상온센터|센타', '센터', name)
    name = name.replace('센터센터', '센터')
    return name

def extract_numbers_only(val):
    """어떤 찌꺼기가 있어도 순수 숫자(정수)만 완벽하게 뽑아내는 함수"""
    s = re.sub(r'[^0-9]', '', str(val))
    return int(s) if s else 0

uploaded_file = st.file_uploader("📥 EDI Raw Data (라떼는.xlsx 등) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("데이터 정제 및 숫자 포맷 변환 중입니다... 🚀"):
        try:
            # 1. EDI Raw 데이터 파싱
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            df_edi = df_edi.dropna(how='all')
            
            parsed_list = []
            curr_center = ""
            curr_doc_no = 0 
            curr_delivery_date = ""
            
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            for i, row in df_edi.iterrows():
                r = [str(x).strip() for x in row.tolist()]
                
                # 'ORDERS' 행 추출
                if r[0] == 'ORDERS':
                    curr_doc_no = extract_numbers_only(r[1])
                    curr_center = clean_center_name(r[5])
                    raw_date = r[7] if len(r) > 7 else ""
                    curr_delivery_date = re.sub(r'[^0-9-]', '', raw_date) 
                    continue
                
                # 상품 행 추출 (880으로 시작하는지 확인)
                raw_barcode = r[1].replace('.0', '')
                if raw_barcode.startswith('880'):
                    barcode_num = extract_numbers_only(r[1]) # 순수 숫자 바코드로 변환
                    
                    qty = extract_numbers_only(r[6])
                    ipsu = extract_numbers_only(r[5])
                    if ipsu == 0: ipsu = 1
                    
                    unit_qty = qty * ipsu
                    
                    if unit_qty > 0:
                        price = extract_numbers_only(r[7])
                        
                        parsed_list.append({
                            '발주_fallback': curr_doc_no,
                            '센터': curr_center,
                            '납품일자': curr_delivery_date,
                            '바코드_num': barcode_num, # 맵핑용 순수 숫자
                            '품명': r[2],
                            'UNIT수량': unit_qty,
                            'UNIT단가': price
                        })
            
            if not parsed_list:
                st.warning("⚠️ 유효한 발주 내역이 없습니다.")
                st.stop()

            df_parsed = pd.DataFrame(parsed_list)

            # 2. ME코드 매핑 (서식 파일 1번 시트 활용)
            df_map_sheet = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            
            # 컬럼 이름이 유동적일 수 있으므로 동적 탐색
            panmae_col = [c for c in df_map_sheet.columns if '판매코드' in str(c) or '바코드' in str(c)]
            me_col = [c for c in df_map_sheet.columns if '상품코드' in str(c) or 'ME' in str(c).upper()]
            
            p_col = panmae_col[0] if panmae_col else df_map_sheet.columns[3]
            m_col = me_col[-1] if me_col else df_map_sheet.columns[13]

            mapping_dict = df_map_sheet[[p_col, m_col]].dropna()
            mapping_dict.columns = ['바코드', 'ME코드']
            
            # 매핑 시트의 바코드도 '순수 숫자'로 강제 변환하여 100% 매칭 보장
            mapping_dict['바코드_num'] = mapping_dict['바코드'].apply(extract_numbers_only)
            mapping_dict['ME코드'] = mapping_dict['ME코드'].astype(str).str.strip()
            mapping_dict = mapping_dict.drop_duplicates(subset=['바코드_num'])

            df_final = pd.merge(df_parsed, mapping_dict, on='바코드_num', how='left')
            df_final['ME코드'] = df_final['ME코드'].fillna(df_final['바코드_num'].astype(str))

            # 3. 센터 + ME코드 기준 중복 수량 합산
            agg_funcs = {
                '발주_fallback': 'first',
                '납품일자': 'first',
                '품명': 'first',
                'UNIT단가': 'first',
                'UNIT수량': 'sum'
            }
            df_grouped = df_final.groupby(['센터', 'ME코드'], as_index=False).agg(agg_funcs)

            # 4. 고정 코드 및 날짜/금액 적용
            df_grouped['발주코드'] = df_grouped['센터'].map(CENTER_CODE_MAP).fillna(df_grouped['발주_fallback'])
            df_grouped['배송코드'] = df_grouped['발주코드']
            df_grouped['수주일자'] = today_str
            df_grouped['Total Amount'] = df_grouped['UNIT수량'] * df_grouped['UNIT단가']
            
            # 숫자형 변환 확인 사살 (판다스 내부에서 명시적 숫자 처리)
            df_grouped['발주코드'] = pd.to_numeric(df_grouped['발주코드'], errors='coerce')
            df_grouped['배송코드'] = pd.to_numeric(df_grouped['배송코드'], errors='coerce')
            df_grouped['UNIT수량'] = pd.to_numeric(df_grouped['UNIT수량'], errors='coerce').fillna(0).astype(int)
            df_grouped['UNIT단가'] = pd.to_numeric(df_grouped['UNIT단가'], errors='coerce').fillna(0).astype(int)
            df_grouped['Total Amount'] = pd.to_numeric(df_grouped['Total Amount'], errors='coerce').fillna(0).astype(int)
            
            # 5. 최종 양식 구성
            df_grouped[' '] = ""
            df_grouped['  '] = ""
            df_grouped['   '] = ""
            df_grouped = df_grouped.rename(columns={'ME코드': '상품코드'})
            
            final_columns = [
                '수주일자', '납품일자', '발주코드', '배송코드', 
                '센터', '상품코드', '품명', 
                'UNIT수량', 'UNIT단가', 'Total Amount', '   '
            ]
            
            result_df = df_grouped[[c for c in final_columns if c in df_grouped.columns]]

            st.success(f"✨ 완료! {len(result_df)}건의 정제된 수주 내역을 생성했습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 6. 💡 [핵심] 엑셀 다운로드 시 '숫자 서식'을 강제 적용하여 저장
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
                
                workbook = writer.book
                worksheet = writer.sheets['롯데마트 수주']
                
                # 숫자 포맷 정의 (천 단위 콤마)
                num_format = workbook.add_format({'num_format': '#,##0'})
                # 일반 텍스트 및 기본 포맷
                center_format = workbook.add_format({'align': 'center'})
                
                # 각 컬럼을 순회하며 숫자로 들어가야 할 열에 포맷 강제 지정
                for col_idx, col_name in enumerate(result_df.columns):
                    if col_name in ['발주코드', '배송코드', 'UNIT수량', 'UNIT단가', 'Total Amount']:
                        worksheet.set_column(col_idx, col_idx, 12, num_format) # 숫자 서식 적용
                    elif col_name in ['수주일자', '납품일자']:
                        worksheet.set_column(col_idx, col_idx, 12, center_format)
                    elif col_name == '품명':
                        worksheet.set_column(col_idx, col_idx, 30)
                    else:
                        worksheet.set_column(col_idx, col_idx, 15)
            
            st.download_button(
                label="📥 숫자 포맷 완벽 적용 최종 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_최종_{today_str}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.info("데이터 처리 중 오류가 발생했습니다. 로그를 확인해주세요.")
