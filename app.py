import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (숫자 포맷 & 마스터 단가 적용)")
st.write("EDI Raw Data 업로드 시, 마스터 단가를 매핑하고 모든 수치 데이터를 '순수 숫자'로 엑셀에 저장합니다.")

# 💡 깃허브 고정 서식 파일
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 💡 센터별 고정 발주/배송코드 매핑
CENTER_CODE_MAP = {
    '오산센터': '81030907',
    '김해센터': '81030908'
}

def clean_center_name(name):
    name = str(name).strip()
    name = re.sub(r'상온센타|상온센터|센타', '센터', name)
    return name.replace('센터센터', '센터')

def clean_code(val):
    """바코드 등에 붙는 '.0' 소수점과 공백을 완벽히 제거하여 100% 매칭 보장"""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def clean_number(val):
    """콤마나 특수문자가 섞인 값에서 순수 정수만 추출"""
    s = str(val).replace(',', '').strip()
    if s.endswith('.0'):
        s = s[:-2]
    s = re.sub(r'[^0-9]', '', s)
    return int(s) if s else 0

uploaded_file = st.file_uploader("📥 EDI Raw Data (xls 변환) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("마스터 단가 매핑 및 숫자 포맷 변환 중입니다... 🚀"):
        try:
            # 1. EDI 데이터 파싱
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            df_edi = df_edi.dropna(how='all')
            
            parsed_list = []
            curr_center = ""
            curr_doc_no = "" 
            curr_delivery_date = ""
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            for i, row in df_edi.iterrows():
                r = [str(x).strip() for x in row.tolist()]
                
                # ORDERS 행
                if r[0] == 'ORDERS':
                    curr_doc_no = clean_code(r[1])
                    curr_center = clean_center_name(r[5])
                    raw_date = r[7] if len(r) > 7 else ""
                    curr_delivery_date = re.sub(r'[^0-9-]', '', raw_date) 
                    continue
                
                # 상품 행 (880 바코드)
                barcode = clean_code(r[1])
                if barcode.startswith('880'):
                    qty = clean_number(r[6])
                    ipsu = clean_number(r[5])
                    if ipsu == 0: ipsu = 1
                    
                    unit_qty = qty * ipsu
                    
                    if unit_qty > 0:
                        # CSV 줄바꿈으로 인해 단가가 유실되어도 일단 추출 시도
                        edi_price = clean_number(r[7] if len(r) > 7 else 0)
                        
                        parsed_list.append({
                            '발주번호': curr_doc_no,
                            '센터': curr_center,
                            '납품일자': curr_delivery_date,
                            '바코드': barcode,
                            'EDI_품명': r[2],
                            'UNIT수량': unit_qty,
                            'EDI_단가': edi_price
                        })
            
            if not parsed_list:
                st.warning("⚠️ 유효한 발주 내역이 없습니다.")
                st.stop()

            df_parsed = pd.DataFrame(parsed_list)

            # 2. 템플릿(서식 파일)에서 정보 긁어오기
            df_map_sheet = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            df_price_sheet = pd.read_excel(TEMPLATE_FILE, sheet_name=1)

            # [시트 1] 바코드 -> ME코드 매핑 사전
            p_col = df_map_sheet.columns[3]
            m_col = df_map_sheet.columns[13]
            mapping_dict = df_map_sheet[[p_col, m_col]].copy()
            mapping_dict.columns = ['바코드', 'ME코드']
            mapping_dict['바코드'] = mapping_dict['바코드'].apply(clean_code)
            mapping_dict['ME코드'] = mapping_dict['ME코드'].astype(str).str.strip()
            mapping_dict = mapping_dict.drop_duplicates(subset=['바코드'])

            # [시트 2] ME코드 -> 마스터 품명 및 단가 사전 (CSV 오류 방지용 핵심 로직)
            c_me = [c for c in df_price_sheet.columns if '상품코드' in str(c) or 'ME' in str(c).upper()][0]
            c_name = [c for c in df_price_sheet.columns if '품명' in str(c) or '상품명' in str(c)][0]
            c_price = [c for c in df_price_sheet.columns if '단가' in str(c)][0]
            
            price_dict = df_price_sheet[[c_me, c_name, c_price]].dropna(subset=[c_me]).copy()
            price_dict.columns = ['ME코드', '마스터_품명', '마스터_단가']
            price_dict['ME코드'] = price_dict['ME코드'].apply(clean_code)
            price_dict['마스터_단가'] = price_dict['마스터_단가'].apply(clean_number)
            price_dict = price_dict.drop_duplicates(subset=['ME코드'])

            # 3. 데이터 병합 (ME코드 입히고 -> 마스터 단가 입히기)
            df_final = pd.merge(df_parsed, mapping_dict, on='바코드', how='left')
            df_final['ME코드'] = df_final['ME코드'].fillna(df_final['바코드'])
            
            df_final = pd.merge(df_final, price_dict, on='ME코드', how='left')
            df_final['품명'] = df_final['마스터_품명'].fillna(df_final['EDI_품명'])
            # 템플릿의 단가가 있으면 무조건 우선 사용! (CSV 줄바꿈으로 인한 0원 처리 방지)
            df_final['UNIT단가'] = df_final['마스터_단가'].fillna(df_final['EDI_단가'])

            # 4. 발주번호, 센터, 납품일자, ME코드 기준으로 합산 
            # (발주번호를 분리하여 서로 다른 날짜의 오산센터 발주가 하나로 뭉치는 것 방지)
            agg_funcs = {
                '품명': 'first',
                'UNIT단가': 'first',
                'UNIT수량': 'sum'
            }
            df_grouped = df_final.groupby(['발주번호', '센터', '납품일자', 'ME코드'], as_index=False).agg(agg_funcs)

            # 5. 고정 코드 및 금액 계산
            df_grouped['발주코드'] = df_grouped['센터'].map(CENTER_CODE_MAP).fillna(df_grouped['발주번호'])
            df_grouped['배송코드'] = df_grouped['발주코드']
            df_grouped['수주일자'] = today_str
            df_grouped['Total Amount'] = df_grouped['UNIT수량'] * df_grouped['UNIT단가']
            
            # 6. 최종 양식 구성
            df_grouped[' '] = ""
            df_grouped['  '] = ""
            df_grouped['   '] = ""
            df_grouped = df_grouped.rename(columns={'ME코드': '상품코드'})
            
            final_columns = [
                '수주일자', '납품일자', '발주코드', '배송코드', 
                '센터', '상품코드', '품명', 
                'UNIT수량', 'UNIT단가', 'Total Amount', '   '
            ]
            result_df = df_grouped[final_columns].copy()

            # 💡 [핵심] 엑셀이 '일반' 텍스트가 아닌 '숫자'로 완벽하게 인식하도록 강제 변환
            numeric_cols = ['발주코드', '배송코드', 'UNIT수량', 'UNIT단가', 'Total Amount']
            for col in numeric_cols:
                result_df[col] = pd.to_numeric(result_df[col], errors='coerce').fillna(0).astype('int64')

            st.success(f"✨ 완료! 숫자 포맷이 적용된 {len(result_df)}건의 리스트를 생성했습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 7. 엑셀 다운로드 (천 단위 콤마 숫자 서식 추가)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
                
                workbook = writer.book
                worksheet = writer.sheets['롯데마트 수주']
                
                num_format = workbook.add_format({'num_format': '#,##0'})
                center_format = workbook.add_format({'align': 'center'})
                
                for col_idx, col_name in enumerate(result_df.columns):
                    if col_name in numeric_cols:
                        worksheet.set_column(col_idx, col_idx, 12, num_format)
                    elif col_name in ['수주일자', '납품일자']:
                        worksheet.set_column(col_idx, col_idx, 12, center_format)
                    elif col_name == '품명':
                        worksheet.set_column(col_idx, col_idx, 30)
                    else:
                        worksheet.set_column(col_idx, col_idx, 15)
            
            st.download_button(
                label="📥 숫자 포맷 강제 적용 완료 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_완벽정제_{today_str}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
