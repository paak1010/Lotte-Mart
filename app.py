import streamlit as st
import pandas as pd
import io
import os
import re
from datetime import datetime

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (날짜 자동화 버전)")
st.write("EDI Raw Data를 업로드하면 수주일자(오늘)와 납품일자를 포함한 수주 양식을 생성합니다.")

# 💡 깃허브의 고정 서식 파일 (ME코드 매핑용)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 센터별 고정 발주/배송코드
CENTER_CODE_MAP = {
    '오산상온센타': '81030907',
    '김해상온센타': '81030908'
}

uploaded_file = st.file_uploader("📥 EDI Raw Data (라떼는.xlsx 등) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("날짜 정보를 추출하고 수량을 합산 중입니다... 🚀"):
        try:
            # 1. EDI Raw 데이터 파싱
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            df_edi = df_edi.dropna(how='all')
            
            parsed_list = []
            curr_center = ""
            curr_doc_no = "" 
            curr_delivery_date = "" # Raw 데이터에서 추출할 납품일자
            
            # 오늘 날짜 (YYYY-MM-DD 형식)
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            for i, row in df_edi.iterrows():
                r = [str(x).strip() for x in row.tolist()]
                
                # 'ORDERS' 행에서 센터, 문서번호, 납품일자 추출
                if r[0] == 'ORDERS':
                    curr_doc_no = r[1].replace('.0', '')
                    curr_center = r[5]
                    # 인덱스 7번에 보통 납품일자(2026-04-22 등)가 위치함
                    raw_date = r[7] if len(r) > 7 else ""
                    # 날짜 형식 정제 (YYYY-MM-DD)
                    curr_delivery_date = re.sub(r'[^0-9-]', '', raw_date) 
                    continue
                
                # 상품 행 추출 (880 바코드)
                barcode = r[1].replace('.0', '')
                if barcode.startswith('880'):
                    qty_str = re.sub(r'[^0-9]', '', r[6])
                    qty = int(qty_str) if qty_str else 0
                    
                    ipsu_str = r[5].replace(',', '')
                    ipsu = int(float(ipsu_str)) if ipsu_str.replace('.', '').isdigit() else 1
                    
                    unit_qty = qty * ipsu
                    
                    if unit_qty > 0:
                        price_str = r[7].replace(',', '')
                        price = float(price_str) if price_str.replace('.', '').isdigit() else 0.0
                        
                        parsed_list.append({
                            '발주_fallback': curr_doc_no,
                            '센터': curr_center,
                            '납품일자': curr_delivery_date,
                            '바코드': barcode,
                            '품명': r[2],
                            'UNIT수량': unit_qty,
                            'UNIT단가': price
                        })
            
            if not parsed_list:
                st.warning("⚠️ 추출된 유효 발주 건수(0 초과)가 없습니다.")
                st.stop()

            df_parsed = pd.DataFrame(parsed_list)

            # 2. 서식 파일에서 매핑 정보 가져오기
            df_map_sheet = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            mapping_dict = df_map_sheet.iloc[:, [3, 13]].dropna() 
            mapping_dict.columns = ['바코드', 'ME코드']
            mapping_dict['바코드'] = mapping_dict['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            mapping_dict['ME코드'] = mapping_dict['ME코드'].astype(str).str.strip()
            mapping_dict = mapping_dict.drop_duplicates()

            df_final = pd.merge(df_parsed, mapping_dict, on='바코드', how='left')
            df_final['ME코드'] = df_final['ME코드'].fillna(df_final['바코드'])

            # 3. 센터와 ME코드 기준 수량 합산
            agg_funcs = {
                '발주_fallback': 'first',
                '납품일자': 'first',
                '품명': 'first',
                'UNIT단가': 'first',
                'UNIT수량': 'sum'
            }
            df_grouped = df_final.groupby(['센터', 'ME코드'], as_index=False).agg(agg_funcs)

            # 4. 최종 데이터 정리 및 날짜 부여
            df_grouped['발주코드'] = df_grouped['센터'].map(CENTER_CODE_MAP).fillna(df_grouped['발주_fallback'])
            df_grouped['배송코드'] = df_grouped['발주코드']
            df_grouped['수주일자'] = today_str # 오늘 날짜
            df_grouped['Total Amount'] = df_grouped['UNIT수량'] * df_grouped['UNIT단가']
            
            # 5. 양식에 맞게 컬럼 배치
            df_grouped[' '] = ""
            df_grouped['  '] = ""
            df_grouped['   '] = ""
            
            df_grouped = df_grouped.rename(columns={'ME코드': '상품코드'})
            
            # 수주일자와 납품일자를 센터 정보 근처에 배치 (양식에 따라 조정 가능)
            final_columns = [
                '수주일자', '납품일자', '발주코드', '배송코드', '센터',
             '상품코드', '품명', 
                'UNIT수량', 'UNIT단가', 'Total Amount', '   '
            ]
            
            # 만약 서식 파일 2번째 시트의 정확한 컬럼 순서가 있다면 그에 맞춰 자동으로 필터링
            result_df = df_grouped[[c for c in final_columns if c in df_grouped.columns]]

            st.success(f"✨ 완료! 오늘 날짜({today_str})와 납품일자가 적용된 {len(result_df)}건의 리스트를 생성했습니다.")
            st.dataframe(result_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 날짜 포함 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_{today_str}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
