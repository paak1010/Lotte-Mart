import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (백지 생성 + 수량 합산)")
st.write("EDI Raw Data 추출 후, 동일 센터 및 동일 상품(ME코드)의 수량을 자동으로 합산합니다.")

# 💡 깃허브의 고정 서식 파일 (ME코드 매핑용)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 💡 센터별 고정 발주/배송코드
CENTER_CODE_MAP = {
    '오산상온센타': '81030907',
    '김해상온센타': '81030908'
}

uploaded_file = st.file_uploader("📥 EDI Raw Data (라떼는.xlsx 등) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("발주 내역 추출 및 중복 수량 합산 중입니다... 🚀"):
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
            
            for i, row in df_edi.iterrows():
                r = [str(x).strip() for x in row.tolist()]
                
                # 'ORDERS' 행에서 센터와 문서번호 추출
                if r[0] == 'ORDERS':
                    curr_doc_no = r[1].replace('.0', '')
                    curr_center = r[5]
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
                        item_name = r[2]
                        
                        parsed_list.append({
                            '발주_fallback': curr_doc_no,
                            '센터': curr_center,
                            '바코드': barcode,
                            '품명': item_name,
                            'UNIT수량': unit_qty,
                            'UNIT단가': price
                        })
            
            if not parsed_list:
                st.warning("⚠️ 추출된 유효 발주 건수(0 초과)가 없습니다.")
                st.stop()

            df_parsed = pd.DataFrame(parsed_list)

            # 2. 서식 파일에서 '바코드 -> ME코드' 매핑 정보 가져오기
            df_map_sheet = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            mapping_dict = df_map_sheet.iloc[:, [3, 13]].dropna() 
            mapping_dict.columns = ['바코드', 'ME코드']
            mapping_dict['바코드'] = mapping_dict['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            mapping_dict['ME코드'] = mapping_dict['ME코드'].astype(str).str.strip()
            mapping_dict = mapping_dict.drop_duplicates()

            # Raw 데이터에 ME코드 입히기
            df_final = pd.merge(df_parsed, mapping_dict, on='바코드', how='left')
            df_final['ME코드'] = df_final['ME코드'].fillna(df_final['바코드'])

            # 3. 💡 [핵심 추가] 센터와 ME코드가 같으면 수량을 합산(Groupby)
            agg_funcs = {
                '발주_fallback': 'first', # 발주번호는 첫 번째 값 유지
                '품명': 'first',         # 품명도 첫 번째 값 유지
                'UNIT단가': 'first',     # 단가 유지
                'UNIT수량': 'sum'        # 수량은 합치기!
            }
            df_grouped = df_final.groupby(['센터', 'ME코드'], as_index=False).agg(agg_funcs)

            # 합산된 수량을 바탕으로 총 금액(Total Amount) 다시 계산
            df_grouped['Total Amount'] = df_grouped['UNIT수량'] * df_grouped['UNIT단가']

            # 4. 센터별 고정 코드 부여 (오산/김해)
            df_grouped['발주코드'] = df_grouped['센터'].map(CENTER_CODE_MAP).fillna(df_grouped['발주_fallback'])
            df_grouped['배송코드'] = df_grouped['발주코드']
            
            # 5. 새 도화지에 2번 시트 양식(껍데기)만 그대로 그리기 
            df_grouped[' '] = ""
            df_grouped['  '] = ""
            df_grouped['   '] = ""
            
            df_grouped = df_grouped.rename(columns={'ME코드': '상품코드'})
            
            final_columns = [
                ' ', '발주코드', '  ', '배송코드', '센터', 
                '상품코드', '품명', 'UNIT수량', 'UNIT단가', 'Total Amount', '   '
            ]
            
            result_df = df_grouped[final_columns]

            st.success(f"✨ 완료! 중복 상품을 하나로 합쳐 총 {len(result_df)}건의 발주 라인이 생성되었습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 다운로드 생성
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 중복 합산 완료된 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_완료.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
