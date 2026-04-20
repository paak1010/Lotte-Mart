import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (백지 생성 버전)")
st.write("EDI Raw Data에 있는 상품들만 추출하여 롯데마트 수주 양식으로 새롭게 그려냅니다.")

# 💡 깃허브의 고정 서식 파일 (오직 ME코드 매핑용으로만 사용)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 💡 요청하신 센터별 고정 발주/배송코드
CENTER_CODE_MAP = {
    '오산상온센타': '81030907',
    '김해상온센타': '81030908'
}

uploaded_file = st.file_uploader("📥 EDI Raw Data (라떼는.xlsx 등) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("Raw 데이터에서 순수 발주 내역만 뽑아내는 중입니다... 🚀"):
        try:
            # 1. EDI Raw 데이터 파싱 (오직 여기서 나온 데이터만 씁니다)
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            df_edi = df_edi.dropna(how='all')
            
            parsed_list = []
            curr_center = ""
            curr_doc_no = "" # 센터가 맵핑 안 될 경우를 대비한 기본 발주번호
            
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
                        # 가격 및 품명 추출
                        price_str = r[7].replace(',', '')
                        price = float(price_str) if price_str.replace('.', '').isdigit() else 0.0
                        item_name = r[2]
                        
                        parsed_list.append({
                            '발주_fallback': curr_doc_no,
                            '센터': curr_center,
                            '바코드': barcode,
                            '품명': item_name,
                            'UNIT수량': unit_qty,
                            'UNIT단가': price,
                            'Total Amount': unit_qty * price
                        })
            
            if not parsed_list:
                st.warning("⚠️ 추출된 유효 발주 건수(0 초과)가 없습니다.")
                st.stop()

            df_parsed = pd.DataFrame(parsed_list)

            # 2. 서식 파일의 1번 시트에서 '바코드 -> ME코드' 공식만 훔쳐오기
            df_map_sheet = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            mapping_dict = df_map_sheet.iloc[:, [3, 13]].dropna() # 3:판매코드, 13:상품코드(ME)
            mapping_dict.columns = ['바코드', 'ME코드']
            mapping_dict['바코드'] = mapping_dict['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            mapping_dict['ME코드'] = mapping_dict['ME코드'].astype(str).str.strip()
            mapping_dict = mapping_dict.drop_duplicates()

            # Raw 데이터에 ME코드 입히기
            df_final = pd.merge(df_parsed, mapping_dict, on='바코드', how='left')
            df_final['ME코드'] = df_final['ME코드'].fillna(df_final['바코드']) # 매핑 안되면 원본 바코드

            # 3. 센터별 고정 코드 부여 (오산/김해)
            df_final['발주코드'] = df_final['센터'].map(CENTER_CODE_MAP).fillna(df_final['발주_fallback'])
            df_final['배송코드'] = df_final['발주코드']
            
            # 4. 새 도화지에 2번 시트 양식(껍데기)만 그대로 그리기 
            # (기존 서식 파일의 데이터는 1도 섞이지 않음)
            df_final[' '] = ""
            df_final['  '] = ""
            df_final['   '] = ""
            
            # 컬럼 순서 및 이름을 두 번째 시트 양식과 100% 동일하게 배치
            df_final = df_final.rename(columns={'ME코드': '상품코드'})
            
            final_columns = [
                ' ', '발주코드', '  ', '배송코드', '센터', 
                '상품코드', '품명', 'UNIT수량', 'UNIT단가', 'Total Amount', '   '
            ]
            
            result_df = df_final[final_columns]

            st.success(f"✨ 완료! Raw 데이터에 있던 순수 {len(result_df)}건의 발주가 완벽하게 정리되었습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 다운로드 생성
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 100% 순수 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_완료.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
