import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기")
st.write("EDI Raw Data만 업로드하세요. 센터별 고정 코드와 ME코드를 매핑하여 최종 양식을 생성합니다.")

# 💡 깃허브의 고정 서식 파일
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 센터별 고정 코드 설정
CENTER_CODE_MAP = {
    '오산상온센타': '81030907',
    '김해상온센타': '81030908'
}

@st.cache_data
def load_template_sheets():
    # 1번 시트: 바코드-ME코드 매핑용, 2번 시트: 출력 양식용
    df_map = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
    df_out = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
    return df_map, df_out

uploaded_file = st.file_uploader("📥 EDI Raw Data 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("센터별 코드를 매핑하고 수주서를 생성 중입니다..."):
        try:
            df_map_sheet, df_out_sheet = load_template_sheets()

            # 1. EDI Raw 데이터 파싱 (고속 처리)
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            df_edi = df_edi.dropna(how='all')
            
            parsed_list = []
            curr_center = ""
            
            for i, row in df_edi.iterrows():
                r = [str(x).strip() for x in row.tolist()]
                # 센터명 추출
                if r[0] == 'ORDERS':
                    curr_center = r[5]
                    continue
                
                # 상품 행 (880 바코드)
                barcode = r[1].replace('.0', '')
                if barcode.startswith('880'):
                    # 주문수 숫자만 추출 (BOX 제거)
                    qty_str = re.sub(r'[^0-9]', '', r[6])
                    qty = int(qty_str) if qty_str else 0
                    
                    # 입수량 곱하기
                    ipsu_str = r[5].replace(',', '')
                    ipsu = int(float(ipsu_str)) if ipsu_str.replace('.', '').isdigit() else 1
                    
                    unit_qty = qty * ipsu
                    
                    if unit_qty > 0:
                        parsed_list.append({
                            '센터': curr_center,
                            '바코드': barcode,
                            'UNIT수량': unit_qty
                        })
            
            if not parsed_list:
                st.warning("⚠️ 유효한 발주 수량이 없습니다.")
                st.stop()

            df_parsed = pd.DataFrame(parsed_list)

            # 2. ME코드 매핑 (1번 시트 활용)
            # 바코드 열(index 3)과 ME코드 열(index 13) 추출
            mapping_dict = df_map_sheet.iloc[:, [3, 13]].dropna()
            mapping_dict.columns = ['바코드', 'ME코드']
            mapping_dict['바코드'] = mapping_dict['바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            
            df_final_data = pd.merge(df_parsed, mapping_dict, on='바코드', how='left')
            df_final_data['ME코드'] = df_final_data['ME코드'].fillna(df_final_data['바코드'])

            # 3. 2번 시트 양식에 데이터 입히기
            # 템플릿의 모든 품목 리스트를 가져와서 업로드한 데이터와 합침
            df_out_sheet['센터_key'] = df_out_sheet['센터'].astype(str).str.strip()
            df_out_sheet['상품코드_key'] = df_out_sheet['상품코드'].astype(str).str.strip()
            df_final_data['센터_key'] = df_final_data['센터'].astype(str).str.strip()
            df_final_data['ME코드_key'] = df_final_data['ME코드'].astype(str).str.strip()

            # 데이터 병합
            merged = pd.merge(
                df_out_sheet, 
                df_final_data[['센터_key', 'ME코드_key', 'UNIT수량']], 
                left_on=['센터_key', '상품코드_key'], 
                right_on=['센터_key', 'ME코드_key'], 
                how='inner'
            )

            # 4. 고정 코드 적용 (오산 81030907, 김해 81030908)
            merged['발주코드'] = merged['센터_key'].map(CENTER_CODE_MAP)
            merged['배송코드'] = merged['발주코드']
            merged['UNIT수량'] = merged['UNIT수량_y']
            
            # 금액 계산
            merged['단가_num'] = pd.to_numeric(merged['UNIT단가'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            merged['Total Amount'] = merged['UNIT수량'] * merged['단가_num']

            # 5. 최종 양식 정리 (컬럼 순서 유지 및 빈 열 복원)
            orig_cols = [c for c in df_out_sheet.columns if '_key' not in str(c)]
            result_df = merged[orig_cols].copy()
            
            # 시각적 피드백을 위해 Unnamed 컬럼을 공백으로 변환
            new_cols = []
            space_idx = 1
            for c in result_df.columns:
                if 'Unnamed' in str(c):
                    new_cols.append(" " * space_idx)
                    space_idx += 1
                else:
                    new_cols.append(c)
            result_df.columns = new_cols

            st.success(f"✨ 변환 성공! {len(result_df)}건의 유효 수주를 추출했습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 다운로드
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_최종_결과.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
