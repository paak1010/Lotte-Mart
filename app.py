import streamlit as st
import pandas as pd
import duckdb
import io
import re

st.set_page_config(page_title="롯데마트 수주 자동화 V3", layout="wide")

st.title("📦 롯데마트 수주서 변환기 (ME 코드 완벽 매핑)")
st.write("EDI 데이터를 업로드하면 '상품코드'를 ME 코드로 변환하여 최종 수주 양식을 생성합니다.")

# 1. ME 코드 매핑 데이터 로드
@st.cache_data
def load_mapping():
    try:
        # 이전에 제공해주신 '제품코드' csv 파일을 mapping.csv로 저장하여 사용
        df = pd.read_csv('mapping.csv')
        # 바코드 정제 (숫자 뒤 .0 제거 및 공백 제거)
        df['바코드'] = df['바코드'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        return df
    except:
        st.warning("⚠️ 'mapping.csv'를 찾을 수 없어 원본 바코드가 출력될 수 있습니다.")
        return pd.DataFrame(columns=['바코드', 'ME코드'])

mapping_df = load_mapping()

# 2. 파일 업로드
uploaded_file = st.file_uploader("롯데마트 RAW 데이터 또는 EDI 발주서 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("ME 코드로 매핑 및 양식 변환 중..."):
        try:
            if uploaded_file.name.endswith('.csv'):
                df_all = pd.read_csv(uploaded_file, header=None)
            else:
                df_all = pd.read_excel(uploaded_file, header=None)

            final_rows = []
            current_center = ""
            current_doc_no = ""
            
            for i, row in df_all.iterrows():
                row_list = row.tolist()
                
                # ORDERS 행에서 센터 및 문서번호 추출
                if str(row_list[0]) == 'ORDERS':
                    current_doc_no = str(row_list[1]).replace('.0', '').strip()
                    current_center = str(row_list[5]).strip()
                    continue
                
                # 상품 행 추출 (바코드 880 시작 기준)
                val_code = str(row_list[1]).replace('.0', '').strip()
                if val_code.startswith('880'):
                    # 주문수 숫자 추출
                    order_qty_raw = str(row_list[6])
                    order_qty_val = int(re.sub(r'[^0-9]', '', order_qty_raw))
                    
                    # 입수량 및 총 UNIT 수량 계산
                    case_size = int(row_list[5]) if pd.notnull(row_list[5]) else 1
                    total_units = order_qty_val * case_size
                    
                    # 데이터 시트 끝에 ME 코드가 이미 있는 경우 (인덱스 13)
                    me_code_in_row = str(row_list[13]).strip() if len(row_list) > 13 and pd.notnull(row_list[13]) else None
                    
                    if total_units > 0:
                        final_rows.append({
                            '발주코드': current_doc_no,
                            '배송코드': current_doc_no,
                            '센터': current_center,
                            '바코드': val_code,
                            'ME코드_직접': me_code_in_row,
                            '품명': str(row_list[2]).strip(),
                            'UNIT수량': total_units,
                            'UNIT단가': row_list[7],
                            'Total Amount': row_list[8]
                        })

            df_extracted = pd.DataFrame(final_rows)

            # 3. DuckDB를 사용하여 최종 ME 코드 확정 및 양식 구성
            # 로직: 1순위(데이터 끝에 있는 ME코드), 2순위(mapping.csv 매핑), 3순위(원본 바코드)
            query = """
                SELECT 
                    '' as " ", 
                    a.발주코드, 
                    '' as "  ", 
                    a.배송코드, 
                    a.센터, 
                    COALESCE(NULLIF(a.ME코드_직접, 'nan'), m.ME코드, a.바코드) as 상품코드, 
                    a.품명, 
                    a.UNIT수량, 
                    a.UNIT단가, 
                    a."Total Amount",
                    '' as "   "
                FROM df_extracted a
                LEFT JOIN mapping_df m ON a.바코드 = m.바코드
            """
            
            result_df = duckdb.query(query).df()

            st.success(f"변환 성공! 총 {len(result_df)}건의 상품이 ME 코드로 변환되었습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 4. 엑셀 생성 및 다운로드
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 ME 코드 적용 수주 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_최종_{current_doc_no}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
