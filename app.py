import streamlit as st
import pandas as pd
import duckdb
import io
import re

st.set_page_config(page_title="롯데마트 수주 자동화 V2", layout="wide")

st.title("📦 롯데마트 수주서 변환기 (최종 양식 맞춤)")
st.write("EDI RAW 데이터를 업로드하면 '롯데마트 수주' 시트 양식으로 즉시 변환합니다.")

# 1. 매핑 데이터 로드 (GitHub에 올린 mapping.csv 활용)
@st.cache_data
def load_mapping():
    try:
        df = pd.read_csv('mapping.csv')
        # 바코드를 문자열로 통일
        df['바코드'] = df['바코드'].astype(str)
        return df
    except:
        st.warning("⚠️ 'mapping.csv' 파일을 찾을 수 없습니다. 상품코드 매핑 없이 진행합니다.")
        return pd.DataFrame(columns=['바코드', 'ME코드'])

mapping_df = load_mapping()

# 2. 파일 업로드
uploaded_file = st.file_uploader("EDI 발주서 (라떼는.xlsx 등) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("데이터 매핑 및 변환 중..."):
        try:
            # 데이터 로드
            if uploaded_file.name.endswith('.csv'):
                df_all = pd.read_csv(uploaded_file, header=None)
            else:
                df_all = pd.read_excel(uploaded_file, header=None)

            # --- EDI 데이터 파싱 로직 ---
            final_rows = []
            current_center = ""
            current_doc_no = ""
            
            for i, row in df_all.iterrows():
                row_list = row.tolist()
                
                # 'ORDERS'로 시작하는 행에서 센터명과 문서번호 추출
                if str(row_list[0]) == 'ORDERS':
                    current_doc_no = str(row_list[1])
                    current_center = str(row_list[5]) # 점포(센터)
                    continue
                
                # 실제 상품 행 추출 (판매코드가 880으로 시작)
                val_code = str(row_list[1])
                if val_code.startswith('880'):
                    # 주문수에서 숫자만 추출 (예: '1 (BOX)' -> 1)
                    order_qty_raw = str(row_list[6])
                    order_qty_val = int(re.sub(r'[^0-9]', '', order_qty_raw))
                    
                    # 입수량 추출 및 총 수량 계산 (UNIT수량)
                    case_size = int(row_list[5]) if pd.notnull(row_list[5]) else 1
                    total_units = order_qty_val * case_size
                    
                    if total_units > 0:
                        final_rows.append({
                            '발주코드': current_doc_no, # 또는 센터코드 매핑 필요 시 수정
                            '배송코드': current_doc_no,
                            '센터': current_center,
                            '바코드': val_code,
                            '품명': row_list[2],
                            'UNIT수량': total_units,
                            'UNIT단가': row_list[7],
                            'Total Amount': row_list[8]
                        })

            df_extracted = pd.DataFrame(final_rows)

            # 3. DuckDB를 활용한 ME 코드 매핑 및 최종 양식 구성
            query = """
                SELECT 
                    '' as " ", 
                    a.발주코드, 
                    '' as "  ", 
                    a.배송코드, 
                    a.센터, 
                    COALESCE(m.ME코드, a.바코드) as 상품코드, 
                    a.품명, 
                    a.UNIT수량, 
                    a.UNIT단가, 
                    a."Total Amount",
                    '' as "   "
                FROM df_extracted a
                LEFT JOIN mapping_df m ON a.바코드 = m.바코드
            """
            
            result_df = duckdb.query(query).df()

            st.success(f"처리 완료! {len(result_df)}개의 품목이 생성되었습니다.")
            st.subheader("📋 수주 시트 미리보기")
            st.dataframe(result_df, use_container_width=True)

            # 4. 엑셀 다운로드 (양식 유지)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 롯데마트 수주 양식 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_{current_doc_no}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
