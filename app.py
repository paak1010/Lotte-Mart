import streamlit as st
import pandas as pd
import duckdb
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주서 변환기 (계층 데이터 파싱)")
st.write("EDI Raw Data를 업로드하면 센터별로 데이터를 분리하고 수량을 추출하여 최종 수주 양식을 생성합니다.")

# 💡 깃허브에 올려두신 서식 파일 (템플릿용)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ '{TEMPLATE_FILE}' 파일을 찾을 수 없습니다. 깃허브 파일명을 확인해주세요.")
    st.stop()

# 1. Raw 데이터 업로드
uploaded_file = st.file_uploader("📥 EDI Raw Data (라떼는.xlsx 등)를 업로드하세요.", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("데이터 분석 및 센터별 파싱 중..."):
        try:
            # 2. Raw 데이터 로드 (헤더 없이 읽어서 파싱)
            if uploaded_file.name.endswith('.csv'):
                df_raw_full = pd.read_csv(uploaded_file, header=None)
            else:
                df_raw_full = pd.read_excel(uploaded_file, header=None)

            # 3. 센터별/상품별 데이터 추출 로직
            parsed_data = []
            current_center = ""
            current_doc_no = ""
            
            for i, row in df_raw_full.iterrows():
                row_list = row.tolist()
                
                # 'ORDERS' 행을 만나면 센터와 문서번호 업데이트
                if str(row_list[0]) == 'ORDERS':
                    current_doc_no = str(row_list[1]).replace('.0', '').strip()
                    current_center = str(row_list[5]).strip()
                    continue
                
                # 상품 행 추출 (판매코드가 880으로 시작하는 행)
                barcode = str(row_list[1]).replace('.0', '').strip()
                if barcode.startswith('880'):
                    # 주문수에서 숫자만 추출 (예: '1 (BOX)' -> 1)
                    qty_str = str(row_list[6])
                    qty_val = int(re.sub(r'[^0-9]', '', qty_str)) if qty_str != 'nan' else 0
                    
                    # 입수량 곱해서 UNIT수량 계산
                    case_size = int(row_list[5]) if pd.notnull(row_list[5]) else 1
                    unit_qty = qty_val * case_size
                    
                    if unit_qty > 0:
                        parsed_data.append({
                            '발주코드': current_doc_no,
                            '배송코드': current_doc_no,
                            '센터': current_center,
                            '바코드': barcode,
                            'UNIT수량': unit_qty,
                            '단가': row_list[7],
                            '금액': row_list[8]
                        })

            df_extracted = pd.DataFrame(parsed_data)

            # 4. 템플릿(두 번째 시트) 로드하여 상품명 및 ME코드 매핑
            # 템플릿은 ME코드와 바코드 매핑 기준점으로 사용합니다.
            df_template = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
            
            # 5. DuckDB를 사용하여 최종 양식으로 조립
            query = """
                SELECT 
                    '' as " ",
                    e.발주코드,
                    '' as "  ",
                    e.배송코드,
                    e.센터,
                    t.상품코드, -- 여기서 템플릿의 ME코드가 들어감
                    t.품명,
                    e.UNIT수량,
                    e.단가 as "UNIT단가",
                    e.금액 as "Total Amount",
                    '' as "   "
                FROM df_extracted e
                JOIN df_template t ON e.센터 = t.센터 AND e.바코드 = t.바코드 (또는 t.상품코드 매핑)
                -- 💡 참고: 템플릿에 바코드가 없다면 '상품명'이나 다른 기준점을 JOIN 조건으로 쓸 수 있습니다.
            """
            
            # 만약 템플릿에 바코드 열이 없다면, 1번 시트(제품코드)를 가져와서 ME코드로 치환 후 JOIN
            df_mapping = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            df_mapping.columns = [f"m_{i}" for i in range(len(df_mapping.columns))]
            
            final_query = """
                SELECT 
                    '' as " ", e.발주코드, '' as "  ", e.배송코드, e.센터,
                    m.m_13 as "상품코드",
                    e.바코드 as "원본바코드", -- 확인용
                    e.UNIT수량, e.단가, e.금액, '' as "   "
                FROM df_extracted e
                LEFT JOIN df_mapping m ON e.바코드 = CAST(m.m_3 AS VARCHAR)
                WHERE e.UNIT수량 > 0
            """
            
            result_df = duckdb.query(final_query).df()

            # 6. 결과 출력 및 다운로드
            st.success(f"변환 완료! {len(result_df)}건의 유효 수주를 추출했습니다.")
            st.dataframe(result_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_변환_{current_doc_no}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
