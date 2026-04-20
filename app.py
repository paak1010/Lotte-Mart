import streamlit as st
import pandas as pd
import duckdb
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동 변환", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기")
st.write("Raw Data만 업로드하면 센터별 상품 데이터를 추출하여 최종 수주 양식을 생성합니다.")

# 💡 깃허브에 미리 올려둔 서식 파일 이름
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

# 1. Raw Data 파일 업로드
uploaded_file = st.file_uploader("📥 EDI Raw Data 파일을 업로드하세요.", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("데이터를 처리 중입니다..."):
        try:
            # 2. Raw 데이터 읽기 (전체 파싱을 위해 헤더 없이 로드)
            if uploaded_file.name.endswith('.csv'):
                df_raw_full = pd.read_csv(uploaded_file, header=None)
            else:
                df_raw_full = pd.read_excel(uploaded_file, header=None)

            parsed_rows = []
            current_center = ""
            current_doc_no = ""
            
            # 3. 센터별 데이터 순회 및 추출 (가장 중요한 로직)
            for i, row in df_raw_full.iterrows():
                row_list = row.tolist()
                
                # 'ORDERS' 행을 만나면 센터와 문서번호(발주/배송코드용) 저장
                if str(row_list[0]) == 'ORDERS':
                    current_doc_no = str(row_list[1]).replace('.0', '').strip()
                    current_center = str(row_list[5]).strip()
                    continue
                
                # 바코드로 시작하는 실제 상품 행인지 확인
                item_code = str(row_list[1]).replace('.0', '').strip()
                if item_code.startswith('880'):
                    # 주문수에서 '(BOX)' 등 문자 제거 후 숫자만 추출
                    qty_str = str(row_list[6])
                    qty_val = int(re.sub(r'[^0-9]', '', qty_str)) if qty_str != 'nan' else 0
                    
                    # 입수량(5번 인덱스)을 곱해서 최종 UNIT 수량 계산
                    case_size = int(row_list[5]) if pd.notnull(row_list[5]) else 1
                    unit_qty = qty_val * case_size
                    
                    # 0보다 큰 것들만 리스트에 추가
                    if unit_qty > 0:
                        parsed_rows.append({
                            '발주코드': current_doc_no,
                            '배송코드': current_doc_no,
                            '센터': current_center,
                            '상품코드': item_code, # 일단 바코드를 담고 아래에서 ME코드로 치환
                            '품명': str(row_list[2]).strip(),
                            'UNIT수량': unit_qty,
                            'UNIT단가': row_list[7],
                            'Total Amount': row_list[8]
                        })

            df_extracted = pd.DataFrame(parsed_rows)

            # 4. ME코드 매핑 (깃허브에 있는 서식파일의 1번째 시트 활용)
            # 제품코드 시트에서 바코드(인덱스 3)와 ME코드(인덱스 13) 매핑
            df_mapping = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            df_mapping.columns = [f"m_{i}" for i in range(len(df_mapping.columns))]
            
            # DuckDB로 최종 조립
            query = """
                SELECT 
                    '' as " ", 
                    e.발주코드, 
                    '' as "  ", 
                    e.배송코드, 
                    e.센터, 
                    m.m_13 as "상품코드", -- ME코드로 치환
                    e.품명, 
                    e.UNIT수량, 
                    e.UNIT단가, 
                    e."Total Amount",
                    '' as "   "
                FROM df_extracted e
                LEFT JOIN df_mapping m ON e.상품코드 = CAST(m.m_3 AS VARCHAR)
            """
            
            result_df = duckdb.query(query).df()

            # 5. 결과 출력 및 다운로드
            st.success(f"변환 완료! {len(result_df)}건의 유효 발주를 추출했습니다.")
            st.dataframe(result_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_완료_{current_doc_no}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
