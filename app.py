import streamlit as st
import pandas as pd
import duckdb
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기")
st.write("Raw Data만 업로드하세요. 서식이 완료된 두 번째 시트에서 유효 데이터만 추출합니다.")

# 💡 깃허브에 있는 고정 서식 파일
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

uploaded_file = st.file_uploader("📥 EDI Raw Data 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("서식 적용 및 수주 데이터 추출 중..."):
        try:
            # 1. 업로드된 Raw Data 파싱 (센터별 데이터 구조 분해)
            df_edi = pd.read_excel(uploaded_file, header=None) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file, header=None)
            
            parsed_data = []
            curr_center = ""
            curr_doc = ""
            
            for i, row in df_edi.iterrows():
                r = row.tolist()
                if str(r[0]) == 'ORDERS':
                    curr_doc = str(r[1]).replace('.0', '')
                    curr_center = str(r[5])
                    continue
                
                if str(r[1]).startswith('880'):
                    qty_raw = str(r[6])
                    qty_num = int(re.sub(r'[^0-9]', '', qty_raw)) if qty_raw != 'nan' else 0
                    
                    if qty_num > 0:
                        parsed_data.append({
                            '점포(센터)': curr_center,
                            '판매코드': r[1],
                            '상품명': r[2],
                            '입수': r[5],
                            '주문수': qty_num,
                            '단가': r[7],
                            '주문금액': r[8],
                            '발주번호': curr_doc
                        })
            
            df_parsed = pd.DataFrame(parsed_data)

            # 2. 템플릿의 두 번째 시트(수주) 구조 가져오기
            # 이미 ME코드와 수량이 서식으로 연결되어 있으므로, 
            # 파싱된 데이터를 기준으로 템플릿 양식에 값을 매칭합니다.
            df_template_order = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
            
            # 3. DuckDB를 사용하여 두 번째 시트 양식 그대로 출력 (UNIT수량 > 0)
            # 엑셀의 서식 결과값인 ME코드와 품명을 그대로 가져오기 위해 JOIN 처리
            query = """
                SELECT 
                    '' as " ", 
                    p.발주번호 as "발주코드", 
                    '' as "  ", 
                    p.발주번호 as "배송코드", 
                    t.센터, 
                    t.상품코드, -- 템플릿 시트의 ME코드
                    t.품명, 
                    (p.주문수 * CAST(t.입수 AS INTEGER)) as "UNIT수량", 
                    t.UNIT단가, 
                    (p.주문수 * CAST(t.입수 AS INTEGER) * t.UNIT단가) as "Total Amount",
                    '' as "   "
                FROM df_template_order t
                JOIN df_parsed p ON t.센터 = p."점포(센터)" AND CAST(t.바코드 AS VARCHAR) = CAST(p.판매코드 AS VARCHAR)
                WHERE (p.주문수 * CAST(t.입수 AS INTEGER)) > 0
            """
            
            result_df = duckdb.query(query).df()

            # 4. 결과 출력 및 다운로드
            st.success(f"변환 완료! 유효 수주 {len(result_df)}건이 추출되었습니다.")
            st.subheader("📋 롯데마트 수주 시트 (최종 출력)")
            st.dataframe(result_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_최종_{curr_doc}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
