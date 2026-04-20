import streamlit as st
import pandas as pd
import duckdb
import io

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 데이터 자동 생성기")
st.write("첫 번째 시트(RAW)의 데이터를 바탕으로 '수주' 양식을 자동 생성합니다.")

# 파일 업로드
uploaded_file = st.file_uploader("롯데마트 서식 파일을 업로드해주세요.", type=['xlsx'])

if uploaded_file is not None:
    with st.spinner("데이터를 분석하고 수주 시트를 생성 중입니다..."):
        try:
            # 1. 첫 번째 시트(RAW) 읽기 (헤더는 0번 행 기준)
            df_raw = pd.read_excel(uploaded_file, sheet_name=0)
            
            # 데이터 컬럼 인덱스 확인 및 매핑
            # 사용자 피드백 반영: 상품코드는 가장 끝쪽(인덱스 13)에 있는 것을 사용
            # 시트 구조에 따른 컬럼 추출:
            # [0]점포코드, [1]점포(센터), [4]상품명, [9]단가, [10]주문금액, [12]납품수량, [13]상품코드(ME), [18]배송코드
            
            # 2. DuckDB를 사용하여 수량이 0보다 큰 데이터만 추출하고 양식에 맞게 변환
            # 컬럼명이 중복될 경우를 대비해 인덱스 기반으로 쿼리를 구성하거나 컬럼명을 정제합니다.
            df_raw.columns = [f"col_{i}" for i in range(len(df_raw.columns))]
            
            query = """
                SELECT 
                    '' as " ",
                    col_0 as "발주코드",
                    '' as "  ",
                    col_18 as "배송코드",
                    col_1 as "센터",
                    col_13 as "상품코드",
                    col_4 as "품명",
                    CAST(col_12 AS INTEGER) as "UNIT수량",
                    col_9 as "UNIT단가",
                    col_10 as "Total Amount",
                    '' as "   "
                FROM df_raw
                WHERE col_12 IS NOT NULL 
                  AND CAST(col_12 AS INTEGER) > 0
            """
            
            result_df = duckdb.query(query).df()
            
            # 3. 결과 출력 및 다운로드
            st.success(f"변환 완료! 유효 수주 {len(result_df)}건을 찾았습니다.")
            
            st.subheader("📋 생성된 수주 시트 미리보기")
            st.dataframe(result_df, use_container_width=True)
            
            # 엑셀 파일 생성 (두 번째 시트 이름은 '롯데마트 수주')
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
                
                # 엑셀 서식 간단 조정 (선택 사항)
                workbook = writer.book
                worksheet = writer.sheets['롯데마트 수주']
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                for col_num, value in enumerate(result_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            
            st.download_button(
                label="📥 변환된 수주 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_자동생성.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("첫 번째 시트의 컬럼 순서가 [0:점포코드, 12:납품수량, 13:상품코드]인지 확인해주세요.")
