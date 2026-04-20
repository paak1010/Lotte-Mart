import streamlit as st
import pandas as pd
import duckdb
import io

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 데이터 자동 변환기")
st.write("RAW 데이터를 업로드하면 수량이 0이 아닌 유효 발주 건만 필터링하여 수주 내역을 생성합니다.")

# 1. 엑셀 파일 업로드 (매번 바뀌는 Raw Data)
uploaded_file = st.file_uploader("롯데마트 RAW 데이터 엑셀 파일을 업로드해주세요.", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("DuckDB가 데이터를 처리 중입니다..."):
        try:
            # CSV로 업로드 된 경우와 엑셀로 업로드 된 경우 모두 처리
            if uploaded_file.name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded_file)
            else:
                # 첫 번째 시트(RAW)를 데이터프레임으로 읽기
                df_raw = pd.read_excel(uploaded_file, sheet_name=0) 
            
            # 2. DuckDB를 활용한 초고속 필터링 SQL 쿼리
            # 컬럼명에 공백이나 특수문자가 있을 수 있으므로 큰따옴표("")로 묶거나 정제해서 사용
            # 업로드된 데이터 구조에 따라 '납품수량' 또는 '주문수' 컬럼을 타겟으로 잡습니다.
            
            query = """
                SELECT *
                FROM df_raw
                WHERE "납품수량" IS NOT NULL 
                  AND CAST("납품수량" AS INTEGER) > 0
            """
            
            # DuckDB 쿼리 실행
            result_df = duckdb.query(query).df()
            
            st.success(f"데이터 처리 완료! 총 {len(df_raw)}건 중 유효 수주 {len(result_df)}건을 추출했습니다.")
            
            # 3. 결과 화면 출력
            st.subheader("✅ 0건 제외 수주 리스트 (미리보기)")
            st.dataframe(result_df, use_container_width=True)
            
            # 4. 엑셀 다운로드 기능
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 엑셀 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_필터링완료.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
            st.info("컬럼명이 '납품수량'이 맞는지 확인해주세요. 엑셀 내 실제 컬럼명에 맞춰 SQL 쿼리를 수정해야 할 수 있습니다.")
