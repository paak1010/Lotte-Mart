import streamlit as st
import pandas as pd
import duckdb
import io

st.set_page_config(page_title="롯데마트 수주 필터링", layout="wide")

st.title("📦 롯데마트 수주서 자동 필터링 (0건 제외)")
st.write("작업하신 엑셀 파일을 올리면 두 번째 시트에서 'UNIT수량'이 0보다 큰 건들만 깔끔하게 뽑아줍니다.")

uploaded_file = st.file_uploader("롯데마트 작업 엑셀 파일 업로드", type=['xlsx'])

if uploaded_file is not None:
    with st.spinner("0건 제외 필터링 중..."):
        try:
            # 1. 두 번째 시트 (인덱스 1) 읽기
            # (만약 시트 순서가 바뀌었다면 sheet_name='롯데마트 수주' 처럼 이름으로 지정해도 됩니다)
            df_order = pd.read_excel(uploaded_file, sheet_name=1)
            
            # 2. DuckDB를 이용한 필터링 (UNIT수량이 0보다 큰 것만)
            query = """
                SELECT * FROM df_order
                WHERE UNIT수량 IS NOT NULL 
                  AND CAST(UNIT수량 AS INTEGER) > 0
            """
            
            result_df = duckdb.query(query).df()
            
            # 3. 판다스가 임의로 붙인 빈 컬럼 이름('Unnamed: 0' 등)을 다시 공백으로 원복 (양식 유지)
            clean_cols = []
            space_count = 1
            for col in result_df.columns:
                if "Unnamed" in str(col):
                    clean_cols.append(" " * space_count)
                    space_count += 1
                else:
                    clean_cols.append(col)
            result_df.columns = clean_cols

            st.success(f"필터링 완료! 총 {len(result_df)}건의 유효 발주가 추출되었습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 4. 엑셀 다운로드 (최종 제출용)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_최종제출용.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("엑셀 파일의 두 번째 시트에 'UNIT수량' 컬럼이 정확히 있는지 확인해주세요.")
