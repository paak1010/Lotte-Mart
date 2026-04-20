import streamlit as st
import pandas as pd
import duckdb
import io

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 데이터 자동 생성기")
st.write("첫 번째 시트(RAW)의 데이터를 바탕으로 '수주' 양식을 자동 생성합니다.")

uploaded_file = st.file_uploader("롯데마트 서식 파일을 업로드해주세요.", type=['xlsx'])

if uploaded_file is not None:
    with st.spinner("데이터를 분석 중입니다..."):
        try:
            # 1. 첫 번째 시트 읽기
            df_raw = pd.read_excel(uploaded_file, sheet_name=0)
            
            # 🚨 디버깅용: 화면에 현재 읽어들인 데이터 띄우기
            st.warning(f"🔍 현재 파이썬이 인식한 총 컬럼(열) 수: {len(df_raw.columns)}개")
            st.write("아래 미리보기를 통해 원본 데이터가 정상적으로 들어왔는지 확인해주세요.")
            st.dataframe(df_raw.head(), use_container_width=True)
            
            # 열이 14개 미만이면 변환을 중단하고 안내 메시지 출력
            if len(df_raw.columns) < 14:
                st.error("앗! 데이터의 열이 부족합니다. '납품수량'과 '상품코드'가 있는 롯데마트 RAW 시트가 맞는지 확인해주세요.")
            else:
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
                
                st.success(f"변환 완료! 유효 수주 {len(result_df)}건을 찾았습니다.")
                st.subheader("📋 생성된 수주 시트 미리보기")
                st.dataframe(result_df, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
                
                st.download_button(
                    label="📥 변환된 수주 엑셀 다운로드",
                    data=buffer.getvalue(),
                    file_name="롯데마트_수주_자동생성.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
