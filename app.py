import streamlit as st
import pandas as pd
import duckdb
import io
import os

st.set_page_config(page_title="롯데마트 수주 자동화 (Raw Data 전용)", layout="wide")

st.title("📦 롯데마트 수주서 자동 생성기")
st.write("오늘 다운받은 **Raw Data**만 업로드하세요. 서버에 저장된 템플릿 양식에 맞춰 자동으로 매핑 후 0건을 제외하고 출력합니다.")

# 💡 깃허브에 이미 올려두신 서식 파일 이름
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ 깃허브에서 '{TEMPLATE_FILE}' 파일을 찾을 수 없습니다. 파일명이 정확한지 확인해주세요.")
    st.stop()

# 1. 사용자는 오직 Raw Data 하나만 업로드
uploaded_file = st.file_uploader("📥 오늘 작업할 Raw Data 파일을 업로드해주세요.", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("서식 파일과 Raw Data를 매칭하여 0건을 걸러내는 중입니다..."):
        try:
            # 2. 업로드한 Raw Data 읽기
            if uploaded_file.name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded_file)
            else:
                df_raw = pd.read_excel(uploaded_file, sheet_name=0)
            
            # 컬럼 이름을 인덱스 번호로 변환 (col_1: 센터, col_12: 수량, col_13: ME코드)
            df_raw.columns = [f"col_{i}" for i in range(len(df_raw.columns))]
            
            # 3. 깃허브에 있는 서식 파일의 '두 번째 시트' 읽기 (템플릿 기준점)
            df_template = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
            
            # 엑셀의 빈 컬럼(열) 양식 그대로 살리기
            clean_cols = []
            space_count = 1
            for col in df_template.columns:
                if "Unnamed" in str(col):
                    clean_cols.append(" " * space_count)
                    space_count += 1
                else:
                    clean_cols.append(str(col).strip())
            df_template.columns = clean_cols

            # 4. DuckDB를 활용해 두 번째 시트 기준 + Raw Data 매핑 + 0 이상 필터링
            query = """
                WITH RawAgg AS (
                    -- Raw Data에서 센터별, ME코드별 납품수량 합계 추출
                    SELECT 
                        col_1 AS center_name,
                        col_13 AS me_code,
                        MAX(col_0) AS order_code,     -- 발주코드 갱신용
                        MAX(col_18) AS delivery_code, -- 배송코드 갱신용
                        SUM(CAST(col_12 AS INTEGER)) AS total_qty
                    FROM df_raw
                    WHERE col_12 IS NOT NULL AND CAST(col_12 AS INTEGER) > 0
                    GROUP BY col_1, col_13
                )
                -- 두 번째 시트(t)를 바탕으로 매핑된 수량(r) 업데이트 및 0건 제외
                SELECT 
                    t." ",
                    COALESCE(r.order_code, t."발주코드") AS "발주코드",
                    t."  ",
                    COALESCE(r.delivery_code, t."배송코드") AS "배송코드",
                    t."센터",
                    t."상품코드",
                    t."품명",
                    COALESCE(r.total_qty, 0) AS "UNIT수량",
                    t."UNIT단가",
                    t."Total Amount",
                    t."   "
                FROM df_template t
                LEFT JOIN RawAgg r 
                       ON t."센터" = r.center_name 
                      AND t."상품코드" = r.me_code
                WHERE COALESCE(r.total_qty, 0) > 0
            """
            
            result_df = duckdb.query(query).df()

            st.success(f"처리 완료! 수주 내역 {len(result_df)}건이 성공적으로 추출되었습니다.")
            st.dataframe(result_df, use_container_width=True)

            # 5. 최종 제출용 엑셀 다운로드 (양식 완벽 유지)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 롯데마트 수주 엑셀 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_완료건.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("올려주신 Raw Data 파일이 기존 양식과 맞는지 확인해주세요.")
