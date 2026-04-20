import streamlit as st
import pandas as pd
import duckdb
import io

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 발주서 자동 변환기")
st.write("EDI에서 다운로드한 RAW 데이터를 업로드하면 유효 발주 건(주문수 > 0)만 추출합니다.")

uploaded_file = st.file_uploader("롯데마트 RAW 데이터 (예: 라떼는.xlsx) 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("DuckDB가 데이터를 처리 중입니다..."):
        try:
            # 1. 헤더 없이 데이터 전체 읽기 (양식이 복잡하므로)
            if uploaded_file.name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded_file, header=None)
            else:
                df_raw = pd.read_excel(uploaded_file, header=None, sheet_name=0)
            
            # 2. 가장 넓은 행을 기준으로 10개 컬럼 이름 강제 지정
            # B2B 양식의 실제 데이터 행 구조: 
            # 상품코드, 판매코드, 상품명, 점포명, 규격, 입수, 주문수, 단가, 주문금액, 입고허용일
            df_raw = df_raw.iloc[:, :10] # 혹시 모를 추가 빈 컬럼 방지
            df_raw.columns = ['상품코드', '판매코드', '상품명', '점포명', '규격', '입수', '주문수', '단가', '주문금액', '입고허용일']
            
            # 3. Pandas로 1차 클렌징: '판매코드'가 880으로 시작하는 진짜 상품 행만 살리기
            df_clean = df_raw[df_raw['판매코드'].astype(str).str.startswith('880')].copy()
            
            # 4. '1 (BOX)' 처럼 문자가 섞인 주문수에서 숫자만 추출
            df_clean['주문수_숫자'] = df_clean['주문수'].astype(str).str.replace(r'[^0-9]', '', regex=True)
            df_clean['주문수_숫자'] = pd.to_numeric(df_clean['주문수_숫자'], errors='coerce').fillna(0).astype(int)
            
            # 5. DuckDB를 활용한 고속 필터링 (주문수가 0보다 큰 것만)
            query = """
                SELECT * EXCLUDE (주문수_숫자)
                FROM df_clean
                WHERE 주문수_숫자 > 0
            """
            
            result_df = duckdb.query(query).df()
            
            st.success(f"데이터 정제 완료! 총 {len(result_df)}건의 유효 발주가 추출되었습니다.")
            
            st.subheader("✅ 정제된 수주 리스트 (미리보기)")
            st.dataframe(result_df, use_container_width=True)
            
            # 6. 엑셀 다운로드 (두 번째 시트 '수주' 포맷으로 저장)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 엑셀 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_변환완료.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        except Exception as e:
            st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
