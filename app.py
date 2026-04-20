import streamlit as st
import pandas as pd
import duckdb
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기")
st.write("Raw Data만 업로드하세요. 두 번째 시트의 서식 결과값 중 수량이 있는 것만 추출합니다.")

# 💡 깃허브에 있는 고정 서식 파일 (템플릿)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

uploaded_file = st.file_uploader("📥 EDI Raw Data 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("서식 적용 및 유효 수주 추출 중..."):
        try:
            # 1. 업로드된 EDI Raw 데이터 파싱 (센터명과 주문수 추출)
            # 엑셀/CSV 대응
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            parsed_list = []
            curr_center = ""
            
            for i, row in df_edi.iterrows():
                r = row.tolist()
                # 'ORDERS' 행에서 센터명 추출
                if str(r[0]) == 'ORDERS':
                    curr_center = str(r[5]).strip()
                    continue
                
                # 상품 행 (880 바코드 기준)
                val_code = str(r[1]).replace('.0', '').strip()
                if val_code.startswith('880'):
                    qty_str = str(r[6])
                    qty_num = int(re.sub(r'[^0-9]', '', qty_str)) if qty_str != 'nan' else 0
                    
                    if qty_num > 0:
                        parsed_list.append({
                            '센터_매칭': curr_center,
                            '바코드_매칭': val_code,
                            '실제주문수': qty_num
                        })
            
            df_parsed = pd.DataFrame(parsed_list)

            # 2. 깃허브에 있는 템플릿의 '두 번째 시트' 가져오기
            # 이 시트에는 이미 'ME코드', '품명', '단가' 등이 다 적혀 있습니다.
            df_template = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
            
            # 3. DuckDB로 템플릿과 Raw 데이터를 연결하여 "수량 > 0"인 행만 추출
            # 템플릿에 있는 양식(컬럼 순서)을 그대로 유지합니다.
            query = """
                SELECT 
                    t.* EXCLUDE (UNIT수량, "Total Amount"), -- 기존 수식 컬럼 제외하고 새로 계산
                    (p.실제주문수 * CAST(t.입수 AS INTEGER)) AS "UNIT수량",
                    (p.실제주문수 * CAST(t.입수 AS INTEGER) * CAST(t.UNIT단가 AS INTEGER)) AS "Total Amount"
                FROM df_template t
                JOIN df_parsed p ON t.센터 = p.센터_매칭 AND CAST(t.바코드 AS VARCHAR) = p.바코드_매칭
                WHERE p.실제주문수 > 0
            """
            
            # 만약 템플릿 시트의 컬럼명이 위와 다르다면 직접 인덱스로 지정하여 결과 생성
            result_df = duckdb.query(query).df()

            # 4. 결과 출력 및 다운로드
            st.success(f"변환 완료! 유효 수주 {len(result_df)}건이 추출되었습니다.")
            
            # 컬럼 순서 재배치 (이미지와 동일하게: 발주코드, 배송코드, 센터, 상품코드, 품명, 수량, 단가, 금액)
            final_cols = ['발주코드', '배송코드', '센터', '상품코드', '품명', 'UNIT수량', 'UNIT단가', 'Total Amount']
            # 존재하는 컬럼만 필터링해서 보여줌
            display_df = result_df[[c for c in final_cols if c in result_df.columns]]
            
            st.dataframe(display_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                display_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_최종.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류 발생: {e}")
            st.info("템플릿의 컬럼명(센터, 바코드, 입수, UNIT단가 등)이 코드와 일치하는지 확인해주세요.")
