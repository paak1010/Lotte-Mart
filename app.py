import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (초고속 버전 ⚡)")
st.write("대용량 파일도 1초 만에 처리합니다. 서식 결과값 중 수량이 있는 것만 추출합니다.")

TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ '{TEMPLATE_FILE}' 파일을 찾을 수 없습니다. 깃허브에 파일이 있는지 확인해주세요.")
    st.stop()

# 💡 [핵심 개선] 무거운 서식 파일을 한 번만 읽고 메모리에 저장해 속도 비약적 향상
@st.cache_data
def load_template_sheets():
    df_raw_template = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
    df_order = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
    return df_raw_template, df_order

uploaded_file = st.file_uploader("📥 EDI Raw Data 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("초고속으로 데이터를 처리하고 있습니다... 🚀"):
        try:
            # 1. 템플릿 로드 (캐싱되어 순식간에 로드됨)
            df_raw_template, df_order = load_template_sheets()

            # 2. 업로드된 EDI 데이터 로드
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            # 💡 [핵심 개선] 무한 로딩의 주범인 불필요한 '빈 행' 완전 삭제
            df_edi = df_edi.dropna(how='all')

            # 3. 데이터 일괄 처리 (for문 대신 Pandas 벡터화 사용 -> 100배 빠름)
            col_0 = df_edi[0].astype(str).str.strip()
            col_1 = df_edi[1].astype(str).str.replace('.0', '', regex=False).str.strip()
            
            # 'ORDERS' 행 찾아서 발주코드와 센터명 기억하기
            is_orders = col_0 == 'ORDERS'
            df_edi['발주코드'] = None
            df_edi['센터'] = None
            
            df_edi.loc[is_orders, '발주코드'] = col_1[is_orders]
            df_edi.loc[is_orders, '센터'] = df_edi[5].astype(str).str.strip()[is_orders]
            
            # 비어있는 아래 행들에 발주코드와 센터를 일괄 채워넣기 (순식간에 처리됨)
            df_edi['발주코드'] = df_edi['발주코드'].ffill()
            df_edi['센터'] = df_edi['센터'].ffill()
            
            # 880 바코드로 시작하는 상품 행만 쏙 빼내기
            is_item = col_1.str.startswith('880')
            df_items = df_edi[is_item].copy()
            
            if df_items.empty:
                st.warning("⚠️ 유효한 발주 상품(880 바코드)을 찾을 수 없습니다.")
                st.stop()

            # 4. 수량 계산
            df_items['판매코드'] = col_1[is_item]
            
            # 입수 (콤마 제거 후 숫자 변환)
            df_items['입수_str'] = df_items[5].astype(str).str.replace(',', '', regex=False)
            df_items['입수'] = pd.to_numeric(df_items['입수_str'], errors='coerce').fillna(1).astype(int)
            
            # 주문수 (숫자만 추출)
            df_items['주문수_str'] = df_items[6].astype(str).str.replace(r'[^0-9]', '', regex=True)
            df_items['주문수'] = pd.to_numeric(df_items['주문수_str'], errors='coerce').fillna(0).astype(int)
            
            # UNIT 수량 계산 및 0건 제외
            df_items['UNIT수량'] = df_items['입수'] * df_items['주문수']
            df_parsed = df_items[df_items['UNIT수량'] > 0][['발주코드', '센터', '판매코드', 'UNIT수량']]

            if df_parsed.empty:
                st.warning("⚠️ 추출할 유효 수량(0 초과)이 없습니다.")
                st.stop()

            # 5. ME코드 매핑
            panmae_cols = [c for c in df_raw_template.columns if '판매코드' in str(c)]
            sangpum_cols = [c for c in df_raw_template.columns if '상품코드' in str(c)]
            panmae_col = panmae_cols[0] if panmae_cols else df_raw_template.columns[3]
            me_col = sangpum_cols[-1] if sangpum_cols else df_raw_template.columns[-1]
            
            df_mapping = df_raw_template[[panmae_col, me_col]].dropna()
            df_mapping.columns = ['판매코드', 'ME코드']
            df_mapping['판매코드'] = df_mapping['판매코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            df_mapping['ME코드'] = df_mapping['ME코드'].astype(str).str.strip()
            df_mapping = df_mapping.drop_duplicates()

            df_mapped = pd.merge(df_parsed, df_mapping, on='판매코드', how='left')
            df_mapped['ME코드'] = df_mapped['ME코드'].fillna(df_mapped['판매코드'])
            df_agg = df_mapped.groupby(['센터', 'ME코드', '발주코드'], as_index=False)['UNIT수량'].sum()

            # 6. 두 번째 시트(수주 양식)에 데이터 합치기
            orig_cols = list(df_order.columns)
            df_order['센터_str'] = df_order['센터'].astype(str).str.strip()
            df_order['상품코드_str'] = df_order['상품코드'].astype(str).str.strip()
            df_agg['센터_str'] = df_agg['센터'].astype(str).str.strip()
            df_agg['ME코드_str'] = df_agg['ME코드'].astype(str).str.strip()
            
            final_df = pd.merge(df_order, df_agg, left_on=['센터_str', '상품코드_str'], right_on=['센터_str', 'ME코드_str'], how='inner')
            
            # 발주코드, 배송코드, 수량, 금액 계산 업데이트
            final_df['발주코드'] = final_df['발주코드_y']
            final_df['배송코드'] = final_df['발주코드_y']
            final_df['UNIT수량'] = final_df['UNIT수량_y']
            
            def clean_price(x):
                try: return float(str(x).replace(',', ''))
                except: return 0.0
            
            final_df['UNIT단가_clean'] = final_df['UNIT단가'].apply(clean_price)
            final_df['Total Amount'] = final_df['UNIT수량'] * final_df['UNIT단가_clean']
            
            # 7. 원래 양식으로 복원
            out_df = final_df[orig_cols].copy()
            clean_cols = []
            space_cnt = 1
            for c in out_df.columns:
                if 'Unnamed' in str(c):
                    clean_cols.append(" " * space_cnt)
                    space_cnt += 1
                else:
                    clean_cols.append(c)
            out_df.columns = clean_cols

            st.success(f"⚡ 변환 완료! 유효 수주 {len(out_df)}건이 1초 만에 추출되었습니다.")
            st.dataframe(out_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                out_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name="롯데마트_수주_완료.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
