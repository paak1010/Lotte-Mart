import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기 (초고속 안정화 버전)")
st.write("EDI Raw Data를 업로드하면 0건을 제외한 수주 내역을 추출합니다.")

TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ '{TEMPLATE_FILE}' 파일을 찾을 수 없습니다. 깃허브에 파일이 있는지 확인해주세요.")
    st.stop()

@st.cache_data
def load_template_sheets():
    df_raw_template = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
    df_order = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
    return df_raw_template, df_order

uploaded_file = st.file_uploader("📥 EDI Raw Data 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("충돌 없이 데이터를 완벽하게 병합 중입니다... 🚀"):
        try:
            # 1. 템플릿 로드
            df_raw_template, df_order = load_template_sheets()

            # 2. 업로드된 EDI 데이터 파싱
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            df_edi = df_edi.dropna(how='all')

            # 3. 데이터 일괄 파싱 (초고속 벡터화)
            col_0 = df_edi[0].astype(str).str.strip()
            col_1 = df_edi[1].astype(str).str.replace('.0', '', regex=False).str.strip()
            
            is_orders = col_0 == 'ORDERS'
            df_edi['발주코드'] = None
            df_edi['센터'] = None
            
            df_edi.loc[is_orders, '발주코드'] = col_1[is_orders]
            df_edi.loc[is_orders, '센터'] = df_edi[5].astype(str).str.strip()[is_orders]
            
            df_edi['발주코드'] = df_edi['발주코드'].ffill()
            df_edi['센터'] = df_edi['센터'].ffill()
            
            is_item = col_1.str.startswith('880')
            df_items = df_edi[is_item].copy()
            
            if df_items.empty:
                st.warning("⚠️ 유효한 발주 상품(880 바코드)을 찾을 수 없습니다.")
                st.stop()

            # 4. 수량 계산
            df_items['판매코드'] = col_1[is_item]
            
            df_items['입수_str'] = df_items[5].astype(str).str.replace(',', '', regex=False)
            df_items['입수'] = pd.to_numeric(df_items['입수_str'], errors='coerce').fillna(1).astype(int)
            
            df_items['주문수_str'] = df_items[6].astype(str).str.replace(r'[^0-9]', '', regex=True)
            df_items['주문수'] = pd.to_numeric(df_items['주문수_str'], errors='coerce').fillna(0).astype(int)
            
            df_items['UNIT수량'] = df_items['입수'] * df_items['주문수']
            df_parsed = df_items[df_items['UNIT수량'] > 0][['발주코드', '센터', '판매코드', 'UNIT수량']]

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

            # 6. 두 번째 시트(수주 양식)에 데이터 합치기 (충돌 완벽 방지 🛡️)
            orig_cols = list(df_order.columns) # 원래 엑셀 양식 기억하기
            
            # 유연한 컬럼 찾기 (양식이 조금 바뀌어도 알아서 찾아냄)
            def find_col(df, keywords):
                for kw in keywords:
                    for c in df.columns:
                        if kw in str(c): return c
                return None
            
            c_center = find_col(df_order, ['센터', '점포'])
            c_item = find_col(df_order, ['상품코드', 'ME코드'])
            c_qty = find_col(df_order, ['수량', 'UNIT'])
            c_price = find_col(df_order, ['단가', 'Price'])
            c_total = find_col(df_order, ['Total Amount', '금액'])
            c_order_no = find_col(df_order, ['발주코드', '주문번호'])
            c_deliv_no = find_col(df_order, ['배송코드'])

            df_order['센터_str'] = df_order[c_center].astype(str).str.strip()
            df_order['상품코드_str'] = df_order[c_item].astype(str).str.strip()
            
            # 병합 전 충돌을 막기 위해 추출한 데이터의 컬럼 이름을 안전하게 변경!
            df_agg_merge = df_agg.copy()
            df_agg_merge['센터_str'] = df_agg_merge['센터'].astype(str).str.strip()
            df_agg_merge['ME코드_str'] = df_agg_merge['ME코드'].astype(str).str.strip()
            
            # 핵심! 기존 양식의 '센터'와 이름이 겹치지 않도록 골라내서 가져옴
            df_agg_safe = df_agg_merge[['센터_str', 'ME코드_str', 'UNIT수량', '발주코드']].rename(
                columns={'UNIT수량': 'EDI_수량', '발주코드': 'EDI_발주'}
            )
            
            final_df = pd.merge(df_order, df_agg_safe, left_on=['센터_str', '상품코드_str'], right_on=['센터_str', 'ME코드_str'], how='inner')
            
            # 발주코드, 배송코드, 수량, 금액 덮어쓰기
            if c_order_no: final_df[c_order_no] = final_df['EDI_발주']
            if c_deliv_no: final_df[c_deliv_no] = final_df['EDI_발주']
            if c_qty: final_df[c_qty] = final_df['EDI_수량']
            
            def clean_price(x):
                try: return float(str(x).replace(',', ''))
                except: return 0.0
            
            if c_qty and c_price and c_total:
                final_df['UNIT단가_clean'] = final_df[c_price].apply(clean_price)
                final_df[c_total] = final_df[c_qty] * final_df['UNIT단가_clean']
            
            # 7. 원래 양식으로 복원 (이제 에러 날 일이 없습니다)
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

            st.success(f"✨ 변환 완료! 유효 수주 {len(out_df)}건이 성공적으로 추출되었습니다.")
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
