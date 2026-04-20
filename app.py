import streamlit as st
import pandas as pd
import io
import os
import re

st.set_page_config(page_title="롯데마트 수주 자동화", layout="wide")

st.title("📦 롯데마트 수주 자동 변환기")
st.write("EDI Raw Data(라떼는.xlsx)를 업로드하면 0건을 제외한 수주 내역을 추출합니다.")

# 💡 깃허브에 있는 고정 서식 파일 (템플릿)
TEMPLATE_FILE = '2022 롯데마트 서식파일 260417납품.xlsx'

if not os.path.exists(TEMPLATE_FILE):
    st.error(f"⚠️ '{TEMPLATE_FILE}' 파일을 찾을 수 없습니다. 깃허브에 파일이 있는지 확인해주세요.")
    st.stop()

uploaded_file = st.file_uploader("📥 EDI Raw Data 업로드", type=['xlsx', 'csv'])

if uploaded_file is not None:
    with st.spinner("데이터 매칭 및 유효 건 추출 중..."):
        try:
            # 1. 업로드된 EDI Raw 데이터 파싱 (안전한 텍스트 기반 추출)
            if uploaded_file.name.endswith('.csv'):
                df_edi = pd.read_csv(uploaded_file, header=None)
            else:
                df_edi = pd.read_excel(uploaded_file, header=None)
            
            parsed_list = []
            curr_center = ""
            curr_doc = ""
            
            for i, row in df_edi.iterrows():
                r = [str(x).strip() for x in row.tolist()]
                
                # 'ORDERS' 행에서 문서번호와 센터명 추출
                if r[0] == 'ORDERS':
                    curr_doc = r[1].replace('.0', '')
                    curr_center = r[5]
                    continue
                
                # 상품 행 (880 바코드 기준)
                val_code = r[1].replace('.0', '')
                if val_code.startswith('880'):
                    # 입수 (콤마 제거 후 숫자 변환)
                    ipsu_str = r[5].replace(',', '')
                    ipsu = int(float(ipsu_str)) if ipsu_str.replace('.', '').isdigit() else 1
                    
                    # 주문수 (글자 제거 후 숫자만 추출)
                    qty_str = r[6]
                    qty_nums = re.sub(r'[^0-9]', '', qty_str)
                    qty = int(qty_nums) if qty_nums else 0
                    
                    unit_qty = ipsu * qty
                    
                    if unit_qty > 0:
                        parsed_list.append({
                            '발주코드': curr_doc,
                            '센터': curr_center,
                            '판매코드': val_code,
                            'UNIT수량': unit_qty
                        })
            
            df_parsed = pd.DataFrame(parsed_list)
            
            if df_parsed.empty:
                st.warning("⚠️ 추출할 유효 수량(0 초과)이 없습니다.")
                st.stop()

            # 2. 템플릿 1번 시트(RAW)에서 바코드 -> ME코드 매핑 정보 가져오기
            df_raw_template = pd.read_excel(TEMPLATE_FILE, sheet_name=0)
            
            # 컬럼 이름이 유동적일 수 있으므로 '판매코드', '상품코드' 키워드로 열 찾기
            panmae_cols = [c for c in df_raw_template.columns if '판매코드' in str(c)]
            sangpum_cols = [c for c in df_raw_template.columns if '상품코드' in str(c)]
            
            panmae_col = panmae_cols[0] if panmae_cols else df_raw_template.columns[3]
            me_col = sangpum_cols[-1] if sangpum_cols else df_raw_template.columns[-1]
            
            df_mapping = df_raw_template[[panmae_col, me_col]].dropna()
            df_mapping.columns = ['판매코드', 'ME코드']
            df_mapping['판매코드'] = df_mapping['판매코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
            df_mapping['ME코드'] = df_mapping['ME코드'].astype(str).str.strip()
            df_mapping = df_mapping.drop_duplicates()

            # EDI 데이터에 ME코드 붙이기
            df_mapped = pd.merge(df_parsed, df_mapping, on='판매코드', how='left')
            # 맵핑 실패 시 원본 바코드 유지
            df_mapped['ME코드'] = df_mapped['ME코드'].fillna(df_mapped['판매코드'])
            
            # 센터와 ME코드 기준으로 수량 합산 (중복 방지)
            df_agg = df_mapped.groupby(['센터', 'ME코드', '발주코드'], as_index=False)['UNIT수량'].sum()

            # 3. 템플릿 2번 시트(수주)와 매칭하여 최종 결과 만들기
            df_order = pd.read_excel(TEMPLATE_FILE, sheet_name=1)
            orig_cols = list(df_order.columns) # 원본 양식(빈 열 포함) 기억하기
            
            # 매칭을 위해 공백 제거
            df_order['센터_str'] = df_order['센터'].astype(str).str.strip()
            df_order['상품코드_str'] = df_order['상품코드'].astype(str).str.strip()
            
            df_agg['센터_str'] = df_agg['센터'].astype(str).str.strip()
            df_agg['ME코드_str'] = df_agg['ME코드'].astype(str).str.strip()
            
            # 유효 수량이 있는 데이터만 2번 시트 양식과 이너 조인(교집합)
            final_df = pd.merge(df_order, df_agg, left_on=['센터_str', '상품코드_str'], right_on=['센터_str', 'ME코드_str'], how='inner')
            
            # 4. 값 갱신 및 Total Amount 계산
            final_df['발주코드'] = final_df['발주코드_y'] # EDI에서 뽑은 문서번호로 덮어쓰기
            final_df['배송코드'] = final_df['발주코드_y']
            final_df['UNIT수량'] = final_df['UNIT수량_y']
            
            # 단가 계산 (콤마 등 불순물 제거)
            def clean_price(x):
                try:
                    return float(str(x).replace(',', ''))
                except:
                    return 0.0
            
            final_df['UNIT단가_clean'] = final_df['UNIT단가'].apply(clean_price)
            final_df['Total Amount'] = final_df['UNIT수량'] * final_df['UNIT단가_clean']
            
            # 5. 원래 2번 시트 양식(컬럼)으로 되돌리기
            out_df = final_df[orig_cols].copy()
            
            # Unnamed(빈 열) 이름을 다시 공백으로 변환하여 엑셀 양식 완벽 유지
            clean_cols = []
            space_cnt = 1
            for c in out_df.columns:
                if 'Unnamed' in str(c):
                    clean_cols.append(" " * space_cnt)
                    space_cnt += 1
                else:
                    clean_cols.append(c)
            out_df.columns = clean_cols

            st.success(f"변환 완료! 유효 수주 {len(out_df)}건이 추출되었습니다.")
            st.dataframe(out_df, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                out_df.to_excel(writer, index=False, sheet_name='롯데마트 수주')
            
            st.download_button(
                label="📥 최종 수주 파일 다운로드",
                data=buffer.getvalue(),
                file_name=f"롯데마트_수주_완료_{curr_doc}.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
            st.info("데이터 파싱 중 문제가 발생했습니다. 관리자에게 에러 메시지를 알려주세요.")
