import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="옵션 판매 비중 분석기",
    page_icon="📊",
    layout="wide"
)

st.title("📊 상품별 묶음 & 옵션 판매 비중 분석기")
st.caption("카페24 주문 목록의 복잡한 옵션 텍스트를 정제하여 인기 옵션 조합 및 단품별 판매 비중을 분석합니다.")

with st.expander("ℹ️ 사용 가이드 및 유의사항"):
    st.markdown("""
    ### 💡 **주요 기능**
    * **옵션 텍스트 자동 파싱**: 카페24의 텍스트형 옵션(`색상: RED / 사이즈: XL`)에서 추가금 및 불필요한 특수문자 정제
    * **인기 옵션 조합 TOP 20**: 가장 높은 매출과 수량을 기록한 효자 옵션 순위 집계
    * **상품별 단품 옵션 비중 분석**: 특정 상품 내에서 어떤 색상/사이즈가 잘 팔리는지 퍼센트(%) 비중 산출
    """)

uploaded_file = st.file_uploader("카페24 주문 목록 엑셀 파일 업로드 (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        # 열 유연 매핑
        prod_col = next((c for c in ['상품명', '상품명(필수)', '품목명'] if c in df.columns), None)
        opt_col = next((c for c in ['옵션정보', '옵션', '상품옵션', '품목명'] if c in df.columns), None)
        qty_col = next((c for c in ['수량', '주문수량', '구매수량'] if c in df.columns), None)
        price_col = next((c for c in ['총 결제금액', '결제금액', '상품구매금액', '주문금액'] if c in df.columns), None)

        if not prod_col or not qty_col:
            st.error("❌ 필수 열('상품명', '수량')을 찾을 수 없습니다. 엑셀 형식을 확인해 주세요.")
            st.stop()

        # 전처리
        df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(1)
        if price_col:
            df[price_col] = pd.to_numeric(df[price_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        else:
            df['결제금액'] = 0
            price_col = '결제금액'

        if opt_col and opt_col in df.columns:
            df['정제옵션'] = df[opt_col].astype(str).fillna('옵션없음')
            # 괄호 안 추가금 텍스트 제거 정규식
            df['정제옵션'] = df['정제옵션'].str.replace(r'\(\+[\d,]+원\)', '', regex=True).str.strip()
        else:
            df['정제옵션'] = '단일옵션'

        # 1. 전체 인기 옵션 TOP 20
        top_options = df.groupby([prod_col, '정제옵션']).agg(
            총판매수량=(qty_col, 'sum'),
            총매출액=(price_col, 'sum'),
            주문건수=(qty_col, 'count')
        ).reset_index().sort_values(by='총판매수량', ascending=False)

        # 2. 상품별 옵션 비중(%) 산출
        prod_totals = df.groupby(prod_col)[qty_col].sum().to_dict()
        top_options['상품내_판매비중(%)'] = top_options.apply(
            lambda row: round((row['총판매수량'] / prod_totals[row[prod_col]]) * 100, 1) if prod_totals.get(row[prod_col], 0) > 0 else 0,
            axis=1
        )

        st.divider()
        st.subheader("🔥 인기 옵션 조합 TOP 20")
        st.dataframe(top_options.head(20), use_container_width=True)

        st.divider()
        st.subheader("🔍 특정 상품별 옵션 비중 상세 조회")
        selected_product = st.selectbox("분석할 상품을 선택하세요", df[prod_col].unique())

        if selected_product:
            detail_df = top_options[top_options[prod_col] == selected_product].copy()
            
            c1, c2 = st.columns([1, 1])
            with c1:
                st.write(f"**[{selected_product}] 옵션별 판매 수량/비중**")
                st.dataframe(detail_df[['정제옵션', '총판매수량', '상품내_판매비중(%)', '총매출액']], use_container_width=True)

            with c2:
                st.write("**옵션 판매 비중 차트**")
                chart_data = detail_df.set_index('정제옵션')['총판매수량']
                st.bar_chart(chart_data)

        # 엑셀 다운로드
        st.divider()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            top_options.to_excel(writer, sheet_name='옵션별판매비중분석', index=False)

        st.download_button(
            label="📄 옵션 비중 분석 리포트 엑셀 다운로드",
            data=output.getvalue(),
            file_name="cafe24_option_share_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"데이터 분석 중 오류가 발생했습니다: {e}")