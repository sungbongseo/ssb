# -*- coding: utf-8 -*-
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import plotly.io as pio
import warnings

warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="HR 퇴직자 설문조사 대시보드",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@100;300;400;500;700;900&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
</style>
""", unsafe_allow_html=True)
pio.templates["custom"] = go.layout.Template(
    layout=go.Layout(font=dict(family="Noto Sans KR, sans-serif"))
)
pio.templates.default = "custom"

EVAL_COLS = [
    '직무경험', '재입사여부', '지인 입사권유', '경영철학', '다양성',
    '리더십', '커뮤니케이션', '성과관리 제도', '상사역량', '직속상사 관계',
    '동일부서 직원 관계', '타부서 직원 관계', '역할과 책임', '역량개발 기회',
    '보상', '복리후생', '근무환경', '근무시간', '회사위치'
]
REQUIRED_COLS = ['부서', '성별', '직위', '근속연수'] + EVAL_COLS

def validate_columns(df):
    missing = [col for col in REQUIRED_COLS if col not in df.columns]
    return missing

@st.cache_data
def load_sample_data():
    try:
        df = pd.read_excel('퇴직자설문조사1.0.xlsx', engine='openpyxl')
        return df
    except Exception:
        return None

st.markdown('<h1 style="font-weight:700;">👥 HR 퇴직자 설문조사 대시보드</h1>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("1. [필수] 퇴직자 설문조사 엑셀파일을 업로드하세요", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(f"파일을 읽는 도중 오류 발생: {e}")
        st.stop()
else:
    df = load_sample_data()
    if df is None:
        st.warning("엑셀 파일을 업로드해 주세요.")
        st.stop()

missing_cols = validate_columns(df)
if missing_cols:
    st.error(f"엑셀에 아래 컬럼이 누락되어 있습니다: {', '.join(missing_cols)}")
    st.stop()

for col in (['근속연수'] + EVAL_COLS):
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

st.sidebar.header("🔍 필터 설정")
dept_selection = st.sidebar.multiselect(
    "📁 부서 선택:", options=df["부서"].dropna().unique().tolist(), default=list(df["부서"].dropna().unique())
)
gender_selection = st.sidebar.multiselect(
    "👤 성별 선택:", options=df["성별"].dropna().unique().tolist(), default=list(df["성별"].dropna().unique())
)
position_selection = st.sidebar.multiselect(
    "💼 직위 선택:", options=df["직위"].dropna().unique().tolist(), default=list(df["직위"].dropna().unique())
)

df_filtered = df[
    df["부서"].isin(dept_selection) &
    df["성별"].isin(gender_selection) &
    df["직위"].isin(position_selection)
].copy()

if df_filtered.empty:
    st.warning("선택한 조건에 맞는 데이터가 없습니다.")
    st.stop()

for col in (['근속연수'] + EVAL_COLS):
    if col in df_filtered.columns:
        df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce')

# ====== KPI 지표 ======
st.markdown("### 📊 주요 지표")
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("총 응답자 수", f"{len(df_filtered)}명",
        f"{(len(df_filtered)/len(df)*100):.1f}%" if len(df) else "0%")
with col2:
    avg_tenure = df_filtered['근속연수'].mean()
    st.metric("평균 근속연수", f"{avg_tenure:.1f}년" if not np.isnan(avg_tenure) else "데이터 없음")
with col3:
    m = df_filtered[df_filtered["성별"] == "남"]
    st.metric("남성 비율", f"{(len(m)/len(df_filtered)*100):.1f}%" if len(df_filtered) else "0%")
with col4:
    f = df_filtered[df_filtered["성별"] == "여"]
    st.metric("여성 비율", f"{(len(f)/len(df_filtered)*100):.1f}%" if len(df_filtered) else "0%")
st.divider()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 기본 분석", "🎯 상세 분석", "🕸️ 레이더 차트", "📊 비교 분석", "🤖 K-Means 군집분석"
])

with tab1:
    c1, c2, c3 = st.columns(3)
    with c1:
        if not df_filtered["성별"].dropna().empty:
            counts = df_filtered["성별"].value_counts()
            fig = px.pie(values=counts.values, names=counts.index, title="성별 분포", hole=0.6)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("성별 데이터가 없습니다.")
    with c2:
        if not df_filtered["부서"].dropna().empty:
            counts = df_filtered["부서"].value_counts().head(10)
            fig = px.bar(
                x=counts.values, y=counts.index, orientation="h",
                title="부서별 인원(Top10)", color=counts.values, color_continuous_scale="Blues"
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("부서 데이터가 없습니다.")
    with c3:
        if not df_filtered["직위"].dropna().empty:
            counts = df_filtered["직위"].value_counts()
            fig = px.pie(values=counts.values, names=counts.index, title="직위별 분포")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("직위 데이터가 없습니다.")

with tab2:
    st.markdown("### 📊 평가 항목별 상세 분석")
    selected_items = st.multiselect("분석할 항목을 선택하세요:", options=EVAL_COLS, default=EVAL_COLS[:5])
    if selected_items:
        col1, col2 = st.columns(2)
        with col1:
            avg_scores = df_filtered[selected_items].apply(pd.to_numeric, errors='coerce').mean().sort_values(ascending=True)
            fig_avg = px.bar(
                x=avg_scores.values,
                y=avg_scores.index,
                orientation='h',
                title="선택 항목 평균 점수",
                color=avg_scores.values,
                color_continuous_scale='RdYlGn',
                range_color=[1, 5],
                height=500
            )
            fig_avg.add_vline(x=3, line_dash="dash", line_color="gray", annotation_text="기준선 (3점)")
            st.plotly_chart(fig_avg, use_container_width=True)
        with col2:
            fig_box = go.Figure()
            for item in selected_items:
                y = pd.to_numeric(df_filtered[item], errors='coerce')
                fig_box.add_trace(go.Box(y=y, name=item, boxmean='sd'))
            fig_box.update_layout(
                title="선택 항목 점수 분포",
                yaxis_title="점수",
                showlegend=False,
                font=dict(family="Noto Sans KR, sans-serif"),
                height=500
            )
            st.plotly_chart(fig_box, use_container_width=True)
    st.markdown("### 🔥 평가 항목 간 상관관계")
    corr_items = st.multiselect("상관분석 항목 선택", options=EVAL_COLS, default=EVAL_COLS[:8])
    if len(corr_items) > 1:
        corr_matrix = df_filtered[corr_items].apply(pd.to_numeric, errors='coerce').corr()
        fig_heatmap = px.imshow(
            corr_matrix,
            labels=dict(color="상관계수"),
            color_continuous_scale='RdBu',
            aspect="auto",
            title="평가 항목 간 상관관계 히트맵"
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

with tab3:
    st.markdown("### 🕸️ 레이더 차트 분석")
    c1, c2 = st.columns(2)
    with c1:
        avg = df_filtered[EVAL_COLS].apply(pd.to_numeric, errors='coerce').mean()
        fig = go.Figure(go.Scatterpolar(
            r=avg.values, theta=avg.index, fill="toself", name="전체 평균"
        ))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,5])), title="전체 평균 레이더", showlegend=True)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        sel_dept = st.multiselect("비교 부서 선택", options=list(df["부서"].dropna().unique()), default=list(df["부서"].dropna().unique())[:3])
        if sel_dept:
            fig = go.Figure()
            color_list = px.colors.qualitative.Plotly
            for i, d in enumerate(sel_dept):
                ddata = df[df["부서"] == d]
                if not ddata.empty:
                    avg_dept = ddata[EVAL_COLS].apply(pd.to_numeric, errors='coerce').mean()
                    fig.add_trace(go.Scatterpolar(
                        r=avg_dept.values, theta=avg_dept.index, fill="toself", name=f"{d}(n={len(ddata)})",
                        line_color=color_list[i%len(color_list)]
                    ))
            fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,5])), title="부서별 레이더", showlegend=True)
            st.plotly_chart(fig, use_container_width=True)

with tab4:
    st.markdown("### 📊 그룹별 비교 분석")
    group_by = st.selectbox("그룹 기준 선택", options=["부서", "성별", "직위"])
    analysis_items = st.multiselect("비교 평가 항목 선택", options=EVAL_COLS, default=EVAL_COLS[:6])
    if analysis_items:
        numeric_df = df_filtered.copy()
        for col in analysis_items:
            numeric_df[col] = pd.to_numeric(numeric_df[col], errors='coerce')
        grp = numeric_df.groupby(group_by)[analysis_items].mean()
        if not grp.empty:
            from plotly.subplots import make_subplots
            rows = int(np.ceil(len(analysis_items) / 2))
            fig = make_subplots(rows=rows, cols=2, subplot_titles=analysis_items)
            for idx, item in enumerate(analysis_items):
                row = idx // 2 + 1
                col_ = idx % 2 + 1
                fig.add_trace(
                    go.Bar(x=grp.index.astype(str), y=grp[item], name=item, marker=dict(line=dict(width=0))),
                    row=row, col=col_
                )
            fig.update_layout(
                title=f"{group_by}별 평가 비교",
                height=350*rows,
                showlegend=False,
                font=dict(family="Noto Sans KR, sans-serif"),
                margin=dict(t=60, b=20, r=10, l=10)
            )
            st.plotly_chart(fig, use_container_width=True)
            st.markdown(f"### 📋 {group_by}별 상세 점수")
            st.dataframe(grp.round(2), use_container_width=True)
        else:
            st.info("그룹별 평균 계산 결과 데이터가 없습니다.")

with tab5:
    st.markdown("### 🤖 K-Means 군집분석 (평가항목 기준)")
    # scikit-learn import 체크
    try:
        from sklearn.cluster import KMeans
        from sklearn.preprocessing import StandardScaler
        km_error = None
    except ImportError:
        KMeans = None
        km_error = "scikit-learn 패키지가 설치되어 있지 않습니다. 터미널에서 'pip install scikit-learn'을 실행해 주세요."
    if km_error:
        st.error(km_error)
    else:
        km_items = st.multiselect("군집화에 사용할 평가항목 선택", options=EVAL_COLS, default=EVAL_COLS[:8], key='kmeans')
        if len(km_items) >= 2:
            km_df = df_filtered[km_items].apply(pd.to_numeric, errors='coerce').dropna()
            if len(km_df) > 2:
                scaler = StandardScaler()
                X_scaled = scaler.fit_transform(km_df)
                inertias = []
                for k in range(1, min(11, len(km_df))):
                    kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
                    kmeans.fit(X_scaled)
                    inertias.append(kmeans.inertia_)
                st.markdown("#### Elbow Method (군집수 결정 참고)")
                fig_elbow = px.line(x=range(1, len(inertias)+1), y=inertias, markers=True,
                                    labels={"x": "군집수(K)", "y": "Inertia"},
                                    title="Elbow Method")
                st.plotly_chart(fig_elbow, use_container_width=True)
                n_clusters = st.slider("K-Means 군집수(K) 선택", min_value=2, max_value=min(8, len(km_df)), value=3)
                kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
                labels = kmeans.fit_predict(X_scaled)
                km_df_cluster = km_df.copy()
                km_df_cluster["군집"] = labels.astype(str)
                st.markdown("#### 2D Scatter (주요 2개 항목 기준)")
                main_x, main_y = km_items[0], km_items[1]
                fig_km = px.scatter(
                    km_df_cluster, x=main_x, y=main_y, color="군집", symbol="군집",
                    title="K-Means 2D Scatter",
                    labels={main_x: main_x, main_y: main_y}
                )
                st.plotly_chart(fig_km, use_container_width=True)
                st.markdown("#### 각 군집별 평균점수")
                st.dataframe(km_df_cluster.groupby("군집")[km_items].mean().round(2))
            else:
                st.info("군집분석에 사용할 데이터가 부족합니다.")
        else:
            st.info("군집분석은 2개 이상의 항목을 선택해야 합니다.")

st.divider()
st.markdown("### 💡 주요 인사이트")
col1, col2, col3 = st.columns(3)
avg_scores = df_filtered[EVAL_COLS].apply(pd.to_numeric, errors='coerce').mean()

# 1~3위, 하위 1~3위
top3 = avg_scores.sort_values(ascending=False).head(3)
bottom3 = avg_scores.sort_values().head(3)
with col1:
    st.info(f"**가장 높은 평가 항목 TOP 3**\n\n"
            + "\n".join([f"{i+1}위: {idx} ({val:.2f}점)" for i, (idx, val) in enumerate(top3.items())])
    )
with col2:
    st.warning(f"**개선 필요 항목 하위 3위**\n\n"
            + "\n".join([f"{i+1}위: {idx} ({val:.2f}점)" for i, (idx, val) in enumerate(bottom3.items())])
    )
with col3:
    corr = df_filtered[["근속연수"] + EVAL_COLS].apply(pd.to_numeric, errors='coerce').corr()["근속연수"].drop("근속연수")
    corr3 = corr.abs().sort_values(ascending=False).head(3)
    st.success("**근속연수와 상관 높은 항목 TOP 3**\n\n" +
        "\n".join([f"{i+1}위: {idx} ({corr[idx]:.2f})" for i, idx in enumerate(corr3.index)])
    )

st.divider()
st.markdown(
    "<div style='text-align:center; color:#666; font-size:0.9em;'>HR 퇴직자 설문조사 대시보드 | 데이터 기준: 업로드 파일</div>",
    unsafe_allow_html=True
)
