import pandas as pd
import streamlit as st
import plotly.express as px

# 페이지 설정
st.set_page_config(
    page_title="HR Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 1) 파일 업로드
st.sidebar.header("📂 HR 데이터 파일 업로드")
uploaded_file = st.sidebar.file_uploader(
    "CSV 또는 XLSX 파일을 업로드하세요",
    type=["csv", "xlsx"],
    help="예: 퇴직자설문조사1.0.xlsx"
)

if uploaded_file is not None:
    # 2) 파일 읽기
    if uploaded_file.name.endswith("xlsx"):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)
    
    # 3) 설문 항목 리스트 (숫자평가된 변수들)
    survey_vars = [
        "직무경험", "재입사여부", "지인 입사권유", "경영철학", "다양성", "리더십",
        "커뮤니케이션", "성과관리 제도", "상사역량", "직속상사 관계", "동일부서 직원 관계",
        "타부서 직원 관계", "역할과 책임", "역량개발 기회", "보상", "복리후생",
        "근무환경", "근무시간", "회사위치"
    ]
    # 4) 해당 변수들 숫자형으로 변환 (문자열→NaN 처리)
    for var in survey_vars:
        if var in df.columns:
            df[var] = pd.to_numeric(df[var], errors="coerce")
    
    # 5) 사이드바 필터
    st.sidebar.header("🔎 필터링")
    dept_sel = st.sidebar.multiselect("부서 선택", options=df["부서"].unique(), default=df["부서"].unique())
    gender_sel = st.sidebar.multiselect("성별 선택", options=df["성별"].unique(), default=df["성별"].unique())
    pos_sel = st.sidebar.multiselect("직위 선택", options=df["직위"].unique(), default=df["직위"].unique())
    df = df[df["부서"].isin(dept_sel) & df["성별"].isin(gender_sel) & df["직위"].isin(pos_sel)]
    
    # 6) 그룹화 기준 선택
    group_by = st.sidebar.radio("🔘 그룹화 기준 선택", ("부서", "성별", "직위"))
    
    # 7) 화면 타이틀
    st.title("👥 HR Dashboard")
    st.write("종합 인사 데이터 분석 대시보드")
    
    # 8) 그룹화 변수 분포 도넛 차트
    counts = df[group_by].value_counts().reset_index()
    counts.columns = [group_by, "count"]
    fig_donut = px.pie(
        counts, names=group_by, values="count",
        hole=0.5,
        title=f"{group_by} 분포"
    )
    fig_donut.update_traces(textinfo="percent+label")
    st.plotly_chart(fig_donut, use_container_width=True)
    
    st.markdown("---")
    
    # 9) 3열 레이아웃으로 평균값 바 차트 표시
    for i in range(0, len(survey_vars), 3):
        cols = st.columns(3)
        for j, var in enumerate(survey_vars[i:i+3]):
            with cols[j]:
                if var in df.columns:
                    grouped = df.groupby(group_by)[var].mean().reset_index()
                    fig = px.bar(
                        grouped, x=group_by, y=var,
                        labels={var: f"평균 {var}"}, title=f"{var} 평균"
                    )
                    fig.update_layout(xaxis_title=None, yaxis_title=f"평균 {var}")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.write(f"⚠️ 데이터에 '{var}' 열이 없습니다.")
else:
    st.info("좌측 사이드바에서 HR 데이터(CSV 또는 XLSX)를 업로드해주세요.")
