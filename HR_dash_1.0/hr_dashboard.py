# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import sys
import locale

# UTF-8 인코딩 설정
if sys.platform.startswith('win'):
    locale.setlocale(locale.LC_ALL, 'Korean_Korea.949')

# 페이지 설정
st.set_page_config(
    page_title="HR 퇴직자 설문조사 대시보드",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 스타일 적용 - 한글 폰트 추가
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
    
    * {
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    .main > div {
        padding-top: 2rem;
    }
    
    /* 메트릭 카드 스타일 - 밝은 배경 */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 15px 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    /* 메트릭 라벨 색상 */
    div[data-testid="metric-container"] label {
        color: #555555 !important;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    /* 메트릭 값 색상 */
    div[data-testid="metric-container"] div[data-testid="metric-value"] {
        color: #1f77b4 !important;
        font-weight: bold;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    /* 사이드바 스타일 */
    .css-1d391kg {
        background-color: #f8f9fa;
    }
    
    /* 제목 색상 */
    h1, h2, h3 {
        color: #2c3e50 !important;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    /* 일반 텍스트 색상 */
    p, span, div {
        color: #333333 !important;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    /* 차트 배경 */
    .js-plotly-plot {
        background-color: #ffffff !important;
    }
    
    /* 정보 박스 스타일 개선 */
    .stAlert {
        background-color: #ffffff;
        border: 1px solid;
        color: #333333;
    }
    
    /* Plotly 차트 텍스트 색상 */
    .plotly text {
        fill: #333333 !important;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    /* 탭 스타일 */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #f8f9fa;
    }
    
    .stTabs [data-baseweb="tab"] {
        color: #333333 !important;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
    
    /* 멀티셀렉트 스타일 */
    .stMultiSelect label {
        color: #333333 !important;
        font-family: 'Noto Sans KR', sans-serif !important;
    }
</style>
""", unsafe_allow_html=True)

# Plotly 한글 폰트 설정
font_dict = dict(
    family="Noto Sans KR, sans-serif",
    size=12,
    color="#333333"
)

# 데이터 로드 함수
@st.cache_data
def load_data(file):
    if file is not None:
        # Excel 파일 읽기 시 인코딩 처리
        df = pd.read_excel(file, engine='openpyxl')
        
        # 컬럼명 정리
        df.columns = df.columns.str.strip()
        
        # 숫자형 컬럼 변환
        numeric_columns = [
            '근속연수', '직무경험', '재입사여부', '지인 입사권유',
            '경영철학', '다양성', '리더십', '커뮤니케이션', '성과관리 제도',
            '상사역량', '직속상사 관계', '동일부서 직원 관계', '타부서 직원 관계',
            '역할과 책임', '역량개발 기회', '보상', '복리후생', '근무환경', '근무시간', '회사위치'
        ]
        
        for col in numeric_columns:
            if col in df.columns:
                # 문자열을 숫자로 변환, 변환할 수 없는 값은 NaN으로 처리
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    return None

# 메인 타이틀
st.title("👥 HR Dashboard")
st.markdown("### 퇴직자 설문조사 분석 대시보드")

# 사이드바 - 파일 업로드
with st.sidebar:
    st.header("📊 데이터 업로드")
    uploaded_file = st.file_uploader("Excel 파일을 선택하세요", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        
        if df is not None:
            st.success("✅ 데이터가 성공적으로 로드되었습니다!")
            
            # 필터 섹션
            st.markdown("---")
            st.header("🔍 필터")
            
            # 부서 필터
            departments = st.multiselect(
                "부서 선택:",
                options=df['부서'].unique().tolist(),
                default=df['부서'].unique().tolist()
            )
            
            # 성별 필터
            genders = st.multiselect(
                "성별 선택:",
                options=df['성별'].unique().tolist(),
                default=df['성별'].unique().tolist()
            )
            
            # 직위 필터
            positions = st.multiselect(
                "직위 선택:",
                options=df['직위'].unique().tolist(),
                default=df['직위'].unique().tolist()
            )
            
            # 필터 적용
            filtered_df = df[
                (df['부서'].isin(departments)) & 
                (df['성별'].isin(genders)) & 
                (df['직위'].isin(positions))
            ]
    else:
        st.info("👈 Excel 파일을 업로드해주세요")
        filtered_df = None

# 메인 대시보드
if filtered_df is not None and len(filtered_df) > 0:
    # KPI 메트릭
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="📋 총 응답 수",
            value=f"{len(filtered_df):,}",
            delta=None
        )
    
    with col2:
        if '근속연수' in filtered_df.columns and not filtered_df['근속연수'].isna().all():
            avg_tenure = filtered_df['근속연수'].mean()
            st.metric(
                label="⏱️ 평균 근속연수",
                value=f"{avg_tenure:.1f}년" if not pd.isna(avg_tenure) else "N/A",
                delta=None
            )
        else:
            st.metric(
                label="⏱️ 평균 근속연수",
                value="N/A",
                delta=None
            )
    
    with col3:
        # 재입사 의향 평균 (5점 만점)
        if '재입사여부' in filtered_df.columns and not filtered_df['재입사여부'].isna().all():
            avg_return = filtered_df['재입사여부'].mean()
            if not pd.isna(avg_return):
                st.metric(
                    label="🔄 재입사 의향",
                    value=f"{avg_return:.1f}/5.0",
                    delta=f"{(avg_return-3)/3*100:.1f}%" if avg_return > 3 else f"{(avg_return-3)/3*100:.1f}%"
                )
            else:
                st.metric(label="🔄 재입사 의향", value="N/A", delta=None)
        else:
            st.metric(label="🔄 재입사 의향", value="N/A", delta=None)
    
    with col4:
        # 추천 의향 평균
        if '지인 입사권유' in filtered_df.columns and not filtered_df['지인 입사권유'].isna().all():
            avg_recommend = filtered_df['지인 입사권유'].mean()
            if not pd.isna(avg_recommend):
                st.metric(
                    label="👥 추천 의향",
                    value=f"{avg_recommend:.1f}/5.0",
                    delta=f"{(avg_recommend-3)/3*100:.1f}%" if avg_recommend > 3 else f"{(avg_recommend-3)/3*100:.1f}%"
                )
            else:
                st.metric(label="👥 추천 의향", value="N/A", delta=None)
        else:
            st.metric(label="👥 추천 의향", value="N/A", delta=None)
    
    st.markdown("---")
    
    # 첫 번째 행: 기본 분포 차트
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # 성별 분포
        gender_dist = filtered_df['성별'].value_counts()
        fig_gender = px.pie(
            values=gender_dist.values,
            names=gender_dist.index,
            title="성별 분포",
            hole=0.4,
            color_discrete_map={'남': '#3498db', '여': '#e74c3c'}
        )
        fig_gender.update_traces(textposition='inside', textinfo='percent+label')
        fig_gender.update_layout(
            height=300, 
            margin=dict(t=40, b=0, l=0, r=0),
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=font_dict
        )
        st.plotly_chart(fig_gender, use_container_width=True)
    
    with col2:
        # 부서별 직원 분포
        dept_dist = filtered_df['부서'].value_counts()
        fig_dept = px.pie(
            values=dept_dist.values,
            names=dept_dist.index,
            title="부서별 직원 분포",
            hole=0.0
        )
        fig_dept.update_traces(textposition='inside', textinfo='percent+label')
        fig_dept.update_layout(
            height=300, 
            margin=dict(t=40, b=0, l=0, r=0),
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=font_dict
        )
        st.plotly_chart(fig_dept, use_container_width=True)
    
    with col3:
        # 직위별 분포
        position_dist = filtered_df['직위'].value_counts()
        fig_position = px.bar(
            x=position_dist.index,
            y=position_dist.values,
            title="직위별 분포",
            labels={'x': '직위', 'y': '인원 수'},
            color=position_dist.values,
            color_continuous_scale='Reds'
        )
        fig_position.update_layout(
            height=300, 
            margin=dict(t=40, b=0, l=0, r=0), 
            showlegend=False,
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=font_dict,
            xaxis=dict(gridcolor='#e0e0e0'),
            yaxis=dict(gridcolor='#e0e0e0')
        )
        st.plotly_chart(fig_position, use_container_width=True)
    
    # 두 번째 행: 만족도 분석
    st.markdown("---")
    st.subheader("📊 항목별 만족도 분석")
    
    # 만족도 항목 정의
    satisfaction_cols = [
        '경영철학', '다양성', '리더십', '커뮤니케이션', '성과관리 제도',
        '상사역량', '직속상사 관계', '동일부서 직원 관계', '타부서 직원 관계',
        '역할과 책임', '역량개발 기회', '보상', '복리후생', '근무환경', '근무시간', '회사위치'
    ]
    
    # 카테고리별 그룹핑
    categories = {
        '조직문화': ['경영철학', '다양성', '리더십', '커뮤니케이션'],
        '인사제도': ['성과관리 제도', '역량개발 기회', '보상', '복리후생'],
        '업무환경': ['역할과 책임', '근무환경', '근무시간', '회사위치'],
        '인간관계': ['상사역량', '직속상사 관계', '동일부서 직원 관계', '타부서 직원 관계']
    }
    
    col1, col2 = st.columns(2)
    
    with col1:
        # 전체 항목별 평균 점수
        avg_scores = filtered_df[satisfaction_cols].mean().sort_values(ascending=True)
        
        fig_scores = px.bar(
            y=avg_scores.index,
            x=avg_scores.values,
            orientation='h',
            title="항목별 평균 만족도",
            labels={'x': '평균 점수 (5점 만점)', 'y': '항목'},
            color=avg_scores.values,
            color_continuous_scale='RdYlGn',
            range_color=[1, 5]
        )
        fig_scores.update_layout(
            height=500, 
            margin=dict(t=40, b=0, l=0, r=0),
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=font_dict,
            xaxis=dict(gridcolor='#e0e0e0'),
            yaxis=dict(gridcolor='#e0e0e0')
        )
        fig_scores.add_vline(x=3, line_dash="dash", line_color="gray", opacity=0.5)
        st.plotly_chart(fig_scores, use_container_width=True)
    
    with col2:
        # 카테고리별 평균 점수
        category_scores = {}
        for cat, cols in categories.items():
            valid_cols = [col for col in cols if col in filtered_df.columns]
            if valid_cols:
                category_scores[cat] = filtered_df[valid_cols].mean().mean()
        
        cat_df = pd.DataFrame(list(category_scores.items()), columns=['카테고리', '평균점수'])
        
        fig_cat = px.bar(
            cat_df,
            x='카테고리',
            y='평균점수',
            title="카테고리별 평균 만족도",
            color='평균점수',
            color_continuous_scale='RdYlGn',
            range_color=[1, 5]
        )
        fig_cat.update_layout(
            height=500, 
            margin=dict(t=40, b=0, l=0, r=0),
            plot_bgcolor='white',
            paper_bgcolor='white',
            font=font_dict,
            xaxis=dict(gridcolor='#e0e0e0'),
            yaxis=dict(gridcolor='#e0e0e0')
        )
        fig_cat.add_hline(y=3, line_dash="dash", line_color="gray", opacity=0.5)
        st.plotly_chart(fig_cat, use_container_width=True)
    
    # 세 번째 행: 상세 분석
    st.markdown("---")
    st.subheader("🔍 상세 분석")
    
    tab1, tab2, tab3 = st.tabs(["부서별 분석", "직위별 분석", "근속연수별 분석"])
    
    with tab1:
        # 부서별 주요 지표 비교
        valid_satisfaction_cols = [col for col in satisfaction_cols if col in filtered_df.columns]
        if valid_satisfaction_cols:
            dept_analysis = filtered_df.groupby('부서')[valid_satisfaction_cols].mean()
            
            # 히트맵
            fig_heatmap = px.imshow(
                dept_analysis.T,
                labels=dict(x="부서", y="평가 항목", color="평균 점수"),
                title="부서별 만족도 히트맵",
                color_continuous_scale='RdYlGn',
                aspect="auto"
            )
            fig_heatmap.update_layout(
                height=600,
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=font_dict
            )
            st.plotly_chart(fig_heatmap, use_container_width=True)
    
    with tab2:
        # 직위별 분석
        if valid_satisfaction_cols:
            position_analysis = filtered_df.groupby('직위')[valid_satisfaction_cols].mean()
            
            # 선택된 항목들에 대한 직위별 비교
            default_items = ['보상', '역량개발 기회', '근무환경', '직속상사 관계']
            default_items = [item for item in default_items if item in valid_satisfaction_cols][:4]
            
            selected_items = st.multiselect(
                "비교할 항목 선택:",
                valid_satisfaction_cols,
                default=default_items if default_items else valid_satisfaction_cols[:4]
            )
            
            if selected_items:
                fig_position_comp = go.Figure()
                
                for col in selected_items:
                    if col in position_analysis.columns:
                        fig_position_comp.add_trace(go.Bar(
                            name=col,
                            x=position_analysis.index,
                            y=position_analysis[col]
                        ))
                
                fig_position_comp.update_layout(
                    title="직위별 항목 비교",
                    xaxis_title="직위",
                    yaxis_title="평균 점수",
                    barmode='group',
                    height=400,
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=font_dict,
                    xaxis=dict(gridcolor='#e0e0e0'),
                    yaxis=dict(gridcolor='#e0e0e0')
                )
                st.plotly_chart(fig_position_comp, use_container_width=True)
    
    with tab3:
        # 근속연수별 분석
        if '근속연수' in filtered_df.columns and not filtered_df['근속연수'].isna().all():
            # 근속연수 구간 생성
            filtered_df['근속연수_구간'] = pd.cut(
                filtered_df['근속연수'],
                bins=[0, 2, 5, 10, float('inf')],
                labels=['2년 이하', '2-5년', '5-10년', '10년 이상']
            )
            
            tenure_cols = []
            if '재입사여부' in filtered_df.columns:
                tenure_cols.append('재입사여부')
            if '지인 입사권유' in filtered_df.columns:
                tenure_cols.append('지인 입사권유')
            
            if tenure_cols:
                tenure_analysis = filtered_df.groupby('근속연수_구간')[tenure_cols].mean()
                
                fig_tenure = go.Figure()
                
                if '재입사여부' in tenure_cols:
                    fig_tenure.add_trace(go.Scatter(
                        x=tenure_analysis.index.astype(str),
                        y=tenure_analysis['재입사여부'],
                        mode='lines+markers',
                        name='재입사 의향',
                        line=dict(width=3)
                    ))
                
                if '지인 입사권유' in tenure_cols:
                    fig_tenure.add_trace(go.Scatter(
                        x=tenure_analysis.index.astype(str),
                        y=tenure_analysis['지인 입사권유'],
                        mode='lines+markers',
                        name='추천 의향',
                        line=dict(width=3)
                    ))
                
                fig_tenure.update_layout(
                    title="근속연수별 재입사/추천 의향",
                    xaxis_title="근속연수",
                    yaxis_title="평균 점수 (5점 만점)",
                    height=400,
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=font_dict,
                    xaxis=dict(gridcolor='#e0e0e0'),
                    yaxis=dict(gridcolor='#e0e0e0')
                )
                st.plotly_chart(fig_tenure, use_container_width=True)
    
    # 네 번째 행: 인사이트
    st.markdown("---")
    st.subheader("💡 주요 인사이트")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # 가장 만족도가 높은 항목
        if len(avg_scores) >= 3:
            top_items = avg_scores.nlargest(3)
            st.info(f"""
            **✅ 만족도 상위 3개 항목**
            1. {top_items.index[0]}: {top_items.values[0]:.2f}점
            2. {top_items.index[1]}: {top_items.values[1]:.2f}점
            3. {top_items.index[2]}: {top_items.values[2]:.2f}점
            """)
    
    with col2:
        # 가장 만족도가 낮은 항목
        if len(avg_scores) >= 3:
            bottom_items = avg_scores.nsmallest(3)
            st.warning(f"""
            **⚠️ 개선 필요 항목**
            1. {bottom_items.index[0]}: {bottom_items.values[0]:.2f}점
            2. {bottom_items.index[1]}: {bottom_items.values[1]:.2f}점
            3. {bottom_items.index[2]}: {bottom_items.values[2]:.2f}점
            """)
    
    with col3:
        # 재입사/추천 의향과 상관관계가 높은 항목
        if '재입사여부' in filtered_df.columns and valid_satisfaction_cols:
            correlations = filtered_df[valid_satisfaction_cols].corrwith(filtered_df['재입사여부']).sort_values(ascending=False)
            if len(correlations) >= 3:
                top_corr = correlations.head(3)
                st.success(f"""
                **🔗 재입사 의향과 상관관계**
                1. {top_corr.index[0]}: {top_corr.values[0]:.3f}
                2. {top_corr.index[1]}: {top_corr.values[1]:.3f}
                3. {top_corr.index[2]}: {top_corr.values[2]:.3f}
                """)

else:
    if uploaded_file is None:
        st.info("👈 사이드바에서 Excel 파일을 업로드해주세요")
    else:
        st.warning("선택된 필터에 해당하는 데이터가 없습니다. 필터를 조정해주세요.")

# 푸터
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray;'>
        <p>HR 퇴직자 설문조사 대시보드 | Powered by Streamlit & Plotly</p>
    </div>
    """,
    unsafe_allow_html=True
)