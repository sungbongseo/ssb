# -*- coding: utf-8 -*-
# ================================================================
# HR 설문 Dash 대시보드
#  - A: KPI 카드
#  - C: 레이더차트 비교
#  - D: Top/Bottom 지표
#  - E: 시계열/근속기간 추세(Line Chart)
# ================================================================
# 설치:  pip install dash plotly pandas openpyxl
# 실행:  cd C:/Users/rcnd/Documents/HR_Dash
#        .\venv\Scripts\Activate
#        python app.py
# ================================================================

import pathlib
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from dash import Dash, dcc, html, Input, Output, State, ctx, no_update

# -----------------------------
# 높이 설정
# -----------------------------
GRAPH_H_SMALL = 220   # 필터/상세
GRAPH_H_MED   = 320   # 평균 막대
GRAPH_H_RADAR = 360   # 레이더
GRAPH_H_TREND = 320   # 시계열/근속 추세

# -----------------------------
# 데이터 로드
# -----------------------------
DATA_PATH = pathlib.Path(__file__).parent / "퇴직자설문조사1.0.xlsx"
df_raw = pd.read_excel(DATA_PATH)

# -----------------------------
# 컬럼 정의
# -----------------------------
FILTER_COLS = ["부서", "성별", "직위"]
METRIC_COLS = [
    "직무경험", "재입사여부", "지인 입사권유", "경영철학", "다양성", "리더십", "커뮤니케이션",
    "성과관리 제도", "상사역량", "직속상사 관계", "동일부서 직원 관계", "타부서 직원 관계",
    "역할과 책임", "역량개발 기회", "보상", "복리후생", "근무환경", "근무시간", "회사위치"
]

# (선택) 시계열/근속 분석용 컬럼명 지정
# 실제 파일에 맞게 바꿔 주세요. 없으면 None 그대로 두면 기능이 비활성.
DATE_COL   = "설문일자"    # 예: '2025-06-30' 형식 날짜열
TENURE_COL = "근속기간(년)"  # 숫자형 근속(년) 열

# 숫자 변환
for c in METRIC_COLS:
    df_raw[c] = pd.to_numeric(df_raw[c], errors="coerce")

# 날짜/근속 처리
if DATE_COL in df_raw.columns:
    df_raw[DATE_COL] = pd.to_datetime(df_raw[DATE_COL], errors="coerce")
else:
    DATE_COL = None

if TENURE_COL in df_raw.columns:
    df_raw[TENURE_COL] = pd.to_numeric(df_raw[TENURE_COL], errors="coerce")
else:
    TENURE_COL = None

# -----------------------------
# Dash 앱
# -----------------------------
external_css = ["https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css"]
app = Dash(__name__, external_stylesheets=external_css, title="HR Dashboard", suppress_callback_exceptions=True)

# -----------------------------
# 레이아웃
# -----------------------------
app.layout = html.Div(
    className="container-fluid",
    children=[
        html.H3("HR 설문 대시보드", className="text-center mt-3 mb-4"),

        # 상태 저장
        dcc.Store(id="filter_state", data={"부서": [], "성별": [], "직위": []}),
        dcc.Store(id="current_metric", data=None),

        # 상단 필터 그래프
        html.Div(className="row", children=[
            html.Div(className="col-md-4 mb-3", children=[
                html.H6("부서 분포 (클릭 → 필터)", className="text-center"),
                dcc.Graph(id="dept_chart", style={"height": f"{GRAPH_H_SMALL}px"})
            ]),
            html.Div(className="col-md-4 mb-3", children=[
                html.H6("성별 분포 (클릭 → 필터)", className="text-center"),
                dcc.Graph(id="gender_chart", style={"height": f"{GRAPH_H_SMALL}px"})
            ]),
            html.Div(className="col-md-4 mb-3", children=[
                html.H6("직위 분포 (클릭 → 필터)", className="text-center"),
                dcc.Graph(id="position_chart", style={"height": f"{GRAPH_H_SMALL}px"})
            ]),
        ]),

        # KPI 카드 + 필터 상태 + 평균 막대
        html.Div(className="row", children=[
            html.Div(className="col-md-3 mb-3", children=[
                html.H6("현재 선택된 필터", className="mb-2"),
                html.Ul(id="current_filters", className="small"),
                html.Button("필터 초기화", id="reset_filters", className="btn btn-secondary btn-sm mt-2"),
                html.Hr(className="my-3"),
                html.H6("KPI", className="mb-2"),
                html.Div(id="kpi_cards", className="row gx-2")
            ]),
            html.Div(className="col-md-9 mb-3", children=[
                html.H6("지표별 평균 점수 (막대 클릭 → 상세)", className="text-center"),
                dcc.Graph(id="metric_chart", style={"height": f"{GRAPH_H_MED}px"})
            ])
        ]),

        # 탭 영역 : Radar / TopBottom / Trend
        dcc.Tabs(id="extra_tabs", value="radar", children=[
            dcc.Tab(label="레이더차트 비교", value="radar"),
            dcc.Tab(label="Top/Bottom 지표", value="topbottom"),
            dcc.Tab(label="시계열 · 근속 추세", value="trend"),
        ], className="mb-3"),
        html.Div(id="extra_tab_content"),

        # 상세 영역
        html.Div(className="row", children=[
            html.Div(className="col-12", children=[
                html.H6("선택 지표 분포 및 기초통계", className="text-center mb-2")
            ]),
            html.Div(className="col-md-6 mb-3", children=[
                dcc.Graph(id="metric_dist", style={"height": f"{GRAPH_H_SMALL}px"})
            ]),
            html.Div(className="col-md-6 mb-3 overflow-auto", children=[
                html.Div(id="metric_summary_table", className="mt-2")
            ]),
        ]),

        html.Footer("© 2025 HR Dashboard", className="text-center text-muted mb-3")
    ]
)

# -----------------------------
# 헬퍼
# -----------------------------
def make_count_bar(df, col, title):
    count_df = df[col].value_counts(dropna=False).reset_index()
    count_df.columns = [col, "count"]
    fig = px.bar(count_df, x=col, y="count", text="count", title=title)
    fig.update_traces(hovertemplate=f"{col}: %{{x}}<br>인원수: %{{y}}")
    fig.update_layout(margin=dict(l=5, r=5, t=40, b=60), xaxis_title="", yaxis_title="인원수", height=GRAPH_H_SMALL)
    return fig

def apply_filters(df, filters):
    out = df.copy()
    for k, vals in filters.items():
        if vals:
            out = out[out[k].isin(vals)]
    return out

def kpi_card(title, value):
    return html.Div(className="col-4 mb-2", children=[
        html.Div(className="card text-center p-2", children=[
            html.Div(title, className="small text-muted"),
            html.H5(value, className="mb-0")
        ], style={"minHeight": "70px"})
    ])

# -----------------------------
# 콜백 1) 상단 필터 그래프
# -----------------------------
@app.callback(
    Output("dept_chart", "figure"),
    Output("gender_chart", "figure"),
    Output("position_chart", "figure"),
    Input("filter_state", "data")
)
def update_filter_charts(_):
    return (
        make_count_bar(df_raw, "부서", "부서 분포"),
        make_count_bar(df_raw, "성별", "성별 분포"),
        make_count_bar(df_raw, "직위", "직위 분포")
    )

# -----------------------------
# 콜백 2) 필터 상태 업데이트 & 표시
# -----------------------------
@app.callback(
    Output("filter_state", "data"),
    Output("current_filters", "children"),
    Input("dept_chart", "clickData"),
    Input("gender_chart", "clickData"),
    Input("position_chart", "clickData"),
    Input("reset_filters", "n_clicks"),
    State("filter_state", "data"),
    prevent_initial_call=True
)
def update_filters(dept_click, gender_click, pos_click, reset_click, filter_state):
    trig = ctx.triggered_id

    if trig == "reset_filters":
        cleared = {"부서": [], "성별": [], "직위": []}
        return cleared, [html.Li("필터 없음")]

    def get_x(cd):
        if cd and "points" in cd and cd["points"]:
            return cd["points"][0]["x"]
        return None

    col_map = {"dept_chart": "부서", "gender_chart": "성별", "position_chart": "직위"}
    clicked_col = col_map.get(trig)
    clicked_val = None
    if trig == "dept_chart":
        clicked_val = get_x(dept_click)
    elif trig == "gender_chart":
        clicked_val = get_x(gender_click)
    elif trig == "position_chart":
        clicked_val = get_x(pos_click)

    if clicked_col and clicked_val is not None:
        cur = filter_state.get(clicked_col, [])
        if clicked_val in cur:
            cur.remove(clicked_val)
        else:
            cur.append(clicked_val)
        filter_state[clicked_col] = cur

    items = []
    for k, v in filter_state.items():
        if v:
            items.append(html.Li(f"{k}: {', '.join(map(str, v))}"))
    if not items:
        items = [html.Li("필터 없음")]

    return filter_state, items

# -----------------------------
# 콜백 3) KPI 카드
# -----------------------------
@app.callback(
    Output("kpi_cards", "children"),
    Input("filter_state", "data")
)
def update_kpi(filter_state):
    df = apply_filters(df_raw, filter_state)
    if df.empty:
        return [html.Div("필터 결과가 없습니다.", className="text-danger")]
    means = df[METRIC_COLS].mean()
    top3 = means.sort_values(ascending=False).head(3).round(2)
    bottom3 = means.sort_values().head(3).round(2)
    overall = round(means.mean(), 2)

    cards = [
        kpi_card("전체 평균", overall),
        kpi_card(f"최고: {top3.index[0]}", top3.iloc[0]),
        kpi_card(f"최저: {bottom3.index[0]}", bottom3.iloc[0]),
    ]
    return cards

# -----------------------------
# 콜백 4) 지표 평균 막대
# -----------------------------
@app.callback(
    Output("metric_chart", "figure"),
    Input("filter_state", "data")
)
def update_metric_chart(filter_state):
    df = apply_filters(df_raw, filter_state)
    if df.empty:
        fig = px.bar(title="필터 결과가 없습니다.")
        fig.update_layout(height=GRAPH_H_MED)
        return fig

    mean_series = df[METRIC_COLS].mean().sort_values(ascending=False).round(2)
    mean_df = mean_series.reset_index()
    mean_df.columns = ["지표", "평균점수"]

    fig = px.bar(mean_df, x="지표", y="평균점수", text="평균점수", title="지표 평균 점수")
    fig.update_traces(hovertemplate="지표: %{x}<br>평균: %{y}")
    fig.update_layout(margin=dict(l=5, r=5, t=40, b=100), xaxis_tickangle=-45, height=GRAPH_H_MED)
    return fig

# -----------------------------
# 콜백 5) 막대 클릭 -> metric 저장
# -----------------------------
@app.callback(
    Output("current_metric", "data"),
    Input("metric_chart", "clickData"),
    prevent_initial_call=True
)
def store_metric(metric_click):
    if metric_click and "points" in metric_click and metric_click["points"]:
        return metric_click["points"][0].get("x")
    return no_update

# -----------------------------
# 콜백 6) 지표 상세 (분포 + 통계)
# -----------------------------
@app.callback(
    Output("metric_dist", "figure"),
    Output("metric_summary_table", "children"),
    Input("current_metric", "data"),
    Input("filter_state", "data")
)
def update_metric_detail(metric, filter_state):
    try:
        df = apply_filters(df_raw, filter_state)

        if metric not in METRIC_COLS:
            fig = go.Figure()
            fig.update_layout(title_text="지표를 클릭하면 분포가 표시됩니다.", height=GRAPH_H_SMALL)
            msg = html.Div("지표를 선택해주세요.", className="text-muted")
            return fig, msg

        if df.empty:
            fig = go.Figure()
            fig.update_layout(title_text="필터 결과가 없습니다.", height=GRAPH_H_SMALL)
            msg = html.Div("필터 결과가 없습니다.", className="text-danger")
            return fig, msg

        series = df[metric].dropna()
        if series.empty:
            fig = go.Figure()
            fig.update_layout(title_text=f"{metric}: 유효 데이터가 없습니다.", height=GRAPH_H_SMALL)
            msg = html.Div("해당 지표에 값이 존재하지 않습니다.", className="text-warning")
            return fig, msg

        if pd.api.types.is_numeric_dtype(series):
            fig = px.histogram(series, x=series, nbins=5, title=f"{metric} 분포(필터 적용)")
            fig.update_traces(hovertemplate=f"{metric}: %{{x}}<br>Count: %{{y}}")
            fig.update_layout(margin=dict(l=5, r=5, t=40, b=40), xaxis_title=metric, yaxis_title="빈도", height=GRAPH_H_SMALL)

            desc = series.describe().round(2)
            rows = [html.Tr([html.Td(idx), html.Td(val)]) for idx, val in desc.items()]
            table = html.Table([
                html.Thead(html.Tr([html.Th("항목"), html.Th("값")], className="table-light")),
                html.Tbody(rows)
            ], className="table table-striped table-bordered table-hover table-sm")
        else:
            count_df = series.value_counts().reset_index()
            count_df.columns = [metric, "count"]
            fig = px.bar(count_df, x=metric, y="count", text="count", title=f"{metric} 분포(범주형)")
            fig.update_traces(hovertemplate=f"{metric}: %{{x}}<br>Count: %{{y}}")
            fig.update_layout(margin=dict(l=5, r=5, t=40, b=40), xaxis_title=metric, yaxis_title="빈도", height=GRAPH_H_SMALL)
            table = html.Div("숫자형 통계를 계산할 수 없습니다.", className="text-muted")

        return fig, table

    except Exception as e:
        err_fig = go.Figure()
        err_fig.update_layout(title_text="콜백 처리 중 오류 발생", height=GRAPH_H_SMALL)
        err_msg = html.Pre(f"Error: {type(e).__name__}: {e}")
        return err_fig, err_msg

# -----------------------------
# 콜백 7) Tabs 컨텐츠 (Radar / TopBottom / Trend)
# -----------------------------
@app.callback(
    Output("extra_tab_content", "children"),
    Input("extra_tabs", "value"),
    Input("filter_state", "data")
)
def render_extra_tab(tab, filter_state):
    df = apply_filters(df_raw, filter_state)
    if df.empty:
        return html.Div("필터 결과가 없습니다.", className="text-danger")

    # -------- Radar --------
    if tab == "radar":
        return html.Div([
            html.Div(className="row mb-2", children=[
                html.Div(className="col-md-3", children=[
                    html.Label("비교 기준"),
                    dcc.Dropdown(
                        id="radar_group_col",
                        options=[{"label": c, "value": c} for c in FILTER_COLS],
                        value="부서",
                        clearable=False
                    )
                ]),
                html.Div(className="col-md-9", children=[
                    html.Label("비교할 그룹(최대 3개)"),
                    dcc.Dropdown(
                        id="radar_groups",
                        options=[],  # 콜백에서 채움
                        value=[],
                        multi=True
                    )
                ])
            ]),
            dcc.Graph(id="radar_chart", style={"height": f"{GRAPH_H_RADAR}px"})
        ])

    # -------- Top/Bottom --------
    if tab == "topbottom":
        means = df[METRIC_COLS].mean().round(2)
        top5 = means.sort_values(ascending=False).head(5)
        bottom5 = means.sort_values().head(5)

        def make_table(title, s):
            rows = [html.Tr([html.Td(idx), html.Td(val)]) for idx, val in s.items()]
            return html.Div(className="col-md-6 mb-3", children=[
                html.H6(title, className="text-center"),
                html.Table([
                    html.Thead(html.Tr([html.Th("지표"), html.Th("평균")], className="table-light")),
                    html.Tbody(rows)
                ], className="table table-striped table-bordered table-hover table-sm")
            ])

        return html.Div(className="row", children=[
            make_table("Top 5 지표", top5),
            make_table("Bottom 5 지표", bottom5)
        ])

    # -------- Trend --------
    if tab == "trend":
        # Trend UI
        controls = []
        if DATE_COL:
            controls.append(html.Div(className="col-md-4 mb-2", children=[
                html.Label("기간 그룹핑 (월/분기/연)"),
                dcc.Dropdown(
                    id="trend_date_group",
                    options=[
                        {"label": "월별", "value": "M"},
                        {"label": "분기별", "value": "Q"},
                        {"label": "연도별", "value": "Y"},
                    ],
                    value="M",
                    clearable=False
                )
            ]))
        if TENURE_COL:
            controls.append(html.Div(className="col-md-4 mb-2", children=[
                html.Label("근속 구간 폭(년)"),
                dcc.Input(id="tenure_bin", type="number", min=0.5, step=0.5, value=1, className="form-control")
            ]))

        controls.append(html.Div(className="col-md-4 mb-2", children=[
            html.Label("표시 지표 (1개)"),
            dcc.Dropdown(
                id="trend_metric",
                options=[{"label": m, "value": m} for m in METRIC_COLS],
                value=METRIC_COLS[0],
                clearable=False
            )
        ]))

        graphs = [
            dcc.Graph(id="trend_time_graph", style={"height": f"{GRAPH_H_TREND}px"}) if DATE_COL else html.Div(),
            dcc.Graph(id="trend_tenure_graph", style={"height": f"{GRAPH_H_TREND}px"}) if TENURE_COL else html.Div()
        ]

        return html.Div([
            html.Div(className="row", children=controls),
            html.Div(graphs)
        ])

    return html.Div("탭 오류", className="text-danger")

# -----------------------------
# 콜백 8-1) 레이더 그룹 옵션
# -----------------------------
@app.callback(
    Output("radar_groups", "options"),
    Output("radar_groups", "value"),
    Input("radar_group_col", "value"),
    Input("filter_state", "data"),
    prevent_initial_call=True
)
def update_radar_group_options(group_col, filter_state):
    df = apply_filters(df_raw, filter_state)
    if group_col not in FILTER_COLS or df.empty:
        return [], []
    opts = sorted(df[group_col].dropna().unique().tolist())
    default = opts[:2]
    return [{"label": o, "value": o} for o in opts], default

# -----------------------------
# 콜백 8-2) 레이더 차트
# -----------------------------
@app.callback(
    Output("radar_chart", "figure"),
    Input("radar_group_col", "value"),
    Input("radar_groups", "value"),
    Input("filter_state", "data"),
    prevent_initial_call=True
)
def update_radar_chart(group_col, groups, filter_state):
    df = apply_filters(df_raw, filter_state)
    fig = go.Figure()
    fig.update_layout(height=GRAPH_H_RADAR, title="레이더차트")
    if not group_col or not groups or df.empty:
        return fig

    # 스케일을 0~5 혹은 데이터 max 기준으로 잡을지 결정
    r_max = df[METRIC_COLS].max().max()
    fig.update_layout(polar=dict(radialaxis=dict(range=[0, r_max])))

    for g in groups[:3]:
        mean_vals = df[df[group_col] == g][METRIC_COLS].mean()
        fig.add_trace(go.Scatterpolar(
            r=mean_vals.values,
            theta=METRIC_COLS,
            fill='toself',
            name=str(g)
        ))
    return fig

# -----------------------------
# 콜백 9) Trend 그래프(시간/근속)
# -----------------------------
@app.callback(
    Output("trend_time_graph", "figure"),
    Output("trend_tenure_graph", "figure"),
    Input("trend_metric", "value"),
    Input("trend_date_group", "value"),
    Input("tenure_bin", "value"),
    Input("filter_state", "data"),
    prevent_initial_call=True
)
def update_trend(metric, date_group, tenure_bin, filter_state):
    df = apply_filters(df_raw, filter_state)

    # 기본 빈 Figure
    empty_fig = go.Figure()
    empty_fig.update_layout(height=GRAPH_H_TREND)

    # 시간 그래프
    if DATE_COL and metric in METRIC_COLS and not df.empty:
        df_date = df[[DATE_COL, metric]].dropna()
        if not df_date.empty:
            # 리샘플링
            df_date = df_date.set_index(DATE_COL).sort_index()
            rule = date_group if date_group in ["M", "Q", "Y"] else "M"
            agg = df_date.resample(rule)[metric].mean().reset_index()
            fig_time = px.line(agg, x=DATE_COL, y=metric, markers=True, title=f"{metric} 시계열 추세")
            fig_time.update_layout(height=GRAPH_H_TREND, margin=dict(l=5, r=5, t=40, b=40))
        else:
            fig_time = empty_fig
    else:
        fig_time = empty_fig

    # 근속 그래프
    if TENURE_COL and metric in METRIC_COLS and not df.empty:
        df_ten = df[[TENURE_COL, metric]].dropna()
        if not df_ten.empty:
            # 구간 나누기
            bin_size = tenure_bin if tenure_bin and tenure_bin > 0 else 1
            max_ten = df_ten[TENURE_COL].max()
            bins = list(range(0, int(max_ten // bin_size + 2) * int(bin_size), int(bin_size)))
            labels = [f"{b}~{b+bin_size}" for b in bins[:-1]]
            df_ten["bin"] = pd.cut(df_ten[TENURE_COL], bins=bins, labels=labels, right=False, include_lowest=True)
            g = df_ten.groupby("bin")[metric].mean().reset_index()
            fig_ten = px.line(g, x="bin", y=metric, markers=True, title=f"{metric} 근속기간별 추세")
            fig_ten.update_layout(height=GRAPH_H_TREND, margin=dict(l=5, r=5, t=40, b=40), xaxis_title="근속구간(년)")
        else:
            fig_ten = empty_fig
    else:
        fig_ten = empty_fig

    return fig_time, fig_ten

# -----------------------------
# 실행
# -----------------------------
if __name__ == "__main__":
    run_fn = getattr(app, "run", None)  # Dash 3.x
    if callable(run_fn):
        run_fn(debug=True, host="0.0.0.0", port=8050)
    else:
        app.run_server(debug=True, host="0.0.0.0", port=8050)
