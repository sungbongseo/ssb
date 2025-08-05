# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import math

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 📁 파일 경로
file_path = r"C:\Users\rcnd\OneDrive\바탕 화면\회사\2025\설계\데이터분석자료\퇴직자설문조사1.0.xlsx"
df = pd.read_excel(file_path)

# 🔍 컬럼 확인
print("📋 전체 컬럼:", df.columns.tolist())

# 🔎 퇴사 관련 컬럼 자동 감지
resign_cols = [c for c in df.columns if '퇴사' in c or '재직' in c]
dept_col = [c for c in df.columns if '부서' in c][0]
resign_col = resign_cols[0] if resign_cols else None
print(f"🔍 부서컬럼: {dept_col}, 퇴사컬럼: {resign_col}")

# ✔️ 퇴사자 필터링
if resign_col:
    df_resigned = df[df[resign_col].astype(str).str.contains('퇴사|Y|1|yes|True', case=False, na=False)]
else:
    df_resigned = df.copy()

df_finance_resigned = df_resigned[df_resigned[dept_col]=='재무팀']

# 📊 분석 대상 컬럼
num_cols = df_finance_resigned.select_dtypes(include=[np.number]).columns.tolist()
satisfaction_cols = [c for c in df_finance_resigned.columns if '만족' in str(c)]
analysis_cols = list(set(num_cols + satisfaction_cols))

# 🎯 평균 비교 분석
mean_scores = df_finance_resigned[analysis_cols].mean()
overall_avg = df[df[dept_col]=='재무팀'][analysis_cols].mean()
diff = mean_scores - overall_avg

comparative = pd.DataFrame({
    '퇴사자 평균': mean_scores,
    '재무팀 전체 평균': overall_avg,
    '차이(퇴사자–전체)': diff
}).sort_values('차이(퇴사자–전체)')

low5 = comparative.head(5).copy()
low5['중요도순위'] = low5['재무팀 전체 평균'].rank(ascending=False)
low5['우선순위점수'] = low5['중요도순위'] * abs(low5['차이(퇴사자–전체)'])

# 📈 레이더 차트
labels = low5.index.tolist()
resigned_vals = low5['퇴사자 평균'].tolist()
overall_vals = low5['재무팀 전체 평균'].tolist()
resigned_vals += resigned_vals[:1]; overall_vals += overall_vals[:1]
angles = [n/float(len(labels))*2*math.pi for n in range(len(labels)+1)]

plt.figure(figsize=(8,8))
ax = plt.subplot(111, polar=True)
ax.plot(angles, resigned_vals, label="퇴사자 평균", marker='o')
ax.plot(angles, overall_vals, label="전체 평균", marker='o') 
ax.fill(angles, resigned_vals, alpha=0.25)
ax.fill(angles, overall_vals, alpha=0.25)
ax.set_xticks(angles[:-1])
ax.set_xticklabels(labels)
ax.set_title("재무팀 퇴사자 vs 전체 평균 레이더 차트", pad=20)
ax.legend()
plt.tight_layout()
plt.show()

# 🔗 상관관계 히트맵
plt.figure(figsize=(12, 8))
sns.heatmap(df_finance_resigned[analysis_cols].corr(), annot=True, cmap='coolwarm', fmt=".2f", square=True)
plt.title("📌 재무팀 퇴사자 상관관계 히트맵")
plt.tight_layout()
plt.show()

# 📊 피벗 테이블 (직책별 평균)
pivot_col = [c for c in df.columns if '직급' in c or '직책' in c]
if pivot_col:
    pivot_summary = df_finance_resigned.pivot_table(values=analysis_cols, index=pivot_col[0], aggfunc='mean')
    print("\n🧩 직급/직책별 평균 피벗 테이블:")
    print(pivot_summary)

# 🔍 인사이트 출력
print("\n🎯 개선이 시급한 항목 Top 3:")
for item in low5.sort_values('우선순위점수', ascending=False).head(3).index:
    row = low5.loc[item]
    print(f"- {item}: 퇴사자 {row['퇴사자 평균']:.2f}, 전체 {row['재무팀 전체 평균']:.2f}, 격차 {row['차이(퇴사자–전체)']:.2f}")
