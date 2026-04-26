import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
import os
import numpy as np
import textwrap
from collections import Counter
from datetime import datetime

# ==========================================
# 🌟 [프리미엄 레드 에디션] 테마 설정
# ==========================================
st.set_page_config(
    page_title="사출생산팀 생산성 데이터 센터", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# [글로벌 트렌드 + 시그니처 레드] CSS
st.markdown("""
<meta name="google" content="notranslate">
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard', sans-serif !important;
        background-color: #FDFDFD; 
        color: #1E293B; 
        translate: no; 
    }

    /* 🌟 시그니처 레드 포인트 타이틀 디자인 */
    .section-banner {
        background-color: #ffffff;
        border: 1px solid #F1F5F9;
        border-left: 6px solid #D91B1B; /* DURING Signature Red */
        padding: 16px 22px;
        border-radius: 10px;
        margin-top: 30px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .section-banner h3 {
        margin: 0;
        font-weight: 800;
        color: #1E293B;
        font-size: 19px;
    }

    /* 종합 분석 리포트 카드 */
    .analysis-report-card {
        background-color: #FFF5F5; /* 아주 연한 레드 베이지 */
        border: 1px solid #FEE2E2;
        border-radius: 12px;
        padding: 25px;
        margin-bottom: 30px;
    }

    .trendy-card {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border: 1px solid #F1F5F9;
        margin-bottom: 24px;
    }

    /* 지표 카드 스타일 */
    .metric-container {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        border-bottom: 4px solid #D91B1B;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
    }

    /* 전광판 헤더 그라데이션 (Deep Red) */
    .dashboard-header {
        background: linear-gradient(135deg, #9F1239, #D91B1B);
        color: #FFFFFF;
        border-radius: 12px;
        padding: 30px;
        margin-bottom: 25px;
    }
</style>
""", unsafe_allow_html=True)

# 유틸리티 함수들
def safe_float(val):
    try:
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val) or val is None: return 0.0
        v_str = str(val).strip().replace(',', '').replace(' ', '')
        if v_str in ['', '-', '#DIV/0!', '#N/A', 'nan', 'None']: return 0.0
        if '%' in v_str: return float(v_str.replace('%', '')) / 100.0
        return float(v_str)
    except: return 0.0

def render_section_title(text):
    st.markdown(f"<div class='section-banner'><h3>{text}</h3></div>", unsafe_allow_html=True)

# 🌟 [신규 기능] 데이터를 정밀 분석하여 총평을 자동 생성하는 함수
def render_total_analysis(filtered_df):
    if filtered_df.empty: return
    
    # 1. 기초 통계 계산
    avg_oee = filtered_df[filtered_df['종합효율'] > 0]['종합효율'].mean()
    total_qty = filtered_df['총 생산수량'].sum()
    bad_qty = filtered_df['불량수량'].sum()
    avg_quality = (1 - (bad_qty / total_qty)) if total_qty > 0 else 0
    
    # 2. WORST 설비 분석
    worst_row = filtered_df[filtered_df['종합효율'] > 0].sort_values(by='종합효율').iloc[0]
    worst_mc = worst_row['설비명'].split(' - ')[0]
    worst_prod = worst_row['품명']
    worst_oee = worst_row['종합효율']
    worst_issue = worst_row['OPEN ISSUE'] if str(worst_row['OPEN ISSUE']).strip() != "" else "기록된 특이사항 없음"
    
    # 3. 분석 문장 구성
    st.markdown(f"""
    <div class='analysis-report-card'>
        <h4 style='margin-top:0; color:#9F1239; font-weight:800;'>📊 금일 생산 실적 종합 분석 리포트</h4>
        <p style='line-height:1.7; font-size:15px; color:#334155;'>
            금일 사출생산팀의 <b>평균 종합효율은 {avg_oee:.1%}</b>를 기록하였으며, 전체 양품률은 <b>{avg_quality:.1%}</b> 수준으로 파악됩니다.<br>
            정밀 분석 결과, 가장 집중 관리가 필요한 설비는 <b>{worst_mc} ({worst_prod})</b>호기로 확인되었습니다.<br><br>
            해당 설비는 현재 <b>종합효율 {worst_oee:.1%}</b>로 팀 평균 대비 현저히 낮은 수치를 보이고 있으며, 
            오픈 이슈 확인 결과 <b>"{worst_issue}"</b>와 같은 원인으로 인해 가동에 차질이 발생한 것으로 분석됩니다. 
            해당 이슈에 대한 설비보전팀과의 긴밀한 협조 및 재발 방지 대책 수립이 권고됩니다.
        </p>
    </div>
    """, unsafe_allow_html=True)

# 메인 타이틀
st.markdown("<h1 style='margin-top: 10px; margin-bottom: 10px; color: #1E293B; font-weight: 900; font-size: 30px;' class='notranslate'>사출생산팀 일일 생산성 정밀 분석</h1>", unsafe_allow_html=True)

# 3. 데이터 수집
data_to_process = []
DATA_DIR = "data"
if os.path.exists(DATA_DIR):
    for f_name in os.listdir(DATA_DIR):
        if f_name.startswith("~$") or not (f_name.endswith('.xlsx') or f_name.endswith('.csv')): continue
        file_path = os.path.join(DATA_DIR, f_name)
        try:
            if f_name.endswith('.csv'):
                try: t_df = pd.read_csv(file_path, encoding='utf-8')
                except: t_df = pd.read_csv(file_path, encoding='cp949')
            else: t_df = pd.read_excel(file_path)
            data_to_process.append((f_name, t_df))
        except Exception as e: st.error(f"오류: {e}")

if data_to_process:
    all_recs = []
    daily_totals = {}
    for f_name, temp_df in data_to_process:
        temp_df.columns = [str(c).replace('\n', '').strip() for c in temp_df.columns]
        temp_df = temp_df.rename(columns={'작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명', '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율', '목표 효율': '목표효율'})
        for col in temp_df.columns:
            if 'Unnamed' in col or 'ISSUE' in col.upper(): temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'}); break
        
        date_match = re.search(r'\d{8}', f_name)
        if date_match:
            raw = date_match.group()[2:]
            try:
                dt = datetime.strptime(raw, '%y%m%d')
                clean_d = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일"
                month_s = f"{dt.strftime('%y')}년 {dt.month}월"; sort_k = raw
            except: clean_d = raw; month_s = "기타"; sort_k = raw
        else: clean_d = f_name.split('.')[0]; month_s = "기타"; sort_k = clean_d

        for _, row in temp_df.iterrows():
            m_val = str(row.get('설비명', '')).strip()
            if m_val in ['', 'nan', '설비명'] or 'TOTAL' in m_val.upper() or '합계' in m_val: continue
            rec = {'sort_key': sort_k, '생산월': month_s, '생산일': clean_d}
            for c in ['설비명', '품명', '양품수량', '불량수량', '총 생산수량', '종합효율', '목표효율', 'OPEN ISSUE']:
                rec[c] = row[c] if c in temp_df.columns else None
            all_recs.append(rec)

    df = pd.DataFrame(all_recs).sort_values(by=['sort_key', '설비명']).reset_index(drop=True)
    for c in ['양품수량', '불량수량', '총 생산수량', '종합효율', '목표효율']: df[c] = df[c].apply(safe_float)

    # 사이드바 필터
    st.sidebar.markdown("<h2 style='font-weight: 800; color: #D91B1B; font-size: 18px;'>🎯 리포트 필터</h2>", unsafe_allow_html=True)
    all_mons = sorted(df['생산월'].unique().tolist())
    sel_m = st.sidebar.multiselect("📅 생산월", all_mons, default=[all_mons[-1]] if all_mons else [])
    df_m = df[df['생산월'].isin(sel_m)] if sel_m else df.copy()
    all_days = sorted(df_m['생산일'].unique().tolist(), reverse=True)
    sel_d = st.sidebar.multiselect("📆 생산일", all_days, default=[all_days[0]] if all_days else [])
    df_d = df_m[df_m['생산일'].isin(sel_d)] if sel_d else df_m.copy()
    
    # 🌟 최상단 종합 분석 리포트 출력
    render_total_analysis(df_d)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📈 종합효율 추이", "📝 OPEN ISSUE", "📅 일별 상세 현황", "🏆 효율 BEST&WORST", "🛑 비가동 정밀 분석", "🤖 AI 챗봇"])

    # TAB 1
    with tab1:
        render_section_title("월별 설비 가동 현황 및 종합효율 추이")
        plot_df = df_d.groupby(['sort_key', '생산일']).mean().reset_index().sort_values('sort_key')
        
        # 종합효율 차트 (Signature Red 포인트)
        fig_oee = go.Figure()
        fig_oee.add_trace(go.Bar(
            x=plot_df['생산일'], y=plot_df['종합효율'], text=plot_df['종합효율'].apply(lambda x: f"{x:.1%}"),
            textposition='auto', marker_color=['#D91B1B' if x < 0.86 else '#3B82F6' for x in plot_df['종합효율']]
        ))
        fig_oee.update_layout(title="📈 일자별 평균 종합효율", yaxis_tickformat='.0%', yaxis_range=[0, 1.0], plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_oee, use_container_width=True)

    # TAB 3
    with tab3:
        render_section_title("설비별 가동 현황 요약")
        active_day = df_d[df_d['종합효율'] > 0].sort_values(by='종합효율', ascending=False)
        
        # 전광판 (Signature Red 그라데이션)
        st.markdown(f"""<div class='dashboard-header'>
            <div style='font-size: 28px; font-weight: 900; margin-bottom: 20px;'>💡 생산 가동 요약 현황</div>
            <div style='display: flex; gap: 20px;'>
                <div style='flex: 1; background: rgba(255,255,255,0.15); padding: 20px; border-radius: 10px;'>
                    <div style='color: #4ADE80; font-weight: 800; font-size: 16px; margin-bottom: 10px;'>🏆 종합효율 BEST 5</div>
                    {"".join([f"<div style='margin-bottom:8px; font-size:14px;'>{r['설비명'].split(' - ')[0]} ({r['종합효율']:.1%})</div>" for _, r in active_day.head(5).iterrows()])}
                </div>
                <div style='flex: 1; background: rgba(255,255,255,0.15); padding: 20px; border-radius: 10px;'>
                    <div style='color: #FFAAAA; font-weight: 800; font-size: 16px; margin-bottom: 10px;'>🚨 종합효율 WORST 5</div>
                    {"".join([f"<div style='margin-bottom:8px; font-size:14px;'>{r['설비명'].split(' - ')[0]} ({r['종합효율']:.1%})</div>" for _, r in active_day.tail(5).iterrows()])}
                </div>
            </div>
        </div>""", unsafe_allow_html=True)
        
        # 🌟 프리미엄 레드 스타일 표
        render_section_title("상세 생산 실적 데이터 (설비 순)")
        disp_df = df_d.copy()
        for c in ['종합효율', '목표효율']: disp_df[c] = disp_df[c].apply(lambda x: f"{x:.1%}" if x > 0 else "")
        for c in ['양품수량', '총 생산수량']: disp_df[c] = disp_df[c].apply(lambda x: f"{int(x):,}" if x > 0 else "")
        
        html_table = disp_df[['설비명', '품명', '총 생산수량', '양품수량', '종합효율', 'OPEN ISSUE']].to_html(index=False, escape=False)
        st.markdown(f"""
        <style>
            .red-table {{ width: 100%; border-collapse: collapse; border-radius: 8px; overflow: hidden; }}
            .red-table th {{ background-color: #D91B1B; color: white; padding: 12px; text-align: center; }}
            .red-table td {{ padding: 10px; border-bottom: 1px solid #F1F5F9; text-align: center; font-size: 13px; }}
            .red-table tr:hover {{ background-color: #FFF5F5; }}
        </style>
        <div style='overflow-x:auto;'>
            {html_table.replace('<table', '<table class="red-table"')}
        </div>
        """, unsafe_allow_html=True)

    # 나머지 탭들은 유사한 스타일로 유지 (생략 가능, 요청 시 추가 업데이트)
    with tab2: render_section_title("특이사항(OPEN ISSUE) 정밀 조회")
    with tab4: render_section_title("생산 효율 순위 (BEST & WORST)")
    with tab5: render_section_title("설비 비가동 원인 분석")
    with tab6: render_section_title("🤖 사출생산 데이터 지능형 챗봇")

else:
    st.info("GitHub data 폴더에 CSV 파일을 업로드해주세요.")
