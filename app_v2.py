import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import os
from datetime import datetime

# ==========================================
# 🌟 1. 기본 설정 및 테마
# ==========================================
st.set_page_config(page_title="설비별 정밀 분석 대시보드", layout="wide", initial_sidebar_state="collapsed")

components.html(
    """<script>
        const parent = window.parent.document;
        parent.documentElement.lang = 'ko';
        parent.documentElement.setAttribute('translate', 'no');
        parent.body.classList.add('notranslate');
    </script>""", width=0, height=0
)

st.markdown("""
<meta name="google" content="notranslate">
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css');
    html, body, [class*="css"] { font-family: 'Pretendard', sans-serif !important; background-color: #F8FAFC; color: #0F172A; translate: no; }
    
    /* 섹션 배너 디자인 */
    .section-banner { background-color: #ffffff; border: 1px solid #E2E8F0; border-left: 8px solid #D91B1B; padding: 18px 24px; border-radius: 12px; margin-top: 35px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
    .section-banner h3 { margin: 0; font-weight: 900; color: #0F172A; font-size: 22px; letter-spacing: -0.5px; }
    
    /* 핵심 지표 (Metric) 카드 디자인 */
    .metric-card-container { background-color: #FFFFFF; border-radius: 16px; padding: 25px 20px; text-align: center; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
    .metric-title { font-size: 16px; color: #475569; margin-bottom: 12px; font-weight: 700; }
    .metric-value-box { display: flex; align-items: center; justify-content: center; gap: 8px; }
    .metric-value { font-size: 42px; font-weight: 900; letter-spacing: -1.5px; line-height: 1; }
    .metric-icon { font-size: 24px; }
    
    /* 동별 구분 헤더 */
    .building-header { font-size: 18px; font-weight: 800; color: #1E293B; margin-top: 25px; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 2px solid #E2E8F0; }
    
    /* 컬럼 및 버튼 간격 */
    div[data-testid="column"] { padding: 0 4px !important; }
    div.stButton > button {
        width: 100%;
        height: 60px;
        background-color: #FFFFFF;
        border: 2px solid #CBD5E1;
        color: #1E293B;
        font-size: 18px !important;
        font-weight: 800 !important;
        border-radius: 8px;
        margin: 0 !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.02);
    }
    div.stButton > button:hover { border-color: #3B82F6; color: #1D4ED8; background-color: #EFF6FF; }
    div.stButton > button:active { background-color: #2563EB !important; color: white !important; border-color: #2563EB; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🌟 2. 팝업창(Modal) 및 렌더링 함수 정의
# ==========================================
def safe_float(val):
    try:
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val) or val is None: return 0.0
        v_str = str(val).strip().replace(',', '').replace(' ', '')
        if v_str in ['', '-', '#DIV/0!', '#N/A', 'nan', 'None']: return 0.0
        if '%' in v_str: return float(v_str.replace('%', '')) / 100.0
        return float(v_str)
    except: return 0.0

def render_section_title(text): st.markdown(f"<div class='section-banner'><h3>{text}</h3></div>", unsafe_allow_html=True)
def render_tab_insight(title, content): st.markdown(f"<div style='background-color:#F1F5F9; border-left:5px solid #3B82F6; border-radius:8px; padding:20px 25px; margin-bottom:25px;'><h4 style='margin-top:0; color:#1E293B; font-weight:800; font-size:17px; margin-bottom:10px;'>{title}</h4><div style='line-height:1.6; font-size:15px; color:#334155;'>{content}</div></div>", unsafe_allow_html=True)
def render_trendy_metric(title, value_str, color, icon): st.markdown(f"<div class='metric-card-container'><div class='metric-title'>{title}</div><div class='metric-value-box'><span class='metric-value' style='color: {color};'>{value_str}</span><span class='metric-icon' style='color: {color};'>{icon}</span></div></div>", unsafe_allow_html=True)

def split_issue_to_columns(issue_text):
    lines = [line.strip() for line in str(issue_text).split('\n') if line.strip()]
    if not lines or str(issue_text).strip() in ['nan', '0', '0.0', 'None']: return "<div style='font-size:13px; color:#94A3B8; padding:8px; font-weight: 500; text-align:center;'>✔ 특이사항 없음</div>"
    d_l, n_l, g_l = [], [], []; has_s = False; curr = g_l
    for line in lines:
        cl = line.replace(' ', '')
        if '*주간' in cl or line.startswith('주간'): curr = d_l; has_s = True; line = re.sub(r'^\*?\s*주간\s*', '', line).strip()
        elif '*야간' in cl or line.startswith('야간'): curr = n_l; has_s = True; line = re.sub(r'^\*?\s*야간\s*', '', line).strip()
        if line: curr.append(line)
    if not has_s: return f"<div style='font-size:14px; font-weight: 600; color:#334155; text-align:left;'>{'<br>'.join(lines)}</div>"
    d_h = '<br>'.join(g_l + d_l) if (g_l + d_l) else "-"; n_h = '<br>'.join(n_l) if n_l else "-"
    return f"<div style='display: flex; gap: 8px; text-align:left;'><div style='flex: 1; background-color: #FFFBEB; padding: 10px; border-radius: 6px; border-top: 3px solid #F59E0B;'><div style='font-size:12px; font-weight:900; color:#B45309; margin-bottom:4px;'>☀️ 주간</div><div style='font-size:13px; font-weight:600;'>{d_h}</div></div><div style='flex: 1; background-color: #F1F5F9; padding: 10px; border-radius: 6px; border-top: 3px solid #334155;'><div style='font-size:12px; font-weight:900; color:#1E293B; margin-bottom:4px;'>🌙 야간</div><div style='font-size:13px; font-weight:600;'>{n_h}</div></div></div>"

def format_issue(text):
    val = str(text).strip()
    if val in ['', '0', '0.0', 'nan', 'None']: return ""
    val = val.replace('\r\n', '\n'); val = re.sub(r'(?<!\n)\*', '\n*', val); val = re.sub(r'(?<!\n)-\.', '\n-.', val); val = re.sub(r'(?<!\n)→', '\n→ ', val)
    return val.strip()

def render_styler_to_html(styler):
    try: html_str = styler.to_html(escape=False)
    except: html_str = styler.to_html()
    custom_css = """
    <style>
        .custom-table { width: 100% !important; border-collapse: collapse !important; font-size: 14px !important; background-color: white !important; }
        .custom-table th { background-color: #1E293B !important; color: #FFFFFF !important; border: 1px solid #334155 !important; padding: 14px !important; font-weight: 700 !important; font-size: 15px !important; text-align: center !important; }
        .custom-table td { border: 1px solid #E2E8F0 !important; padding: 12px !important; font-weight: 500 !important; color: #334155 !important; text-align: center !important; vertical-align: middle !important; }
        .custom-table td:last-child { text-align: left !important; padding-left: 20px !important; } /* 오픈이슈 좌측 정렬 */
    </style>
    """
    html_str = html_str.replace('<table', '<table class="custom-table"')
    st.markdown(custom_css + f"<div style='width:100%; overflow-x:auto; border:1px solid #CBD5E1; border-radius:10px; margin-bottom:25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);'>{html_str}</div>", unsafe_allow_html=True)

def get_building_group(mach_name):
    try:
        num = int(re.search(r'\d+', mach_name).group())
        if 4 <= num <= 21: return "창조동 A"
        elif 22 <= num <= 39: return "창조동 B"
        elif 40 <= num <= 46: return "창조동 C"
        elif 47 <= num <= 52: return "혁신동"
        elif 53 <= num <= 58: return "미래동"
        else: return "기타 구역"
    except: return "기타 구역"

# 탭 1 그래프 클릭 팝업
@st.dialog("📅 일일 가동 상세 현황", width="large")
def show_daily_summary_popup(clicked_date, f_df, daily_df):
    st.markdown(f"<h3 style='text-align:center; color:#0F172A; font-weight:900;'>{clicked_date} 생산 요약</h3><hr>", unsafe_allow_html=True)
    
    day_df = f_df[f_df['생산일'] == clicked_date].copy().sort_values('설비명')
    active_day = day_df[day_df['종합효율'] > 0]
    
    if not active_day.empty:
        active_count = len(active_day)
        total_down = active_day['비가동시간'].apply(safe_float).sum()
        
        matching_daily = daily_df[daily_df['생산일'] == clicked_date]
        if not matching_daily.empty: day_total_val = matching_daily['공장종합효율'].iloc[0]
        else: day_total_val = active_day['종합효율'].apply(safe_float).mean() 
        
        c1, c2, c3 = st.columns(3)
        with c1: render_trendy_metric("실가동 설비", f"{active_count}대", "#10B981", "🏭")
        with c2: render_trendy_metric("공장 종합효율", f"{day_total_val:.1%}", "#2563EB" if day_total_val >= 0.86 else "#DC2626", "📈")
        with c3: render_trendy_metric("총 비가동시간", f"{total_down:.1f}h", "#DC2626" if total_down > 0 else "#10B981", "🛑")
        
        st.markdown("<br><h4 style='font-weight:800; color:#0F172A; margin-bottom:15px;'>📋 설비별 상세 가동 내역</h4>", unsafe_allow_html=True)
        
        # 🚨 에러 방지용 안전한 컬럼 추출
        req_cols = ['설비명', '품명', '종합효율', '비가동시간', 'OPEN ISSUE']
        safe_cols = [c for c in req_cols if c in active_day.columns]
        disp_day = active_day[safe_cols].copy()
        
        for idx, row in disp_day.iterrows():
            if '종합효율' in disp_day.columns: disp_day.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
            if '비가동시간' in disp_day.columns: disp_day.at[idx, '비가동시간'] = f"{safe_float(row['비가동시간']):.1f}h"
        if 'OPEN ISSUE' in disp_day.columns: disp_day['OPEN ISSUE'] = disp_day['OPEN ISSUE'].apply(split_issue_to_columns)
        
        render_styler_to_html(disp_day.style.hide(axis="index"))
    else:
        st.info("해당 일자의 설비 가동 데이터가 존재하지 않습니다.")
        
    if st.button("창 닫기", key="close_daily_popup", use_container_width=True):
        st.rerun()

# 탭 2 설비 클릭 팝업 (🚨 WORST 5 기능 탑재 완료)
@st.dialog("💻 설비 집중 분석 리포트", width="large")
def show_machine_popup(tgt_mach, t7_df):
    st.markdown(f"<h3 style='text-align:center; color:#0F172A; font-weight:900;'>{tgt_mach}</h3><hr>", unsafe_allow_html=True)
    
    valid_t7 = t7_df[t7_df['종합효율'] > 0].copy()
    
    avg_oee = valid_t7['종합효율'].apply(safe_float).mean() if not valid_t7.empty else 0.0
    total_down = t7_df['비가동시간'].apply(safe_float).sum()
    issue_count = t7_df['OPEN ISSUE'].apply(lambda x: 0 if str(x).strip() in ['', 'nan', '0', '0.0'] else 1).sum()
    
    c1, c2, c3 = st.columns(3)
    with c1: render_trendy_metric("기간 내 평균 OEE", f"{avg_oee:.1%}", "#2563EB" if avg_oee >= 0.86 else "#DC2626", "📈")
    with c2: render_trendy_metric("누적 비가동 손실", f"{total_down:.1f}h", "#DC2626" if total_down > 0 else "#059669", "🛑")
    with c3: render_trendy_metric("이슈 발생 일수", f"{issue_count}일", "#D97706" if issue_count > 0 else "#059669", "📝")
    st.markdown("<br>", unsafe_allow_html=True)
    
    st.markdown("<h4 style='font-weight:800; color:#0F172A; margin-bottom:15px;'>📊 일자별 OEE 흐름도</h4>", unsafe_allow_html=True)
    fig7 = go.Figure(go.Scatter(
        x=t7_df['생산일'], y=t7_df['종합효율'], mode='lines+markers+text',
        text=t7_df['종합효율'].apply(lambda x: f"{x:.1%}"), textposition="top center",
        line=dict(color='#2563EB', width=4), marker=dict(size=12, color='#0F172A'), textfont=dict(size=14, weight='bold')
    ))
    fig7.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=350, yaxis=dict(tickformat='.0%', range=[0, 1.1]), margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig7, use_container_width=True)
    
    # 🚨 요청 1번 반영: 전체 표 대신 효율 가장 낮은 WORST 5일 추려서 보여주기 (에러 방어막 포함)
    st.markdown("<br><h4 style='font-weight:900; color:#DC2626; margin-bottom:15px;'>🚨 WORST 5 생산일 (종합효율 최하위 5건)</h4>", unsafe_allow_html=True)
    
    if not valid_t7.empty:
        worst_5 = valid_t7.sort_values(by='종합효율', ascending=True).head(5)
        
        req_cols = ['생산일', '품명', '종합효율', '비가동시간', '총 생산수량', 'OPEN ISSUE']
        safe_cols = [c for c in req_cols if c in worst_5.columns]
        disp_worst = worst_5[safe_cols].copy()
        
        for idx, row in disp_worst.iterrows():
            if '종합효율' in disp_worst.columns: disp_worst.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
            if '총 생산수량' in disp_worst.columns: disp_worst.at[idx, '총 생산수량'] = f"{int(safe_float(row['총 생산수량'])):,}"
            if '비가동시간' in disp_worst.columns: disp_worst.at[idx, '비가동시간'] = f"{safe_float(row['비가동시간']):.1f}h"
            
        if 'OPEN ISSUE' in disp_worst.columns:
            disp_worst['OPEN ISSUE'] = disp_worst['OPEN ISSUE'].apply(split_issue_to_columns)
            
        render_styler_to_html(disp_worst.style.hide(axis="index"))
    else:
        st.info("해당 설비의 유효한 가동 데이터가 없습니다.")
    
    if st.button("창 닫기", key="close_mach_popup", use_container_width=True):
        st.rerun()

# ==========================================
# 🌟 3. 데이터 로드 (🚨 2번 에러: 공백 찌꺼기 원천 차단)
# ==========================================
target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '종합효율', '양품율', 'OPEN ISSUE']
data_to_process = []
DATA_DIR = "data"

if os.path.exists(DATA_DIR):
    for file_name in os.listdir(DATA_DIR):
        if file_name.startswith("~$"): continue 
        if file_name.endswith(('.xlsx', '.csv')):
            file_path = os.path.join(DATA_DIR, file_name)
            try:
                if file_name.endswith('.csv'): 
                    try: t_df = pd.read_csv(file_path, encoding='utf-8')
                    except: t_df = pd.read_csv(file_path, encoding='cp949')
                else: t_df = pd.read_excel(file_path)
                data_to_process.append((file_name, t_df))
            except: pass

if data_to_process:
    all_records = []
    daily_totals_data = {} 
    week_arr = ['(월)', '(화)', '(수)', '(목)', '(금)', '(토)', '(일)']
    
    for file_name, temp_df in data_to_process:
        temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
        temp_df = temp_df.rename(columns={'작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명', '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율'})
        for col in temp_df.columns:
            if 'Unnamed' in col or 'ISSUE' in col.upper(): temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'}); break
        
        date_match = re.search(r'\d{8}', file_name)
        if date_match:
            raw_date = date_match.group()[2:]
            try:
                dt = datetime.strptime(raw_date, '%y%m%d')
                clean_date = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일 {week_arr[dt.weekday()]}"
                month_str = f"{dt.strftime('%y')}년 {dt.month}월"; sort_key = raw_date 
            except: clean_date = raw_date; month_str = "분류 안됨"; sort_key = raw_date
        else: clean_date = file_name.split('.')[0]; month_str = "분류 안됨"; sort_key = clean_date
        
        d_total_oee, daily_val, fallback_val, last_valid_val = 0.0, 0.0, 0.0, 0.0
        for _, row in temp_df.iterrows():
            row_str = "".join([str(v).replace(' ', '').upper() for v in row.values])
            val = safe_float(row.get('종합효율', 0.0))
            if val > 0:
                if any(kw in row_str for kw in ['당월', '월간', '누계', 'MONTH']): continue
                if any(kw in row_str for kw in ['금일', '당일', 'TODAY', 'DAILY']): daily_val = val
                elif any(kw in row_str for kw in ['TOTAL', '합계', '총합', '전체', '총계']): fallback_val = val
                last_valid_val = val
        if daily_val > 0: d_total_oee = daily_val
        elif fallback_val > 0: d_total_oee = fallback_val
        else: d_total_oee = last_valid_val
            
        if sort_key not in daily_totals_data: daily_totals_data[sort_key] = {'생산일': clean_date, '생산월': month_str, '공장종합효율': d_total_oee}

        for _, row in temp_df.iterrows():
            # 🚨 엑셀의 설비명 글자 뒤에 숨은 띄어쓰기(공백)를 완벽하게 제거하여 충돌 원천 방어!
            m_val = str(row.get('설비명')).strip()
            if m_val.lower() in ['', 'nan', 'none', '#n/a'] or any(kw in m_val.upper() for kw in ['TOTAL', '합계']): continue
            
            record = {'sort_key': sort_key, '생산월': month_str, '생산일': clean_date}
            for col in target_cols:
                if col == '설비명':
                    record[col] = m_val
                elif col != '생산일': 
                    record[col] = row[col] if col in temp_df.columns else None
            all_records.append(record)

    df = pd.DataFrame(all_records).sort_values(by='sort_key').reset_index(drop=True)
    date_mapping = dict(zip(df['생산일'], df['sort_key']))
    daily_df = pd.DataFrame([{'sort_key': k, **v} for k, v in daily_totals_data.items()]).sort_values(by='sort_key').reset_index(drop=True)
    for col in ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '종합효율', '양품율']:
        if col in df.columns: df[col] = df[col].apply(safe_float)
    df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

    # =========================================================
    # 🌟 4. 레이아웃 및 필터
    # =========================================================
    title_col1, title_col2 = st.columns([0.6, 9.4])
    with title_col1:
        try: st.image("logo.png", width=80) 
        except: st.markdown("<div style='font-size: 50px; margin-top: 5px;'>🏭</div>", unsafe_allow_html=True)
    with title_col2: st.markdown("<h1 style='margin-top: 15px; margin-bottom: 25px; font-weight: 900; font-size: 34px; color: #0F172A;'>사출생산팀 현장 설비 모니터링</h1>", unsafe_allow_html=True)

    f1, f2 = st.columns(2)
    all_months = [m for m in df['생산월'].unique() if str(m).strip() != ""]
    
    with f1: sel_m_side = st.multiselect("📅 생산월 선택", all_months, default=[all_months[-1]] if all_months else [])
    
    m_f_df = df[df['생산월'].isin(sel_m_side)].copy() if sel_m_side else df.copy()
    all_dates = list(m_f_df['생산일'].unique())
    all_dates.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
    
    with f2: sel_d_side = st.multiselect("📆 생산일 선택", all_dates, default=[])
    f_df = m_f_df[m_f_df['생산일'].isin(sel_d_side)].copy() if sel_d_side else m_f_df.copy()

    st.markdown("<div style='margin-bottom: 25px;'></div>", unsafe_allow_html=True)
    
    # 🌟 5. 메인 탭 설정 
    tab1, tab2 = st.tabs(["📈 사출생산팀 종합효율 추이", "🎯 설비별 정밀 분석 (팝업 호출)"])

    # -----------------------------------------------------
    # TAB 1: 종합 효율 추이 
    # -----------------------------------------------------
    with tab1:
        p_df = daily_df[daily_df['생산월'].isin(sel_m_side)].copy() if sel_m_side else daily_df.copy()
        if sel_d_side: p_df = p_df[p_df['생산일'].isin(sel_d_side)]
            
        if sel_m_side:
            title_month = ", ".join([m.split(' ')[-1] for m in sel_m_side])
            render_section_title(f"사출생산팀 ({title_month}) 종합효율 추이")
        else:
            render_section_title("사출생산팀 전체 종합효율 추이")
            
        if not p_df.empty:
            avg_oee = p_df['공장종합효율'].mean()
            render_tab_insight("💡 현장 운영 가이드", f"조회하신 기간 동안 사출 공정의 평균 OEE는 <b><span style='color:#3B82F6; font-size:18px;'>{avg_oee:.1%}</span></b>를 기록했습니다. <b>아래 막대그래프를 클릭하시면 해당 일자의 '진짜 가동 데이터'가 상세 팝업으로 나타납니다.</b>")
            
            bar_colors = ['#3B82F6' if safe_float(row['공장종합효율']) >= 0.86 else '#EF4444' for _, row in p_df.iterrows()]
            fig_oee = go.Figure(go.Bar(x=p_df['생산일'], y=p_df['공장종합효율'], text=p_df['공장종합효율'].apply(lambda x: f"{x:.1%}"), textposition='auto', marker_color=bar_colors, textfont=dict(size=14, weight='bold', color='white')))
            fig_oee.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=450, yaxis=dict(tickformat='.0%', range=[0, 1.0]), margin=dict(t=20))
            
            try:
                event = st.plotly_chart(fig_oee, use_container_width=True, on_select="rerun", selection_mode="points")
                if event and "selection" in event and event["selection"]["points"]:
                    clicked_date = event["selection"]["points"][0]["x"]
                    show_daily_summary_popup(clicked_date, f_df, daily_df)
            except TypeError:
                st.plotly_chart(fig_oee, use_container_width=True)
                st.info("💡 팁: 막대 그래프를 클릭하여 일일 상세 내역을 보시려면 시스템을 최신 버전으로 업데이트해주세요.")
                
        else: st.info("조건에 해당하는 데이터가 없습니다.")

    # -----------------------------------------------------
    # TAB 2: 설비별 정밀 분석 (버튼 중복 생성 방어막 추가)
    # -----------------------------------------------------
    with tab2:
        render_section_title("👆 점검할 설비 버튼을 터치하여 상세 리포트를 확인하세요")
        
        # 🚨 에러 방어: 리스트의 중복을 100% 제거
        raw_machines = [str(m).strip() for m in f_df['설비명'].unique() if pd.notna(m) and str(m).strip() != 'nan']
        machine_list = sorted(list(set(raw_machines)))
        
        if machine_list:
            building_dict = {"창조동 A": [], "창조동 B": [], "창조동 C": [], "혁신동": [], "미래동": [], "기타 구역": []}
            for mach in machine_list:
                b_name = get_building_group(mach)
                if b_name in building_dict: building_dict[b_name].append(mach)
                else: building_dict["기타 구역"].append(mach)
                
            for b_name, m_list in building_dict.items():
                if not m_list: continue
                
                st.markdown(f"<div class='building-header'>🏭 {b_name}</div>", unsafe_allow_html=True)
                
                cols = st.columns(8) 
                for i, mach in enumerate(m_list):
                    short_name = mach.split(' - ')[0].strip()
                    # 🚨 에러 방어: 버튼 키(Key) 값에 고유 인덱스를 부여하여 충돌 원천 차단
                    if cols[i % 8].button(short_name, key=f"btn_{b_name}_{i}_{mach}"):
                        t7_df = f_df[f_df['설비명'] == mach].copy().sort_values('sort_key')
                        show_machine_popup(mach, t7_df)

        else: st.info("분석할 설비 데이터가 존재하지 않습니다.")

else: st.info("GitHub data 폴더에 CSV/Excel 파일을 넣어주세요.")
