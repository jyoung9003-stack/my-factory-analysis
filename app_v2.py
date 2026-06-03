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
st.set_page_config(page_title="듀링 사출생산팀 설비 분석 대시보드", layout="wide", initial_sidebar_state="collapsed")

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
    
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; max-width: 96% !important; }
    header[data-testid="stHeader"] { display: none !important; }
    [data-testid="column"] { display: flex; flex-direction: column; justify-content: center; }
    
    .section-banner { background-color: #ffffff; border: 1px solid #E2E8F0; border-left: 8px solid #D91B1B; padding: 18px 24px; border-radius: 12px; margin-top: 15px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
    .section-banner h3 { margin: 0; font-weight: 900; color: #0F172A; font-size: 22px; letter-spacing: -0.5px; }
    
    .building-header { font-size: 18px; font-weight: 800; color: #1E293B; margin-top: 25px; margin-bottom: 10px; padding-bottom: 5px; border-bottom: 2px solid #E2E8F0; }
    div[data-testid="column"] { padding: 0 6px !important; }
    
    div.stButton > button {
        width: 100%; min-height: 75px; height: auto !important; background-color: #FFFFFF; border: 2px solid #CBD5E1; color: #1E293B; 
        font-size: 16px !important; font-weight: 800 !important; border-radius: 8px; margin: 0 !important; box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.02);
        white-space: normal !important; word-break: keep-all !important; line-height: 1.4 !important; padding: 10px !important;
    }
    div.stButton > button:hover { border-color: #3B82F6; color: #1D4ED8; background-color: #EFF6FF; transform: translateY(-1px); }
    div.stButton > button:active { background-color: #2563EB !important; color: white !important; border-color: #2563EB; }

    div[data-testid="stRadio"] div[role="radiogroup"] { gap: 20px !important; flex-wrap: wrap !important; margin-top: 15px; }
    div[data-testid="stRadio"] label[data-baseweb="radio"] {
        background-color: #FFFFFF !important; border: 2px solid #CBD5E1 !important; padding: 16px 40px !important;
        border-radius: 12px !important; margin: 0 !important; box-shadow: 0 2px 4px rgba(0,0,0,0.02); cursor: pointer; transition: all 0.2s ease;
        display: flex; justify-content: center; align-items: center; min-width: 160px !important;
    }
    div[data-testid="stRadio"] label[data-baseweb="radio"]:hover { border-color: #D91B1B !important; background-color: #FFF5F5 !important; }
    div[data-testid="stRadio"] label[data-baseweb="radio"] > div:first-child { display: none !important; }
    div[data-testid="stRadio"] label[data-baseweb="radio"] div[data-testid="stMarkdownContainer"] p {
        display: block !important; visibility: visible !important; font-size: 22px !important; font-weight: 900 !important;
        color: #475569 !important; margin: 0 !important; padding: 0 !important; text-align: center !important; width: 100% !important;
    }
    div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background-color: #D91B1B !important; border-color: #D91B1B !important; }
    div[data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) div[data-testid="stMarkdownContainer"] p { color: #FFFFFF !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🌟 2. 공통 함수
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
def render_tab_insight(title, content): st.markdown(f"<div style='background-color:#F8FAFC; border-left:6px solid #3B82F6; border-radius:8px; padding:20px 25px; margin-bottom:25px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);'><h4 style='margin-top:0; color:#1E293B; font-weight:900; font-size:18px; margin-bottom:12px;'>{title}</h4><div style='line-height:1.7; font-size:15.5px; color:#334155;'>{content}</div></div>", unsafe_allow_html=True)

def render_scoreboard_metric(title, value_str, glow_color):
    html = f"""
    <div style="background-color: #000000; border: 4px solid #1E293B; border-radius: 12px; padding: 20px 10px; text-align: center; box-shadow: inset 0px 0px 20px rgba(0,0,0,1);">
        <div style="color: #94A3B8; font-size: 15px; font-weight: 800; margin-bottom: 5px;">{title}</div>
        <div style="color: {glow_color}; font-size: 42px; font-weight: 900; letter-spacing: 2px; font-family: 'Courier New', monospace; text-shadow: 0px 0px 15px {glow_color}; line-height: 1.1;">
            {value_str}
        </div>
    </div>
    """
    st.markdown(html.replace('\n', ''), unsafe_allow_html=True)

def render_rank_cards(df, title, is_worst, name_col):
    border_color = "#DC2626" if is_worst else "#2563EB"
    icon = "🚨" if is_worst else "🏆"
    html = f"<div style='background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 15px; margin-bottom: 20px;'>"
    html += f"<h4 style='margin-top: 0; margin-bottom: 15px; color: #0F172A; font-weight: 900;'>{icon} {title}</h4>"
    html += "<div style='display: flex; flex-direction: column; gap: 8px;'>"
    
    if df.empty: html += "<div style='padding: 10px; text-align: center; color: #64748B; font-weight: 600;'>해당 데이터가 없습니다.</div>"
    else:
        for i, (_, row) in enumerate(df.iterrows()):
            name = str(row.get(name_col, '')).split(' - ')[0]
            oee = safe_float(row.get('종합효율', 0.0))
            down = safe_float(row.get('비가동시간', 0.0))
            html += f"""
            <div style="display: flex; justify-content: space-between; align-items: center; background: white; padding: 12px 15px; border-radius: 8px; border-left: 5px solid {border_color}; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <div style="font-weight: 900; color: #1E293B; font-size: 16px;">{i+1}. {name}</div>
                <div style="display: flex; gap: 20px; align-items: center;">
                    <div style="text-align: right;"><span style="font-size: 12px; color: #64748B; margin-right: 5px; font-weight: 600;">종합효율</span><span style="color: {border_color}; font-weight: 900; font-size: 17px;">{oee:.1%}</span></div>
                    <div style="text-align: right; min-width: 60px;"><span style="font-size: 12px; color: #64748B; margin-right: 5px; font-weight: 600;">비가동</span><span style="color: #475569; font-weight: 900; font-size: 16px;">{down:.1f}h</span></div>
                </div>
            </div>"""
    html += "</div></div>"
    st.markdown(html.replace('\n', ''), unsafe_allow_html=True)

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
    return f"<div style='display: flex; gap: 8px; text-align:left; width:100%; min-width:300px;'><div style='flex: 1; background-color: #FFFBEB; padding: 10px; border-radius: 6px; border-top: 3px solid #F59E0B;'><div style='font-size:12px; font-weight:900; color:#B45309; margin-bottom:4px;'>☀️ 주간</div><div style='font-size:13px; font-weight:600;'>{d_h}</div></div><div style='flex: 1; background-color: #F1F5F9; padding: 10px; border-radius: 6px; border-top: 3px solid #334155;'><div style='font-size:12px; font-weight:900; color:#1E293B; margin-bottom:4px;'>🌙 야간</div><div style='font-size:13px; font-weight:600;'>{n_h}</div></div></div>"

def format_issue(text):
    val = str(text).strip()
    if val in ['', '0', '0.0', 'nan', 'None']: return ""
    val = val.replace('\r\n', '\n'); val = re.sub(r'(?<!\n)\*', '\n*', val); val = re.sub(r'(?<!\n)-\.', '\n-.', val); val = re.sub(r'(?<!\n)→', '\n→ ', val)
    return val.strip()

def render_styler_to_html(styler):
    try: raw_html = styler.to_html(escape=False)
    except: raw_html = styler.to_html()
    clean_html = re.sub(r'<style.*?</style>', '', raw_html, flags=re.DOTALL | re.IGNORECASE)
    custom_css = "<style>.custom-table { width: 100% !important; border-collapse: collapse !important; font-size: 14px !important; background-color: white !important; } .custom-table th { background-color: #1E293B !important; color: #FFFFFF !important; border: 1px solid #334155 !important; padding: 14px !important; font-weight: 700 !important; font-size: 15px !important; text-align: center !important; white-space: nowrap; } .custom-table td { border: 1px solid #E2E8F0 !important; padding: 12px !important; font-weight: 500 !important; color: #334155 !important; text-align: center !important; vertical-align: middle !important; } .custom-table td:last-child { text-align: left !important; padding-left: 20px !important; white-space: normal; }</style>"
    clean_html = clean_html.replace('<table', '<table class="custom-table"')
    final_html = custom_css + f"<div style='width:100%; overflow-x:auto; border:1px solid #CBD5E1; border-radius:10px; margin-bottom:25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);'>{clean_html}</div>"
    final_html = final_html.replace('\n', '').replace('\r', '')
    st.markdown(final_html, unsafe_allow_html=True)

def prepare_popup_table_with_diff(df, full_df, current_date, is_tab1=False):
    safe_cols = [c for c in ['생산일', '설비명', '품명', '종합효율', '비가동시간', 'OPEN ISSUE'] if c in df.columns]
    disp_df = df[safe_cols].copy()
    var_list = []
    
    for idx, row in disp_df.iterrows():
        mach_name = row.get('설비명', '')
        if not mach_name: 
            var_list.append("<span style='color:#94A3B8;'>-</span>")
            continue
            
        mach_df = full_df[full_df['설비명'] == mach_name].sort_values('sort_key')
        valid_mach_df = mach_df[mach_df['종합효율'].apply(safe_float) > 0]
        
        curr_row = mach_df[mach_df['생산일'] == current_date]
        if curr_row.empty or safe_float(curr_row.iloc[0]['종합효율']) <= 0: 
            var_list.append("<span style='color:#94A3B8;'>-</span>")
        else:
            curr_sort_key = curr_row.iloc[0]['sort_key']
            past_df = valid_mach_df[valid_mach_df['sort_key'] < curr_sort_key] 
            
            if past_df.empty: var_list.append("<span style='color:#94A3B8;'>-</span>")
            else:
                prev_oee = safe_float(past_df.iloc[-1]['종합효율'])
                curr_oee = safe_float(curr_row.iloc[0]['종합효율'])
                diff = curr_oee - prev_oee
                if diff > 0: var_list.append(f"<span style='color:#2563EB; font-weight:900;'>▲ +{diff*100:.1f}%p</span>")
                elif diff < 0: var_list.append(f"<span style='color:#DC2626; font-weight:900;'>▼ {abs(diff)*100:.1f}%p</span>")
                else: var_list.append("<span style='color:#64748B; font-weight:900;'>-</span>")
    
    disp_df['전일 대비 증감율'] = var_list
    for idx, row in disp_df.iterrows():
        if '종합효율' in disp_df.columns: disp_df.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
        if '비가동시간' in disp_df.columns: disp_df.at[idx, '비가동시간'] = f"{safe_float(row['비가동시간']):.1f}h"
    if 'OPEN ISSUE' in disp_df.columns: disp_df['OPEN ISSUE'] = disp_df['OPEN ISSUE'].apply(split_issue_to_columns)
    
    if is_tab1: final_cols = [c for c in ['설비명', '품명', '종합효율', '전일 대비 증감율', '비가동시간', 'OPEN ISSUE'] if c in disp_df.columns]
    else: final_cols = [c for c in ['생산일', '품명', '종합효율', '전일 대비 증감율', '비가동시간', 'OPEN ISSUE'] if c in disp_df.columns]
    return disp_df[final_cols]

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

# ==========================================
# 🌟 3. 팝업창 
# ==========================================
@st.dialog("📅 일일 가동 상세 현황", width="large")
def show_daily_summary_popup(clicked_date, f_df, daily_df, full_df):
    st.markdown(f"<h2 style='text-align:center; color:#0F172A; font-weight:900; font-size:32px; margin-bottom:10px;'><span style='color:#D91B1B;'>{clicked_date}</span> 사출생산팀 생산 실적 및 오픈이슈 현황</h2><hr style='border-top: 3px solid #E2E8F0; margin-bottom: 30px;'>", unsafe_allow_html=True)
    
    day_df = f_df[f_df['생산일'] == clicked_date].copy().sort_values('설비명')
    active_day = day_df[day_df['종합효율'] > 0]
    
    if not active_day.empty:
        active_count = len(active_day)
        total_down = active_day['비가동시간'].apply(safe_float).sum()
        matching_daily = daily_df[daily_df['생산일'] == clicked_date]
        if not matching_daily.empty: day_total_val = matching_daily['공장종합효율'].iloc[0]
        else: day_total_val = active_day['종합효율'].apply(safe_float).mean() 
        
        c1, c2, c3 = st.columns(3)
        with c1: render_scoreboard_metric("💡 가동 설비 대수", f"{active_count}대", "#32CD32")
        with c2: render_scoreboard_metric("📊 당일 공장 종합효율", f"{day_total_val:.1%}", "#3B82F6" if day_total_val >= 0.86 else "#FF3131")
        with c3: render_scoreboard_metric("🛑 총 비가동시간", f"{total_down:.1f}h", "#FF3131" if total_down > 0 else "#32CD32")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        best_5_mach = active_day.sort_values(by='종합효율', ascending=False).head(5)
        worst_5_mach = active_day.sort_values(by='종합효율', ascending=True).head(5)
        col_best, col_worst = st.columns(2)
        with col_best: render_rank_cards(best_5_mach, "BEST 5", is_worst=False, name_col="설비명")
        with col_worst: render_rank_cards(worst_5_mach, "WORST 5", is_worst=True, name_col="설비명")

        st.markdown("<h4 style='font-weight:800; color:#0F172A; margin-bottom:15px;'>📋 전체 설비 상세 가동 내역</h4>", unsafe_allow_html=True)
        disp_day = prepare_popup_table_with_diff(active_day, full_df, clicked_date, is_tab1=True)
        render_styler_to_html(disp_day.style.hide(axis="index"))
    else:
        st.info("해당 일자의 설비 가동 데이터가 존재하지 않습니다.")

@st.dialog("💻 설비 집중 분석 리포트", width="large")
def show_machine_popup(tgt_mach, t7_df, full_df):
    st.markdown(f"<h2 style='text-align:center; color:#0F172A; font-weight:900; font-size:32px; margin-bottom:10px;'><span style='color:#2563EB;'>{tgt_mach}</span> 이력 모니터링</h2><hr style='border-top: 3px solid #E2E8F0; margin-bottom: 30px;'>", unsafe_allow_html=True)
    
    valid_t7 = t7_df[t7_df['종합효율'] > 0].copy()
    avg_oee = valid_t7['종합효율'].apply(safe_float).mean() if not valid_t7.empty else 0.0
    total_down = t7_df['비가동시간'].apply(safe_float).sum()
    issue_count = t7_df['OPEN ISSUE'].apply(lambda x: 0 if str(x).strip() in ['', 'nan', '0', '0.0'] else 1).sum()
    
    c1, c2, c3 = st.columns(3)
    with c1: render_scoreboard_metric("📊 누적 평균 종합효율", f"{avg_oee:.1%}", "#3B82F6" if avg_oee >= 0.86 else "#FF3131")
    with c2: render_scoreboard_metric("🛑 누적 비가동 손실", f"{total_down:.1f}h", "#FF3131" if total_down > 0 else "#32CD32")
    with c3: render_scoreboard_metric("📝 이슈 발생 일수", f"{issue_count}일", "#FF9900" if issue_count > 0 else "#32CD32") 
    
    st.markdown("<br><h4 style='font-weight:800; color:#0F172A; margin-bottom:15px;'>📊 일자별 종합효율 흐름도</h4>", unsafe_allow_html=True)
    fig7 = go.Figure(go.Scatter(
        x=t7_df['생산일'], y=t7_df['종합효율'], mode='lines+markers+text',
        text=t7_df['종합효율'].apply(lambda x: f"{x:.1%}"), textposition="top center",
        line=dict(color='#2563EB', width=4), marker=dict(size=12, color='#0F172A'), textfont=dict(size=14, weight='bold')
    ))
    fig7.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=350, yaxis=dict(tickformat='.0%', range=[0, 1.1]), margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig7, use_container_width=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    if not valid_t7.empty:
        best_5_days = valid_t7.sort_values(by='종합효율', ascending=False).head(5)
        worst_5_days = valid_t7.sort_values(by='종합효율', ascending=True).head(5)
        col_best, col_worst = st.columns(2)
        with col_best: render_rank_cards(best_5_days, "BEST 5", is_worst=False, name_col="생산일")
        with col_worst: render_rank_cards(worst_5_days, "WORST 5", is_worst=True, name_col="생산일")
    else: 
        st.info("해당 설비의 유효한 가동 데이터가 없습니다.")

    mach_short = str(tgt_mach).split(' - ')[0].strip()
    prod_name = valid_t7['품명'].dropna().iloc[0] if not valid_t7['품명'].dropna().empty else ""
    title_dynamic = f"📋 {mach_short} {prod_name} 생산성 및 오픈 이슈 현황"
    st.markdown(f"<h4 style='font-weight:900; color:#0F172A; margin-bottom:15px; margin-top:20px;'>{title_dynamic}</h4>", unsafe_allow_html=True)
    
    disp_t7 = prepare_popup_table_with_diff(t7_df, full_df, None, is_tab1=False) 
    
    safe_cols = [c for c in ['생산일', 'sort_key', '품명', '종합효율', '비가동시간', 'OPEN ISSUE'] if c in t7_df.columns]
    raw_disp_t7 = t7_df[safe_cols].copy()
    var_list_t2 = []
    valid_mach_df = full_df[full_df['설비명'] == tgt_mach].sort_values('sort_key')
    valid_mach_df = valid_mach_df[valid_mach_df['종합효율'].apply(safe_float) > 0]
    
    for idx, row in raw_disp_t7.iterrows():
        curr_oee = safe_float(row.get('종합효율', 0.0))
        if curr_oee <= 0:
            var_list_t2.append("<span style='color:#94A3B8;'>-</span>")
            continue
        curr_sk = row.get('sort_key', '')
        past_df = valid_mach_df[valid_mach_df['sort_key'] < curr_sk]
        if past_df.empty: var_list_t2.append("<span style='color:#94A3B8;'>-</span>")
        else:
            prev_oee = safe_float(past_df.iloc[-1]['종합효율'])
            diff = curr_oee - prev_oee
            if diff > 0: var_list_t2.append(f"<span style='color:#2563EB; font-weight:900;'>▲ +{diff*100:.1f}%p</span>")
            elif diff < 0: var_list_t2.append(f"<span style='color:#DC2626; font-weight:900;'>▼ {abs(diff)*100:.1f}%p</span>")
            else: var_list_t2.append("<span style='color:#64748B; font-weight:900;'>-</span>")
            
    raw_disp_t7['전일 대비 증감율'] = var_list_t2
    for idx, row in raw_disp_t7.iterrows():
        if '종합효율' in raw_disp_t7.columns: raw_disp_t7.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
        if '비가동시간' in raw_disp_t7.columns: raw_disp_t7.at[idx, '비가동시간'] = f"{safe_float(row['비가동시간']):.1f}h"
    if 'OPEN ISSUE' in raw_disp_t7.columns: raw_disp_t7['OPEN ISSUE'] = raw_disp_t7['OPEN ISSUE'].apply(split_issue_to_columns)
    
    final_cols_t7 = [c for c in ['생산일', '품명', '종합효율', '전일 대비 증감율', '비가동시간', 'OPEN ISSUE'] if c in raw_disp_t7.columns]
    render_styler_to_html(raw_disp_t7[final_cols_t7].style.hide(axis="index"))

# ==========================================
# 🌟 4. 데이터 로드 및 정제 
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
            m_val = str(row.get('설비명')).strip()
            if m_val.lower() in ['', 'nan', 'none', '#n/a'] or any(kw in m_val.upper() for kw in ['TOTAL', '합계']): continue
            
            record = {'sort_key': sort_key, '생산월': month_str, '생산일': clean_date}
            for col in target_cols:
                if col == '설비명': record[col] = m_val
                elif col != '생산일': record[col] = row[col] if col in temp_df.columns else None
            all_records.append(record)

    df = pd.DataFrame(all_records).sort_values(by='sort_key').reset_index(drop=True)
    date_mapping = dict(zip(df['생산일'], df['sort_key']))
    daily_df = pd.DataFrame([{'sort_key': k, **v} for k, v in daily_totals_data.items()]).sort_values(by='sort_key').reset_index(drop=True)
    for col in ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '종합효율', '양품율']:
        if col in df.columns: df[col] = df[col].apply(safe_float)
    df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

    # =========================================================
    # 🌟 5. 레이아웃 및 필터 (정렬 최적화 완료)
    # =========================================================
    title_col1, title_col2, title_col3 = st.columns([1, 5.5, 3.5], gap="small", vertical_alignment="center")
    
    with title_col1:
        try: st.image("logo.png", width=120) 
        except: st.markdown("<div style='font-size: 50px;'>🏭</div>", unsafe_allow_html=True)
        
    with title_col2: 
        st.markdown("<h1 style='margin: 0; padding: 0; font-weight: 900; font-size: 26px; color: #0F172A; white-space: nowrap;'>사출생산팀 생산성 및 오픈 이슈 분석 및 관리 리포트</h1>", unsafe_allow_html=True)

    with title_col3:
        if not daily_df.empty:
            recent_row = daily_df.iloc[-1]
            rec_date = recent_row['생산일']
            rec_oee = safe_float(recent_row['공장종합효율'])
            
            diff_str = "<span style='color:#64748B;'>-</span>"
            if len(daily_df) > 1:
                prev_oee = safe_float(daily_df.iloc[-2]['공장종합효율'])
                diff = rec_oee - prev_oee
                if diff > 0: diff_str = f"<span style='color:#2563EB;'>▲ +{diff*100:.1f}%p</span>"
                elif diff < 0: diff_str = f"<span style='color:#DC2626;'>▼ {abs(diff)*100:.1f}%p</span>"

            st.markdown(f"""
            <div style='background-color: #F8FAFC; border: 2px solid #E2E8F0; border-radius: 12px; padding: 12px 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); width: 100%;'>
                <div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; border-bottom: 1px solid #E2E8F0; padding-bottom: 8px;'>
                    <span style='font-size: 14px; color: #475569; font-weight: 800;'>{rec_date}</span>
                    <div style='font-size: 22px; font-weight: 900; color: #0F172A; display: flex; align-items: baseline;'>
                        {rec_oee:.1%} <span style='font-size: 11px; font-weight:800; color:#64748B; margin-left:12px; margin-right:4px;'>전일대비 증감율</span><span style='font-size: 15px;'>{diff_str}</span>
                    </div>
                </div>
                <div style='display: flex; justify-content: space-between; align-items: center;'>
                    <span style='font-size: 14px; color: #475569; font-weight: 800;' id='month_name_placeholder'>선택월 평균</span>
                    <div style='font-size: 20px; font-weight: 800; color: #1E293B;' id='month_avg_placeholder'>
                        {daily_df['공장종합효율'].mean():.1%}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<hr style='border: 1px solid #E2E8F0; margin-top: 10px; margin-bottom: 10px;'>", unsafe_allow_html=True)

    f1, f2, f3 = st.columns([1, 1, 2])
    all_months = [m for m in df['생산월'].unique() if str(m).strip() != ""]
    with f1: sel_m_side = st.multiselect("📅 생산월 선택", all_months, default=[all_months[-1]] if all_months else [])
    
    m_f_df = df[df['생산월'].isin(sel_m_side)].copy() if sel_m_side else df.copy()
    all_dates = list(m_f_df['생산일'].unique())
    all_dates.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
    with f2: sel_d_side = st.multiselect("📆 생산일 선택", all_dates, default=[])
    
    f_df = m_f_df[m_f_df['생산일'].isin(sel_d_side)].copy() if sel_d_side else m_f_df.copy()

    title_month_str = ", ".join(sel_m_side) if sel_m_side else "전체"
    if not daily_df.empty:
        p_df_for_summary = daily_df[daily_df['생산월'].isin(sel_m_side)] if sel_m_side else daily_df
        month_oee = p_df_for_summary['공장종합효율'].mean() if not p_df_for_summary.empty else 0.0
        components.html(f"""
        <script>
            setTimeout(function() {{
                const doc = window.parent.document;
                const avgElem = doc.getElementById('month_avg_placeholder');
                const nameElem = doc.getElementById('month_name_placeholder');
                if(avgElem) avgElem.innerText = '{month_oee:.1%}';
                if(nameElem) nameElem.innerText = '{title_month_str} 평균';
            }}, 150);
        </script>
        """, width=0, height=0)

    st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
    
    # 🚨 신규 탭 추가: 현장 도면(사진) 기반 모니터링
    tab1, tab2, tab3 = st.tabs(["📈 사출생산팀 종합효율 추이", "🎯 설비별 생산성 및 오픈이슈 분석", "🗺️ 도면 기반 설비 모니터링 (BETA)"])

    # -----------------------------------------------------
    # TAB 1: 종합 효율 추이 
    # -----------------------------------------------------
    with tab1:
        p_df = daily_df[daily_df['생산월'].isin(sel_m_side)].copy() if sel_m_side else daily_df.copy()
        if sel_d_side: p_df = p_df[p_df['생산일'].isin(sel_d_side)]
            
        if sel_m_side: render_section_title(f"사출생산팀 ({title_month_str}) 종합효율 추이")
        else: render_section_title("사출생산팀 전체 종합효율 추이")
            
        if not p_df.empty:
            avg_oee = p_df['공장종합효율'].mean()
            is_achieved = avg_oee >= 0.86
            status_color = "#2563EB" if is_achieved else "#DC2626"
            status_text = "목표 효율 대비 달성 중입니다." if is_achieved else "목표 효율 대비 미달성 중입니다."
            
            guide_text = f"사출생산팀 ({title_month_str}) 종합 효율은 <b>{avg_oee:.1%}</b>로 <span style='color:{status_color}; font-weight:900;'>{status_text}</span><br>분석할 생산일의 그래프를 선택하면 해당 일의 생산성 자료 및 오픈이슈를 확인 가능합니다. 또한 <b>탭2 (설비별 생산성 및 오픈이슈 분석)</b>에서 각 설비의 생산일 별 생산성 및 오픈이슈를 확인 가능합니다."
            
            render_tab_insight("💡 사출생산팀 생산성 추이 분석", guide_text)
            
            bar_colors = ['#3B82F6' if safe_float(row['공장종합효율']) >= 0.86 else '#EF4444' for _, row in p_df.iterrows()]
            fig_oee = go.Figure(go.Bar(x=p_df['생산일'], y=p_df['공장종합효율'], text=p_df['공장종합효율'].apply(lambda x: f"{x:.1%}"), textposition='auto', marker_color=bar_colors, textfont=dict(size=14, weight='bold', color='white')))
            fig_oee.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=450, yaxis=dict(tickformat='.0%', range=[0, 1.0]), margin=dict(t=20))
            
            if 'last_chart_click' not in st.session_state: st.session_state.last_chart_click = None
            if 'trigger_daily_popup' not in st.session_state: st.session_state.trigger_daily_popup = None

            try:
                event = st.plotly_chart(fig_oee, use_container_width=True, on_select="rerun", selection_mode="points")
                curr_click = event["selection"]["points"][0]["x"] if event and "selection" in event and event["selection"]["points"] else None
                
                if curr_click != st.session_state.last_chart_click:
                    st.session_state.last_chart_click = curr_click
                    if curr_click is not None: st.session_state.trigger_daily_popup = curr_click
                
                if st.session_state.trigger_daily_popup:
                    date_to_show = st.session_state.trigger_daily_popup
                    st.session_state.trigger_daily_popup = None 
                    show_daily_summary_popup(date_to_show, f_df, daily_df, df)
                    
            except TypeError:
                st.plotly_chart(fig_oee, use_container_width=True)
                st.info("💡 팁: 막대 그래프를 클릭하여 일일 상세 내역을 보시려면 시스템을 최신 버전으로 업데이트해주세요.")
        else: st.info("조건에 해당하는 데이터가 없습니다.")

    # -----------------------------------------------------
    # TAB 2: 설비별 정밀 분석 
    # -----------------------------------------------------
    with tab2:
        render_section_title("👆 점검할 구역(동)을 선택하고 설비 리포트를 확인하세요")
        
        raw_machines = [str(m).strip() for m in f_df['설비명'].unique() if pd.notna(m) and str(m).strip() != 'nan']
        machine_list = sorted(list(set(raw_machines)))
        
        if machine_list:
            building_dict = {"창조동 A": [], "창조동 B": [], "창조동 C": [], "혁신동": [], "미래동": [], "기타 구역": []}
            for mach in machine_list:
                b_name = get_building_group(mach)
                if b_name in building_dict: building_dict[b_name].append(mach)
                else: building_dict["기타 구역"].append(mach)
            
            active_buildings = [b for b, m in building_dict.items() if len(m) > 0]
            selected_building = st.radio("🏭 조회를 원하는 구역(동)을 선택하세요", active_buildings, horizontal=True, label_visibility="collapsed")
            
            st.markdown("<hr style='border: 2px dashed #CBD5E1; margin: 25px 0;'>", unsafe_allow_html=True)
            
            m_list = building_dict[selected_building]
            if m_list:
                cols = st.columns(4) 
                for i, mach in enumerate(m_list):
                    parts = [p.strip() for p in mach.split('-')]
                    if len(parts) >= 2: display_name = f"{parts[0]} - {parts[1]}"
                    else: display_name = parts[0]
                        
                    if cols[i % 4].button(display_name, key=f"btn_{selected_building}_{i}_{mach}"):
                        t7_df = f_df[f_df['설비명'] == mach].copy().sort_values('sort_key')
                        show_machine_popup(mach, t7_df, df)

        else: st.info("분석할 설비 데이터가 존재하지 않습니다.")

    # -----------------------------------------------------
    # 🚨 TAB 3: 현장 도면(사진) 기반 모니터링 (BETA)
    # -----------------------------------------------------
    with tab3:
        render_section_title("🗺️ 현장 레이아웃 도면 매핑")
        
        # GitHub에 layout.jpg 또는 layout.png를 올리면 이 공간에 그림이 나타납니다.
        if os.path.exists("layout.jpg"):
            st.image("layout.jpg", use_container_width=True)
        elif os.path.exists("layout.png"):
            st.image("layout.png", use_container_width=True)
        else:
            st.info("💡 GitHub 데이터 폴더에 현장 도면 사진을 `layout.jpg` 또는 `layout.png` 이름으로 올려주시면 이곳에 나타납니다.")
            
        st.markdown("<hr style='border: 1px dashed #CBD5E1; margin: 20px 0;'>", unsafe_allow_html=True)
        st.markdown("<h4 style='font-weight:800; color:#0F172A; text-align:center;'>👇 도면에서 확인한 설비를 아래에서 즉시 선택하세요</h4><br>", unsafe_allow_html=True)
        
        # 탭 2의 버튼 로직을 재사용하여 도면 바로 밑에서 클릭할 수 있게 지원
        if machine_list:
            cols = st.columns(6) 
            for i, mach in enumerate(machine_list):
                parts = [p.strip() for p in mach.split('-')]
                short_name = parts[0]
                if cols[i % 6].button(short_name, key=f"btn_map_{i}_{mach}"):
                    t7_df = f_df[f_df['설비명'] == mach].copy().sort_values('sort_key')
                    show_machine_popup(mach, t7_df, df)

else: st.info("GitHub data 폴더에 CSV/Excel 파일을 넣어주세요.")
