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
        const meta = parent.createElement('meta');
        meta.name = "google"; meta.content = "notranslate";
        parent.head.appendChild(meta);
    </script>""", width=0, height=0
)

st.markdown("""
<meta name="google" content="notranslate">
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css');
    html, body, [class*="css"] { font-family: 'Pretendard', sans-serif !important; background-color: #FFFFFF; color: #1E293B; translate: no; }
    .section-banner { background-color: #ffffff; border: 1px solid #F1F5F9; border-left: 6px solid #D91B1B; padding: 16px 22px; border-radius: 10px; margin-top: 30px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
    .section-banner h3 { margin: 0; font-weight: 800; color: #1E293B; font-size: 19px; }
    .analysis-report-card { background-color: #FFF5F5; border: 1px solid #FEE2E2; border-radius: 12px; padding: 20px 25px; margin-bottom: 25px; box-shadow: 0 2px 4px rgba(217, 27, 27, 0.05); }
    .metric-card-container { background-color: #FFFFFF; border-radius: 12px; padding: 20px; text-align: center; border: 1px solid #F1F5F9; box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.03); }
    .metric-title { font-size: 13px; color: #64748B; margin-bottom: 10px; font-weight: 500; }
    .metric-value-box { display: flex; align-items: center; justify-content: center; gap: 6px; }
    .metric-value { font-size: 32px; font-weight: 800; letter-spacing: -1px; }
    .metric-icon { font-size: 18px; }
    
    /* 설비 선택 버튼 디자인 커스텀 */
    div.stButton > button {
        width: 100%;
        height: 60px;
        background-color: #F8FAFC;
        border: 1px solid #CBD5E1;
        color: #334155;
        font-weight: 700;
        border-radius: 8px;
        transition: all 0.3s;
    }
    div.stButton > button:hover {
        border-color: #3B82F6;
        color: #3B82F6;
        background-color: #EFF6FF;
    }
    div.stButton > button:active {
        background-color: #3B82F6 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🌟 2. 함수 정의
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
def render_tab_insight(title, content): st.markdown(f"<div class='analysis-report-card'><h4 style='margin-top:0; color:#9F1239; font-weight:800; font-size:16px; margin-bottom:12px;'>{title}</h4><div style='line-height:1.7; font-size:14.5px; color:#334155; margin-bottom:0;'>{content}</div></div>", unsafe_allow_html=True)
def get_status_color(oee, tgt=0.86): return ("#3B82F6", "#EBF5FF", "✓") if oee >= tgt else ("#D91B1B", "#FFF5F5", "⚠")
def render_trendy_metric(title, value_str, color, icon): st.markdown(f"<div class='metric-card-container'><div class='metric-title'>{title}</div><div class='metric-value-box'><span class='metric-value' style='color: {color};'>{value_str}</span><span class='metric-icon' style='color: {color};'>{icon}</span></div></div>", unsafe_allow_html=True)

def split_issue_to_columns(issue_text):
    lines = [line.strip() for line in str(issue_text).split('\n') if line.strip()]
    if not lines or str(issue_text).strip() in ['nan', '0', '0.0', 'None']: return "<div style='font-size:12px; color:#ADB5BD; padding:8px;'>📝 특이사항 없음</div>"
    d_l, n_l, g_l = [], [], []; has_s = False; curr = g_l
    for line in lines:
        cl = line.replace(' ', '')
        if '*주간' in cl or line.startswith('주간'): curr = d_l; has_s = True; line = re.sub(r'^\*?\s*주간\s*', '', line).strip()
        elif '*야간' in cl or line.startswith('야간'): curr = n_l; has_s = True; line = re.sub(r'^\*?\s*야간\s*', '', line).strip()
        if line: curr.append(line)
    if not has_s: return f"<div style='font-size:13px;'>{'<br>'.join(lines)}</div>"
    d_h = '<br>'.join(g_l + d_l) if (g_l + d_l) else "없음"; n_h = '<br>'.join(n_l) if n_l else "없음"
    return f"<div style='display: flex; gap: 8px;'><div style='flex: 1; background-color: #F8FAFC; padding: 10px; border-top: 3px solid #FBBF24;'><div style='font-size:11px; font-weight:bold; color:#B45309;'>☀️ 주간</div><div style='font-size:13px;'>{d_h}</div></div><div style='flex: 1; background-color: #F8FAFC; padding: 10px; border-top: 3px solid #1E293B;'><div style='font-size:11px; font-weight:bold; color:#1E293B;'>🌙 야간</div><div style='font-size:13px;'>{n_h}</div></div></div>"

def format_issue(text):
    val = str(text).strip()
    if val in ['', '0', '0.0', 'nan', 'None']: return ""
    val = val.replace('\r\n', '\n'); val = re.sub(r'(?<!\n)\*', '\n*', val); val = re.sub(r'(?<!\n)-\.', '\n-.', val); val = re.sub(r'(?<!\n)→', '\n→ ', val)
    return val.strip()

def render_styler_to_html(styler):
    try: html_str = styler.to_html(escape=False)
    except: html_str = styler.to_html()
    html_str = html_str.replace('<table', '<table style="width: 100%; border-collapse: collapse; font-size: 13px; text-align: center;"')
    html_str = html_str.replace('<th', '<th style="background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 12px; font-weight: 600;"')
    html_str = html_str.replace('<td', '<td style="border: 1px solid #F1F5F9; padding: 10px;"')
    st.markdown(f"<div style='width:100%; overflow-x:auto; border:1px solid #E2E8F0; border-radius:8px; margin-bottom:20px;'>{html_str}</div>", unsafe_allow_html=True)

# ==========================================
# 🌟 3. 데이터 로드 (철통 필터 엔진 유지)
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
                clean_date = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일 ({['월','화','수','목','금','토','일'][dt.weekday()]})"
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
                if col != '생산일': record[col] = row[col] if col in temp_df.columns else None
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
        except: st.markdown("<div style='font-size: 50px; margin-top: 10px;'>🏢</div>", unsafe_allow_html=True)
    with title_col2: st.markdown("<h1 style='margin-top: 20px; margin-bottom: 20px; font-weight: 900; font-size: 30px;'>사출생산팀 설비 모니터링 시스템</h1>", unsafe_allow_html=True)

    f1, f2 = st.columns(2)
    all_months = [m for m in df['생산월'].unique() if str(m).strip() != ""]
    with f1: sel_m_side = st.multiselect("📅 조회할 월 선택", all_months, default=[all_months[-1]] if all_months else [])
    
    m_f_df = df[df['생산월'].isin(sel_m_side)].copy() if sel_m_side else df.copy()
    all_dates = list(m_f_df['생산일'].unique())
    all_dates.sort(key=lambda x: date_mapping.get(x, ""), reverse=True)
    with f2: sel_d_side = st.multiselect("📆 특정 일자만 선택 (선택 안하면 월 전체)", all_dates, default=[])
    f_df = m_f_df[m_f_df['생산일'].isin(sel_d_side)].copy() if sel_d_side else m_f_df.copy()

    st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)
    
    # 🌟 5. 탭 설정 (딱 2개만!)
    tab1, tab2 = st.tabs(["📈 사출생산팀 종합효율 추이", "⚙️ 설비별 정밀 분석 (버튼 선택)"])

    # -----------------------------------------------------
    # TAB 1: 종합 효율 추이 
    # -----------------------------------------------------
    with tab1:
        p_df = daily_df[daily_df['생산월'].isin(sel_m_side)].copy() if sel_m_side else daily_df.copy()
        if sel_d_side: p_df = p_df[p_df['생산일'].isin(sel_d_side)]
            
        render_section_title("공장 전체 종합효율(OEE) 추이")
        if not p_df.empty:
            avg_oee = p_df['공장종합효율'].mean()
            render_tab_insight("📊 종합 요약", f"조회 기간의 사출 공정 평균 OEE는 <b>{avg_oee:.1%}</b> 입니다.")
            
            bar_colors = ['#3B82F6' if safe_float(row['공장종합효율']) >= 0.86 else '#D91B1B' for _, row in p_df.iterrows()]
            fig_oee = go.Figure(go.Bar(x=p_df['생산일'], y=p_df['공장종합효율'], text=p_df['공장종합효율'].apply(lambda x: f"{x:.1%}"), textposition='auto', marker_color=bar_colors))
            fig_oee.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=400, yaxis=dict(tickformat='.0%', range=[0, 1.0]))
            st.plotly_chart(fig_oee, use_container_width=True)
        else: st.info("조건에 해당하는 데이터가 없습니다.")

    # -----------------------------------------------------
    # TAB 2: 설비별 정밀 분석 (버튼 배열 구조)
    # -----------------------------------------------------
    with tab2:
        render_section_title("🎯 설비를 클릭하여 정밀 분석을 확인하세요")
        
        machine_list = sorted([m for m in f_df['설비명'].unique() if m and str(m).strip() != 'nan'])
        
        if machine_list:
            # 세션 스테이트(기억장치)에 선택된 설비를 저장
            if 'selected_mach' not in st.session_state:
                st.session_state.selected_mach = machine_list[0]
                
            # 설비 버튼을 6칸 가로 배열로 쫙 깔아주기
            cols = st.columns(6)
            for i, mach in enumerate(machine_list):
                # 버튼에 표시될 이름 (예: '51호기')
                short_name = mach.split(' - ')[0].strip()
                
                # 버튼을 클릭하면 세션에 해당 설비 이름을 저장함
                if cols[i % 6].button(short_name, key=f"btn_{mach}"):
                    st.session_state.selected_mach = mach

            st.markdown("<hr style='border: 1px dashed #E2E8F0; margin: 25px 0;'>", unsafe_allow_html=True)
            
            # 선택된 설비 데이터만 뽑아오기
            tgt_mach = st.session_state.selected_mach
            st.markdown(f"<h3 style='color:#1E293B; font-weight:900;'>💻 {tgt_mach} 집중 분석</h3>", unsafe_allow_html=True)
            
            t7_df = f_df[f_df['설비명'] == tgt_mach].copy().sort_values('sort_key')
            
            if not t7_df.empty:
                # 지표 계산
                avg_oee = t7_df['종합효율'].apply(safe_float).mean()
                total_down = t7_df['비가동시간'].apply(safe_float).sum()
                issue_count = t7_df['OPEN ISSUE'].apply(lambda x: 0 if str(x).strip() in ['', 'nan', '0', '0.0'] else 1).sum()
                
                # 3단 요약 카드
                c1, c2, c3 = st.columns(3)
                with c1: render_trendy_metric("해당 기간 평균 OEE", f"{avg_oee:.1%}", "#3B82F6" if avg_oee >= 0.86 else "#D91B1B", "📊")
                with c2: render_trendy_metric("누적 비가동시간", f"{total_down:.1f}h", "#EF4444" if total_down > 0 else "#10B981", "🛑")
                with c3: render_trendy_metric("이슈 발생 일수", f"{issue_count}일", "#F59E0B" if issue_count > 0 else "#10B981", "📝")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # 차트와 표를 좌우로 배치 (모니터링 최적화)
                chart_col, table_col = st.columns([1, 1.2])
                
                with chart_col:
                    st.markdown("#### 📈 효율(OEE) 추이")
                    fig7 = go.Figure(go.Scatter(
                        x=t7_df['생산일'], y=t7_df['종합효율'], mode='lines+markers+text',
                        text=t7_df['종합효율'].apply(lambda x: f"{x:.1%}"), textposition="top center",
                        line=dict(color='#3B82F6', width=3), marker=dict(size=10, color='#1E293B')
                    ))
                    fig7.update_layout(plot_bgcolor='rgba(0,0,0,0)', height=350, yaxis=dict(tickformat='.0%', range=[0, 1.0]), margin=dict(l=0, r=0, t=30, b=0))
                    st.plotly_chart(fig7, use_container_width=True)
                
                with table_col:
                    st.markdown("#### 📋 생산 실적 및 OPEN ISSUE 이력")
                    disp_t7 = t7_df[['생산일', '품명', '종합효율', '비가동시간', '총 생산수량', 'OPEN ISSUE']].copy()
                    
                    for idx, row in disp_t7.iterrows():
                        disp_t7.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
                        disp_t7.at[idx, '총 생산수량'] = f"{int(safe_float(row['총 생산수량'])):,}"
                        disp_t7.at[idx, '비가동시간'] = f"{safe_float(row['비가동시간']):.1f}h"
                        
                    disp_t7['OPEN ISSUE'] = disp_t7['OPEN ISSUE'].apply(split_issue_to_columns)
                    render_styler_to_html(disp_t7.style.hide(axis="index"))
                    
            else: st.warning("해당 설비는 선택한 기간 내 가동 데이터가 없습니다.")
        else: st.info("분석할 설비 데이터가 존재하지 않습니다.")

else: st.info("GitHub data 폴더에 CSV/Excel 파일을 넣어주세요.")
