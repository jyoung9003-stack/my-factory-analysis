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
# 🌟 [트렌디 디자인] UI3 프리미엄 테마 설정
# ==========================================
st.set_page_config(
    page_title="사출생산팀 일일 생산성 정밀 분석", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# [글로벌 트렌드] 소프트 UI, 카드 기반 레이아웃, 세련된 컬러 CSS
st.markdown("""
<meta name="google" content="notranslate">
<style>
    /* 폰트 및 기본 배경 (소프트 오프화이트) */
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css');
    html, body, [class*="css"] {
        font-family: 'Pretendard', sans-serif !important;
        background-color: #F8FAFC; /* 글로벌 트렌드 배경색 */
        color: #1E293B; /* 딥 네이비 텍스트 */
        translate: no; 
    }

    /* 사이드바 스타일링 */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }

    /* 🌟 [글로벌 트렌드] 공통 카드 스타일 (둥근 모서리 + 소프트 섀도우) */
    .trendy-card {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05), 0 1px 2px -1px rgba(0, 0, 0, 0.05); /* Soft Shadow */
        border: 1px solid #F1F5F9;
        margin-bottom: 24px;
    }

    /* 🌟 [트렌디 지표] 커스텀 Metric 카드 디자인 */
    .metric-card-container {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        border: 1px solid #F1F5F9;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.03);
    }
    .metric-title {
        font-size: 13px;
        color: #64748B; /* 서브 텍스트 */
        margin-bottom: 10px;
        font-weight: 500;
    }
    .metric-value-box {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
    }
    .metric-value {
        font-size: 32px;
        font-weight: 800;
        letter-spacing: -1px;
    }
    .metric-icon {
        font-size: 18px;
    }

    /* 전광판 요약판 스타일 개선 */
    .dashboard-header {
        background: linear-gradient(135deg, #2D3E5E, #3B82F6);
        color: #FFFFFF;
        border-radius: 12px;
        padding: 30px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.08);
        margin-bottom: 24px;
    }
</style>
""", unsafe_allow_html=True)

# 지표 컬러링을 위한 유틸리티 함수
def get_status_color(oee, tgt=0.86):
    color = "#3B82F6" if oee >= tgt else "#F97316" # 테크 블루 vs 모던 오렌지
    bg = "#EBF5FF" if oee >= tgt else "#FFF1E6"
    icon = "✓" if oee >= tgt else "⚠"
    return color, bg, icon

# 🌟 [트렌디 지표] 세련된 지표 카드를 직접 그려주는 함수 생성
def render_trendy_metric(title, value_str, color, icon):
    st.markdown(f"""
    <div class="metric-card-container">
        <div class="metric-title">{title}</div>
        <div class="metric-value-box">
            <span class="metric-value" style="color: {color};">{value_str}</span>
            <span class="metric-icon" style="color: {color};">{icon}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

# 섹션 타이틀 디자인 개선
def render_section_title(title_text):
    st.markdown(f"""
    <div style="background-color: #ffffff; border: 1px solid #F1F5F9; border-left: 5px solid #3B82F6; padding: 18px 24px; border-radius: 12px; margin-top: 30px; margin-bottom: 20px; box-shadow: 0 1px 3px 0 rgba(0,0,0,0.03);">
        <h3 style="margin: 0; font-weight: 800; color: #1E293B; font-size: 18px; letter-spacing: -0.5px;">{title_text}</h3>
    </div>
    """, unsafe_allow_html=True)

# 2. 메인 타이틀 영역 개선 (Reference 2 그라데이션 느낌 반영)
st.markdown("""
<div style='background: linear-gradient(90deg, #F0F2F6, #FFFFFF); padding: 15px 20px; border-radius: 10px; margin-bottom: 30px;'>
    <h1 style='margin: 0; color: #1E293B; font-weight: 900; font-size: 28px; letter-spacing: -1.5px;' class='notranslate'>사출생산팀 일일 생산성 정밀 분석</h1>
</div>
""", unsafe_allow_html=True)

def safe_float(val):
    try:
        if isinstance(val, pd.Series): val = val.iloc[0]
        if pd.isna(val) or val is None: return 0.0
        v_str = str(val).strip().replace(',', '').replace(' ', '')
        if v_str in ['', '-', '#DIV/0!', '#N/A', 'nan', 'None']: return 0.0
        if '%' in v_str: return float(v_str.replace('%', '')) / 100.0
        return float(v_str)
    except: return 0.0

target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']
target_order = ['생산일', '설비명', '품명', '종합효율', '양품율', '성능가동율', '시간가동율', '총 생산수량', '양품수량', '불량수량', 'OPEN ISSUE']
multi_cols = [
    ('구분', '생산일'), ('구분', '설비명'), ('구분', '품명'),
    ('생산성', '종합효율'), ('생산성', '양품율'), ('생산성', '성능가동율'), ('생산성', '시간가동율'),
    ('생산실적', '총 생산수량'), ('생산실적', '양품수량'), ('생산실적', '불량수량'),
    ('OPEN ISSUE', 'OPEN ISSUE')
]

# 3. 데이터 수집 로직
data_to_process = []
DATA_DIR = "data"
if os.path.exists(DATA_DIR):
    for file_name in os.listdir(DATA_DIR):
        if file_name.startswith("~$"): continue 
        if file_name.endswith('.xlsx') or file_name.endswith('.csv'):
            file_path = os.path.join(DATA_DIR, file_name)
            try:
                if file_name.endswith('.csv'): 
                    try: df = pd.read_csv(file_path, encoding='utf-8')
                    except UnicodeDecodeError: df = pd.read_csv(file_path, encoding='cp949')
                else: df = pd.read_excel(file_path)
                data_to_process.append((file_name, df))
            except Exception as e: st.error(f"데이터 읽기 오류 ({file_name}): {e}")

uploaded_files = st.file_uploader("📂 새로운 일일 생산성 파일이 있다면 추가하세요", type=['xlsx', 'csv'], accept_multiple_files=True)
if uploaded_files:
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'): 
                try: df = pd.read_csv(file, encoding='utf-8')
                except UnicodeDecodeError:
                    file.seek(0)
                    df = pd.read_csv(file, encoding='cp949')
            else: df = pd.read_excel(file)
            data_to_process.append((file.name, df))
        except Exception as e: st.error(f"업로드 오류: {e}")

if data_to_process:
    all_records = []
    daily_totals_data = {} 
    
    for file_name, temp_df in data_to_process:
        temp_df = temp_df.loc[:, ~temp_df.columns.duplicated(keep='first')]
        temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
        name_map = {'작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명', '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율', '목표 효율': '목표효율'}
        temp_df = temp_df.rename(columns=name_map)
        
        for col in temp_df.columns:
            if 'Unnamed' in col or 'ISSUE' in col.upper():
                temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'})
                break
        
        date_match = re.search(r'\d{8}', file_name)
        if date_match:
            raw_date = date_match.group()[2:]
            try:
                dt = datetime.strptime(raw_date, '%y%m%d')
                weekdays = ['월', '화', '수', '목', '금', '토', '일']
                clean_date = f"{dt.strftime('%y')}년 {dt.month}월 {dt.day}일 ({weekdays[dt.weekday()]})"
                month_str = f"{dt.strftime('%y')}년 {dt.month}월"; sort_key = raw_date 
            except: clean_date = raw_date; month_str = "분류 안됨"; sort_key = raw_date
        else: clean_date = file_name.split('.')[0]; month_str = "분류 안됨"; sort_key = clean_date
        
        daily_total_oee = 0.0; daily_total_perf = 0.0; daily_total_avail = 0.0; daily_total_qual = 0.0
        for _, row in temp_df.iterrows():
            m_val = str(row.get('설비명', ''))
            if 'TOTAL' in m_val.upper() or '합계' in m_val or 'GRAND' in m_val.upper():
                daily_total_oee = safe_float(row.get('종합효율', 0.0))
                daily_total_perf = safe_float(row.get('성능가동율', 0.0))
                daily_total_avail = safe_float(row.get('시간가동율', 0.0))
                daily_total_qual = safe_float(row.get('양품율', 0.0))
                break
        if daily_total_oee == 0.0:
            try: 
                daily_total_oee = safe_float(temp_df['종합효율'].iloc[45])
                daily_total_perf = safe_float(temp_df['성능가동율'].iloc[45])
                daily_total_avail = safe_float(temp_df['시간가동율'].iloc[45])
                daily_total_qual = safe_float(temp_df['양품율'].iloc[45])
            except: pass
        
        if sort_key not in daily_totals_data:
            daily_totals_data[sort_key] = {'생산일': clean_date, '생산월': month_str, '공장종합효율': daily_total_oee, '공장성능가동율': daily_total_perf, '공장시간가동율': daily_total_avail, '공장양품율': daily_total_qual}
        else:
            if daily_total_oee > 0:
                daily_totals_data[sort_key].update({'공장종합효율': daily_total_oee, '공장성능가동율': daily_total_perf, '공장시간가동율': daily_total_avail, '공장양품율': daily_total_qual})

        for _, row in temp_df.iterrows():
            m_val = str(row.get('설비명', '')).strip()
            if m_val in ['', 'nan', 'NaN', 'None', '#N/A'] or 'Unnamed' in m_val or m_val == '설비명': continue
            if 'TOTAL' in m_val.upper() or '합계' in m_val or 'GRAND' in m_val.upper(): continue
            
            record = {'sort_key': sort_key, '생산월': month_str, '생산일': clean_date}
            for col in target_cols:
                if col != '생산일': record[col] = row[col] if col in temp_df.columns else None
            all_records.append(record)

    if all_records:
        df = pd.DataFrame(all_records).sort_values(by='sort_key').reset_index(drop=True)
        daily_df = pd.DataFrame([{'sort_key': k, **v} for k, v in daily_totals_data.items()]).sort_values(by='sort_key').reset_index(drop=True)
        
        num_cols = ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            if col in df.columns: df[col] = df[col].apply(safe_float)

        def format_issue(text):
            val = str(text).strip()
            if val in ['', '0', '0.0', 'nan', 'NaN', 'None']: return ""
            val = val.replace('\r\n', '\n')
            val = re.sub(r'(?<!\n)\*', '\n*', val); val = re.sub(r'(?<!\n)-\.', '\n-.', val); val = re.sub(r'(?<!\n)→', '\n→ ', val)
            return val.strip()
        df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

        def split_issue_to_columns(issue_text):
            lines = [line.strip() for line in str(issue_text).split('\n') if line.strip()]
            if not lines: return "<div style='font-size:12px; color:#ADB5BD; margin-top:6px; background-color:#F8FAFC; padding:8px; border-radius:4px; border:1px dashed #E9ECEF;'>📝 기록된 특이사항(OPEN ISSUE) 없음</div>"
            day_lines, night_lines, general_lines = [], [], []
            has_shift = False
            curr = general_lines
            for line in lines:
                cl = line.replace(' ', '')
                if '*주간' in cl or line.startswith('주간'): curr = day_lines; has_shift = True; line = re.sub(r'^\*?\s*주간\s*', '', line).strip()
                elif '*야간' in cl or line.startswith('야간'): curr = night_lines; has_shift = True; line = re.sub(r'^\*?\s*야간\s*', '', line).strip()
                if line: curr.append(line)
            if not has_shift: return f"<div style='font-size:13px; color:#495057; line-height:1.6;'>{'<br>'.join(lines)}</div>"
            day_html = '<br>'.join(general_lines + day_lines) if (general_lines + day_lines) else "<span style='color:#ADB5BD; font-size:12px;'>특이사항 없음</span>"
            night_html = '<br>'.join(night_lines) if night_lines else "<span style='color:#ADB5BD; font-size:12px;'>특이사항 없음</span>"
            return f"<div style='display: flex; gap: 10px; margin-top: 5px; width: 100%; min-width: 400px;'><div style='flex: 1; background-color: #F8FAFC; border: 1px solid #E9ECEF; border-radius: 6px; padding: 12px; border-top: 3px solid #FBBF24;'><div style='font-size:11px; font-weight:bold; color:#B45309; margin-bottom:4px;'>☀️ 주간</div><div style='font-size:13px; color:#495057; line-height:1.6;'>{day_html}</div></div><div style='flex: 1; background-color: #F8FAFC; border: 1px solid #E9ECEF; border-radius: 6px; padding: 12px; border-top: 3px solid #1E293B;'><div style='font-size:11px; font-weight:bold; color:#1E293B; margin-bottom:4px;'>🌙 야간</div><div style='font-size:13px; color:#495057; line-height:1.6;'>{night_html}</div></div></div>"

        # 사이드바 필터 (트렌디 룩 반영)
        st.sidebar.markdown("<h2 style='font-weight: 800; color: #1E293B; font-size: 18px; margin-bottom: 20px;'>🎯 정밀 필터링</h2>", unsafe_allow_html=True)
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_months = [m for m in df['생산월'].unique() if str(m).strip() != ""]
        sel_m_side = st.sidebar.multiselect("📅 생산월 선택", all_months, default=[], placeholder="전체 월")
        m_f_df = df[df['생산월'].isin(sel_m_side)].copy() if sel_m_side else df.copy()
        all_dates = [d for d in m_f_df['생산일'].unique() if str(d).strip() != ""]
        sel_d_side = st.sidebar.multiselect("📆 생산일 선택", all_dates, default=[], placeholder="전체 생산일")
        d_f_df = m_f_df[m_f_df['생산일'].isin(sel_d_side)].copy() if sel_d_side else m_f_df.copy()
        all_machines = sorted([m for m in d_f_df['설비명'].unique() if m.strip() != ""])
        sel_mach_side = st.sidebar.multiselect("⚙️ 설비 선택", all_machines, default=[], placeholder="전체 설비")
        pool_df = d_f_df[d_f_df['설비명'].isin(sel_mach_side)].copy() if sel_mach_side else d_f_df.copy()
        actual_prods = sorted([p for p in pool_df['품명'].fillna("").astype(str).str.strip().unique() if p not in ["", "0", "0.0", "nan", "NaN", "None"]])
        sel_prod = st.sidebar.selectbox("📦 품목 선택", ["전체 품목"] + actual_prods)
        f_df = pool_df[pool_df['품명'].str.strip() == sel_prod].copy() if sel_prod != "전체 품목" else pool_df.copy()

        # [트렌디 표 스타일 개선]
        def render_styler_to_html(styler, is_multi=False):
            try: html_str = styler.to_html(escape=False)
            except: html_str = styler.to_html()
            html_str = html_str.replace('<table', '<table class="custom-table notranslate"')
            wrapped_html = f"""<div style='width: 100%; max-height: 500px; overflow: auto; border: 1px solid #E2E8F0; border-radius: 8px; box-shadow: 0 1px 3px 0 rgba(0,0,0,0.03); background-color: white; margin-bottom: 24px;'><style>.custom-table {{ width: 100%; border-collapse: collapse; font-size: 13px; color: #1E293B; background-color: white; }}.custom-table th {{ background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 12px 16px; text-align: center !important; font-weight: 600; color: #64748B; position: sticky; top: 0; z-index: 2; }}.custom-table thead tr:nth-child(2) th {{ top: 40px; }}.custom-table td {{ border: 1px solid #F1F5F9; padding: 10px 16px; text-align: center !important; }}.custom-table td:last-child {{ text-align: left !important; min-width: 450px; line-height: 1.5; }}</style>{html_str}</div>"""
            if is_multi:
                wrapped_html = re.sub(r'<th class=\"col_heading level0 col10\".*?>OPEN ISSUE</th>', r'<th class=\"col_heading level0 col10\" rowspan=\"2\" style=\"vertical-align: middle;\">OPEN ISSUE</th>', wrapped_html)
                wrapped_html = re.sub(r'<th class=\"col_heading level1 col10\".*?>OPEN ISSUE</th>', '', wrapped_html)
            st.markdown(wrapped_html, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📈 종합 효율 추이", "📝 OPEN ISSUE", "📅 일별 상세 현황", "🏆 효율 BEST&WORST", "🛑 비가동 정밀 분석", "🤖 AI 챗봇"])

        # TAB 1
        with tab1:
            is_fac = (not sel_mach_side and sel_prod == "전체 품목")
            if is_fac: p_df = daily_df.copy(); tgt = 0.86; y_v = '공장종합효율'
            else: act_oee = f_df[f_df['종합효율'] > 0]; p_df = act_oee.groupby(['sort_key', '생산월', '생산일']).mean().reset_index().sort_values('sort_key'); tgt = 0.86; y_v = '종합효율'
            
            render_section_title("최근 5일 종합효율 요약")
            if not p_df.empty:
                r5 = p_df.sort_values('sort_key').tail(5); r_cols = st.columns(5)
                for i, (_, r) in enumerate(r5.iterrows()):
                    oee = safe_float(r[y_v])
                    clr, _, icn = get_status_color(oee, tgt)
                    with r_cols[i]: 
                        # 🌟 [트렌디 지표] 커스텀 Metric 함수 사용
                        render_trendy_metric(r['생산일'], f"{oee:.1%}", clr, icn)
            
            render_section_title("월별 설비 가동 현황 및 종합효율 추이")
            mons = list(dict.fromkeys(p_df['생산월'].tolist()))
            sel_mons = st.multiselect("📅 조회할 월을 선택하세요", mons, default=[mons[-1]] if mons else [], key='t1_m')
            
            for m in sel_mons:
                st.markdown(f"""<div class='trendy-card'>""", unsafe_allow_html=True) # 카드 시작
                
                m_p = p_df[p_df['생산월'] == m].copy()
                if m_p.empty: continue
                m_f_df = f_df[f_df['생산월'] == m].copy()
                mc = m_f_df[m_f_df['종합효율'] > 0].groupby(['sort_key', '생산일'])['설비명'].nunique().reset_index()
                mc.rename(columns={'설비명': '가동대수'}, inplace=True)
                combo = pd.merge(m_p, mc, on=['sort_key', '생산일'], how='left').fillna(0); combo['x_label'] = combo['생산일']
                
                # 🌟 [수정 2] 종합효율 바(Bar) 차트 커스터마이징 (Trendy 컬러 적용)
                text_colors = ['#3B82F6' if safe_float(row[y_v]) >= 0.86 else '#F97316' for _, row in combo.iterrows()]
                fig_oee = go.Figure()
                fig_oee.add_trace(go.Bar(
                    x=combo['x_label'], y=combo[y_v], name='종합효율',
                    text=combo[y_v].apply(lambda x: f"{x:.1%}"), textposition='auto',
                    marker_color=text_colors,
                    textfont=dict(weight='bold', color='white', size=11),
                    hovertemplate='생산일: %{x}<br>종합효율: %{y:.1%}<extra></extra>'
                ))
                fig_oee.add_hline(y=0.86, line_dash='dash', line_color='#F97316', annotation_text='목표 86%')
                # 글로벌 Trendy Look (Transparent BG, 깔끔한 격자)
                fig_oee.update_layout(
                    title=dict(text=f"📊 {m} 종합효율 추이", font=dict(size=16, weight=800, color='#1E293B')),
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=10, r=10, t=50, b=10),
                    height=350,
                    xaxis=dict(gridcolor='#F1F5F9'), yaxis=dict(gridcolor='#F1F5F9', tickformat='.0%', range=[0, 1.0])
                )
                st.plotly_chart(fig_oee, use_container_width=True)
                
                # 가동 대수 개별 차트
                fig_mac = go.Figure()
                fig_mac.add_trace(go.Bar(
                    x=combo['x_label'], y=combo['가동대수'], name='가동 대수',
                    marker_color='#94A3B8', text=combo['가동대수'], textposition='outside', # 옅은 블루그레이
                    textfont=dict(weight='bold', size=10, color='#64748B')
                ))
                fig_mac.update_layout(
                    title=dict(text=f"⚙️ {m} 설비 가동 현황", font=dict(size=16, weight=800, color='#1E293B')),
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=10, r=10, t=50, b=10),
                    height=250,
                    xaxis=dict(gridcolor='#F1F5F9'), yaxis=dict(gridcolor='#F1F5F9', title_text="대수", range=[0, combo['가동대수'].max() * 1.5])
                )
                st.plotly_chart(fig_mac, use_container_width=True)
                st.markdown("""</div>""", unsafe_allow_html=True) # 카드 끝

        # TAB 2
        with tab2:
            render_section_title("OPEN ISSUE 현황 정밀 조회")
            st.markdown(f"""<div class='trendy-card'>""", unsafe_allow_html=True) # 카드 시작
            mons2 = list(dict.fromkeys(f_df['생산월'].tolist()))
            sel_m2 = st.multiselect("📅 월 선택", mons2, default=[mons2[-1]] if mons2 else [], key='t2_m')
            t2_df = f_df[f_df['생산월'].isin(sel_m2)].copy()
            all_d2 = list(reversed([d for d in t2_df['생산일'].unique() if str(d).strip() != ""]))
            if all_d2:
                sd2 = st.selectbox("조회할 생산일 선택", all_d2, key='tab2_date')
                issue_df = t2_df[t2_df['생산일'] == sd2].copy()
                if not issue_df.empty:
                    issue_disp = issue_df[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].copy().sort_values(by='설비명').reset_index(drop=True)
                    for idx, row in issue_disp.iterrows():
                        prod = str(row['품명']).strip()
                        if prod in ['', 'nan', '0', '0.0']: issue_disp.at[idx, '품명'] = ""; issue_disp.at[idx, '종합효율'] = ""
                        else: issue_disp.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
                    issue_disp['OPEN ISSUE'] = issue_disp['OPEN ISSUE'].apply(split_issue_to_columns)
                    render_styler_to_html(issue_disp.style.hide(axis="index"))
                else: st.info("데이터가 없습니다.")
            st.markdown("""</div>""", unsafe_allow_html=True) # 카드 끝

        # TAB 3 (가독성 최적화 적용됨)
        def get_vert_summary_label(r, rank):
            try:
                m_short = str(r.get('설비명', '')).split(' - ')[0].strip()
                p_name = str(r.get('품명', '')).strip()
                p_str = f"<div style='font-size: 13px; color: rgba(255,255,255,0.7); margin-left: 18px; margin-top: 2px;'>{p_name}</div>" if p_name and p_name not in ['', 'nan', '0', '0.0'] else ""
                oee = safe_float(r.get('종합효율', 0))
                return f"<div style='margin-bottom: 12px;'><span style='font-size: 14px;'><b>{rank}. {m_short}</b></span> <span style='float: right; font-weight: 900;'>{oee:.1%}</span>{p_str}</div>"
            except Exception: return ""

        with tab3:
            render_section_title("일일 생산성 상세 현황")
            st.markdown(f"""<div class='trendy-card'>""", unsafe_allow_html=True) # 카드 시작
            mons3 = list(dict.fromkeys(f_df['생산월'].tolist()))
            sel_m3 = st.multiselect("📅 월 선택", mons3, default=[mons3[-1]] if mons3 else [], key='t3_m')
            t3_df = f_df[f_df['생산월'].isin(sel_m3)].copy()
            all_d3 = list(reversed([d for d in t3_df['생산일'].unique() if str(d).strip() != ""]))
            if all_d3:
                sd3 = st.selectbox("📅 생산일 선택", all_d3, key='tab3_date')
                day_df = t3_df[t3_df['생산일'] == sd3].copy().sort_values(by='설비명').reset_index(drop=True)
                day_df['설비_짧은명'] = day_df['설비명'].apply(lambda x: str(x).split(' - ')[0].strip())
                active_day = day_df[day_df['종합효율'] > 0].sort_values(by='종합효율', ascending=False)
                active_count = active_day['설비명'].nunique(); total_count = day_df['설비명'].nunique()
                
                # 🌟 [트렌디 요약판] 좌우 분할 및 세로 나열 (안전벨트 장착)
                best_df = active_day.head(5)
                worst_df = active_day[~active_day.index.isin(best_df.index)].tail(5).sort_values(by='종합효율', ascending=True)
                
                best_html = "".join([get_vert_summary_label(r, i+1) for i, (_, r) in enumerate(best_df.iterrows())])
                if worst_df.empty: worst_html = "<div style='font-size: 13px; color: rgba(255,255,255,0.6); margin-top: 10px;'>해당 없음 (가동 설비 부족)</div>"
                else: worst_html = "".join([get_vert_summary_label(r, i+1) for i, (_, r) in enumerate(worst_df.iterrows())])
                
                st.markdown(f"""<div class='dashboard-header'>
<div style='font-size: 16px; opacity: 0.9; margin-bottom: 5px; font-weight: 500;'>💡 {sd3} 생산 요약</div>
<div style='font-size: 32px; font-weight: 900; margin-bottom: 25px; letter-spacing: -1.5px;'>총 <span style='color: #E2E8F0;'>{total_count}</span>대 중 <span style='color: #FBBF24;'>{active_count}</span>대 가동 중</div>
<div style='display: flex; gap: 30px; background-color: rgba(255,255,255,0.1); padding: 25px; border-radius: 10px; text-align: left; box-sizing: border-box; line-height: 1.4; border: 1px solid rgba(255,255,255,0.05);'>
<div style='flex: 1;'><div style='color: #20C997; font-weight: 900; font-size: 17px; margin-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.15); padding-bottom: 10px; letter-spacing:-0.5px;'>🏆 종합효율 BEST 5</div>{best_html}</div>
<div style='width: 1px; background-color: rgba(255,255,255,0.15);'></div>
<div style='flex: 1;'><div style='color: #FF7E7E; font-weight: 900; font-size: 17px; margin-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.15); padding-bottom: 10px; letter-spacing:-0.5px;'>🚨 종합효율 WORST 5</div>{worst_html}</div>
</div></div>""", unsafe_allow_html=True)
                
                st.markdown(f"#### 📊 {sd3} 설비별 종합효율 비교")
                if not active_day.empty:
                    bar_colors = ['#3B82F6' if safe_float(row['종합효율']) >= 0.86 else '#F97316' for _, row in active_day.iterrows()]
                    fig3 = px.bar(active_day, x='설비_짧은명', y='종합효율', text_auto='.1%', hover_data=['품명'])
                    fig3.update_traces(marker_color=bar_colors, textposition="auto", textfont=dict(size=11, weight="bold", color='white'))
                    fig3.update_xaxes(title="", tickangle=0, gridcolor='#F1F5F9')
                    # 🌟 [수정 2] Y축 100% 제한
                    fig3.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.0], showgrid=True, gridcolor='#F1F5F9')
                    fig3.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=400, margin=dict(l=10, r=10, t=10, b=10))
                    st.plotly_chart(fig3, use_container_width=True)
                else: st.info("해당 일자에 가동된 설비가 없습니다.")
                
                st.write("---")
                downtime_day = day_df[day_df['비가동시간']>0].sort_values(by='비가동시간', ascending=False)
                fig4 = px.bar(downtime_day, x='설비_짧은명', y='비가동시간', text_auto='.1f', title=f"🛑 {sd3} 설비별 비가동 현황 (시간)")
                fig4.update_xaxes(title="", tickangle=0, gridcolor='#F1F5F9')
                fig4.update_yaxes(title="시간(h)", gridcolor='#F1F5F9')
                fig4.update_traces(marker_color='#EF4444') # 비가동은 강렬한 레드
                fig4.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', height=350, margin=dict(l=10, r=10, t=50, b=10), title_font=dict(size=16, weight=800, color='#1E293B'))
                st.plotly_chart(fig4, use_container_width=True)

                disp_day = day_df[target_order].copy()
                for idx, row in disp_day.iterrows():
                    prod = str(row['품명']).strip()
                    if prod in ['', 'nan', '0', '0.0']:
                        disp_day.at[idx, '품명'] = ""
                        for c in ['종합효율', '양품율', '성능가동율', '시간가동율', '양품수량', '불량수량', '총 생산수량']: if c in disp_day.columns: disp_day.at[idx, c] = ""
                    else:
                        for c in ['종합효율', '양품율', '성능가동율', '시간가동율']: if c in disp_day.columns: disp_day.at[idx, c] = f"{safe_float(row[c]):.1%}"
                        for c in ['양품수량', '불량수량', '총 생산수량']: if c in disp_day.columns: disp_day.at[idx, c] = f"{int(safe_float(row[c])):,}"
                disp_day['OPEN ISSUE'] = disp_day['OPEN ISSUE'].apply(split_issue_to_columns)
                disp_day.columns = pd.MultiIndex.from_tuples(multi_cols)
                
                def style_day_row(row):
                    styles = [''] * len(row)
                    idx = row.name
                    try:
                        if 0 < safe_float(day_df.loc[idx, '종합효율']) < safe_float(day_df.loc[idx, '목표효율']):
                            pos = row.index.get_loc(('생산성', '종합효율')); if isinstance(pos, np.ndarray): pos = np.where(pos)[0][0]
                            styles[pos] = 'color: #F97316; font-weight: 800;' # 미달시 모던 오렌지
                    except: pass
                    return styles
                render_styler_to_html(disp_day.style.apply(style_day_row, axis=1).hide(axis="index"), is_multi=True)
            st.markdown("""</div>""", unsafe_allow_html=True) # 카드 끝

        # TAB 4
        with tab4:
            render_section_title("종합효율 BEST 5 & WORST 5")
            st.markdown(f"""<div class='trendy-card'>""", unsafe_allow_html=True) # 카드 시작
            mons4 = list(dict.fromkeys(f_df['생산월'].tolist()))
            sel_m4 = st.multiselect("📅 월 선택", mons4, default=[mons4[-1]] if mons4 else [], key='t4_m')
            t4_df = f_df[(f_df['생산월'].isin(sel_m4)) & (f_df['종합효율'] > 0)].copy()
            if not t4_df.empty:
                for label, asc in [("🏆 BEST 5 (최고 효율)", False), ("🚨 WORST 5 (최저 효율)", True)]:
                    # 표 제목 스타일 개선
                    color = "#20C997" if not asc else "#FF7E7E"
                    st.markdown(f"<h4 style='font-weight: 800; color: #1E293B; margin-top: 15px; margin-bottom: 15px; font-size: 16px;'><span style='color: {color}; margin-right: 8px;'>■</span>{label}</h4>", unsafe_allow_html=True)
                    res = t4_df.sort_values(by='종합효율', ascending=asc).head(5)
                    res_disp = res[['생산일', '설비명', '품명', '종합효율', 'OPEN ISSUE']].copy()
                    for idx, row in res_disp.iterrows():
                        prod = str(row['품명']).strip()
                        if prod in ['', 'nan', '0', '0.0']: res_disp.at[idx, '품명'] = ""; res_disp.at[idx, '종합효율'] = ""
                        else: res_disp.at[idx, '종합효율'] = f"{safe_float(row['종합효율']):.1%}"
                    res_disp['OPEN ISSUE'] = res_disp['OPEN ISSUE'].apply(split_issue_to_columns)
                    render_styler_to_html(res_disp.style.hide(axis="index"))
            st.markdown("""</div>""", unsafe_allow_html=True) # 카드 끝

        # TAB 5
        with tab5:
            render_section_title("비가동시간 요인 정밀 분석")
            st.markdown(f"""<div class='trendy-card'>""", unsafe_allow_html=True) # 카드 시작
            mons5 = list(dict.fromkeys(f_df['생산월'].tolist()))
            sel_m5 = st.multiselect("📅 월 선택", mons5, default=[mons5[-1]] if mons5 else [], key='t5_m')
            t5_df = f_df[f_df['생산월'].isin(sel_m5)].copy()
            if not t5_df.empty:
                st.markdown("<h4 style='font-weight: 800; color: #1E293B; margin-top: 15px; margin-bottom: 15px; font-size: 16px;'><span style='color: #EF4444; margin-right: 8px;'>■</span>🚨 WORST 10 (비가동 최장 설비)</h4>", unsafe_allow_html=True)
                res = t5_df.sort_values(by='비가동시간', ascending=False).head(10)
                res_disp = res[['생산일', '설비명', '품명', '비가동시간', 'OPEN ISSUE']].copy()
                for idx, row in res_disp.iterrows():
                    prod = str(row['품명']).strip()
                    if prod in ['', 'nan', '0', '0.0']: res_disp.at[idx, '품명'] = ""
                res_disp['비가동시간'] = res_disp['비가동시간'].apply(lambda x: f"{safe_float(x):.1f}h")
                res_disp['OPEN ISSUE'] = res_disp['OPEN ISSUE'].apply(split_issue_to_columns)
                render_styler_to_html(res_disp.style.hide(axis="index"))
            else: st.info("해당 월에 분석할 비가동 데이터가 없습니다.")
            st.markdown("""</div>""", unsafe_allow_html=True) # 카드 끝

        # TAB 6
        with tab6:
            render_section_title("🤖 AI 생산 데이터 챗봇 (데이터 분석가)")
            st.markdown(f"""<div class='trendy-card'>""", unsafe_allow_html=True) # 카드 시작
            ak = st.text_input("🔑 OpenAI API Key 입력", type="password", placeholder="sk-...")
            st.markdown("<div style='font-size: 13px; color: #64748B; margin-top: -15px; margin-bottom: 20px;'>💡 입력하신 키는 서버에 저장되지 않으며, 브라우저 종료 시 소멸됩니다.</div>", unsafe_allow_html=True)
            if "msgs" not in st.session_state: st.session_state.msgs = [{"role": "assistant", "content": "사출생산팀 데이터 분석을 도와드리는 AI 챗봇입니다. 무엇을 분석해 드릴까요?"}]
            for m in st.session_state.msgs: st.chat_message(m["role"]).write(m["content"])
            if pr := st.chat_input("설비 가동률 추이에 대해 질문해보세요"):
                st.session_state.msgs.append({"role": "user", "content": pr}); st.chat_message("user").write(pr)
                if not ak: st.chat_message("assistant").write("💡 API 키를 상단에 입력하시면 실제 분석이 시작됩니다.")
                else: st.chat_message("assistant").write("데이터를 분석하고 답변을 구성 중입니다... (키를 확인해주세요)")
            st.markdown("""</div>""", unsafe_allow_html=True) # 카드 끝

else: st.info("GitHub의 'data' 폴더에 CSV 또는 XLSX 생산 실적 파일을 넣어주세요.")
