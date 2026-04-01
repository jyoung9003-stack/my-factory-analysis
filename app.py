import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
import os
from collections import Counter

# 1. 웹 화면 및 차분한 폰트/스타일 설정
st.set_page_config(page_title="사출생산팀 일일 생산성 정밀 분석", layout="wide")

st.markdown("""
<style>
    /* 전체 폰트를 세련되고 가독성 높은 폰트로 고정 */
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
        background-color: #F8F9FA; /* 눈이 편안한 아주 옅은 그레이 배경 */
    }
    
    /* 요약 카드(Best/Worst) 스타일 */
    .metric-card {
        background-color: white;
        border: 1px solid #E9ECEF;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-title { font-size: 14px; color: #6C757D; font-weight: bold; margin-bottom: 5px; }
    .metric-value.best { font-size: 20px; color: #1F77B4; font-weight: 900; }
    .metric-value.worst { font-size: 20px; color: #FF4B4B; font-weight: 900; }
    .metric-date { font-size: 12px; color: #ADB5BD; }
</style>
""", unsafe_allow_html=True)

# 2. 로고 및 타이틀
col1, col2 = st.columns([1, 10])
with col1:
    logo_path = "듀링로고_가로형_빨강_JPG.jpg"
    if os.path.exists(logo_path):
        st.image(logo_path, width=100)
    else:
        st.markdown("<h2 style='color: #FF2A2A; font-weight: 900; margin-top: 10px;'>DÜRING</h2>", unsafe_allow_html=True)
with col2:
    st.markdown("<h1 style='margin-top: 0px; color: #212529;'>사출생산팀 일일 생산성 정밀 분석</h1>", unsafe_allow_html=True)

# 3. 데이터 업로드 및 전처리
uploaded_files = st.file_uploader("분석할 파일들을 선택하세요 (여러 개 선택 가능)", type=['xlsx', 'csv'], accept_multiple_files=True)

target_cols = ['생산일', '설비명', '품명', '양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '양품율', '성능가동율', '시간가동율', '종합효율', '목표효율', 'OPEN ISSUE']

if uploaded_files:
    all_records = []
    daily_totals_data = {} 
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                temp_df = pd.read_csv(file)
            else:
                temp_df = pd.read_excel(file)
            
            temp_df.columns = [str(c).replace('\n', '').replace('\r', '').strip() for c in temp_df.columns]
            
            name_map = {
                '작업장 [설비]': '설비명', '작업장[설비]': '설비명', '품목명': '품명',
                '합계': '총 생산수량', '합게수량': '총 생산수량', '종합 효율': '종합효율', '목표 효율': '목표효율'
            }
            temp_df = temp_df.rename(columns=name_map)
            
            for col in temp_df.columns:
                if 'Unnamed' in col or 'ISSUE' in col.upper():
                    temp_df = temp_df.rename(columns={col: 'OPEN ISSUE'})
                    break
            
            date_match = re.search(r'\d{8}', file.name)
            clean_date = date_match.group()[2:] if date_match else file.name.split('.')[0]
            
            daily_total_oee = None
            for _, row in temp_df.iterrows():
                machine_val = str(row.get('설비명', ''))
                if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
                if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper():
                    val = row.get('종합효율', None)
                    if isinstance(val, pd.Series): val = val.iloc[0]
                    try: daily_total_oee = float(val)
                    except: pass
                    break
            
            if daily_total_oee is None:
                try: daily_total_oee = float(temp_df['종합효율'].iloc[45]) if not isinstance(temp_df['종합효율'].iloc[45], pd.Series) else float(temp_df['종합효율'].iloc[45].iloc[0])
                except: daily_total_oee = 0
            
            daily_totals_data[clean_date] = daily_total_oee

            for _, row in temp_df.iterrows():
                machine_val = str(row.get('설비명', ''))
                if isinstance(machine_val, pd.Series): machine_val = str(machine_val.iloc[0])
                if 'TOTAL' in machine_val.upper() or '합계' in machine_val or 'GRAND' in machine_val.upper():
                    continue
                    
                record = {'생산일': clean_date}
                for col in target_cols:
                    if col == '생산일': continue
                    if col in temp_df.columns:
                        val = row[col]
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        record[col] = val
                    else: record[col] = None
                all_records.append(record)
                
        except Exception as e:
            st.error(f"파일 읽기 오류 ({file.name}): {e}")

    if all_records:
        df = pd.DataFrame(all_records)
        daily_df = pd.DataFrame(list(daily_totals_data.items()), columns=['생산일', '공장종합효율']).sort_values(by='생산일')
        
        num_cols = ['양품수량', '불량수량', '총 생산수량', '투입시간', '가동시간', '비가동시간', '정미시간', '종합효율', '목표효율', '양품율', '성능가동율', '시간가동율']
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        def format_issue(text):
            val = str(text).strip()
            if val in ['', '0', '0.0', 'nan', 'NaN', 'None']: return ""
            val = val.replace('\r\n', '\n')
            val = re.sub(r'(?<!\n)\*', '\n*', val)
            val = re.sub(r'(?<!\n)-\.', '\n-.', val)
            val = re.sub(r'(?<!\n)→', '\n→', val)
            return val.strip()

        df['OPEN ISSUE'] = df['OPEN ISSUE'].apply(format_issue)

        # ---------------------------------------------------------
        # [사이드바 필터]
        # ---------------------------------------------------------
        st.sidebar.header("🎯 정밀 필터링")
        df['설비명'] = df['설비명'].fillna("").astype(str)
        all_machines = sorted([m for m in df['설비명'].unique() if m.strip() != ""])
        selected_machines = st.sidebar.multiselect("설비 선택", all_machines, default=[], placeholder="전체 설비")
        
        if len(selected_machines) == 0: pool_df = df.copy()
        else: pool_df = df[df['설비명'].isin(selected_machines)].copy()
            
        pool_df['품명_필터'] = pool_df['품명'].fillna("").astype(str).str.strip()
        pool_df['품명_필터'] = pool_df['품명_필터'].replace(['0', '0.0', 'nan', 'NaN', 'None'], "")
        actual_prods = sorted([p for p in pool_df['품명_필터'].unique() if p != ""])
        selected_prod = st.sidebar.selectbox("품목 선택 (해당 설비 생산품)", ["전체 품목"] + actual_prods)

        f_df = pool_df.copy()
        if selected_prod != "전체 품목":
            f_df = f_df[f_df['품명_필터'] == selected_prod]

        # ---------------------------------------------------------
        # [그래프 분석 탭]
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📈 생산 추이 및 요약", "📝 OPEN ISSUE 정밀 분석"])

        with tab1:
            is_factory_view = (len(selected_machines) == 0 and selected_prod == "전체 품목")
            
            if is_factory_view:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 사출생산팀 일별 종합 효율(%)</h3>", unsafe_allow_html=True)
                plot_df = daily_df.copy()
                plot_df['목표효율'] = 0.86 # 공장 전체 목표는 86% 고정
                y_val = '공장종합효율'
            else:
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 선택 조건(설비/품목) 일별 평균 종합 효율(%)</h3>", unsafe_allow_html=True)
                active_oee = f_df[f_df['종합효율'] > 0]
                # 설비/품목 필터 시 해당 데이터들의 평균 목표효율을 계산하여 타겟으로 잡음
                plot_df = active_oee.groupby('생산일')[['종합효율', '목표효율']].mean().reset_index().sort_values(by='생산일')
                y_val = '종합효율'

            if not plot_df.empty:
                # 🚨 [수정 1] 미달 시에만 빨간색 적용 (차분한 디자인 유지)
                colors = ['#FF4B4B' if row[y_val] < row['목표효율'] else '#1F77B4' for _, row in plot_df.iterrows()]
                
                plot_df['x_label'] = plot_df.apply(lambda row: f"{row['생산일']}<br><span style='font-size:11px;color:gray;'>({row[y_val]:.1%})</span>", axis=1)
                
                fig1 = go.Figure()
                # 베이스 파란색 라인 (면적 포함)
                fig1.add_trace(go.Scatter(
                    x=plot_df['x_label'], y=plot_df[y_val], mode='lines',
                    line=dict(shape='spline', width=3, color='#1F77B4'),
                    fill='tozeroy', fillcolor='rgba(31, 119, 180, 0.05)', hoverinfo='skip', showlegend=False
                ))
                # 데이터 점(Marker)과 텍스트 - 효율 미달 시에만 빨간색!
                fig1.add_trace(go.Scatter(
                    x=plot_df['x_label'], y=plot_df[y_val], mode='markers+text',
                    text=plot_df[y_val].apply(lambda x: f'{x:.1%}'), textposition="top center",
                    marker=dict(size=10, color='white', line=dict(width=2.5, color=colors)),
                    textfont=dict(size=14, color=colors, weight="bold"), showlegend=False, cliponaxis=False
                ))
                
                # 목표 효율 범례 (우측 상단)
                target_val = 0.86 if is_factory_view else plot_df['목표효율'].mean()
                fig1.add_trace(go.Scatter(
                    x=plot_df['x_label'], y=[target_val] * len(plot_df), mode='lines', 
                    name=f'목표 효율 ({target_val:.1%})', line=dict(color='#ADB5BD', dash='dash', width=2)
                ))
                
                fig1.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(plot_df)-0.5]) 
                fig1.update_yaxes(title="종합효율", tickformat='.0%', range=[0, 1.2], showgrid=True, gridcolor='rgba(230,230,230,0.5)') 
                fig1.update_layout(height=450, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white', legend=dict(yanchor="top", y=1.1, xanchor="right", x=1))
                st.plotly_chart(fig1, use_container_width=True)
                
                # 🚨 [수정 2 & 3] BEST 3 / WORST 3 요약 카드
                sorted_df = plot_df.sort_values(by=y_val, ascending=False)
                best_3 = sorted_df.head(3)
                worst_3 = sorted_df.tail(3).sort_values(by=y_val, ascending=True)
                
                st.markdown("#### 🏆 종합효율 BEST 3 & 🚨 WORST 3")
                b_cols = st.columns(3)
                w_cols = st.columns(3)
                
                for i, (_, r) in enumerate(best_3.iterrows()):
                    with b_cols[i]:
                        st.markdown(f"<div class='metric-card'><div class='metric-title'>BEST {i+1}</div><div class='metric-value best'>{r[y_val]:.1%}</div><div class='metric-date'>{r['생산일']}</div></div>", unsafe_allow_html=True)
                for i, (_, r) in enumerate(worst_3.iterrows()):
                    with w_cols[i]:
                        st.markdown(f"<div class='metric-card'><div class='metric-title'>WORST {i+1}</div><div class='metric-value worst'>{r[y_val]:.1%}</div><div class='metric-date'>{r['생산일']}</div></div>", unsafe_allow_html=True)

            else:
                st.info("종합효율 데이터가 없습니다.")
            
            st.write("---")
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일별 총 비가동 시간(분)</h3>", unsafe_allow_html=True)
            daily_stop = f_df.groupby('생산일')['비가동시간'].sum().reset_index().sort_values(by='생산일')
            fig2 = px.bar(daily_stop, x='생산일', y='비가동시간', text_auto='.1f')
            
            # 비가동 시간 차분한 오렌지/레드 톤 적용
            fig2.update_traces(marker=dict(color='#E07A5F', opacity=0.9), textposition="outside", textfont=dict(size=13, weight="bold", color="#E07A5F"), cliponaxis=False)
            fig2.update_xaxes(type='category', title="", showgrid=False, range=[-0.5, len(daily_stop)-0.5])
            fig2.update_yaxes(title="비가동시간 (분)", showgrid=True, gridcolor='rgba(230,230,230,0.5)')
            fig2.update_layout(height=350, margin=dict(l=40, r=40, t=40, b=40), plot_bgcolor='white', paper_bgcolor='white')
            st.plotly_chart(fig2, use_container_width=True)

        # ---------------------------------------------------------
        # [OPEN ISSUE & 키워드 분석 탭]
        # ---------------------------------------------------------
        with tab2:
            st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 빈출 이슈 (OPEN ISSUE 키워드 분석)</h3>", unsafe_allow_html=True)
            st.markdown("선택된 기간/조건 동안 가장 많이 발생한 특이사항의 핵심 단어입니다. (효율 개선 집중 포인트)")
            
            issue_df = f_df[f_df['OPEN ISSUE'] != ""].copy()
            
            if not issue_df.empty:
                # 🚨 [수정 5] 자연어(NLP) 처리 - 빈출 불량/이슈 단어 랭킹
                all_text = " ".join(issue_df['OPEN ISSUE'].astype(str))
                words = re.findall(r'[가-힣]+', all_text) # 한글 단어만 추출
                # 불용어(의미 없는 단어) 제거
                stopwords = {'주간', '야간', '확인', '점검', '가동', '조치', '이상', '완료', '발생', '후', '및', '설비', '생산', '연속', '특이사항', '동작'}
                filtered_words = [w for w in words if w not in stopwords and len(w) > 1] # 2글자 이상 의미 있는 단어만
                
                if filtered_words:
                    word_counts = Counter(filtered_words).most_common(5)
                    wc_cols = st.columns(len(word_counts))
                    for i, (word, count) in enumerate(word_counts):
                        with wc_cols[i]:
                            st.markdown(f"<div style='background-color:#FFF5F5; padding:10px; border-radius:5px; text-align:center; border:1px solid #FFD6D6;'><div style='font-size:16px; font-weight:bold; color:#D32F2F;'>{word}</div><div style='font-size:12px; color:#555;'>{count}건 발생</div></div>", unsafe_allow_html=True)
                else:
                    st.write("반복되는 유의미한 키워드가 감지되지 않았습니다.")
                
                st.write("---")
                
                # OPEN ISSUE 테이블 커스텀 렌더링
                st.markdown("<h3 style='font-weight: 800; color: #212529;'><span style='color: #FF4B4B;'>■</span> 일별 특이사항 상세 내역</h3>", unsafe_allow_html=True)
                
                issue_display = issue_df[['생산일', '설비명', '품명', '종합효율', '목표효율', 'OPEN ISSUE']].sort_values(by=['생산일', '설비명'])
                
                # HTML 렌더링용 변환
                html_rows = ""
                for _, row in issue_display.iterrows():
                    oee = float(row['종합효율']) if pd.notnull(row['종합효율']) else 0
                    tgt = float(row['목표효율']) if pd.notnull(row['목표효율']) else 0
                    
                    # 목표 미달일 때만 종합효율 숫자를 빨간색으로!
                    oee_str = f"<span style='color: #FF4B4B; font-weight: bold;'>{oee:.1%}</span>" if oee < tgt else f"{oee:.1%}"
                    issue_text = str(row['OPEN ISSUE']).replace('\n', '<br>')
                    
                    html_rows += f"""
                    <tr>
                        <td style='text-align:center;'>{row['생산일']}</td>
                        <td style='text-align:center;'>{row['설비명']}</td>
                        <td style='text-align:center;'>{row['품명']}</td>
                        <td style='text-align:center;'>{oee_str}</td>
                        <td style='white-space:pre-wrap; text-align:left; min-width:350px;'>{issue_text}</td>
                    </tr>
                    """
                    
                table_html = f"""
                <div style="width: 100%; max-height: 500px; overflow: auto; border: 1px solid #DEE2E6; border-radius: 5px;">
                    <table style="width: 100%; border-collapse: collapse; background-color: white; font-size: 13.5px; color: #212529;">
                        <thead>
                            <tr style="background-color: #F8F9FA; position: sticky; top: 0; border-bottom: 2px solid #DEE2E6;">
                                <th style="padding: 10px; text-align: center; border-right: 1px solid #DEE2E6;">생산일</th>
                                <th style="padding: 10px; text-align: center; border-right: 1px solid #DEE2E6;">설비명</th>
                                <th style="padding: 10px; text-align: center; border-right: 1px solid #DEE2E6;">품명</th>
                                <th style="padding: 10px; text-align: center; border-right: 1px solid #DEE2E6;">종합효율</th>
                                <th style="padding: 10px; text-align: center;">OPEN ISSUE</th>
                            </tr>
                        </thead>
                        <tbody>
                            {html_rows}
                        </tbody>
                    </table>
                </div>
                """
                st.markdown(table_html, unsafe_allow_html=True)
            else:
                st.info("선택된 조건에 해당하는 특이사항(OPEN ISSUE)이 없습니다.")

else:
    st.info("왼쪽 상단에서 생산성 파일을 업로드해 주세요.")
