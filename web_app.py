import streamlit as st
import gspread
import datetime
import pandas as pd
import json
from io import BytesIO

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑", layout="wide")

# 2. 구글 시트 연결 (💡 클라우드 배포용 보안 코드 적용 완료!)
@st.cache_resource
def get_gspread_client():
    try:
        # 웹 서버(Streamlit Cloud) 환경일 때: 서버의 비밀 공간(Secrets)에서 읽어옵니다.
        creds_json = st.secrets["google_credentials"]
        credentials = json.loads(creds_json)
        return gspread.service_account_from_dict(credentials)
    except:
        # 내 PC(로컬) 환경일 때: 기존처럼 로컬 파일을 읽어옵니다.
        return gspread.service_account(filename="now_secret.json")

try:
    gc = get_gspread_client()
    sh = gc.open("업무관리_DB")
    w_d = sh.worksheet("일일업무")
    w_p = sh.worksheet("프로젝트")
    w_s = sh.worksheet("세부업무")
    w_st = sh.worksheet("설정")
    w_u = sh.worksheet("계정관리")
except Exception as e:
    st.error(f"구글 시트 연결 실패: {e}")
    st.stop()

# --- [보안 및 로그인 로직] ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_info' not in st.session_state:
    st.session_state['user_info'] = {}

def check_login(user_id, user_pw):
    users = w_u.get_all_records()
    for u in users:
        if str(u.get('아이디')) == str(user_id) and str(u.get('비밀번호')) == str(user_pw):
            st.session_state['logged_in'] = True
            st.session_state['user_info'] = u
            return True
    return False

if not st.session_state['logged_in']:
    st.markdown("<h1 style='text-align: center;'>🔒 NOWSYSTEM 관제탑</h1>", unsafe_allow_html=True)
    _, l_col, _ = st.columns([1, 1, 1])
    with l_col:
        with st.form("login"):
            in_id = st.text_input("아이디")
            in_pw = st.text_input("비밀번호", type="password")
            if st.form_submit_button("접속하기", use_container_width=True):
                if check_login(in_id, in_pw): st.rerun()
                else: st.error("정보가 일치하지 않습니다.")
    st.stop()

u_info = st.session_state['user_info']
u_name = u_info.get('이름', '사용자')
u_role = u_info.get('권한', '일반')

# --- [데이터 로드 및 개인 KPI 필터링] ---
all_daily = w_d.get_all_records()
proj_data = w_p.get_all_records()
sub_data = w_s.get_all_records()

kpi_config = w_st.get_all_records()
for k in kpi_config:
    if '구분' not in k: k['구분'] = '공통'

my_kpi_opts = [str(k.get('KPI명', '')) for k in kpi_config if str(k.get('구분', '공통')).strip() in ['공통', u_name]]

# --- [사이드바 및 날짜 제어] ---
with st.sidebar:
    st.subheader(f"👤 {u_name}님 ({u_role})")
    if st.button("🚪 로그아웃", use_container_width=True): 
        st.session_state['logged_in'] = False
        st.rerun()

st.title("🚀 NOWSYSTEM 통합 업무 관리")
t_date = st.date_input("📅 업무 기준일 선택", datetime.date.today(), key="main_date")
t_str = t_date.strftime("%Y-%m-%d")

if f"lock_{t_str}" not in st.session_state: st.session_state[f"lock_{t_str}"] = False
is_locked = st.session_state[f"lock_{t_str}"]

# --- [탭 구성] ---
if u_role == "마스터":
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str]
    tabs = st.tabs(["📝 전사 업무 관리", "📁 프로젝트 관리", "⚙️ 마스터 설정", "📈 통합 KPI 현황", "📊 데이터 및 보고서"])
    tab_kpi = tabs[3]
    tab_rep = tabs[4]
else:
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str and d.get('담당자') == u_name]
    tabs = st.tabs(["📝 나의 일과", "📁 나의 프로젝트", "📈 나의 KPI 현황", "📊 나의 데이터 및 보고서"])
    tab_kpi = tabs[2]
    tab_rep = tabs[3]

sub_dict = {p.get("프로젝트명", ""): [] for p in proj_data}
for idx, s in enumerate(sub_data):
    if u_role == "마스터" or s.get('담당자') == u_name:
        pn = s.get("프로젝트명", "")
        if pn in sub_dict:
            s['_r'] = idx + 2
            sub_dict[pn].append(s)

# ==========================================
# 탭 1: 일과 관리
# ==========================================
with tabs[0]:
    st.header(f"📝 {t_str} 업무")
    with st.expander("➕ 새 업무 추가", expanded=not is_locked):
        c1, c2, c3 = st.columns([2,1,1])
        n_task = c1.text_area("업무명 (엔터로 줄바꿈 가능)", key="n_d_n", disabled=is_locked)
        n_cat = c2.selectbox("분류", ["경영관리", "재무업무", "입찰업무", "조달업무", "현장업무", "기타"], key="n_d_c", disabled=is_locked)
        
        dropdown_opts = [k.get('KPI명', '') for k in kpi_config] if u_role == "마스터" else my_kpi_opts
        n_kpi = c3.selectbox("연관 KPI", dropdown_opts + ["기타"], key="n_d_k", disabled=is_locked)
        
        if st.button("저장", disabled=is_locked):
            if n_task:
                w_d.append_row([t_str, n_task, 0, "FALSE", n_cat, "", n_kpi, u_name])
                st.rerun()
    
    st.divider()
    for i, row in enumerate(filtered_daily):
        r_idx = all_daily.index(row) + 2
        col1, col2, col3 = st.columns([4, 5, 1])
        disp_name = str(row['업무명']).replace('\n', '<br>')
        badge = f" <small style='color:blue;'>[{row.get('담당자','')}]</small>" if u_role == "마스터" else ""
        col1.markdown(f"**[{row['분류']}]** {disp_name}{badge}", unsafe_allow_html=True)
        new_p = col2.slider("진행", 0, 100, int(row.get('진행률', 0)), 10, key=f"d_sld_{r_idx}", disabled=is_locked, label_visibility="collapsed")
        if not is_locked and new_p != int(row.get('진행률', 0)):
            w_d.update_cell(r_idx, 2, new_p); st.toast("완료")
        if col3.button("🗑️", key=f"d_del_{r_idx}", disabled=is_locked):
            w_d.delete_rows(r_idx); st.rerun()

# ==========================================
# 탭 2: 프로젝트 관리
# ==========================================
with tabs[1]:
    st.header("📁 프로젝트 현황")
    with st.expander("✨ 신규 프로젝트 등록", expanded=False):
        pc1, pc2 = st.columns(2)
        p_name = pc1.text_input("프로젝트명", disabled=is_locked)
        p_cat = pc2.selectbox("분류", ["경영관리", "재무업무", "입찰업무", "조달업무", "현장업무", "기타"], key="p_c", disabled=is_locked)
        p_start = pc1.date_input("시작일", key="p_s", disabled=is_locked)
        p_kpi = pc2.selectbox("연관 KPI", dropdown_opts + ["기타"], key="p_k", disabled=is_locked)
        
        if st.button("프로젝트 저장", disabled=is_locked):
            if p_name:
                w_p.append_row([p_name, str(p_start), "", "FALSE", "FALSE", "FALSE", "", p_cat, p_kpi, u_name])
                st.rerun()

    for i, p in enumerate(proj_data):
        r_idx = i + 2
        if (u_role != "마스터" and p.get('담당자') != u_name) or str(p.get("보관함이동", "FALSE")).upper() == "TRUE": continue
            
        pn = p.get("프로젝트명", "")
        owner = f" ({p.get('담당자','')})" if u_role == "마스터" else ""
        
        with st.expander(f"📂 {pn} [{p.get('분류')}]{owner}"):
            with st.form(key=f"sub_form_{r_idx}"):
                sc1, sc2 = st.columns([3,1])
                new_sub = sc1.text_input("세부 업무명")
                if sc2.form_submit_button("추가", disabled=is_locked):
                    if new_sub:
                        w_s.append_row([pn, new_sub, 0, "TRUE", u_name])
                        st.rerun()
            for s in sub_dict.get(pn, []):
                sl1, sl2 = st.columns([6, 4])
                sl1.markdown(f"· {str(s['세부업무명']).replace('\n','<br>')}", unsafe_allow_html=True)
                sp = sl2.slider("진행", 0, 100, int(s.get('진행률',0)), 10, key=f"s_sld_{s['_r']}", disabled=is_locked, label_visibility="collapsed")
                if not is_locked and sp != int(s.get('진행률',0)):
                    w_s.update_cell(s['_r'], 3, sp); st.toast("완료")
            
            st.write("---")
            ac1, ac2, ac3 = st.columns([2,1,1])
            if ac2.button("📦 보관함 이동", key=f"arc_{r_idx}", disabled=is_locked):
                w_p.update_cell(r_idx, 6, "TRUE"); st.rerun()
            if ac3.button("🗑️ 삭제", key=f"pdel_{r_idx}", disabled=is_locked):
                w_p.delete_rows(r_idx); st.rerun()

# ==========================================
# 탭 3: 마스터 전용 설정
# ==========================================
if u_role == "마스터":
    with tabs[2]:
        st.header("⚙️ 마스터 전용 설정")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("👥 사내 계정 목록")
            u_df = pd.DataFrame(w_u.get_all_records())
            e_u_df = st.data_editor(u_df, num_rows="dynamic", use_container_width=True)
            if st.button("계정 정보 저장"):
                w_u.clear(); w_u.update([e_u_df.columns.values.tolist()] + e_u_df.values.tolist())
                st.success("계정 정보가 저장되었습니다.")
        
        with c2:
            st.subheader("🎯 개인별 & 공통 KPI 세부 관리")
            st.caption("💡 '구분' 칸에 '공통'을 적으면 전 직원에게, 직원이름을 적으면 해당 직원의 전용 KPI로 지정됩니다.")
            k_df = pd.DataFrame(kpi_config)
            
            if '구분' not in k_df.columns:
                k_df['구분'] = '공통'
            cols = k_df.columns.tolist()
            if '구분' in cols:
                cols.insert(0, cols.pop(cols.index('구분')))
                k_df = k_df[cols]

            e_k_df = st.data_editor(k_df, num_rows="dynamic", use_container_width=True)
            if st.button("전체 KPI 지표 저장"):
                w_st.clear(); w_st.update([e_k_df.columns.values.tolist()] + e_k_df.values.tolist())
                st.success("KPI 지표가 동기화되었습니다.")

# ==========================================
# KPI 현황 탭
# ==========================================
with tab_kpi:
    st.header("📈 전사 통합 KPI" if u_role == "마스터" else f"📈 {u_name}님 전용 KPI 현황")
    
    stats = {}
    for d in all_daily:
        k_name = str(d.get('KPI', '기타')).strip()
        if u_role == "마스터" or k_name in my_kpi_opts:
            p = int(d.get('진행률', 0) if d.get('진행률') else 0)
            if k_name not in stats: stats[k_name] = {"sum": 0, "count": 0}
            stats[k_name]["sum"] += p
            stats[k_name]["count"] += 1
            
    if not stats:
        st.info("현재 표시할 KPI 실적이 없습니다.")
    else:
        for k_name, data in stats.items():
            avg = int(data["sum"] / data["count"]) if data["count"] > 0 else 0
            st.write(f"**{k_name}** (총 {data['count']}건의 업무)")
            st.progress(avg / 100, text=f"평균 달성률: {avg}%")

# ==========================================
# 보고서 및 데이터 탭
# ==========================================
with tab_rep:
    st.header("📊 데이터 및 보고서 관리")
    
    with st.expander("📥 과거 업무 일괄 업로드 (CSV)"):
        temp_df = pd.DataFrame(columns=["날짜", "업무명", "진행률", "프로젝트연동", "분류", "연결프로젝트", "KPI", "담당자"])
        st.download_button("📄 업로드 양식 받기", temp_df.to_csv(index=False).encode('utf-8-sig'), "template.csv")
        up_file = st.file_uploader("CSV 파일 선택", type="csv")
        if up_file and st.button("시트로 일괄 전송"):
            df_up = pd.read_csv(up_file).fillna("")
            if u_role != "마스터":
                df_up['담당자'] = u_name 
            new_rows = df_up.values.tolist()
            w_d.append_rows(new_rows)
            st.success("업로드 완료!")

    st.divider()

    st.subheader("🖨️ 맞춤형 보고서 출력")
    r_type = st.radio("보고서 종류 선택", ["일일(HTML)", "기간별(Excel)"], horizontal=True)
    
    if r_type == "일일(HTML)":
        r_d = st.date_input("보고 날짜", t_date, key="r_date_p")
        r_s = r_d.strftime("%Y-%m-%d")
        
        rep_daily = [d for d in all_daily if str(d.get('날짜')) == r_s and (u_role == "마스터" or d.get('담당자') == u_name)]
        rep_proj = [p for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name)]
        
        h_d_html = ""
        for t in rep_daily:
            st_t = "(완료)" if int(t.get('진행률',0))==100 else f"({t.get('진행률',0)}%)"
            name_t = str(t.get('업무명','')).replace('\n', '<br>')
            owner_txt = f" <span style='color:#1976D2; font-size:0.9em;'>[{t.get('담당자','')}]</span>" if u_role == "마스터" and t.get('담당자') else ""
            h_d_html += f"<li style='margin-bottom:8px;'><b>[{t.get('분류','기타')}]</b> {st_t} {name_t}{owner_txt}</li>"
            
        h_p_html = ""
        for p in rep_proj:
            if str(p.get("보관함이동", "FALSE")).upper() == "TRUE": continue
            pn = p.get('프로젝트명', '')
            st_txt = "(완료)" if str(p.get('완료여부','FALSE')).upper() == "TRUE" else ""
            owner_p = f" <span style='color:#555; font-size:0.9em;'>({p.get('담당자','')})</span>" if u_role == "마스터" and p.get('담당자') else ""
            h_p_html += f"<div style='margin-top:15px;'><h4 style='margin-bottom:5px;'>■ [{p.get('분류','기타')}] {pn}{owner_p} <span style='color:#2e7d32;'>{st_txt}</span></h4><ul style='margin-top:0;'>"
            
            for s in sub_dict.get(pn, []):
                prog = int(s.get('진행률',0))
                icon = "✓" if prog==100 else ("▶" if prog>0 else "□")
                s_name = str(s.get('세부업무명','')).replace('\n', '<br>')
                owner_s = f" <span style='color:#1976D2; font-size:0.9em;'>[{s.get('담당자','')}]</span>" if u_role == "마스터" and s.get('담당자') else ""
                h_p_html += f"<li style='margin-bottom:5px;'>{icon} {s_name} ({prog}%){owner_s}</li>"
            h_p_html += "</ul></div>"
            
        title_txt = "전사 업무 보고서" if u_role == "마스터" else f"{u_name} 업무 보고서"
        full_html = f"<html><body style='font-family:sans-serif;'><h2>[{r_s}] {title_txt}</h2><h3>■ 일일 업무</h3><ul style='line-height:1.5;'>{h_d_html}</ul><hr><h3>■ 프로젝트 현황</h3>{h_p_html}</body></html>"
        
        st.components.v1.html(full_html, height=300, scrolling=True)
        c_btn1, c_btn2 = st.columns([1, 3])
        with c_btn1:
            st.download_button("📥 HTML 다운로드", full_html.encode('utf-8'), f"Report_{r_s}.html")
        with st.expander("📋 HTML 코드 복사"):
            st.code(full_html, language="html")

    elif r_type == "기간별(Excel)":
        c_ds1, c_ds2 = st.columns(2)
        s_w = c_ds1.date_input("시작일", t_date - datetime.timedelta(days=7), key="ws")
        e_w = c_ds2.date_input("종료일", t_date, key="we")
        
        rep_proj = [p for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name)]
        
        xls_hr = ""
        for p in rep_proj:
            if str(p.get("보관함이동", "FALSE")).upper() == "TRUE": continue
            pn = p.get('프로젝트명', '')
            cat = p.get('분류', '기타')
            ph = f"<b>[{pn}]</b><br>"
            total_p = 0
            my_s = sub_dict.get(pn, [])
            
            for s in my_s:
                prog = int(s.get('진행률',0))
                total_p += prog
                s_n = str(s.get('세부업무명','')).replace('\n', '<br>')
                owner_s = f" [{s.get('담당자','')}]" if u_role == "마스터" and s.get('담당자') else ""
                ph += f"- {s_n} ({prog}%){owner_s}<br>"
                
            avg_p = int(total_p / len(my_s)) if my_s else 0
            xls_hr += f"<tr><td><b>{cat}</b></td><td>{ph}</td><td style='text-align:center;'><b>{avg_p}%</b></td><td></td></tr>"
            
        th = "<tr><th style='background:#e0f7fa; padding:8px;'>분류</th><th style='background:#e0f7fa; padding:8px;'>업무내역</th><th style='background:#e0f7fa; padding:8px;'>진행률</th><th style='background:#e0f7fa; padding:8px;'>예정사항</th></tr>"
        title_txt = "전사" if u_role == "마스터" else u_name
        xls_html = f"<html><meta charset='utf-8'><style>td {{border: 1px solid #ccc; padding: 8px; vertical-align: top; line-height:1.5;}}</style><body><h2>{title_txt} 업무 보고서 ({s_w} ~ {e_w})</h2><table style='border-collapse:collapse; width:100%; border: 1px solid #ccc;'>{th}{xls_hr}</table></body></html>"
        
        st.download_button("💾 Excel 다운로드 (.xls)", xls_html.encode('utf-8-sig'), f"Report_{s_w}_{e_w}.xls")
