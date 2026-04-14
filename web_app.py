import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑 V16.1", layout="wide")

# 2. 수파베이스 DB 연결
@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_connection()
except Exception as e:
    st.error("데이터베이스 연결 실패. Secrets 설정을 확인하세요.")
    st.stop()

# 💡 데이터 로드 (에러 발생 시 어떤 테이블인지 알려주도록 개선)
@st.cache_data(ttl=30)
def load_db_data():
    try:
        daily = supabase.table('daily').select("*").execute().data
        projects = supabase.table('projects').select("*").execute().data
        sub_tasks = supabase.table('sub_tasks').select("*").execute().data
        routines = supabase.table('routines').select("*").execute().data
        settings = supabase.table('settings').select("*").execute().data
        users = supabase.table('users').select("*").execute().data
        categories = supabase.table('categories').select("*").execute().data
        return daily, projects, sub_tasks, routines, settings, users, categories
    except Exception as e:
        st.warning(f"일부 데이터를 불러오지 못했습니다. DB 테이블 설정을 확인하세요: {e}")
        return [], [], [], [], [], [], []

def apply_changes():
    load_db_data.clear() 
    st.rerun()

# --- [로그인/보안] ---
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_info' not in st.session_state: st.session_state['user_info'] = {}

def check_login(user_id, user_pw):
    _, _, _, _, _, users, _ = load_db_data()
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

# --- [데이터 할당] ---
all_daily, proj_data, sub_data, routine_data, kpi_config, user_data, cat_data = load_db_data()

# 💡 DB categories 테이블에서 분류 목록 추출
cat_list = sorted(list(set([str(c.get('분류명', '')) for c in cat_data if pd.notna(c.get('분류명')) and str(c.get('분류명')).strip() != ""])))
if not cat_list: # DB에 아무것도 없을 때만 나오는 비상용 목록
    cat_list = ["경영관리", "재무업무", "입찰업무", "조달업무", "현장업무", "기타"]

my_kpi_opts = [str(k.get('KPI명', '')) for k in kpi_config if pd.notna(k.get('KPI명')) and str(k.get('구분', '공통')).strip() in ['공통', u_name] and str(k.get('KPI명')).strip() != ""]

# --- [사이드바] ---
with st.sidebar:
    st.subheader(f"👤 {u_name}님 ({u_role})")
    if st.button("🔄 최신 데이터 불러오기", use_container_width=True, type="primary"):
        apply_changes()
    st.divider()
    if st.button("🚪 로그아웃", use_container_width=True): 
        st.session_state['logged_in'] = False
        st.rerun()

st.title("🚀 NOWSYSTEM 통합 업무 관리")
t_date = st.date_input("📅 업무 기준일 선택", datetime.date.today(), key="main_date")
t_str = t_date.strftime("%Y-%m-%d")

if f"lock_{t_str}" not in st.session_state: st.session_state[f"lock_{t_str}"] = False
is_locked = st.session_state[f"lock_{t_str}"]

if u_role == "마스터":
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str]
    tabs = st.tabs(["📝 전사 일과 관리", "📁 프로젝트 관리", "⚙️ 설정1 (KPI/계정)", "⚙️ 설정2 (업무분류)", "📈 통합 KPI", "📊 데이터/보고서"])
    tab_set1 = tabs[2]; tab_set2 = tabs[3]; tab_kpi = tabs[4]; tab_rep = tabs[5]
else:
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str and d.get('담당자') == u_name]
    tabs = st.tabs(["📝 나의 일과", "📁 나의 프로젝트", "📈 나의 KPI", "📊 데이터/보고서"])
    tab_kpi = tabs[2]; tab_rep = tabs[3]

sub_dict = {p.get("프로젝트명", ""): [] for p in proj_data}
for s in sub_data:
    pn = s.get("프로젝트명", "")
    if pn in sub_dict: sub_dict[pn].append(s)

# ==========================================
# 탭 1: 일과 관리
# ==========================================
with tabs[0]:
    st.header(f"📝 {t_str} 업무 리스트")
    dropdown_opts = [k.get('KPI명', '') for k in kpi_config if k.get('KPI명')]
    
    with st.expander("➕ 오늘의 업무 추가", expanded=not is_locked):
        task_type = st.radio("업무 종류 선택", ["일반/데일리 업무", "프로젝트 연동 업무"], horizontal=True)
        
        if task_type == "일반/데일리 업무":
            my_routines = [r.get('업무명') for r in routine_data if r.get('담당자') == u_name]
            sel_opt = st.selectbox("업무명", ["✏️ 직접 입력"] + my_routines)
            n_task = st.text_input("새 업무명 (직접 입력)", disabled=is_locked) if sel_opt == "✏️ 직접 입력" else sel_opt
                
            c1, c2 = st.columns(2)
            n_cat = c1.selectbox("분류", cat_list, disabled=is_locked) 
            n_kpi = c2.selectbox("연관 KPI", dropdown_opts + ["기타"], disabled=is_locked)
            
            if st.button("업무 추가", disabled=is_locked, type="primary"):
                if n_task:
                    supabase.table('daily').insert({"날짜": t_str, "업무명": n_task, "진행률": 0, "프로젝트연동": "FALSE", "분류": n_cat, "연결프로젝트": "", "KPI": n_kpi, "담당자": u_name}).execute()
                    apply_changes()
                    
        else: # 프로젝트 연동
            my_projs = [p.get('프로젝트명') for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name) and str(p.get("보관함이동", "FALSE")).upper() != "TRUE"]
            sel_p = st.selectbox("진행 중인 프로젝트", my_projs) if my_projs else None
            
            if sel_p:
                my_subs = [s.get('세부업무명') for s in sub_data if s.get('프로젝트명') == sel_p]
                sel_s = st.selectbox("프로젝트 하위 세부업무", my_subs) if my_subs else None
                if st.button("프로젝트 업무 당겨오기", disabled=is_locked, type="primary"):
                    if sel_p and sel_s:
                        # 💡 [핵심 해결] 해당 프로젝트의 실제 '분류'와 'KPI'를 찾아옵니다.
                        p_info = next((p for p in proj_data if p.get('프로젝트명') == sel_p), {})
                        p_cat = p_info.get('분류', '프로젝트')
                        p_kpi = p_info.get('KPI', '기타')
                        
                        supabase.table('daily').insert({
                            "날짜": t_str, "업무명": sel_s, "진행률": 0, "프로젝트연동": "TRUE", 
                            "분류": p_cat, # "프로젝트"라는 글자 대신 실제 분류 사용
                            "연결프로젝트": f"{sel_p}::{sel_s}", "KPI": p_kpi, "담당자": u_name
                        }).execute()
                        apply_changes()

    st.divider()
    # ... (이하 동일 로직)
    for i, row in enumerate(filtered_daily):
        r_id = row.get('id', f"temp_d_{i}") 
        col1, col2, col3 = st.columns([4, 5, 1])
        disp_name = str(row.get('업무명','')).replace('\n', '<br>')
        badge = f" <small style='color:blue;'>[{row.get('담당자','')}]</small>" if u_role == "마스터" else ""
        is_proj = str(row.get('프로젝트연동', 'FALSE')).upper() == "TRUE"
        
        if is_proj:
            p_info = str(row.get('연결프로젝트', ''))
            # 프로젝트 연동 업무도 저장된 분류명을 표시함
            col1.markdown(f"**[{row.get('분류', '프로젝트')}]** <span style='color:#555;'>{p_info.replace('::', ' > ')}</span>{badge}", unsafe_allow_html=True)
        else:
            col1.markdown(f"**[{row.get('분류', '기타')}]** {disp_name}{badge}", unsafe_allow_html=True)
            
        cur_p = int(row.get('진행률', 0) if str(row.get('진행률', 0)).isdigit() else 0)
        new_p = col2.slider("진행", 0, 100, cur_p, 10, key=f"d_sld_{r_id}", disabled=is_locked, label_visibility="collapsed")
        
        if not is_locked and new_p != cur_p:
            if isinstance(r_id, int) or str(r_id).isdigit():
                supabase.table('daily').update({"진행률": new_p}).eq('id', r_id).execute()
                if is_proj:
                    p_info = str(row.get('연결프로젝트', ''))
                    if "::" in p_info:
                        p_n, s_n = p_info.split("::", 1)
                        for s in sub_data:
                            if s.get('프로젝트명') == p_n and s.get('세부업무명') == s_n:
                                s_id = s.get('id')
                                if s_id: supabase.table('sub_tasks').update({"진행률": new_p}).eq('id', s_id).execute()
                                break
                st.toast(f"✅ 진행률 {new_p}% 저장 완료!")
                apply_changes()
        
        if col3.button("🗑️", key=f"d_del_{r_id}", disabled=is_locked):
            if isinstance(r_id, int) or str(r_id).isdigit():
                supabase.table('daily').delete().eq('id', r_id).execute()
                apply_changes()

    # (이하 루틴, 프로젝트, 마스터 설정 탭 로직은 V16과 동일하게 유지)
# ... [중략] ...
