import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime
import io

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑 V22", layout="wide")

# 2. 수파베이스 DB 연결
@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_connection()
except Exception as e:
    st.error("데이터베이스 연결에 실패했습니다.")
    st.stop()

# 💡 초고속 데이터 로드
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
        st.warning(f"데이터 로드 실패: {e}")
        return [], [], [], [], [], [], []

def apply_changes():
    load_db_data.clear() 
    st.rerun()

# --- [보안, 로그인 및 세션 메모리] ---
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_info' not in st.session_state: st.session_state['user_info'] = {}
if 'active_proj_id' not in st.session_state: st.session_state['active_proj_id'] = None
if 'edit_d_id' not in st.session_state: st.session_state['edit_d_id'] = None
if 'edit_s_id' not in st.session_state: st.session_state['edit_s_id'] = None

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

# --- [데이터 할당 및 정렬] ---
all_daily, proj_data, sub_data, routine_data, kpi_config, user_data, cat_data = load_db_data()

proj_data = sorted(proj_data, key=lambda x: (int(x.get('정렬순서', 999)), int(x.get('id', 0))))
sub_data = sorted(sub_data, key=lambda x: int(x.get('id', 0)))

cat_list = sorted(list(set([str(c.get('분류명', '')) for c in cat_data if pd.notna(c.get('분류명')) and str(c.get('분류명')).strip() != ""])))
if not cat_list: cat_list = ["경영관리", "재무업무", "기타"]

st.title("🚀 NOWSYSTEM 통합 업무 관리")
t_date = st.date_input("📅 업무 기준일 선택", datetime.date.today(), key="main_date")
t_str = t_date.strftime("%Y-%m-%d")

if f"lock_{t_str}" not in st.session_state: st.session_state[f"lock_{t_str}"] = False
is_locked = st.session_state[f"lock_{t_str}"]

view_target = "전체"
target_user = u_name

# --- [사이드바] ---
with st.sidebar:
    st.subheader(f"👤 {u_name}님 ({u_role})")
    if st.button("🔄 최신 데이터 불러오기", use_container_width=True, type="primary"): apply_changes()
    st.divider()
    
    if is_locked:
        st.success("🔒 현재 업무가 마감되었습니다.")
        if st.button("🔓 일과 마감 취소", use_container_width=True):
            st.session_state[f"lock_{t_str}"] = False
            st.rerun()
    else:
        if st.button("🔒 일과 마감하기", use_container_width=True, type="primary"):
            st.session_state[f"lock_{t_str}"] = True
            st.rerun()
            
    if u_role == "마스터":
        st.divider()
        st.markdown("**👀 직원 모니터링**")
        user_names = sorted(list(set([u.get('이름') for u in user_data if u.get('이름')])))
        view_target = st.selectbox("업무를 확인할 직원 선택", ["전체"] + user_names)
        if view_target != "전체":
            target_user = view_target

    st.divider()
    if st.button("🚪 로그아웃", use_container_width=True): 
        st.session_state['logged_in'] = False
        st.rerun()

my_kpi_opts = [str(k.get('KPI명', '')) for k in kpi_config if pd.notna(k.get('KPI명')) and str(k.get('구분', '공통')).strip() in ['공통', target_user] and str(k.get('KPI명')).strip() != ""]

if u_role == "마스터":
    if view_target == "전체":
        filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str]
        tabs = st.tabs(["📝 전사 일과 관리", "📁 전사 프로젝트", "⚙️ 설정1 (KPI/계정)", "⚙️ 설정2 (업무분류)", "📈 전사 통합 KPI", "📊 데이터/보고서"])
    else:
        filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str and d.get('담당자') == target_user]
        tabs = st.tabs([f"📝 {target_user} 일과", f"📁 {target_user} 프로젝트", "⚙️ 설정1 (KPI/계정)", "⚙️ 설정2 (업무분류)", f"📈 {target_user} KPI", "📊 데이터/보고서"])
    tab_set1 = tabs[2]; tab_set2 = tabs[3]; tab_kpi = tabs[4]; tab_rep = tabs[5]
else:
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str and d.get('담당자') == target_user]
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
    header_title = f"📝 {t_str} {target_user} 업무 리스트" if view_target != "전체" else f"📝 {t_str} 전사 업무 리스트"
    st.header(header_title)
    dropdown_opts = [k.get('KPI명', '') for k in kpi_config if k.get('KPI명')]
    
    with st.expander("➕ 오늘의 업무 추가", expanded=not is_locked):
        task_type = st.radio("업무 종류 선택", ["일반/데일리 업무", "프로젝트 연동 업무"], horizontal=True, disabled=is_locked)
        
        if task_type == "일반/데일리 업무":
            with st.form("add_daily_normal_form", clear_on_submit=True):
                my_routines = [r.get('업무명') for r in routine_data if r.get('담당자') == target_user]
                sel_opt = st.selectbox("업무명 (루틴에서 선택)", ["✏️ 직접 입력"] + my_routines, disabled=is_locked)
                n_task = st.text_area("새 업무명 (직접 입력 시: Alt+Enter 줄바꿈, Ctrl+Enter 저장)", height=100, disabled=is_locked) 
                    
                c1, c2 = st.columns(2)
                n_cat = c1.selectbox("분류", cat_list, disabled=is_locked) 
                n_kpi = c2.selectbox("연관 KPI", my_kpi_opts + ["기타"], disabled=is_locked)
                
                if st.form_submit_button("업무 추가", type="primary", disabled=is_locked):
                    final_task = n_task if sel_opt == "✏️ 직접 입력" else sel_opt
                    if final_task:
                        supabase.table('daily').insert({"날짜": t_str, "업무명": final_task, "진행률": 0, "프로젝트연동": "FALSE", "분류": n_cat, "연결프로젝트": "", "KPI": n_kpi, "담당자": target_user, "보고서제외": False, "진행중": False}).execute()
                        apply_changes()
        else: 
            if u_role == "마스터" and view_target == "전체":
                my_projs = [p.get('프로젝트명') for p in proj_data if not (str(p.get('보관함이동')).upper() == "TRUE" or p.get('보관함이동') == True)]
            else:
                my_projs = [p.get('프로젝트명') for p in proj_data if p.get('담당자') == target_user and not (str(p.get('보관함이동')).upper() == "TRUE" or p.get('보관함이동') == True)]
            
            sel_p = st.selectbox("진행 중인 프로젝트", my_projs, disabled=is_locked) if my_projs else None
            my_subs = [s.get('세부업무명') for s in sub_data if s.get('프로젝트명') == sel_p] if sel_p else []
            sel_s = st.selectbox("프로젝트 하위 세부업무", my_subs, disabled=is_locked) if my_subs else None
            
            if st.button("프로젝트 업무 당겨오기", type="primary", disabled=is_locked):
                if sel_p and sel_s:
                    p_info = next((p for p in proj_data if p.get('프로젝트명') == sel_p), {})
                    p_owner = p_info.get('담당자', target_user) if view_target == "전체" else target_user
                    supabase.table('daily').insert({"날짜": t_str, "업무명": sel_s, "진행률": 0, "프로젝트연동": "TRUE", "분류": p_info.get('분류', '프로젝트'), "연결프로젝트": f"{sel_p}::{sel_s}", "KPI": p_info.get('KPI', '기타'), "담당자": p_owner, "보고서제외": False, "진행중": False}).execute()
                    apply_changes()

    st.divider()
    
    for i, row in enumerate(filtered_daily):
        r_id = row.get('id', f"temp_d_{i}") 
        
        if st.session_state.get('edit_d_id') == r_id:
            with st.container(border=True):
                st.write("🛠️ **업무 직접 수정**")
                e_name = st.text_area("업무명 수정", row.get('업무명',''), height=80)
                ec1, ec2 = st.columns(2)
                e_cat_idx = cat_list.index(row.get('분류')) if row.get('분류') in cat_list else 0
                e_kpi_idx = (my_kpi_opts + ["기타"]).index(row.get('KPI')) if row.get('KPI') in (my_kpi_opts + ["기타"]) else 0
                e_cat = ec1.selectbox("분류 변경", cat_list, index=e_cat_idx, key=f"ec_{r_id}")
                e_kpi = ec2.selectbox("KPI 변경", my_kpi_opts + ["기타"], index=e_kpi_idx, key=f"ek_{r_id}")
                
                eb1, eb2, _ = st.columns([1, 1, 4])
                if eb1.button("저장", type="primary", key=f"esv_{r_id}"):
                    supabase.table('daily').update({"업무명": e_name, "분류": e_cat, "KPI": e_kpi}).eq('id', r_id).execute()
                    st.session_state['edit_d_id'] = None
                    apply_changes()
                if eb2.button("취소", key=f"ecan_{r_id}"):
                    st.session_state['edit_d_id'] = None
                    st.rerun()
            continue 
            
        col1, col2, col3, col4, col5 = st.columns([4, 3, 1, 0.7, 0.7])
        disp_name = str(row.get('업무명','')).replace('\n', '<br>')
        badge = f" <small style='color:blue;'>[{row.get('담당자','')}]</small>" if u_role == "마스터" and view_target == "전체" else ""
        is_proj = str(row.get('프로젝트연동', 'FALSE')).upper() == "TRUE"
        kpi_txt = f" <span style='color:#007BFF; font-size:0.85em;'>(KPI: {row.get('KPI', '미지정')})</span>"
        
        if is_proj: col1.markdown(f"**[{row.get('분류', '프로젝트')}]** <span style='color:#555;'>{str(row.get('연결프로젝트', '')).replace('::', ' > ')}</span>{badge}", unsafe_allow_html=True)
        else: col1.markdown(f"**[{row.get('분류', '기타')}]** {disp_name}{kpi_txt}{badge}", unsafe_allow_html=True)
            
        cur_p = int(row.get('진행률', 0) if str(row.get('진행률', 0)).isdigit() else 0)
        new_p = col2.slider("진행", 0, 100, cur_p, 10, key=f"d_sld_{r_id}", label_visibility="collapsed", disabled=is_locked)
        
        if not is_locked and new_p != cur_p:
            if isinstance(r_id, int) or str(r_id).isdigit():
                supabase.table('daily').update({"진행률": new_p}).eq('id', r_id).execute()
                if is_proj:
                    p_info = str(row.get('연결프로젝트', ''))
                    if "::" in p_info:
                        p_n, s_n = p_info.split("::", 1)
                        s_id_match = next((s.get('id') for s in sub_data if s.get('프로젝트명') == p_n and s.get('세부업무명') == s_n), None)
                        if s_id_match: supabase.table('sub_tasks').update({"진행률": new_p}).eq('id', s_id_match).execute()
                apply_changes()
        
        is_ex = bool(row.get('보고서제외', False))
        new_ex = col3.checkbox("🚫제외", value=is_ex, key=f"d_ex_{r_id}", help="체크 시 보고서에 출력 안됨", disabled=is_locked)
        if not is_locked and new_ex != is_ex:
            supabase.table('daily').update({"보고서제외": new_ex}).eq('id', r_id).execute()
            apply_changes()

        if col4.button("✏️", key=f"d_edt_{r_id}", disabled=is_locked):
            st.session_state['edit_d_id'] = r_id
            st.rerun()

        if col5.button("🗑️", key=f"d_del_{r_id}", disabled=is_locked):
            if isinstance(r_id, int) or str(r_id).isdigit():
                supabase.table('daily').delete().eq('id', r_id).execute()
                apply_changes()

    st.write("---")
    st.subheader(f"📌 {target_user} 데일리 고정 업무 (루틴)" if view_target != "전체" else "📌 전사 데일리 고정 업무 (루틴)")
    c_r1, c_r2 = st.columns([1, 1])
    with c_r1:
        with st.form("add_routine_form", clear_on_submit=True):
            r_task = st.text_input("새 데일리 업무명 등록", disabled=is_locked)
            r_cat = st.selectbox("분류", cat_list, disabled=is_locked) 
            r_kpi = st.selectbox("연관 KPI", my_kpi_opts + ["기타"], disabled=is_locked)
            if st.form_submit_button("루틴 목록에 추가", disabled=is_locked):
                if r_task:
                    supabase.table('routines').insert({"업무명": r_task, "분류": r_cat, "KPI": r_kpi, "담당자": target_user}).execute()
                    apply_changes()
    with c_r2:
        for i, r in enumerate(routine_data):
            if (u_role == "마스터" and view_target == "전체") or r.get('담당자') == target_user:
                r_id = r.get('id', f"temp_r_{i}")
                rr1, rr2 = st.columns([4, 1])
                badge_r = f" [{r.get('담당자','')}]" if u_role == "마스터" and view_target == "전체" else ""
                rr1.write(f"· [{r.get('분류')}] {r.get('업무명')}{badge_r}")
                if rr2.button("삭제", key=f"rdel_{r_id}", disabled=is_locked):
                    if isinstance(r_id, int) or str(r_id).isdigit():
                        supabase.table('routines').delete().eq('id', r_id).execute()
                        apply_changes()

# ==========================================
# 탭 2: 프로젝트 관리 
# ==========================================
with tabs[1]:
    st.header("📁 프로젝트 현황")
    with st.expander("✨ 신규 프로젝트 등록", expanded=False):
        with st.form("new_proj_form", clear_on_submit=True):
            pc1, pc2 = st.columns(2)
            p_name = pc1.text_input("프로젝트명", disabled=is_locked)
            p_cat = pc2.selectbox("분류", cat_list, disabled=is_locked) 
            p_start = pc1.date_input("시작일", disabled=is_locked)
            p_kpi = pc2.selectbox("연관 KPI", my_kpi_opts + ["기타"], disabled=is_locked)
            if st.form_submit_button("프로젝트 저장", type="primary", disabled=is_locked):
                if p_name:
                    supabase.table('projects').insert({"프로젝트명": p_name, "시작일": str(p_start), "분류": p_cat, "KPI": p_kpi, "담당자": target_user, "정렬순서": 999, "보고서제외": False}).execute()
                    st.success(f"[{p_name}] 저장 완료!")
                    apply_changes()

    for i, p in enumerate(proj_data):
        r_id = p.get('id', f"temp_p_{i}")
        is_archived_bool = True if str(p.get("보관함이동")).upper() == "TRUE" or p.get("보관함이동") == True else False
        
        if u_role == "마스터" and view_target != "전체":
            if p.get('담당자') != target_user: continue
        elif u_role != "마스터":
            if p.get('담당자') != target_user: continue
            
        if is_archived_bool: continue
            
        pn = p.get("프로젝트명", "")
        owner = f" ({p.get('담당자','')})" if u_role == "마스터" and view_target == "전체" else ""
        my_s_list = sub_dict.get(pn, [])
        
        total_p = sum(int(s.get('진행률', 0) if str(s.get('진행률',0)).isdigit() else 0) for s in my_s_list)
        avg_p = int(total_p / len(my_s_list)) if len(my_s_list) > 0 else 0
        
        is_expanded = (st.session_state.get('active_proj_id') == r_id)
        
        with st.expander(f"📂 {pn} [{p.get('분류')}] (KPI: {p.get('KPI', '미지정')}) - 📊 전체 진행률: {avg_p}% {owner}", expanded=is_expanded):
            set_c1, set_c2, set_c3 = st.columns([1.5, 1.5, 2])
            cur_ord = int(p.get('정렬순서', 999) if pd.notna(p.get('정렬순서')) else 999)
            new_ord = set_c1.number_input("🔢 순서", value=cur_ord, key=f"pord_{r_id}", disabled=is_locked)
            if not is_locked and new_ord != cur_ord:
                supabase.table('projects').update({"정렬순서": new_ord}).eq('id', r_id).execute()
                st.session_state['active_proj_id'] = r_id 
                apply_changes()
                
            p_ex = bool(p.get('보고서제외', False))
            new_p_ex = set_c2.checkbox("🚫 전체 제외", value=p_ex, key=f"pex_{r_id}", disabled=is_locked)
            if not is_locked and new_p_ex != p_ex:
                supabase.table('projects').update({"보고서제외": new_p_ex}).eq('id', r_id).execute()
                st.session_state['active_proj_id'] = r_id 
                apply_changes()

            with set_c3:
                with st.form(key=f"ren_p_{r_id}"):
                    ren_p = st.text_input("🛠️ 프로젝트명 수정", value=pn, label_visibility="collapsed", disabled=is_locked)
                    if st.form_submit_button("이름 변경 적용", disabled=is_locked):
                        if ren_p and ren_p != pn:
                            supabase.table('projects').update({"프로젝트명": ren_p}).eq('id', r_id).execute()
                            supabase.table('sub_tasks').update({"프로젝트명": ren_p}).eq('프로젝트명', pn).execute()
                            for d in all_daily:
                                p_link = str(d.get('연결프로젝트', ''))
                                if p_link.startswith(pn + "::"):
                                    new_link = p_link.replace(pn + "::", ren_p + "::", 1)
                                    supabase.table('daily').update({"연결프로젝트": new_link}).eq('id', d.get('id')).execute()
                            st.session_state['active_proj_id'] = r_id
                            st.success("수정 완료!")
                            apply_changes()

            st.write("---")

            with st.form(key=f"sub_form_{r_id}", clear_on_submit=True):
                sc1, sc2 = st.columns([4,1])
                new_sub = sc1.text_area("세부 업무명 (Alt+Enter: 줄바꿈 / Ctrl+Enter: 즉시 추가)", height=80, disabled=is_locked)
                if sc2.form_submit_button("하위 업무 추가", disabled=is_locked):
                    if new_sub:
                        supabase.table('sub_tasks').insert({"프로젝트명": pn, "세부업무명": new_sub, "진행률": 0, "담당자": target_user, "보고서제외": False, "진행중": False}).execute()
                        st.session_state['active_proj_id'] = r_id 
                        apply_changes()
                        
            for j, s in enumerate(my_s_list):
                s_id = s.get('id', f"temp_s_{j}")
                
                if st.session_state.get('edit_s_id') == s_id:
                    with st.container(border=True):
                        st.write("🛠️ **하위 업무 수정**")
                        e_s_name = st.text_area("세부업무명 수정", s.get('세부업무명',''), height=80)
                        eb1, eb2, _ = st.columns([1, 1, 4])
                        if eb1.button("저장", type="primary", key=f"esv_s_{s_id}"):
                            supabase.table('sub_tasks').update({"세부업무명": e_s_name}).eq('id', s_id).execute()
                            old_s_name = s.get('세부업무명')
                            for d in all_daily:
                                if d.get('연결프로젝트') == f"{pn}::{old_s_name}":
                                    supabase.table('daily').update({"연결프로젝트": f"{pn}::{e_s_name}", "업무명": e_s_name}).eq('id', d.get('id')).execute()
                            st.session_state['edit_s_id'] = None
                            st.session_state['active_proj_id'] = r_id
                            apply_changes()
                        if eb2.button("취소", key=f"ecan_s_{s_id}"):
                            st.session_state['edit_s_id'] = None
                            st.session_state['active_proj_id'] = r_id
                            st.rerun()
                    continue
                
                sl1, sl2, sl3, sl4, sl5, sl6, sl7 = st.columns([3.5, 2.5, 1, 1.2, 1, 0.6, 0.6])
                sl1.markdown(f"· {str(s.get('세부업무명','')).replace('\n','<br>')}", unsafe_allow_html=True)
                
                cur_sp = int(s.get('진행률',0) if str(s.get('진행률',0)).isdigit() else 0)
                sp = sl2.slider("진행", 0, 100, cur_sp, 10, key=f"s_sld_{s_id}", label_visibility="collapsed", disabled=is_locked)
                if not is_locked and sp != cur_sp:
                    if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').update({"진행률": sp}).eq('id', s_id).execute()
                        st.session_state['active_proj_id'] = r_id 
                        apply_changes()
                
                if sl3.button("✅완료", key=f"sdone_{s_id}", disabled=is_locked):
                     if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').update({"진행률": 100}).eq('id', s_id).execute()
                        st.session_state['active_proj_id'] = r_id 
                        apply_changes()

                s_prog = bool(s.get('진행중', False))
                s_new_prog = sl4.checkbox("▶️진행중", value=s_prog, key=f"s_prg_{s_id}", disabled=is_locked)
                if not is_locked and s_new_prog != s_prog:
                    supabase.table('sub_tasks').update({"진행중": s_new_prog}).eq('id', s_id).execute()
                    st.session_state['active_proj_id'] = r_id 
                    apply_changes()

                s_ex = bool(s.get('보고서제외', False))
                s_new_ex = sl5.checkbox("🚫제외", value=s_ex, key=f"s_ex_{s_id}", disabled=is_locked)
                if not is_locked and s_new_ex != s_ex:
                    supabase.table('sub_tasks').update({"보고서제외": s_new_ex}).eq('id', s_id).execute()
                    st.session_state['active_proj_id'] = r_id 
                    apply_changes()
                
                if sl6.button("✏️", key=f"sedt_{s_id}", disabled=is_locked):
                    st.session_state['edit_s_id'] = s_id
                    st.session_state['active_proj_id'] = r_id
                    st.rerun()

                if sl7.button("🗑️", key=f"sdel_{s_id}", disabled=is_locked):
                    if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').delete().eq('id', s_id).execute()
                        st.session_state['active_proj_id'] = r_id 
                        apply_changes()
            
            st.write("---")
            ac1, ac2, ac3 = st.columns([2,1,1])
            if ac2.button("📦 보관함 이동", key=f"arc_{r_id}", disabled=is_locked):
                if isinstance(r_id, int) or str(r_id).isdigit():
                    supabase.table('projects').update({"보관함이동": True}).eq('id', r_id).execute()
                    st.session_state['active_proj_id'] = None 
                    apply_changes()
            if ac3.button("🗑️ 프로젝트 삭제", key=f"pdel_{r_id}", disabled=is_locked):
                if isinstance(r_id, int) or str(r_id).isdigit():
                    supabase.table('projects').delete().eq('id', r_id).execute()
                    st.session_state['active_proj_id'] = None 
                    apply_changes()

# ==========================================
# 탭 3 & 4: 설정 탭 (마스터)
# ==========================================
if u_role == "마스터":
    with tab_set1:
        st.header("⚙️ 설정 1 (사내 계정 및 KPI 관리)")
        c1, c2 = st.columns(2)
        with c1:
            u_df = pd.DataFrame(user_data)
            e_u_df = st.data_editor(u_df, num_rows="dynamic", use_container_width=True)
            if st.button("계정 정보 저장"):
                orig_u_ids = set(u_df['id'].dropna()) if 'id' in u_df.columns else set()
                new_u_ids = set(e_u_df['id'].dropna()) if 'id' in e_u_df.columns else set()
                for did in orig_u_ids - new_u_ids: supabase.table('users').delete().eq('id', did).execute()
                for r in e_u_df.to_dict('records'):
                    if pd.isna(r.get('id')): r.pop('id', None) 
                    supabase.table('users').upsert(r).execute()
                apply_changes()
        with c2:
            k_df = pd.DataFrame(kpi_config)
            if 'KPI명' not in k_df.columns: k_df['KPI명'] = ""
            if '구분' not in k_df.columns: k_df['구분'] = '공통'
            k_df = k_df[[col for col in k_df.columns if col != '분류명']]
            e_k_df = st.data_editor(k_df, num_rows="dynamic", use_container_width=True)
            if st.button("KPI 지표 저장"):
                orig_k_ids = set(k_df['id'].dropna()) if 'id' in k_df.columns else set()
                new_k_ids = set(e_k_df['id'].dropna()) if 'id' in e_k_df.columns else set()
                for did in orig_k_ids - new_k_ids: supabase.table('settings').delete().eq('id', did).execute()
                for r in e_k_df.to_dict('records'):
                    if pd.isna(r.get('id')): r.pop('id', None)
                    if r.get('KPI명'): supabase.table('settings').upsert(r).execute()
                apply_changes()

    with tab_set2:
        st.header("⚙️ 설정 2 (업무 분류 전용 관리)")
        c_df = pd.DataFrame(cat_data)
        if '분류명' not in c_df.columns: c_df['분류명'] = ""
        e_c_df = st.data_editor(c_df, num_rows="dynamic", use_container_width=False, width=600)
        if st.button("업무 분류(카테고리) 목록 저장", type="primary"):
            orig_c_ids = set(c_df['id'].dropna()) if 'id' in c_df.columns else set()
            new_c_ids = set(e_c_df['id'].dropna()) if 'id' in e_c_df.columns else set()
            for did in orig_c_ids - new_c_ids: supabase.table('categories').delete().eq('id', did).execute()
            for r in e_c_df.to_dict('records'):
                if pd.isna(r.get('id')): r.pop('id', None)
                if r.get('분류명'): supabase.table('categories').upsert(r).execute()
            apply_changes()

# ==========================================
# KPI 현황 탭 
# ==========================================
with tab_kpi:
    st.header(f"📈 {target_user} KPI 현황" if view_target != "전체" else "📈 전사 통합 KPI")
    stats = {}
    for d in all_daily:
        if view_target != "전체" and d.get('담당자') != target_user: continue
        k_name = str(d.get('KPI', '기타')).strip()
        if u_role == "마스터" or k_name in my_kpi_opts:
            p = int(d.get('진행률', 0) if str(d.get('진행률', 0)).isdigit() else 0)
            if k_name not in stats: stats[k_name] = {"sum": 0, "count": 0}
            stats[k_name]["sum"] += p
            stats[k_name]["count"] += 1
    if stats:
        for k_name, data in stats.items():
            avg = int(data["sum"] / data["count"]) if data["count"] > 0 else 0
            st.write(f"**{k_name}** (총 {data['count']}건)")
            st.progress(avg / 100, text=f"평균 달성률: {avg}%")
    else: st.info("데이터가 없습니다.")

# ==========================================
# 탭 5: 데이터/보고서 (💡 보고서 렌더링 V22 핵심 수정)
# ==========================================
with tab_rep:
    st.header("📊 데이터 및 보고서 관리")
    
    with st.expander("🛠️ 등록된 전체 업무 일괄 수정 (오타/내용 변경)"):
        st.info("표 안의 글씨를 더블클릭하여 바로 수정하고, 맨 아래 [일괄 수정 저장] 버튼을 누르세요.")
        t1, t2, t3 = st.tabs(["📝 일일 업무", "📁 프로젝트", "📋 하위 세부업무"])
        with t1: e_d_df = st.data_editor(pd.DataFrame(all_daily), key="ed_d", use_container_width=True)
        with t2: e_p_df = st.data_editor(pd.DataFrame(proj_data), key="ed_p", use_container_width=True)
        with t3: e_s_df = st.data_editor(pd.DataFrame(sub_data), key="ed_s", use_container_width=True)
            
        if st.button("💾 전체 데이터 일괄 수정 저장", type="primary", disabled=is_locked):
            for r in e_d_df.to_dict('records'): supabase.table('daily').upsert(r).execute()
            for r in e_p_df.to_dict('records'): supabase.table('projects').upsert(r).execute()
            for r in e_s_df.to_dict('records'): supabase.table('sub_tasks').upsert(r).execute()
            st.success("모든 수정 사항이 DB에 저장되었습니다!")
            apply_changes()

    st.divider()
    st.subheader("📥 과거 업무 일괄 업로드 (Excel)")
    
    temp_df = pd.DataFrame(columns=["날짜", "업무명", "진행률", "프로젝트연동", "분류", "연결프로젝트", "KPI", "담당자", "보고서제외", "진행중"])
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        temp_df.to_excel(writer, index=False)
    
    c_up1, c_up2 = st.columns([1, 2])
    with c_up1:
        st.write("1. 엑셀 양식을 다운로드하여 작성합니다.")
        st.download_button("📄 업로드 양식(Excel) 받기", data=excel_buffer.getvalue(), file_name="template.xlsx", mime="application/vnd.openxmlformats-officedomedocument.spreadsheetml.sheet")
    
    with c_up2:
        st.write("2. 작성된 엑셀 파일을 업로드합니다.")
        up_file = st.file_uploader("Excel 파일 선택 (.xlsx)", type=["xlsx"])
        if up_file and st.button("시트로 일괄 전송 (Excel)"):
            try:
                df_up = pd.read_excel(up_file).fillna("")
                if u_role != "마스터": df_up['담당자'] = u_name 
                new_rows = df_up.to_dict('records')
                for row in new_rows: supabase.table('daily').insert(row).execute()
                st.success("엑셀 데이터 업로드 완료!")
                apply_changes()
            except Exception as e:
                st.error(f"업로드 중 에러가 발생했습니다: {e}")
                
    st.divider()
    
    st.subheader("🖨️ 맞춤형 보고서 출력")
    r_type = st.radio("보고서 종류 선택", ["일일(HTML)", "기간별(Excel)"], horizontal=True)
    
    if r_type == "일일(HTML)":
        r_d = st.date_input("보고 날짜", t_date, key="r_date_p")
        r_s = r_d.strftime("%Y-%m-%d")
        
        if u_role == "마스터" and view_target == "전체":
            rep_daily = [d for d in all_daily if str(d.get('날짜')) == r_s and not bool(d.get('보고서제외', False))]
            rep_proj = [p for p in proj_data]
            rep_routines = routine_data
            title_txt = "전사 업무 보고서"
        else:
            rep_daily = [d for d in all_daily if str(d.get('날짜')) == r_s and d.get('담당자') == target_user and not bool(d.get('보고서제외', False))]
            rep_proj = [p for p in proj_data if p.get('담당자') == target_user]
            rep_routines = [r for r in routine_data if r.get('담당자') == target_user]
            title_txt = f"{target_user} 업무 보고서"
            
        h_d_html = ""
        grouped_proj = {}
        
        # 💡 [요청 적용] 일일 업무 렌더링 (분류 제거, 아이콘 적용, 프로젝트 참조 문구)
        for t in rep_daily:
            is_proj = str(t.get('프로젝트연동', 'FALSE')).upper() == "TRUE"
            prog = int(t.get('진행률', 0) if str(t.get('진행률',0)).isdigit() else 0)
            is_in_prog = bool(t.get('진행중', False))
            
            if prog == 100:
                icon = "✓"
                prog_txt = "(완료)"
            elif prog == 0:
                icon = "□"
                prog_txt = f"({prog}%)"
                if is_in_prog:
                    icon = "▶"
                    prog_txt += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"
            else:
                icon = "▶"
                prog_txt = f"({prog}%)"
                if is_in_prog:
                    prog_txt += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"
            
            task_name = str(t.get('업무명','')).replace(chr(10), '<br>')
            
            if is_proj:
                p_info = str(t.get('연결프로젝트', ''))
                p_name = p_info.split('::')[0] if '::' in p_info else p_info
                
                # 💡 [요청 1번] (아래 프로젝트명 참조) 문구 삽입
                task_str = f"{task_name} - <b>{p_name}</b> <span style='color:#777; font-size:0.85em;'>(아래 {p_name} 참조)</span>"
                
                if p_name not in grouped_proj:
                    grouped_proj[p_name] = {'tasks': []}
                grouped_proj[p_name]['tasks'].append(f"{icon} {task_str} {prog_txt}")
            else:
                # 💡 [요청 2, 3번] 분류([경영관리] 등) 완전히 제거
                h_d_html += f"<li style='margin-bottom:8px;'>{icon} {task_name} {prog_txt}</li>"

        for p_name, data in grouped_proj.items():
            h_d_html += f"<li style='margin-bottom:8px;'><b>{p_name}</b><ul style='margin-top:4px; margin-bottom:0;'>"
            for sub_t in data['tasks']:
                h_d_html += f"<li style='margin-bottom:4px; list-style-type: none;'>{sub_t}</li>"
            h_d_html += "</ul></li>"
            
        # 💡 [요청 4번] 루틴 업무는 항상 ✓ 아이콘과 (완료) 표기, 분류 제거
        h_r_html = "".join([f"<li style='margin-bottom:8px;'>✓ {str(r.get('업무명','')).replace(chr(10), '<br>')} (완료)</li>" for r in rep_routines])
        
        # 💡 프로젝트 현황 렌더링
        h_p_html = ""
        for p in rep_proj:
            if bool(p.get('보관함이동', False)) or str(p.get('보관함이동')).upper() == "TRUE" or bool(p.get('보고서제외', False)): continue
            pn = p.get('프로젝트명', '')
            st_txt = "(완료)" if str(p.get('완료여부')).upper() == "TRUE" else ""
            
            valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue 
            
            # 💡 [요청 2번] 분류([경영관리] 등) 완전히 제거
            h_p_html += f"<div style='margin-top:15px;'><h4 style='margin-bottom:5px;'>■ {pn} <span style='color:#2e7d32;'>{st_txt}</span></h4><ul style='margin-top:0;'>"
            for s in valid_subs:
                prog = int(s.get('진행률',0))
                is_in_prog = bool(s.get('진행중', False))
                
                # 💡 [요청 3번] 아이콘 로직 정밀 적용
                if prog == 100:
                    icon = "✓"
                    prog_txt = "(완료)"
                elif prog == 0:
                    icon = "□"
                    prog_txt = f"({prog}%)"
                    if is_in_prog:
                        icon = "▶"
                        prog_txt += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"
                else:
                    icon = "▶"
                    prog_txt = f"({prog}%)"
                    if is_in_prog:
                        prog_txt += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"

                h_p_html += f"<li style='margin-bottom:5px;'>{icon} {str(s.get('세부업무명','')).replace(chr(10), '<br>')} {prog_txt}</li>"
            h_p_html += "</ul></div>"
            
        full_html = f"<html><body style='font-family:sans-serif;'><h2>[{r_s}] {title_txt}</h2><h3>■ 일일 업무</h3><ul style='line-height:1.5;'>{h_d_html}</ul><hr><h3>■ 고정 업무 (루틴)</h3><ul style='line-height:1.5;'>{h_r_html}</ul><hr><h3>■ 프로젝트 현황</h3>{h_p_html}</body></html>"
        
        st.components.v1.html(full_html, height=400, scrolling=True)
        
        c_btn1, c_btn2 = st.columns([1, 1])
        with c_btn1: 
            st.download_button("📥 HTML 다운로드", full_html.encode('utf-8'), f"Report_{r_s}.html", use_container_width=True)
        with c_btn2: 
            with st.expander("📋 HTML 코드 복사"): 
                st.code(full_html, language="html")

    elif r_type == "기간별(Excel)":
        c_ds1, c_ds2 = st.columns(2)
        s_w = c_ds1.date_input("시작일", t_date - datetime.timedelta(days=7), key="ws")
        e_w = c_ds2.date_input("종료일", t_date, key="we")
        
        if u_role == "마스터" and view_target == "전체":
            rep_proj = [p for p in proj_data]
            title_txt = "전사"
        else:
            rep_proj = [p for p in proj_data if p.get('담당자') == target_user]
            title_txt = target_user
            
        xls_hr = ""
        for p in rep_proj:
            if bool(p.get('보관함이동', False)) or str(p.get('보관함이동')).upper() == "TRUE" or bool(p.get('보고서제외', False)): continue
            pn = p.get('프로젝트명', '')
            ph = f"<b>{pn}</b><br>"
            
            valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue
            
            total_p = 0
            for s in valid_subs:
                prog = int(s.get('진행률',0))
                total_p += prog
                ph += f"- {str(s.get('세부업무명','')).replace(chr(10), '<br>')} ({prog}%)<br>"
            avg_p = int(total_p / len(valid_subs)) if valid_subs else 0
            xls_hr += f"<tr><td>{ph}</td><td style='text-align:center;'><b>{avg_p}%</b></td><td></td></tr>"
            
        th = "<tr><th style='background:#e0f7fa; padding:8px;'>업무내역</th><th style='background:#e0f7fa; padding:8px;'>진행률</th><th style='background:#e0f7fa; padding:8px;'>예정사항</th></tr>"
        xls_html = f"<html><meta charset='utf-8'><style>td {{border: 1px solid #ccc; padding: 8px; vertical-align: top; line-height:1.5;}}</style><body><h2>{title_txt} 업무 보고서 ({s_w} ~ {e_w})</h2><table style='border-collapse:collapse; width:100%; border: 1px solid #ccc;'>{th}{xls_hr}</table></body></html>"
        st.download_button("💾 Excel 다운로드 (.xls)", xls_html.encode('utf-8-sig'), f"Report_{s_w}_{e_w}.xls")
