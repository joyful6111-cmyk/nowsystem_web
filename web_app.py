import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime
import io
import time
from streamlit_cookies_controller import CookieController

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑 V38", layout="wide")

# 쿠키 컨트롤러
cookie_controller = CookieController()
if 'cookie_init' not in st.session_state:
    st.session_state['cookie_init'] = True
    time.sleep(0.3) 
    st.rerun()

# 오늘 완료한 이월 업무 기억 저장소
if 'finished_today' not in st.session_state:
    st.session_state['finished_today'] = []

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

# 💡 데이터 로드
@st.cache_data(ttl=30)
def load_db_data():
    try:
        daily = supabase.table('daily').select("*").execute().data or []
        projects = supabase.table('projects').select("*").execute().data or []
        sub_tasks = supabase.table('sub_tasks').select("*").execute().data or []
        routines = supabase.table('routines').select("*").execute().data or []
        settings = supabase.table('settings').select("*").execute().data or []
        users = supabase.table('users').select("*").execute().data or []
        categories = supabase.table('categories').select("*").execute().data or []
        return daily, projects, sub_tasks, routines, settings, users, categories
    except Exception as e:
        st.warning(f"데이터 로드 실패: {e}")
        return [], [], [], [], [], [], []

def apply_changes():
    load_db_data.clear() 
    st.rerun()

# --- [보안 및 로그인] ---
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_info' not in st.session_state: st.session_state['user_info'] = {}
if 'active_proj_id' not in st.session_state: st.session_state['active_proj_id'] = None
if 'edit_d_id' not in st.session_state: st.session_state['edit_d_id'] = None
if 'edit_s_id' not in st.session_state: st.session_state['edit_s_id'] = None

def check_login(user_id, user_pw):
    _, _, _, _, _, users, _ = load_db_data()
    for u in users:
        if str(u.get('아이디') or '') == str(user_id) and str(u.get('비밀번호') or '') == str(user_pw):
            st.session_state['logged_in'] = True
            st.session_state['user_info'] = u
            return True
    return False

if not st.session_state['logged_in']:
    try:
        saved_id = cookie_controller.get('now_id')
        saved_pw = cookie_controller.get('now_pw')
        if saved_id and saved_pw:
            if check_login(saved_id, saved_pw):
                st.rerun()
    except Exception: pass

if not st.session_state['logged_in']:
    st.markdown("<h1 style='text-align: center;'>🔒 NOWSYSTEM 관제탑</h1>", unsafe_allow_html=True)
    _, l_col, _ = st.columns([1, 1, 1])
    with l_col:
        with st.form("login"):
            in_id = st.text_input("아이디")
            in_pw = st.text_input("비밀번호", type="password")
            if st.form_submit_button("접속하기", use_container_width=True):
                if check_login(in_id, in_pw):
                    try:
                        cookie_controller.set('now_id', in_id)
                        cookie_controller.set('now_pw', in_pw)
                    except Exception: pass
                    st.rerun()
                else: st.error("정보가 일치하지 않습니다.")
    st.stop()

u_info = st.session_state['user_info']
u_name = u_info.get('이름') or '사용자'
u_role = u_info.get('권한') or '일반'

# --- [데이터 할당 및 정렬] ---
all_daily, proj_data, sub_data, routine_data, kpi_config, user_data, cat_data = load_db_data()

proj_data = sorted(proj_data, key=lambda x: (int(x.get('정렬순서') or 999), int(x.get('id') or 0)))
sub_data = sorted(sub_data, key=lambda x: int(x.get('id') or 0))

cat_list = sorted(list(set([str(c.get('분류명') or '') for c in cat_data if pd.notna(c.get('분류명')) and str(c.get('분류명') or '').strip() != ""])))
if not cat_list: cat_list = ["경영관리", "재무업무", "기타"]

KST = datetime.timezone(datetime.timedelta(hours=9))
today_kst = datetime.datetime.now(KST).date()
t_date = st.date_input("📅 업무 기준일 선택", value=today_kst)
t_str = t_date.strftime("%Y-%m-%d")

# --- [사이드바 & 타겟 유저 결정] ---
with st.sidebar:
    st.subheader(f"👤 {u_name}님 ({u_role})")
    if st.button("🔄 최신 데이터 불러오기", use_container_width=True, type="primary"): apply_changes()
    st.divider()
    
    view_target = u_name
    target_user = u_name
    
    if u_role == "마스터":
        st.markdown("**👀 직원 모니터링**")
        user_names = sorted(list(set([u.get('이름') for u in user_data if u.get('이름')])))
        monitor_opts = [u_name, "전체"] + [n for n in user_names if n != u_name]
        
        if 'current_view' not in st.session_state:
            st.session_state['current_view'] = u_name
            
        view_target = st.selectbox("업무를 확인할 직원 선택", monitor_opts, index=monitor_opts.index(st.session_state['current_view']))
        st.session_state['current_view'] = view_target
        target_user = view_target

    st.divider()
    lock_key = f"lock_{t_str}_{target_user}"
    if lock_key not in st.session_state: st.session_state[lock_key] = False
    is_locked = st.session_state[lock_key]
    
    is_readonly = (u_role == "마스터" and target_user != u_name)
    disable_edit = is_locked or is_readonly

    if is_readonly:
        st.info("👀 모니터링 전용 모드 (수정 불가)")
    elif is_locked:
        st.success(f"🔒 {target_user} 업무 마감됨")
        if st.button("🔓 일과 마감 취소", use_container_width=True):
            st.session_state[lock_key] = False
            st.rerun()
    else:
        if st.button("🔒 일과 마감하기", use_container_width=True, type="primary"):
            st.session_state[lock_key] = True
            st.rerun()

    st.divider()
    if st.button("🚪 로그아웃", use_container_width=True): 
        try:
            cookie_controller.remove('now_id')
            cookie_controller.remove('now_pw')
        except Exception: pass
        st.session_state['logged_in'] = False
        st.rerun()

st.title("🚀 NOWSYSTEM 통합 업무 관리")

# 💡 [요청 3번 해결] KPI 이름표 부착 로직
kpi_dict = {str(k.get('KPI명') or '').strip(): str(k.get('구분') or '공통').strip() for k in kpi_config if pd.notna(k.get('KPI명')) and str(k.get('KPI명') or '').strip() != ""}
kpi_owner = target_user if target_user != "전체" else u_name
my_kpi_opts = sorted(list(set([k for k, v in kpi_dict.items() if v in ['공통', kpi_owner]])))

def format_kpi(kpi_name):
    if kpi_name == "기타": return "기타"
    owner = kpi_dict.get(kpi_name, '공통')
    return f"[{owner}] {kpi_name}"

def is_task_visible(d, target_date_str):
    d_id = str(d.get('id'))
    d_date = str(d.get('날짜') or '')
    if not d_date: return False
    prog_val = str(d.get('진행률') or '0')
    prog = int(prog_val) if prog_val.isdigit() else 0
    if d_date == target_date_str: return True
    if d_id in st.session_state['finished_today'] and target_date_str == t_str: return True
    if d_date < target_date_str and prog < 100: return True 
    return False

filtered_daily = [d for d in all_daily if (target_user == "전체" or d.get('담당자') == target_user) and is_task_visible(d, t_str)]

if u_role == "마스터":
    tab_list = ["📝 전사 일과 관리", "📁 전사 프로젝트", "⚙ 설정1 (KPI/계정)", "⚙ 설정2 (업무분류)", "📈 전사 통합 KPI", "📊 데이터/보고서"] if target_user == "전체" else [f"📝 {target_user} 일과", f"📁 {target_user} 프로젝트", "⚙ 설정1 (KPI/계정)", "⚙ 설정2 (업무분류)", f"📈 {target_user} KPI", "📊 데이터/보고서"]
    tabs = st.tabs(tab_list)
    tab_set1, tab_set2, tab_kpi, tab_rep = tabs[2], tabs[3], tabs[4], tabs[5]
else:
    tabs = st.tabs(["📝 나의 일과", "📁 나의 프로젝트", "📈 나의 KPI", "📊 데이터/보고서"])
    tab_kpi, tab_rep = tabs[2], tabs[3]

sub_dict = {p.get("프로젝트명") or "": [] for p in proj_data}
for s in sub_data:
    pn = s.get("프로젝트명") or ""
    if pn in sub_dict: sub_dict[pn].append(s)

# ==========================================
# 탭 1: 일과 관리
# ==========================================
with tabs[0]:
    st.header(f"📝 {t_str} {target_user} 업무 리스트" if target_user != "전체" else f"📝 {t_str} 전사 업무 리스트")
    if not is_readonly:
        with st.expander("➕ 오늘의 업무 추가", expanded=not disable_edit):
            task_type = st.radio("업무 종류 선택", ["일반/데일리 업무", "프로젝트 연동 업무"], horizontal=True, disabled=disable_edit)
            if task_type == "일반/데일리 업무":
                with st.form("add_daily_normal_form", clear_on_submit=True):
                    my_routines = [r.get('업무명') for r in routine_data if r.get('담당자') == target_user]
                    sel_opt = st.selectbox("업무명", ["✏ 직접 입력"] + my_routines, disabled=disable_edit)
                    n_task = st.text_area("내용 (Alt+Enter: 줄바꿈)", height=100, disabled=disable_edit)
                    c1, c2 = st.columns(2)
                    n_cat = c1.selectbox("분류", cat_list, disabled=disable_edit)
                    # 💡 KPI 선택 시 format_kpi 함수 적용
                    n_kpi = c2.selectbox("KPI", my_kpi_opts + ["기타"], format_func=format_kpi, disabled=disable_edit)
                    if st.form_submit_button("추가", type="primary", disabled=disable_edit):
                        final_task = n_task if sel_opt == "✏ 직접 입력" else sel_opt
                        if final_task:
                            supabase.table('daily').insert({"날짜": t_str, "업무명": final_task, "진행률": 0, "프로젝트연동": "FALSE", "분류": n_cat, "KPI": n_kpi, "담당자": target_user, "보고서제외": False, "진행중": False}).execute()
                            apply_changes()
            else:
                p_filter = [p.get('프로젝트명') for p in proj_data if (target_user == "전체" or p.get('담당자') == target_user) and str(p.get('시작일') or '') <= t_str and not (str(p.get('보관함이동') or 'FALSE').upper() == "TRUE")]
                sel_p = st.selectbox("프로젝트 선택", p_filter, disabled=disable_edit) if p_filter else None
                my_subs = [s.get('세부업무명') for s in sub_data if s.get('프로젝트명') == sel_p] if sel_p else []
                sel_s = st.selectbox("세부업무 선택", my_subs, disabled=disable_edit) if my_subs else None
                if st.button("업무 가져오기", type="primary", disabled=disable_edit):
                    if sel_p and sel_s:
                        p_inf = next((p for p in proj_data if p.get('프로젝트명') == sel_p), {})
                        supabase.table('daily').insert({"날짜": t_str, "업무명": sel_s, "진행률": 0, "프로젝트연동": "TRUE", "분류": p_inf.get('분류', '프로젝트'), "연결프로젝트": f"{sel_p}::{sel_s}", "KPI": p_inf.get('KPI', '기타'), "담당자": p_inf.get('담당자', target_user), "보고서제외": False, "진행중": False}).execute()
                        apply_changes()

    st.divider()
for i, row in enumerate(filtered_daily):
        r_id = row.get('id')
        
        # 💡 [수정 1] 아이디를 문자로 강제 변환하여 수정 버튼 클릭 시 반응 없는 에러 완벽 차단
        if not is_readonly and str(st.session_state.get('edit_d_id')) == str(r_id):
            with st.container(border=True):
                e_name = st.text_area("업무명 수정", row.get('업무명') or '', height=80)
                ec1, ec2 = st.columns(2)
                e_cat = ec1.selectbox("분류", cat_list, index=cat_list.index(row.get('분류')) if row.get('분류') in cat_list else 0)
                e_kpi = ec2.selectbox("KPI", my_kpi_opts + ["기타"], format_func=format_kpi, index=(my_kpi_opts + ["기타"]).index(row.get('KPI')) if row.get('KPI') in (my_kpi_opts + ["기타"]) else 0)
                eb1, eb2, _ = st.columns([1, 1, 4])
                
                if eb1.button("저장", type="primary", key=f"esv_{r_id}"):
                    # 💡 [수정 3] 일일업무에서 내용(이름/분류/KPI) 수정 시 프로젝트 탭의 원본 세부업무 및 부모 프로젝트까지 동시 업데이트 (양방향 동기화)
                    is_proj_task = str(row.get('프로젝트연동') or 'FALSE').upper() == "TRUE"
                    old_p_info = str(row.get('연결프로젝트') or '')
                    
                    if is_proj_task and "::" in old_p_info:
                        p_n, old_s_n = old_p_info.split("::", 1)
                        new_p_info = f"{p_n}::{e_name}"
                        
                        s_id_match = next((s.get('id') for s in sub_data if s.get('프로젝트명') == p_n and s.get('세부업무명') == old_s_n), None)
                        if s_id_match: supabase.table('sub_tasks').update({"세부업무명": e_name}).eq('id', s_id_match).execute()
                            
                        p_id_match = next((p.get('id') for p in proj_data if p.get('프로젝트명') == p_n), None)
                        if p_id_match: supabase.table('projects').update({"분류": e_cat, "KPI": e_kpi}).eq('id', p_id_match).execute()
                            
                        for d in all_daily:
                            if str(d.get('연결프로젝트') or '') == old_p_info:
                                supabase.table('daily').update({"연결프로젝트": new_p_info, "업무명": e_name, "분류": e_cat, "KPI": e_kpi}).eq('id', d.get('id')).execute()
                    else:
                        supabase.table('daily').update({"업무명": e_name, "분류": e_cat, "KPI": e_kpi}).eq('id', r_id).execute()
                        
                    st.session_state['edit_d_id'] = None; apply_changes()
                    
                if eb2.button("취소", key=f"ecan_{r_id}"): st.session_state['edit_d_id'] = None; st.rerun()
            continue
            
        c1, c2, c3, c4, c5 = st.columns([4, 3, 1, 0.7, 0.7])
        d_date = str(row.get('날짜') or '')
        carry_txt = f" <small style='color:#E65100; font-weight:bold;'>[🔥이월: {d_date}]</small>" if d_date < t_str else ""
        badge = f" <small style='color:blue;'>[{row.get('담당자') or ''}]</small>" if target_user == "전체" else ""
        
        if str(row.get('프로젝트연동')).upper() == "TRUE": c1.markdown(f"**[{row.get('분류') or '프로젝트'}]** <span style='color:#555;'>{str(row.get('연결프로젝트')).replace('::', ' > ')}</span>{carry_txt}{badge}", unsafe_allow_html=True)
        else: c1.markdown(f"**[{row.get('분류') or '기타'}]** {str(row.get('업무명') or '').replace(chr(10), '<br>')} <small>(KPI: {row.get('KPI') or '미지정'})</small>{carry_txt}{badge}", unsafe_allow_html=True)
        
        cur_p = int(str(row.get('진행률') or '0')) if str(row.get('진행률') or '0').isdigit() else 0
        new_p = c2.slider("진행", 0, 100, cur_p, 10, key=f"ds_{r_id}", label_visibility="collapsed", disabled=disable_edit)
        
        if not disable_edit and new_p != cur_p:
            supabase.table('daily').update({"진행률": new_p}).eq('id', r_id).execute()
            if new_p == 100 and d_date < t_str:
                if str(r_id) not in st.session_state['finished_today']:
                    st.session_state['finished_today'].append(str(r_id))
            
            # 💡 [수정 3] 일일업무에서 진행률 변경 시 프로젝트 세부업무 진행률 동시 업데이트
            if str(row.get('프로젝트연동') or 'FALSE').upper() == "TRUE":
                p_info = str(row.get('연결프로젝트') or '')
                if "::" in p_info:
                    p_n, s_n = p_info.split("::", 1)
                    s_id_match = next((s.get('id') for s in sub_data if s.get('프로젝트명') == p_n and s.get('세부업무명') == s_n), None)
                    if s_id_match: supabase.table('sub_tasks').update({"진행률": new_p}).eq('id', s_id_match).execute()
            apply_changes()
            
        is_ex = bool(row.get('보고서제외', False))
        if c3.checkbox("🚫제외", value=is_ex, key=f"dex_{r_id}", disabled=disable_edit) != is_ex:
            if not disable_edit: supabase.table('daily').update({"보고서제외": not is_ex}).eq('id', r_id).execute(); apply_changes()
            
if not is_readonly:
            if c4.button("✏", key=f"ded_{r_id}", disabled=disable_edit): st.session_state['edit_d_id'] = r_id; st.rerun()
            if c5.button("🗑", key=f"ddl_{r_id}", disabled=disable_edit): supabase.table('daily').delete().eq('id', r_id).execute(); apply_changes()

    st.write("---")
    st.subheader(f"📌 {target_user} 데일리 고정 업무 (루틴)" if target_user != "전체" else "📌 전사 데일리 고정 업무 (루틴)")
    c_r1, c_r2 = st.columns([1, 1])
    with c_r1:
        if not is_readonly:
            with st.form("add_routine_form", clear_on_submit=True):
                r_task = st.text_input("새 데일리 업무명 등록", disabled=disable_edit)
                r_cat = st.selectbox("분류", cat_list, disabled=disable_edit) 
                r_kpi = st.selectbox("연관 KPI", my_kpi_opts + ["기타"], format_func=format_kpi, disabled=disable_edit)
                sub_r_btn = st.form_submit_button("루틴 추가", disabled=disable_edit)
                if sub_r_btn and r_task:
                    supabase.table('routines').insert({"업무명": r_task, "분류": r_cat, "KPI": r_kpi, "담당자": target_user}).execute(); apply_changes()
    with c_r2:
        for i, r in enumerate(routine_data):
            if (u_role == "마스터" and target_user == "전체") or r.get('담당자') == target_user:
                r_id = r.get('id')
                rr1, rr2 = st.columns([4, 1])
                badge_r = f" [{r.get('담당자') or ''}]" if u_role == "마스터" and target_user == "전체" else ""
                rr1.write(f"· [{r.get('분류')}] {r.get('업무명')}{badge_r}")
                if not is_readonly and rr2.button("삭제", key=f"rdel_{r_id}", disabled=disable_edit):
                    supabase.table('routines').delete().eq('id', r_id).execute(); apply_changes()

# ==========================================
# 탭 2: 프로젝트 관리 
# ==========================================
with tabs[1]:
    st.header("📁 프로젝트 현황")
    if not is_readonly:
        with st.expander("✨ 신규 프로젝트 등록", expanded=False):
            with st.form("new_proj_form", clear_on_submit=True):
                pc1, pc2 = st.columns(2)
                p_name = pc1.text_input("프로젝트명", disabled=disable_edit)
                p_cat = pc2.selectbox("분류", cat_list, disabled=disable_edit) 
                p_start = pc1.date_input("시작일", disabled=disable_edit)
                p_kpi = pc1.selectbox("연관 KPI", my_kpi_opts + ["기타"], format_func=format_kpi, disabled=disable_edit)
                if st.form_submit_button("프로젝트 저장", type="primary", disabled=disable_edit) and p_name:
                    supabase.table('projects').insert({"프로젝트명": p_name, "시작일": str(p_start), "완료일": "", "분류": p_cat, "KPI": p_kpi, "담당자": target_user, "정렬순서": 999, "보고서제외": False}).execute(); apply_changes()

    for i, p in enumerate(proj_data):
        r_id = p.get('id')
        if str(p.get("보관함이동") or 'FALSE').upper() == "TRUE": continue
        if u_role == "마스터" and target_user != "전체" and p.get('담당자') != target_user: continue
        if u_role != "마스터" and p.get('담당자') != target_user: continue
        p_start_str = str(p.get("시작일") or "")
        if p_start_str and p_start_str > t_str: continue
        pn = p.get("프로젝트명") or ""
        owner = f" ({p.get('담당자') or ''})" if u_role == "마스터" and target_user == "전체" else ""
        my_s_list = sub_dict.get(pn, [])
        total_p = sum(int(str(s.get('진행률') or '0')) if str(s.get('진행률') or '0').isdigit() else 0 for s in my_s_list)
        avg_p = int(total_p / len(my_s_list)) if len(my_s_list) > 0 else 0
        is_expanded = (st.session_state.get('active_proj_id') == r_id)
        with st.expander(f"📂 {pn} [{p.get('분류')}] - 📊 {avg_p}% {owner}", expanded=is_expanded):
            set_c1, set_c2, set_c3 = st.columns([1, 1, 1.5])
            cur_ord = int(p.get('정렬순서') or 999)
            new_ord = set_c1.number_input("🔢 순서", value=cur_ord, key=f"pord_{r_id}", disabled=disable_edit)
            if not disable_edit and new_ord != cur_ord:
                supabase.table('projects').update({"정렬순서": new_ord}).eq('id', r_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
            cur_end_str = str(p.get("완료일") or "")
            try: cur_end_date = datetime.datetime.strptime(cur_end_str, "%Y-%m-%d").date() if cur_end_str else today_kst
            except: cur_end_date = today_kst
            new_end = set_c2.date_input("🏁 완료일", value=cur_end_date, key=f"pend_edt_{r_id}", disabled=disable_edit)
            if not disable_edit and str(new_end) != cur_end_str:
                supabase.table('projects').update({"완료일": str(new_end)}).eq('id', r_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
            p_ex = bool(p.get('보고서제외', False))
            if set_c3.checkbox("🚫 제외", value=p_ex, key=f"pex_{r_id}", disabled=disable_edit) != p_ex:
                if not disable_edit: supabase.table('projects').update({"보고서제외": not p_ex}).eq('id', r_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
            if not is_readonly:
                with st.form(key=f"sub_form_{r_id}", clear_on_submit=True):
                    sc1, sc2 = st.columns([4,1])
                    new_sub = sc1.text_area("세부 업무 추가", height=80, disabled=disable_edit)
                    if sc2.form_submit_button("추가", disabled=disable_edit) and new_sub:
                        supabase.table('sub_tasks').insert({"프로젝트명": pn, "세부업무명": new_sub, "진행률": 0, "담당자": target_user, "보고서제외": False, "진행중": False}).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
for j, s in enumerate(my_s_list):
                s_id = s.get('id')
                
                # 💡 [수정 1] 아이디 문자열 변환으로 수정 버튼 에러 해결
                if not is_readonly and str(st.session_state.get('edit_s_id')) == str(s_id):
                    with st.container(border=True):
                        e_s_name = st.text_area("세부업무명 수정", s.get('세부업무명') or '', height=80)
                        eb1, eb2, _ = st.columns([1, 1, 4])
                        if eb1.button("저장", type="primary", key=f"esv_s_{s_id}"):
                            supabase.table('sub_tasks').update({"세부업무명": e_s_name}).eq('id', s_id).execute()
                            old_s_name = s.get('세부업무명') or ''
                            
                            # 💡 [수정 3] 프로젝트에서 하위업무 수정 시 당겨진 일일업무 이름 동시 업데이트
                            for d in all_daily:
                                if str(d.get('연결프로젝트') or '') == f"{pn}::{old_s_name}":
                                    supabase.table('daily').update({"연결프로젝트": f"{pn}::{e_s_name}", "업무명": e_s_name}).eq('id', d.get('id')).execute()
                                    
                            st.session_state['edit_s_id'] = None; st.session_state['active_proj_id'] = r_id; apply_changes()
                        if eb2.button("취소", key=f"ecan_s_{s_id}"): st.session_state['edit_s_id'] = None; st.session_state['active_proj_id'] = r_id; st.rerun()
                    continue
                
                sl1, sl2, sl3, sl4, sl5, sl6, sl7 = st.columns([3.5, 2.5, 1, 1.2, 1, 0.6, 0.6])
                sl1.markdown(f"· {str(s.get('세부업무명') or '').replace('\n','<br>')}", unsafe_allow_html=True)
                
                cur_sp = int(str(s.get('진행률') or '0')) if str(s.get('진행률') or '0').isdigit() else 0
                sp = sl2.slider("진행", 0, 100, cur_sp, 10, key=f"s_sld_{s_id}", label_visibility="collapsed", disabled=disable_edit)
                
                if not disable_edit and sp != cur_sp:
                    supabase.table('sub_tasks').update({"진행률": sp}).eq('id', s_id).execute()
                    
                    # 💡 [수정 3] 프로젝트에서 진행률 수정 시 일일업무 진행률 동시 업데이트
                    for d in all_daily:
                        if str(d.get('연결프로젝트') or '') == f"{pn}::{s.get('세부업무명')}":
                            supabase.table('daily').update({"진행률": sp}).eq('id', d.get('id')).execute()
                            
                    st.session_state['active_proj_id'] = r_id; apply_changes()
                
                if sl3.button("✅", key=f"sdone_{s_id}", disabled=disable_edit):
                    supabase.table('sub_tasks').update({"진행률": 100}).eq('id', s_id).execute()
                    for d in all_daily:
                        if str(d.get('연결프로젝트') or '') == f"{pn}::{s.get('세부업무명')}":
                            supabase.table('daily').update({"진행률": 100}).eq('id', d.get('id')).execute()
                    st.session_state['active_proj_id'] = r_id; apply_changes()

                s_prog = bool(s.get('진행중', False))
                if sl4.checkbox("▶", value=s_prog, key=f"s_prg_{s_id}", disabled=disable_edit) != s_prog:
                    if not disable_edit: supabase.table('sub_tasks').update({"진행중": not s_prog}).eq('id', s_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()

                s_ex = bool(s.get('보고서제외', False))
                if sl5.checkbox("🚫", value=s_ex, key=f"s_ex_{s_id}", disabled=disable_edit) != s_ex:
                    if not disable_edit: supabase.table('sub_tasks').update({"보고서제외": not s_ex}).eq('id', s_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
                
                if not is_readonly:
                    if sl6.button("✏", key=f"sedt_{s_id}", disabled=disable_edit): st.session_state['edit_s_id'] = s_id; st.session_state['active_proj_id'] = r_id; st.rerun()
                    if sl7.button("🗑
            st.write("---")
            if not is_readonly:
                ac1, ac2 = st.columns([1,1])
                can_archive = cur_end_str and t_str >= cur_end_str
                if ac1.button("📦 보관함 이동", key=f"arc_{r_id}", disabled=disable_edit or not can_archive):
                    supabase.table('projects').update({"보관함이동": True}).eq('id', r_id).execute(); st.session_state['active_proj_id'] = None; apply_changes()
                if ac2.button("🗑 삭제", key=f"pdel_{r_id}", disabled=disable_edit):
                    supabase.table('projects').delete().eq('id', r_id).execute(); st.session_state['active_proj_id'] = None; apply_changes()

# ==========================================
# 탭 3~5 (설정 유지)
# ==========================================
if u_role == "마스터":
    with tab_set1:
        st.header("⚙ 설정 1 (사내 계정 및 KPI 관리)")
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
        st.header("⚙ 설정 2 (업무 분류 전용 관리)")
        c_df = pd.DataFrame(cat_data)
        e_c_df = st.data_editor(c_df, num_rows="dynamic", use_container_width=False, width=600)
        if st.button("분류 목록 저장", type="primary"):
            orig_c_ids = set(c_df['id'].dropna()) if 'id' in c_df.columns else set()
            new_c_ids = set(e_c_df['id'].dropna()) if 'id' in e_c_df.columns else set()
            for did in orig_c_ids - new_c_ids: supabase.table('categories').delete().eq('id', did).execute()
            for r in e_c_df.to_dict('records'):
                if pd.isna(r.get('id')): r.pop('id', None)
                if r.get('분류명'): supabase.table('categories').upsert(r).execute()
            apply_changes()

with tab_kpi:
    st.header(f"📈 {target_user} KPI" if target_user != "전체" else "📈 전사 통합 KPI")
    
    # 💡 [요청 2번 해결] 일반 직원용 KPI 지표 열람 기능 추가
    with st.expander("📋 할당된 KPI 지표 목록 보기", expanded=False):
        my_kpi_list = [k for k in kpi_config if str(k.get('구분') or '공통').strip() in ['공통', kpi_owner] and str(k.get('KPI명') or '').strip() != ""]
        if my_kpi_list:
            df_kpi = pd.DataFrame(my_kpi_list)
            cols_to_show = [c for c in df_kpi.columns if c not in ['id', 'created_at']]
            st.dataframe(df_kpi[cols_to_show], use_container_width=True)
        else:
            st.info("할당된 KPI 지표가 없습니다.")
    st.divider()

    stats = {}
    for d in all_daily:
        owner = str(d.get('담당자') or '알수없음')
        if u_role != "마스터" and owner != u_name: continue
        if u_role == "마스터" and target_user != "전체" and owner != target_user: continue
        k_name = str(d.get('KPI') or '기타').strip()
        stat_key = f"[{owner}] {k_name}" if u_role == "마스터" and target_user == "전체" else k_name
        if stat_key not in stats: stats[stat_key] = {"sum": 0, "count": 0}
        p_val = str(d.get('진행률') or '0')
        stats[stat_key]["sum"] += int(p_val) if p_val.isdigit() else 0
        stats[stat_key]["count"] += 1
    if stats:
        for stat_key in sorted(stats.keys()):
            data = stats[stat_key]
            avg = int(data["sum"] / data["count"]) if data["count"] > 0 else 0
            st.write(f"**{stat_key}** (총 {data['count']}건)")
            st.progress(avg / 100, text=f"평균 달성률: {avg}%")
    else: st.info("데이터가 없습니다.")

# ==========================================
# 탭 5: 데이터/보고서 
# ==========================================
with tab_rep:
    st.header("📊 데이터 및 보고서 관리")
    if not is_readonly:
        with st.expander("🛠 등록된 전체 업무 일괄 수정"):
            t1, t2, t3 = st.tabs(["📝 일일 업무", "📁 프로젝트", "📋 하위 세부업무"])
            with t1: e_d_df = st.data_editor(pd.DataFrame(all_daily), key="ed_d", use_container_width=True)
            with t2: e_p_df = st.data_editor(pd.DataFrame(proj_data), key="ed_p", use_container_width=True)
            with t3: e_s_df = st.data_editor(pd.DataFrame(sub_data), key="ed_s", use_container_width=True)
            if st.button("💾 일괄 수정 저장", type="primary", disabled=disable_edit):
                for r in e_d_df.to_dict('records'): supabase.table('daily').upsert(r).execute()
                for r in e_p_df.to_dict('records'): supabase.table('projects').upsert(r).execute()
                for r in e_s_df.to_dict('records'): supabase.table('sub_tasks').upsert(r).execute()
                st.success("저장되었습니다!"); apply_changes()

    st.divider()
    st.subheader("🖨 맞춤형 보고서 출력")
    r_type = st.radio("보고서 종류", ["일일(HTML)", "기간별(Excel)"], horizontal=True)
    
    font_css = "font-family: 'Malgun Gothic', '맑은 고딕', sans-serif;"
    
    if r_type == "일일(HTML)":
        r_d = st.date_input("보고 날짜", today_kst, key="r_date_p")
        r_s = r_d.strftime("%Y-%m-%d")
        
        rep_daily = [d for d in all_daily if is_task_visible(d, r_s) and (target_user == "전체" or d.get('담당자') == target_user) and not bool(d.get('보고서제외', False))]
        rep_proj = [p for p in proj_data if target_user == "전체" or p.get('담당자') == target_user]
        rep_routines = [r for r in routine_data if target_user == "전체" or r.get('담당자') == target_user]
        
        h_d_html, grouped_proj = "", {}
h_d_html, grouped_proj = "", {}
        for t in rep_daily:
            prog = int(str(t.get('진행률') or '0')) if str(t.get('진행률') or '0').isdigit() else 0
            is_in_p = bool(t.get('진행중', False))
            
            if prog == 100: icon, prog_txt = "✓", "(완료)"
            elif prog == 0: icon, prog_txt = "□", f"({prog}%)"
            else: icon, prog_txt = "▶", f"({prog}%)"
            if is_in_p and prog < 100: prog_txt += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"
            
            task_n = str(t.get('업무명') or '').replace(chr(10), '<br>')
            
            # 💡 [수정 4] 일일업무 텍스트가 줄바꿈(여러 줄)일 경우 첫 줄만 진하게(Bold), 나머지는 연하게 처리
            task_lines = task_n.split('<br>')
            if len(task_lines) > 1:
                styled_task_n = f"<b>{task_lines[0]}</b><br><span style='color:#777; font-size:0.9em;'>" + "<br>".join(task_lines[1:]) + "</span>"
            else:
                styled_task_n = f"<b>{task_n}</b>"
                
            if str(t.get('프로젝트연동') or 'FALSE').upper() == "TRUE":
                p_n = str(t.get('연결프로젝트') or '').split('::')[0]
                if p_n not in grouped_proj: grouped_proj[p_n] = {'tasks': []}
                grouped_proj[p_n]['tasks'].append(f"{icon} {styled_task_n} {prog_txt}")
            else: 
                h_d_html += f"<li style='margin-bottom:8px;'>{icon} {styled_task_n} {prog_txt}</li>"
                
        for p_n, d in grouped_proj.items():
            # 💡 [수정 2] 해당 프로젝트의 모든 하위업무가 100%인지 검사하여 프로젝트 이름 옆에 (완료) 표기
            p_sub_list = sub_dict.get(p_n, [])
            is_all_done = False
            if len(p_sub_list) > 0 and all(int(str(s.get('진행률') or '0')) == 100 for s in p_sub_list):
                is_all_done = True
                
            done_badge = " <span style='color:#2e7d32; font-weight:bold;'>(완료)</span>" if is_all_done else ""
            
            h_d_html += f"<li style='margin-bottom:8px;'><b>{p_n}</b>{done_badge} <span style='color:#777; font-size:0.85em;'>(아래 상세 참조)</span><ul style='margin-top:4px; margin-bottom:0;'>"
            for sub_t in d['tasks']: h_d_html += f"<li style='margin-bottom:4px; list-style-type: none;'>{sub_t}</li>"
            h_d_html += "</ul></li>"
            
        h_r_html = "".join([f"<li style='margin-bottom:8px;'>✓ {str(r.get('업무명') or '').replace(chr(10), '<br>')} (완료)</li>" for r in rep_routines])
        
        h_p_html = ""
        for p in rep_proj:
            if bool(p.get('보관함이동', False)) or bool(p.get('보고서제외', False)): continue
            pn = p.get('프로젝트명') or ''; valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue 
            h_p_html += f"<div style='margin-top:15px;'><h4 style='margin-bottom:5px;'>■ {pn}</h4><ul style='margin-top:0;'>"
            for s in valid_subs:
                prog = int(str(s.get('진행률') or '0')); is_in_p = bool(s.get('진행중', False))
                if prog == 100: icon, pr_t = "✓", "(완료)"
                elif prog == 0: icon, pr_t = "□", f"({prog}%)"
                else: icon, pr_t = "▶", f"({prog}%)"
                if is_in_p and prog < 100: pr_t += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"
                h_p_html += f"<li style='margin-bottom:5px;'>{icon} {str(s.get('세부업무명') or '').replace(chr(10), '<br>')} {pr_t}</li>"
            h_p_html += "</ul></div>"
            
        full_html = f"<html><head><meta charset='utf-8'></head><body style='{font_css}'><h2>[{r_s}] 업무 내용</h2><h3>■ 일일 업무</h3><ul style='line-height:1.5;'>{h_d_html}</ul><hr><h3>■ 고정 업무</h3><ul style='line-height:1.5;'>{h_r_html}</ul><hr><h3>■ 프로젝트 현황</h3>{h_p_html}</body></html>"
        st.components.v1.html(full_html, height=400, scrolling=True)
        c_btn1, c_btn2 = st.columns([1, 1])
        with c_btn1: st.download_button("📥 HTML 다운로드", full_html.encode('utf-8'), f"[{r_s}] 업무 내용.html", use_container_width=True)
        with c_btn2: 
            with st.expander("📋 HTML 복사"): st.code(full_html, language="html")

    elif r_type == "기간별(Excel)":
        c_ds1, c_ds2 = st.columns(2)
        s_w, e_w = c_ds1.date_input("시작일", today_kst - datetime.timedelta(days=7), key="ws"), c_ds2.date_input("종료일", today_kst, key="we")
        
        export_items = []
        rep_proj = [p for p in proj_data if (target_user == "전체" or p.get('담당자') == target_user) and not bool(p.get('보관함이동', False)) and not bool(p.get('보고서제외', False))]
        for p in rep_proj:
            pn = p.get('프로젝트명') or ''
            cat = p.get('분류') or '기타'
            valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue
            ph, total_p = f"<b>[프로젝트] {pn}</b><br>", 0
            for s in valid_subs:
                prog = int(str(s.get('진행률') or '0'))
                total_p += prog
                ph += f"- {str(s.get('세부업무명') or '').replace(chr(10), '<br>')} ({prog}%)<br>"
            avg_p = int(total_p / len(valid_subs))
            export_items.append({'cat': cat, 'content': ph, 'prog': avg_p})
            
        s_w_str, e_w_str = s_w.strftime("%Y-%m-%d"), e_w.strftime("%Y-%m-%d")
        daily_period = [d for d in all_daily if s_w_str <= str(d.get('날짜') or '') <= e_w_str and (target_user == "전체" or d.get('담당자') == target_user) and not bool(d.get('보고서제외', False)) and str(d.get('프로젝트연동') or 'FALSE').upper() != "TRUE"]
        for d in daily_period:
            cat = d.get('분류') or '기타'
            task_n = str(d.get('업무명') or '').replace(chr(10), '<br>')
            d_date = str(d.get('날짜') or '')
            prog = int(str(d.get('진행률') or '0'))
            ph = f"{task_n} <span style='color:#777; font-size:0.85em;'>({d_date})</span>"
            export_items.append({'cat': cat, 'content': ph, 'prog': prog})
            
        routines_period = [r for r in routine_data if (target_user == "전체" or r.get('담당자') == target_user)]
        for r in routines_period:
            cat = r.get('분류') or '기타'
            task_n = str(r.get('업무명') or '').replace(chr(10), '<br>')
            export_items.append({'cat': cat, 'content': f"<b>[루틴]</b> {task_n}", 'prog': 100})
            
        export_items = sorted(export_items, key=lambda x: x['cat'])
        
        xls_hr = ""
        for item in export_items:
            xls_hr += f"<tr><td style='vertical-align: middle; {font_css}'>{item['cat']}</td><td style='vertical-align: top; {font_css}'>{item['content']}</td><td style='text-align:center; vertical-align: middle; {font_css}'><b>{item['prog']}%</b></td><td style='{font_css}'></td></tr>"
            
        th = f"<tr><th style='background:#e0f7fa; padding:8px; {font_css}'>업무분류</th><th style='background:#e0f7fa; padding:8px; {font_css}'>업무내역</th><th style='background:#e0f7fa; padding:8px; {font_css}'>진행률</th><th style='background:#e0f7fa; padding:8px; {font_css}'>예정사항</th></tr>"
        xls_html = f"<html><head><meta charset='utf-8'></head><body><h2 style='{font_css}'>[{s_w} ~ {e_w}] 업무 내용</h2><table style='border-collapse:collapse; width:100%; border: 1px solid #ccc;'>{th}{xls_hr}</table></body></html>"
        
        st.download_button("💾 Excel 다운로드 (.xls)", xls_html.encode('utf-8-sig'), f"[{s_w}_{e_w}] 업무 내용.xls")
