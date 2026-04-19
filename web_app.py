import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime
import io
import time
from streamlit_cookies_controller import CookieController

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑 V52 (게시판 새창 UI 개편)", layout="wide")

# 쿠키 컨트롤러
cookie_controller = CookieController()
if 'cookie_init' not in st.session_state:
    st.session_state['cookie_init'] = True
    time.sleep(0.3) 
    st.rerun()

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

# 💡 실시간 반영을 위해 캐시 완전 제거됨 (신규 테이블 포함)
def load_db_data():
    try:
        daily = supabase.table('daily').select("*").execute().data or []
        projects = supabase.table('projects').select("*").execute().data or []
        sub_tasks = supabase.table('sub_tasks').select("*").execute().data or []
        routines = supabase.table('routines').select("*").execute().data or []
        users = supabase.table('users').select("*").execute().data or []
        categories = supabase.table('categories').select("*").execute().data or []
        kpi_targets = supabase.table('kpi_targets').select("*").execute().data or []
        kpi_details = supabase.table('kpi_details').select("*").execute().data or []
        kpi_submissions = supabase.table('kpi_submissions').select("*").execute().data or []
        shared_boards = supabase.table('shared_boards').select("*").execute().data or []
        
        return daily, projects, sub_tasks, routines, users, categories, kpi_targets, kpi_details, kpi_submissions, shared_boards
    except Exception as e:
        st.warning(f"데이터 로드 실패: {e}")
        return [], [], [], [], [], [], [], [], [], []

def apply_changes():
    st.rerun()

# --- [보안 및 로그인 세션] ---
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_info' not in st.session_state: st.session_state['user_info'] = {}
if 'active_proj_id' not in st.session_state: st.session_state['active_proj_id'] = None
if 'edit_d_id' not in st.session_state: st.session_state['edit_d_id'] = None
if 'edit_s_id' not in st.session_state: st.session_state['edit_s_id'] = None
if 'edit_kpi_id' not in st.session_state: st.session_state['edit_kpi_id'] = None
if 'edit_detail_id' not in st.session_state: st.session_state['edit_detail_id'] = None
if 'edit_board_id' not in st.session_state: st.session_state['edit_board_id'] = None
# 💡 [신규] 넓은 화면 게시판 출력을 위한 세션
if 'active_modal_board' not in st.session_state: st.session_state['active_modal_board'] = None

def check_login(user_id, user_pw):
    _, _, _, _, users, _, _, _, _, _ = load_db_data()
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
all_daily, proj_data, sub_data, routine_data, user_data, cat_data, kpi_targets, kpi_details, kpi_subs, shared_boards = load_db_data()

proj_data = sorted(proj_data, key=lambda x: (int(x.get('정렬순서') or 999), int(x.get('id') or 0)))
sub_data = sorted(sub_data, key=lambda x: int(x.get('id') or 0))

cat_list = sorted(list(set([str(c.get('분류명') or '') for c in cat_data if pd.notna(c.get('분류명')) and str(c.get('분류명') or '').strip() != ""])))
if not cat_list: cat_list = ["경영관리", "재무업무", "기타"]

KST = datetime.timezone(datetime.timedelta(hours=9))
today_kst = datetime.datetime.now(KST).date()
t_date = st.date_input("📅 업무 기준일 선택", value=today_kst)
t_str = t_date.strftime("%Y-%m-%d")

# 금주 월요일 계산 헬퍼 함수
def get_monday(date_obj):
    return date_obj - datetime.timedelta(days=date_obj.weekday())
current_monday_str = get_monday(today_kst).strftime("%Y-%m-%d")

# 💡 [신규] 넓은 메인화면 출력용 게시판 렌더링 함수
def render_shared_board_full_width(board_type, title, is_weekly=False):
    c_back, _ = st.columns([2, 8])
    if c_back.button("🔙 메인 화면으로 돌아가기", type="primary", use_container_width=True):
        st.session_state['active_modal_board'] = None
        st.rerun()
        
    st.header(title)
    st.divider()
    
    boards = [b for b in shared_boards if b.get('board_type') == board_type]
    if is_weekly:
        boards = [b for b in boards if b.get('week_start') == current_monday_str]
    
    for b in boards:
        b_id = b.get('id')
        is_author = (b.get('author') == u_name)
        
        if str(st.session_state.get('edit_board_id')) == str(b_id) and is_author:
            with st.container(border=True):
                e_content = st.text_area("내용 수정", value=b.get('content') or '', key=f"eb_{b_id}", height=120)
                eb1, eb2, _ = st.columns([1, 1, 8])
                if eb1.button("저장", key=f"ebs_{b_id}", type="primary"):
                    supabase.table('shared_boards').update({"content": e_content}).eq('id', b_id).execute()
                    st.session_state['edit_board_id'] = None; apply_changes()
                if eb2.button("취소", key=f"ebc_{b_id}"):
                    st.session_state['edit_board_id'] = None; st.rerun()
        else:
            with st.container(border=True):
                # 가독성을 위해 넓게 배치 [내용 8 : 버튼 2]
                sc1, sc2 = st.columns([8, 2])
                status_badge = "✅ 완료" if b.get('status') == '완료' else "▶ 진행중"
                sc1.markdown(f"**👤 {b.get('author')}** | {status_badge}")
                sc1.write(b.get('content'))
                
                if is_author:
                    b1, b2, b3 = sc2.columns(3)
                    if b.get('status') != '완료' and b1.button("✅완료", key=f"bd_{b_id}"):
                        supabase.table('shared_boards').update({"status": "완료"}).eq('id', b_id).execute(); apply_changes()
                    if b2.button("✏️수정", key=f"be_{b_id}"):
                        st.session_state['edit_board_id'] = b_id; st.rerun()
                    if b3.button("🗑️삭제", key=f"bx_{b_id}"):
                        supabase.table('shared_boards').delete().eq('id', b_id).execute(); apply_changes()

    st.subheader("📝 새 내용 작성")
    with st.form(f"form_{board_type}", clear_on_submit=True):
        new_content = st.text_area("내용을 상세히 입력하세요", height=150, key=f"new_content_{board_type}")
        if st.form_submit_button("게시물 등록", type="primary") and new_content:
            supabase.table('shared_boards').insert({
                "board_type": board_type, "content": new_content, "author": u_name, 
                "status": "진행중", "week_start": current_monday_str
            }).execute(); apply_changes()

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
        
        if 'current_view' not in st.session_state: st.session_state['current_view'] = u_name
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
        if st.button("🔓 일과 마감 취소", use_container_width=True): st.session_state[lock_key] = False; st.rerun()
    else:
        if st.button("🔒 일과 마감하기", use_container_width=True, type="primary"): st.session_state[lock_key] = True; st.rerun()
  
    st.divider()
    # 💡 [변경] 사이드바에서는 오직 '버튼'만 제공하여 클릭 시 메인 화면으로 전환
    st.markdown("**📂 부서 공용 게시판**")
    if st.button("📬 타 부서 요청 사항", use_container_width=True):
        st.session_state['active_modal_board'] = ('요청사항', "📬 타 부서 요청 사항 (공용)", False)
        st.rerun()
    if st.button("💰 금주 세금계산서 현황", use_container_width=True):
        st.session_state['active_modal_board'] = ('세금계산서', "💰 금주 세금계산서 현황 (매주 리셋)", True)
        st.rerun()
    if st.button("📄 금주 분납서 이슈 현황", use_container_width=True):
        st.session_state['active_modal_board'] = ('분납서', "📄 금주 분납서 이슈 현황 (매주 리셋)", True)
        st.rerun()
    if st.button("🛡️ 금주 보증보험 이슈 현황", use_container_width=True):
        st.session_state['active_modal_board'] = ('보증보험', "🛡️ 금주 보증보험 이슈 현황 (매주 리셋)", True)
        st.rerun()
    if st.button("💾 서버 저장 위치 정보", use_container_width=True):
        st.session_state['active_modal_board'] = ('서버위치', "💾 서버 저장 위치 정보 (공용)", False)
        st.rerun()
    
    st.divider()
    if st.button("🚪 로그아웃", use_container_width=True): 
        try: cookie_controller.remove('now_id'); cookie_controller.remove('now_pw')
        except Exception: pass
        st.session_state['logged_in'] = False; st.rerun()

# 💡 [핵심] 사이드바 게시판 버튼이 눌렸을 때 메인 화면 전체를 덮어쓰는 로직
if st.session_state.get('active_modal_board'):
    b_type, b_title, b_weekly = st.session_state['active_modal_board']
    render_shared_board_full_width(b_type, b_title, b_weekly)
    st.stop() # st.stop()을 호출하여 하단의 일과/프로젝트/KPI 탭 렌더링을 차단하고 게시판만 보여줍니다.

# ----------------------------------------------------------------------
# 아래부터는 메인 화면 (게시판이 꺼져 있을 때만 보임)
st.title("🚀 NOWSYSTEM 통합 업무 관리")

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
    tab_list = ["📝 전사 일과 관리", "📁 전사 프로젝트", "⚙ 설정 (계정/분류)", "📈 통합 KPI 관리", "📊 데이터/보고서"] if target_user == "전체" else [f"📝 {target_user} 일과", f"📁 {target_user} 프로젝트", "⚙ 설정 (계정/분류)", f"📈 {target_user} KPI", "📊 데이터/보고서"]
    tabs = st.tabs(tab_list)
    tab_set1, tab_kpi, tab_rep = tabs[2], tabs[3], tabs[4]
else:
    tabs = st.tabs(["📝 나의 일과", "📁 나의 프로젝트", "📈 나의 KPI", "📊 데이터/보고서"])
    tab_kpi, tab_rep = tabs[2], tabs[3]
  
sub_dict = {p.get("프로젝트명") or "": [] for p in proj_data}
for s in sub_data:
    pn = s.get("프로젝트명") or ""
    if pn in sub_dict: sub_dict[pn].append(s)

# ==========================================
# 💡 콜백 함수 세팅 
# ==========================================
daily_to_sub = {}
sub_to_daily = {}
for d in all_daily:
    if str(d.get('프로젝트연동') or 'FALSE').upper() == "TRUE":
        p_info = str(d.get('연결프로젝트') or '')
        if "::" in p_info:
            p_n, s_n = p_info.split("::", 1)
            p_n, s_n = p_n.strip(), s_n.strip()
            for s_item in sub_data:
                if str(s_item.get('프로젝트명') or '').strip() == p_n and str(s_item.get('세부업무명') or '').strip() == s_n:
                    d_id = d.get('id'); s_id = s_item.get('id')
                    daily_to_sub[d_id] = s_id
                    if s_id not in sub_to_daily: sub_to_daily[s_id] = []
                    sub_to_daily[s_id].append(d_id)
                    break

def on_daily_slider_change(d_id, d_date, target_str):
    new_val = st.session_state[f"ds_{d_id}"]
    supabase.table('daily').update({"진행률": new_val}).eq('id', d_id).execute()
    if new_val == 100 and d_date < target_str:
        if str(d_id) not in st.session_state['finished_today']: st.session_state['finished_today'].append(str(d_id))
    s_id = daily_to_sub.get(d_id)
    if s_id:
        supabase.table('sub_tasks').update({"진행률": new_val}).eq('id', s_id).execute()
        st.session_state[f"s_sld_{s_id}"] = new_val

def on_sub_slider_change(s_id, p_id):
    new_val = st.session_state[f"s_sld_{s_id}"]
    supabase.table('sub_tasks').update({"진행률": new_val}).eq('id', s_id).execute()
    d_ids = sub_to_daily.get(s_id, [])
    for d_id in d_ids:
        supabase.table('daily').update({"진행률": new_val}).eq('id', d_id).execute()
        st.session_state[f"ds_{d_id}"] = new_val
    st.session_state['active_proj_id'] = str(p_id)

def on_complete_button_click(s_id, p_id):
    supabase.table('sub_tasks').update({"진행률": 100}).eq('id', s_id).execute()
    st.session_state[f"s_sld_{s_id}"] = 100
    d_ids = sub_to_daily.get(s_id, [])
    for d_id in d_ids:
        supabase.table('daily').update({"진행률": 100}).eq('id', d_id).execute()
        st.session_state[f"ds_{d_id}"] = 100
    st.session_state['active_proj_id'] = str(p_id)

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
                    n_cat = st.selectbox("분류", cat_list, disabled=disable_edit)
                    if st.form_submit_button("추가", type="primary", disabled=disable_edit):
                        final_task = n_task if sel_opt == "✏ 직접 입력" else sel_opt
                        if final_task:
                            supabase.table('daily').insert({"날짜": t_str, "업무명": final_task, "진행률": 0, "프로젝트연동": "FALSE", "분류": n_cat, "담당자": target_user, "보고서제외": False, "진행중": False}).execute()
                            apply_changes()
            else:
                p_filter = [p.get('프로젝트명') for p in proj_data if (target_user == "전체" or p.get('담당자') == target_user) and str(p.get('시작일') or '') <= t_str and not (str(p.get('보관함이동') or 'FALSE').upper() == "TRUE")]
                sel_p = st.selectbox("프로젝트 선택", p_filter, disabled=disable_edit) if p_filter else None
                my_subs = [s.get('세부업무명') for s in sub_data if s.get('프로젝트명') == sel_p] if sel_p else []
                sel_s = st.selectbox("세부업무 선택", my_subs, disabled=disable_edit) if my_subs else None
                if st.button("업무 가져오기", type="primary", disabled=disable_edit):
                    if sel_p and sel_s:
                        p_inf = next((p for p in proj_data if p.get('프로젝트명') == sel_p), {})
                        supabase.table('daily').insert({"날짜": t_str, "업무명": sel_s, "진행률": 0, "프로젝트연동": "TRUE", "분류": p_inf.get('분류', '프로젝트'), "연결프로젝트": f"{sel_p}::{sel_s}", "담당자": p_inf.get('담당자', target_user), "보고서제외": False, "진행중": False}).execute()
                        apply_changes()
  
    st.divider()
    for i, row in enumerate(filtered_daily):
        r_id = row.get('id')
        if not is_readonly and str(st.session_state.get('edit_d_id')) == str(r_id):
            with st.container(border=True):
                e_name = st.text_area("업무명 수정", row.get('업무명') or '', height=80, key=f"ed_name_{r_id}")
                e_cat = st.selectbox("분류", cat_list, index=cat_list.index(row.get('분류')) if row.get('분류') in cat_list else 0, key=f"ed_cat_{r_id}")
                eb1, eb2, _ = st.columns([1, 1, 4])
                if eb1.button("저장", type="primary", key=f"esv_{r_id}"):
                    is_proj_task = str(row.get('프로젝트연동') or 'FALSE').upper() == "TRUE"
                    old_p_info = str(row.get('연결프로젝트') or '')
                    if is_proj_task and "::" in old_p_info:
                        p_n, old_s_n = old_p_info.split("::", 1)
                        new_p_info = f"{p_n}::{e_name}"
                        s_id_match = next((s.get('id') for s in sub_data if s.get('프로젝트명') == p_n and s.get('세부업무명') == old_s_n), None)
                        if s_id_match: supabase.table('sub_tasks').update({"세부업무명": e_name}).eq('id', s_id_match).execute()
                        p_id_match = next((p.get('id') for p in proj_data if p.get('프로젝트명') == p_n), None)
                        if p_id_match: supabase.table('projects').update({"분류": e_cat}).eq('id', p_id_match).execute()
                        for d in all_daily:
                            if str(d.get('연결프로젝트') or '') == old_p_info:
                                supabase.table('daily').update({"연결프로젝트": new_p_info, "업무명": e_name, "분류": e_cat}).eq('id', d.get('id')).execute()
                    else: supabase.table('daily').update({"업무명": e_name, "분류": e_cat}).eq('id', r_id).execute()
                    st.session_state['edit_d_id'] = None; apply_changes()
                if eb2.button("취소", key=f"ecan_{r_id}"): st.session_state['edit_d_id'] = None; st.rerun()
            continue
            
        c1, c2, c3, c4, c5, c6 = st.columns([3.5, 2.5, 0.8, 0.8, 0.8, 1.2])
        d_date = str(row.get('날짜') or '')
        carry_txt = f" <small style='color:#E65100; font-weight:bold;'>[🔥이월: {d_date}]</small>" if d_date < t_str else ""
        badge = f" <small style='color:blue;'>[{row.get('담당자') or ''}]</small>" if target_user == "전체" else ""
        
        task_title = str(row.get('연결프로젝트')).replace('::', ' > ') if str(row.get('프로젝트연동')).upper() == "TRUE" else str(row.get('업무명') or '').replace(chr(10), '<br>')
        is_copied_badge = " 🔗(이슈복사됨)" if row.get('is_copied') else ""
        c1.markdown(f"**[{row.get('분류')}]** <span style='color:#555;'>{task_title}</span>{carry_txt}{badge}{is_copied_badge}", unsafe_allow_html=True)
        
        cur_p = int(str(row.get('진행률') or '0')) if str(row.get('진행률') or '0').isdigit() else 0
        c2.slider("진행", 0, 100, cur_p, 10, key=f"ds_{r_id}", on_change=on_daily_slider_change, args=(r_id, d_date, t_str), label_visibility="collapsed", disabled=disable_edit)
            
        is_ex = bool(row.get('보고서제외', False))
        if c3.checkbox("🚫제외", value=is_ex, key=f"dex_{r_id}", disabled=disable_edit) != is_ex:
            if not disable_edit: supabase.table('daily').update({"보고서제외": not is_ex}).eq('id', r_id).execute(); apply_changes()
            
        if not is_readonly:
            if c4.button("✏수정", key=f"ded_{r_id}", disabled=disable_edit): st.session_state['edit_d_id'] = r_id; st.rerun()
            if c5.button("🗑삭제", key=f"ddl_{r_id}", disabled=disable_edit): supabase.table('daily').delete().eq('id', r_id).execute(); apply_changes()
            
            with c6.popover("➡️ 이슈복사"):
                sel_issue = st.selectbox("복사할 게시판 선택", ["세금계산서", "분납서", "보증보험"], key=f"sel_issue_{r_id}")
                if st.button("복사하기", key=f"cpy_{r_id}"):
                    supabase.table('shared_boards').insert({
                        "board_type": sel_issue, "content": f"[{row.get('분류')}] {task_title}", 
                        "author": u_name, "status": "진행중", "week_start": current_monday_str
                    }).execute()
                    supabase.table('daily').update({"is_copied": True}).eq('id', r_id).execute()
                    st.success("게시판에 복사되었습니다!"); time.sleep(0.5); apply_changes()
  
    st.write("---")
    st.subheader(f"📌 {target_user} 데일리 고정 업무 (루틴)" if target_user != "전체" else "📌 전사 데일리 고정 업무 (루틴)")
    c_r1, c_r2 = st.columns([1, 1])
    with c_r1:
        if not is_readonly:
            with st.form("add_routine_form", clear_on_submit=True):
                r_task = st.text_input("새 데일리 업무명 등록", disabled=disable_edit)
                r_cat = st.selectbox("분류", cat_list, disabled=disable_edit) 
                if st.form_submit_button("루틴 추가", disabled=disable_edit) and r_task:
                    supabase.table('routines').insert({"업무명": r_task, "분류": r_cat, "담당자": target_user}).execute(); apply_changes()
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
                if st.form_submit_button("프로젝트 저장", type="primary", disabled=disable_edit) and p_name:
                    supabase.table('projects').insert({"프로젝트명": p_name, "시작일": str(p_start), "완료일": "", "분류": p_cat, "담당자": target_user, "정렬순서": 999, "보고서제외": False}).execute(); apply_changes()
  
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
        
        is_expanded = (str(st.session_state.get('active_proj_id')) == str(r_id))
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
                    new_sub = sc1.text_area("세부 업무 추가", height=80, disabled=disable_edit, key=f"new_sub_{r_id}")
                    if sc2.form_submit_button("추가", disabled=disable_edit) and new_sub:
                        supabase.table('sub_tasks').insert({"프로젝트명": pn, "세부업무명": new_sub, "진행률": 0, "담당자": target_user, "보고서제외": False, "진행중": False}).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
            
            for j, s in enumerate(my_s_list):
                s_id = s.get('id')
                if not is_readonly and str(st.session_state.get('edit_s_id')) == str(s_id):
                    with st.container(border=True):
                        e_s_name = st.text_area("세부업무명 수정", s.get('세부업무명') or '', height=80, key=f"e_s_name_{s_id}")
                        eb1, eb2, _ = st.columns([1, 1, 4])
                        if eb1.button("저장", type="primary", key=f"esv_s_{s_id}"):
                            supabase.table('sub_tasks').update({"세부업무명": e_s_name}).eq('id', s_id).execute()
                            old_s_name = s.get('세부업무명') or ''
                            for d in all_daily:
                                if str(d.get('연결프로젝트') or '') == f"{pn}::{old_s_name}":
                                    supabase.table('daily').update({"연결프로젝트": f"{pn}::{e_s_name}", "업무명": e_s_name}).eq('id', d.get('id')).execute()
                            st.session_state['edit_s_id'] = None; st.session_state['active_proj_id'] = r_id; apply_changes()
                        if eb2.button("취소", key=f"ecan_s_{s_id}"): st.session_state['edit_s_id'] = None; st.session_state['active_proj_id'] = r_id; st.rerun()
                    continue
                
                sl1, sl2, sl3, sl4, sl5, sl6, sl7 = st.columns([2.5, 2.0, 1.2, 1.4, 1.3, 0.9, 0.9])
                sl1.markdown(f"· {str(s.get('세부업무명') or '').replace('\n','<br>')}", unsafe_allow_html=True)
                cur_sp = int(str(s.get('진행률') or '0')) if str(s.get('진행률') or '0').isdigit() else 0
                sl2.slider("진행", 0, 100, cur_sp, 10, key=f"s_sld_{s_id}", on_change=on_sub_slider_change, args=(s_id, r_id), label_visibility="collapsed", disabled=disable_edit)
                sl3.button("✅완료", key=f"sdone_{s_id}", on_click=on_complete_button_click, args=(s_id, r_id), disabled=disable_edit)
  
                s_prog = bool(s.get('진행중', False))
                if sl4.checkbox("▶진행중", value=s_prog, key=f"s_prg_{s_id}", disabled=disable_edit) != s_prog:
                    if not disable_edit: supabase.table('sub_tasks').update({"진행중": not s_prog}).eq('id', s_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
  
                s_ex = bool(s.get('보고서제외', False))
                if sl5.checkbox("🚫출력제외", value=s_ex, key=f"s_ex_{s_id}", disabled=disable_edit) != s_ex:
                    if not disable_edit: supabase.table('sub_tasks').update({"보고서제외": not s_ex}).eq('id', s_id).execute(); st.session_state['active_proj_id'] = r_id; apply_changes()
                
                if not is_readonly:
                    if sl6.button("✏수정", key=f"sedt_{s_id}", disabled=disable_edit): st.session_state['edit_s_id'] = str(s_id); st.session_state['active_proj_id'] = str(r_id); st.rerun()
                    if sl7.button("🗑삭제", key=f"sdel_{s_id}", disabled=disable_edit): supabase.table('sub_tasks').delete().eq('id', s_id).execute(); st.session_state['active_proj_id'] = str(r_id); apply_changes()
            st.write("---")
            if not is_readonly:
                ac1, ac2 = st.columns([1,1])
                can_archive = cur_end_str and t_str >= cur_end_str
                if ac1.button("📦 보관함 이동", key=f"arc_{r_id}", disabled=disable_edit or not can_archive):
                    supabase.table('projects').update({"보관함이동": True}).eq('id', r_id).execute(); st.session_state['active_proj_id'] = None; apply_changes()
                if ac2.button("🗑 삭제", key=f"pdel_{r_id}", disabled=disable_edit):
                    supabase.table('projects').delete().eq('id', r_id).execute(); st.session_state['active_proj_id'] = None; apply_changes()
                    
    st.divider()
    with st.expander("📦 프로젝트 보관함 (종료된 업무)"):
        archived_projs = [p for p in proj_data if str(p.get("보관함이동") or 'FALSE').upper() == "TRUE"]
        if u_role != "마스터" or target_user != "전체":
            archived_projs = [p for p in archived_projs if p.get('담당자') == target_user]

        if not archived_projs:
            st.info("보관함에 있는 프로젝트가 없습니다.")
        else:
            for p in archived_projs:
                arc_id = p.get('id')
                arc_pn = p.get('프로젝트명') or ''
                arc_owner = f" ({p.get('담당자') or ''})" if u_role == "마스터" and target_user == "전체" else ""
                arc_cat = p.get('분류') or '기타'
                arc_end = p.get('완료일') or '미지정'

                arc_c1, arc_c2, arc_c3, arc_c4 = st.columns([4, 2, 1.5, 1.5])
                arc_c1.markdown(f"**[{arc_cat}]** <span style='color:#777; text-decoration: line-through;'>{arc_pn}</span>{arc_owner}", unsafe_allow_html=True)
                arc_c2.write(f"완료일: {arc_end}")

                if not is_readonly:
                    if arc_c3.button("🔄 복구", key=f"unarc_{arc_id}", disabled=disable_edit):
                        supabase.table('projects').update({"보관함이동": False}).eq('id', arc_id).execute()
                        st.session_state['active_proj_id'] = arc_id; apply_changes()
                    if arc_c4.button("🗑 영구삭제", key=f"harddel_{arc_id}", disabled=disable_edit):
                        supabase.table('projects').delete().eq('id', arc_id).execute(); apply_changes()

# ==========================================
# 탭 3: 마스터 설정 (계정/분류)
# ==========================================
if u_role == "마스터":
    with tab_set1:
        st.header("⚙ 설정 (사내 계정 및 분류 관리)")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("계정 관리")
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
            st.subheader("업무 분류 관리")
            c_df = pd.DataFrame(cat_data)
            e_c_df = st.data_editor(c_df, num_rows="dynamic", use_container_width=True)
            if st.button("분류 목록 저장", type="primary"):
                orig_c_ids = set(c_df['id'].dropna()) if 'id' in c_df.columns else set()
                new_c_ids = set(e_c_df['id'].dropna()) if 'id' in e_c_df.columns else set()
                for did in orig_c_ids - new_c_ids: supabase.table('categories').delete().eq('id', did).execute()
                for r in e_c_df.to_dict('records'):
                    if pd.isna(r.get('id')): r.pop('id', None)
                    if r.get('분류명'): supabase.table('categories').upsert(r).execute()
                apply_changes()

# ==========================================
# 탭 4: 전면 개편된 독립형 KPI 시스템
# ==========================================
with tab_kpi:
    def format_target(val, unit):
        try: v = int(val)
        except: v = val
        if unit == "금액(원)": return f"{v:,}원"
        elif unit == "요율(%)": return f"{v}%"
        else: return f"{v}건"

    def calculate_kpi_score(target, submissions):
        t_id = target.get('id')
        t_name = target.get('kpi_name', '')
        t_owner = target.get('owner', '공통')
        total = int(target.get('target_count') or 1)
        weight = int(target.get('weight') or 0)
        
        if t_owner == "공통":
            approved = len([s for s in submissions if str(s.get('kpi_id')) == str(t_id) and s.get('status') == '승인'])
        else:
            approved = len([s for s in submissions if str(s.get('kpi_id')) == str(t_id) and s.get('user_name') == t_owner and s.get('status') == '승인'])
            
        missing = max(0, total - approved)
        rate = (approved / total) * 100 if total > 0 else 0
        
        points = 0
        if "누락" in t_name or "법정" in t_name: 
            if missing == 0: points = 15
            elif 1 <= missing <= 2: points = 8
            else: points = 0
        elif "정시율" in t_name: 
            if rate >= 90: points = 10
            elif rate >= 85: points = 8
            elif rate >= 80: points = 6
            elif rate >= 70: points = 3
            else: points = 0
        else: 
            points = round((rate / 100) * weight, 1)
            
        return approved, total, points, rate, missing

    if u_role == "마스터" and target_user == "전체":
        st.header("👑 전사 KPI 승인 및 지표 할당")
        mt1, mt2 = st.tabs(["✅ 실무자 증빙 승인", "⚙️ KPI 목표 및 상세 할당 설정"])
        
        with mt1:
            st.subheader("대기 중인 확인 요청")
            pending_reqs = [s for s in kpi_subs if s.get('status') == '대기']
            if not pending_reqs:
                st.info("현재 승인 대기 중인 증빙 자료가 없습니다.")
            else:
                for p in pending_reqs:
                    t_info = next((t for t in kpi_targets if str(t.get('id')) == str(p.get('kpi_id'))), {'kpi_name': '삭제된 지표', 'owner': '알수없음'})
                    d_info = next((d for d in kpi_details if str(d.get('id')) == str(p.get('detail_id'))), None)
                    t_name = f"[{t_info['owner']}] {t_info['kpi_name']}"
                    if d_info: t_name += f" ➔ 🔹{d_info['detail_name']}"
                    
                    with st.container(border=True):
                        st.markdown(f"**제출자:** {p['user_name']} | **지표:** {t_name} | **제출 기간/분류:** {p['period']}")
                        st.write(f"**증빙 내역:** {p['evidence']}")
                        ac1, ac2, _ = st.columns([1, 1, 6])
                        if ac1.button("✅ 승인", key=f"app_{p['id']}"):
                            supabase.table('kpi_submissions').update({"status": "승인"}).eq('id', p['id']).execute(); apply_changes()
                        if ac2.button("❌ 반려", key=f"rej_{p['id']}"):
                            supabase.table('kpi_submissions').update({"status": "반려"}).eq('id', p['id']).execute(); apply_changes()
        
        with mt2:
            with st.form("new_kpi_target_form", clear_on_submit=True):
                st.subheader("새로운 메인 KPI 지표 생성")
                sc1, sc2 = st.columns(2)
                t_name = sc1.text_input("KPI 지표명 (예: 법정 제출 누락률)")
                all_users = sorted(list(set([u.get('이름') for u in user_data if u.get('이름')])))
                t_owner = sc2.selectbox("적용 대상", ["공통"] + all_users)
                
                sc3_u, sc3_v, sc4, sc5 = st.columns([1, 1.5, 1, 1.5])
                t_unit = sc3_u.selectbox("목표 단위", ["건수", "요율(%)", "금액(원)"])
                if t_unit == "금액(원)": t_count = sc3_v.number_input("목표 수치", min_value=0, value=0, step=100000)
                elif t_unit == "요율(%)": t_count = sc3_v.number_input("목표 수치", min_value=0, value=100, step=5)
                else: t_count = sc3_v.number_input("목표 수치", min_value=0, value=14, step=1)
                
                t_weight = sc4.number_input("배점", value=15)
                t_cycle = sc5.text_input("측정 주기 (예: 분기, 월)")
                t_desc = st.text_area("산출식 및 배점 설명 (예: 누락 0건 15점)")
                
                if st.form_submit_button("➕ 메인 KPI 생성", type="primary") and t_name:
                    supabase.table('kpi_targets').insert({
                        "kpi_name": t_name, "owner": t_owner, "target_count": t_count, 
                        "unit": t_unit, "weight": t_weight, "cycle": t_cycle, "description": t_desc
                    }).execute(); apply_changes()
            
            st.divider()
            st.subheader("등록된 지표 및 담당자 상세 할당 관리")
            if not kpi_targets: st.info("등록된 지표가 없습니다.")
            for i, target in enumerate(kpi_targets):
                t_id = target.get('id')
                if str(st.session_state.get('edit_kpi_id')) == str(t_id):
                    with st.container(border=True):
                        ec1, ec2 = st.columns(2)
                        e_name = ec1.text_input("KPI명", value=target.get('kpi_name'), key=f"ekn_{t_id}_{i}")
                        cur_own = target.get('owner')
                        e_own = ec2.selectbox("적용 대상", ["공통"] + all_users, index=(["공통"] + all_users).index(cur_own) if cur_own in ["공통"] + all_users else 0, key=f"eko_{t_id}_{i}")
                        ec3_u, ec3_v, ec4, ec5 = st.columns([1, 1.5, 1, 1.5])
                        cur_unit = target.get('unit', '건수')
                        unit_opts = ["건수", "요율(%)", "금액(원)"]
                        e_unit = ec3_u.selectbox("목표 단위", unit_opts, index=unit_opts.index(cur_unit) if cur_unit in unit_opts else 0, key=f"eku_{t_id}_{i}")
                        
                        if e_unit == "금액(원)": e_cnt = ec3_v.number_input("목표 수치", value=int(target.get('target_count') or 0), key=f"ekc_{t_id}_{i}", step=100000)
                        elif e_unit == "요율(%)": e_cnt = ec3_v.number_input("목표 수치", value=int(target.get('target_count') or 0), key=f"ekc_{t_id}_{i}", step=5)
                        else: e_cnt = ec3_v.number_input("목표 수치", value=int(target.get('target_count') or 0), key=f"ekc_{t_id}_{i}", step=1)
                        
                        e_wgt = ec4.number_input("배점", value=int(target.get('weight') or 0), key=f"ekw_{t_id}_{i}")
                        e_cyc = ec5.text_input("주기", value=target.get('cycle') or '', key=f"eky_{t_id}_{i}")
                        e_desc = st.text_area("산출식", value=target.get('description') or '', key=f"ekd_{t_id}_{i}")
                        eb1, eb2, _ = st.columns([1, 1, 4])
                        if eb1.button("💾 저장", type="primary", key=f"esvk_{t_id}_{i}"):
                            supabase.table('kpi_targets').update({
                                "kpi_name": e_name, "owner": e_own, "target_count": e_cnt, "unit": e_unit,
                                "weight": e_wgt, "cycle": e_cyc, "description": e_desc
                            }).eq('id', t_id).execute()
                            st.session_state['edit_kpi_id'] = None; apply_changes()
                        if eb2.button("취소", key=f"ecank_{t_id}_{i}"):
                            st.session_state['edit_kpi_id'] = None; st.rerun()
                else:
                    t_str_format = format_target(target.get('target_count', 0), target.get('unit', '건수'))
                    with st.expander(f"[{target.get('owner')}] {target.get('kpi_name')} (목표: {t_str_format} / 배점: {target.get('weight')}점)"):
                        st.info(f"**📝 산출식 및 가이드:** {target.get('description', '설명글이 없습니다.')}")
                        st.markdown("**🔹 상세 업무(Sub-KPI) 할당 및 리스트**")
                        
                        details_for_this = [d for d in kpi_details if str(d.get('kpi_id')) == str(t_id)]
                        for d in details_for_this:
                            d_id = d.get('id')
                            if str(st.session_state.get('edit_detail_id')) == str(d_id):
                                with st.container(border=True):
                                    edc1, edc2, edc3 = st.columns([4, 2, 2])
                                    ed_name = edc1.text_input("상세 업무명", value=d.get('detail_name'), key=f"edn_{d_id}")
                                    ed_assig = edc2.selectbox("담당자", all_users, index=all_users.index(d.get('assignee')) if d.get('assignee') in all_users else 0, key=f"eda_{d_id}")
                                    ed_cycle = edc3.text_input("주기", value=d.get('cycle') or '', key=f"edc_{d_id}")
                                    ed_desc = st.text_area("상세 설명", value=d.get('description') or '', key=f"edd_{d_id}")
                                    edb1, edb2, _ = st.columns([1, 1, 4])
                                    if edb1.button("💾 상세 저장", type="primary", key=f"edsave_{d_id}"):
                                        supabase.table('kpi_details').update({"detail_name": ed_name, "assignee": ed_assig, "cycle": ed_cycle, "description": ed_desc}).eq('id', d_id).execute()
                                        st.session_state['edit_detail_id'] = None; apply_changes()
                                    if edb2.button("취소", key=f"edcan_{d_id}"):
                                        st.session_state['edit_detail_id'] = None; st.rerun()
                            else:
                                with st.container(border=True):
                                    dc1, dc2, dc3, dc4 = st.columns([4, 2, 0.8, 0.8])
                                    dc1.write(f"**{d.get('detail_name')}** (주기: {d.get('cycle', '미지정')})")
                                    dc1.caption(f"설명: {d.get('description', '')}")
                                    dc2.write(f"👤 담당: {d.get('assignee')}")
                                    if dc3.button("✏️ 수정", key=f"edit_det_{d_id}"):
                                        st.session_state['edit_detail_id'] = d_id; st.rerun()
                                    if dc4.button("🗑 삭제", key=f"del_det_{d_id}"):
                                        supabase.table('kpi_details').delete().eq('id', d_id).execute(); apply_changes()
                        
                        with st.form(key=f"add_det_{t_id}", clear_on_submit=True):
                            st.markdown("**새로운 상세 업무 할당**")
                            c_n, c_a, c_c = st.columns([4, 2, 2])
                            new_d_name = c_n.text_input("상세 업무명", placeholder="예: 1분기 건강검진안내", key=f"ndn_{t_id}")
                            new_d_assig = c_a.selectbox("담당자", all_users, key=f"nda_{t_id}")
                            new_d_cycle = c_c.text_input("주기", placeholder="예: 매월, 1분기", key=f"ndc_{t_id}")
                            new_d_desc = st.text_area("상세 설명 (가이드)", placeholder="담당자가 참고할 설명글을 적어주세요.", key=f"ndd_{t_id}")
                            if st.form_submit_button("상세 할당"):
                                if new_d_name:
                                    supabase.table('kpi_details').insert({
                                        "kpi_id": t_id, "detail_name": new_d_name, "assignee": new_d_assig,
                                        "cycle": new_d_cycle, "description": new_d_desc
                                    }).execute()
                                    apply_changes()
                        
                        st.markdown("---")
                        b1, b2, _ = st.columns([1.5, 1.5, 5])
                        if b1.button("✏ 메인 지표 수정", key=f"kedt_{t_id}_{i}"): st.session_state['edit_kpi_id'] = t_id; st.rerun()
                        if b2.button("🗑 메인 지표 완전삭제", key=f"kdel_{t_id}_{i}"): supabase.table('kpi_targets').delete().eq('id', t_id).execute(); apply_changes()

    else:
        st.header(f"📈 {target_user} KPI 달성 현황 및 증빙 제출")
        my_common = [t for t in kpi_targets if t.get('owner') == '공통']
        my_personal = [t for t in kpi_targets if t.get('owner') == target_user]
        
        st.subheader("🏆 현재 스코어 보드")
        if not my_common and not my_personal:
            st.info("할당된 KPI 지표가 없습니다.")
        else:
            all_my_targets = my_common + my_personal
            for t in all_my_targets:
                app, tot, pts, rate, mis = calculate_kpi_score(t, kpi_subs)
                t_unit = t.get('unit', '건수')
                t_format = format_target(tot, t_unit)
                
                with st.container(border=True):
                    st.markdown(f"### [{t['owner']}] {t['kpi_name']}")
                    st.info(f"**💡 산출식 및 설명:** {t.get('description', '설명글이 없습니다.')}")
                    
                    mc1, mc2, mc3 = st.columns(3)
                    mc1.write(f"- **🎯 목표:** {t_format}")
                    mc2.write(f"- **⏱️ 주기:** {t.get('cycle', '미지정')}")
                    mc3.write(f"- **⭐️ 배점:** {t.get('weight', 0)}점")
                    
                    if "누락" in t['kpi_name'] or "법정" in t['kpi_name']:
                        st.error(f"**📊 달성 현황:** 누락 {mis}건 (부서 총 {app}건 승인) ➔ **{pts}점 획득**")
                    else:
                        st.success(f"**📊 달성 현황:** 달성률 {round(rate, 1)}% (내 승인 {app}건) ➔ **{pts}점 획득**")

        st.divider()
        
        if not is_readonly:
            my_assigned_details = [d for d in kpi_details if d.get('assignee') == target_user]
            if my_assigned_details:
                with st.expander("💡 나에게 할당된 상세 지표 가이드 보기", expanded=False):
                    for md in my_assigned_details:
                        parent_t = next((t for t in kpi_targets if str(t['id']) == str(md['kpi_id'])), None)
                        p_name = parent_t['kpi_name'] if parent_t else "알수없음 지표"
                        st.markdown(f"**[{p_name}] ➔ {md.get('detail_name')}** (주기: {md.get('cycle', '미지정')})")
                        st.write(md.get('description') or '등록된 설명글이 없습니다.')
                        st.markdown("---")
            
            st.subheader("📤 증빙 자료 확인 요청")
            submit_options = []
            for pt in my_personal:
                submit_options.append({"type": "personal", "target": pt, "detail": None, "label": f"[개인] {pt['kpi_name']}"})
            
            for md in my_assigned_details:
                parent_t = next((t for t in kpi_targets if str(t['id']) == str(md['kpi_id'])), None)
                if parent_t:
                    submit_options.append({"type": "detail", "target": parent_t, "detail": md, "label": f"[공통] {parent_t['kpi_name']} ➔ {md['detail_name']}"})
            
            with st.form("kpi_submission_form", clear_on_submit=True):
                s1, s2 = st.columns(2)
                if submit_options:
                    sel_opt = s1.selectbox("할당된 지표 선택", submit_options, format_func=lambda x: x["label"])
                    sub_period = s2.text_input("실제 진행 분류/기간 (예: 1분기, 4월 15일 일계표)")
                    sub_evidence = st.text_area("증빙 내용 및 위치 (예: 그룹웨어 결재 번호 #1234)")
                    
                    if st.form_submit_button("확인 요청 전송", type="primary"):
                        d_id = sel_opt['detail']['id'] if sel_opt['type'] == 'detail' else None
                        supabase.table('kpi_submissions').insert({
                            "user_name": u_name, "kpi_id": sel_opt['target']['id'], "detail_id": d_id,
                            "period": sub_period, "evidence": sub_evidence, "status": "대기"
                        }).execute()
                        st.success("확인 요청이 마스터에게 전송되었습니다."); apply_changes()
                else:
                    st.warning("개인적으로 할당된 지표나 상세 업무가 없습니다.")
                    st.form_submit_button("제출 불가", disabled=True)

        st.divider()
        st.subheader("📜 나의 제출 내역")
        my_history = [s for s in kpi_subs if s.get('user_name') == target_user]
        if not my_history:
            st.info("제출된 증빙 내역이 없습니다.")
        else:
            for s in reversed(my_history):
                t_info = next((t for t in kpi_targets if str(t.get('id')) == str(s.get('kpi_id'))), None)
                d_info = next((d for d in kpi_details if str(d.get('id')) == str(s.get('detail_id'))), None)
                t_name = f"[{t_info['owner']}] {t_info['kpi_name']}" if t_info else "삭제된 지표"
                if d_info: t_name += f" ➔ 🔹{d_info['detail_name']}"
                
                status_color = "#1E88E5" if s['status'] == '대기' else "#43A047" if s['status'] == '승인' else "#E53935"
                with st.container(border=True):
                    st.markdown(f"**{t_name}** | 기간: {s['period']} | 상태: <span style='color:{status_color}; font-weight:bold;'>{s['status']}</span>", unsafe_allow_html=True)
                    st.write(f"- 증빙: {s['evidence']}")
                    if s['status'] == '대기' and not is_readonly:
                        if st.button("요청 취소", key=f"dels_{s['id']}"):
                            supabase.table('kpi_submissions').delete().eq('id', s['id']).execute(); apply_changes()

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
        for t in rep_daily:
            if bool(t.get('is_copied', False)): continue # 이슈 복사된 일일업무 제외
            prog = int(str(t.get('진행률') or '0')) if str(t.get('진행률') or '0').isdigit() else 0
            is_in_p = bool(t.get('진행중', False))
            
            if prog == 100: icon, prog_txt = "✓", "(완료)"
            elif prog == 0: icon, prog_txt = "□", f"({prog}%)"
            else: icon, prog_txt = "▶", f"({prog}%)"
            if is_in_p and prog < 100: prog_txt += " <span style='color:#E65100; font-weight:bold;'>(진행중)</span>"
            
            task_n = str(t.get('업무명') or '').replace(chr(10), '<br>')
            task_lines = task_n.split('<br>')
            if len(task_lines) > 1: styled_task_n = f"<b>{task_lines[0]}</b><br><span style='color:#777; font-size:0.9em;'>" + "<br>".join(task_lines[1:]) + "</span>"
            else: styled_task_n = f"<b>{task_n}</b>"
                
            if str(t.get('프로젝트연동') or 'FALSE').upper() == "TRUE":
                p_n = str(t.get('연결프로젝트') or '').split('::')[0]
                if p_n not in grouped_proj: grouped_proj[p_n] = {'tasks': []}
                grouped_proj[p_n]['tasks'].append(f"{icon} {styled_task_n} {prog_txt}")
            else: h_d_html += f"<li style='margin-bottom:8px;'>{icon} {styled_task_n} {prog_txt}</li>"
                
        for p_n, d in grouped_proj.items():
            p_sub_list = sub_dict.get(p_n, [])
            is_all_done = len(p_sub_list) > 0 and all(int(str(s.get('진행률') or '0')) == 100 for s in p_sub_list)
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
            
            is_all_done = all(int(str(s.get('진행률') or '0')) == 100 for s in valid_subs)
            done_badge = " <span style='color:#2e7d32; font-weight:bold;'>(완료)</span>" if is_all_done else ""
            h_p_html += f"<div style='margin-top:15px;'><h4 style='margin-bottom:5px;'>■ {pn}{done_badge}</h4><ul style='margin-top:0;'>"
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
        s_w_str, e_w_str = s_w.strftime("%Y-%m-%d"), e_w.strftime("%Y-%m-%d")
        
        st.markdown("**📝 업무 분류별 주간 고정 내용 입력** (엑셀 출력용)")
        st.info("입력된 내용은 이번 엑셀 파일에만 포함되며, 데이터베이스에 영구 저장되지 않습니다.")
        
        excel_sort_order = ["경영", "재무", "입찰", "조달", "지원", "R&D", "현장", "요청사항"]
        fixed_notes = {}
        with st.expander("고정 내용 입력 열기"):
            for cat in excel_sort_order:
                fixed_notes[cat] = st.text_area(f"[{cat}] 주간 고정 내용", key=f"fn_{cat}")
        
        daily_period = [d for d in all_daily if s_w_str <= str(d.get('날짜') or '') <= e_w_str and (target_user == "전체" or d.get('담당자') == target_user) and not bool(d.get('보고서제외', False)) and str(d.get('프로젝트연동') or 'FALSE').upper() != "TRUE" and not bool(d.get('is_copied', False))]
        rep_proj = [p for p in proj_data if (target_user == "전체" or p.get('담당자') == target_user) and not bool(p.get('보관함이동', False)) and not bool(p.get('보고서제외', False))]
        routines_period = [r for r in routine_data if (target_user == "전체" or r.get('담당자') == target_user)]
        current_issues = [b for b in shared_boards if b.get('week_start') == current_monday_str]
        
        excel_data = {cat: [] for cat in excel_sort_order}
        excel_data['기타'] = []
        
        def match_category(cat_str):
            for sort_cat in excel_sort_order:
                if sort_cat in cat_str: return sort_cat
            return '기타'
        
        for d in daily_period:
            matched_cat = match_category(d.get('분류') or '기타')
            t_n = str(d.get('업무명') or '').replace(chr(10), '<br>')
            prog = int(str(d.get('진행률') or '0'))
            excel_data[matched_cat].append({'type': 'daily', 'content': f"· {t_n}", 'prog': prog, 'next': ''})
            
        for r in routines_period:
            matched_cat = match_category(r.get('분류') or '기타')
            t_n = str(r.get('업무명') or '').replace(chr(10), '<br>')
            excel_data[matched_cat].append({'type': 'routine', 'content': f"<b>[루틴]</b> {t_n}", 'prog': 100, 'next': ''})
            
        for p in rep_proj:
            matched_cat = match_category(p.get('분류') or '기타')
            pn = p.get('프로젝트명') or ''
            valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue
            is_all_done = all(int(str(s.get('진행률') or '0')) == 100 for s in valid_subs)
            done_text = " (완료)" if is_all_done else ""
            ph, total_p = f"<b>[프로젝트] {pn}{done_text}</b><br>", 0
            for s in valid_subs:
                prog = int(str(s.get('진행률') or '0'))
                total_p += prog
                ph += f"- {str(s.get('세부업무명') or '').replace(chr(10), '<br>')} ({prog}%)<br>"
            avg_p = int(total_p / len(valid_subs))
            excel_data[matched_cat].append({'type': 'project', 'content': ph, 'prog': avg_p, 'next': ''})

        for issue in current_issues:
            b_type = issue.get('board_type')
            c_text = str(issue.get('content')).replace(chr(10), '<br>')
            status_text = f" ({issue.get('status')})"
            prog = 100 if issue.get('status') == '완료' else 50
            if b_type == '요청사항':
                excel_data['요청사항'].append({'type': 'issue', 'content': f"· <b>[타부서요청]</b> {c_text}{status_text}", 'prog': prog, 'next': ''})
            elif b_type == '세금계산서':
                excel_data['재무'].append({'type': 'issue', 'content': f"· <b>[세금계산서]</b> {c_text}{status_text}", 'prog': prog, 'next': ''})
            elif b_type in ['분납서', '보증보험']:
                excel_data['현장'].append({'type': 'issue', 'content': f"· <b>[{b_type}]</b> {c_text}{status_text}", 'prog': prog, 'next': ''})
        
        xls_hr = ""
        for cat in excel_sort_order + ['기타']:
            items = excel_data[cat]
            f_note = fixed_notes.get(cat, "").replace('\n', '<br>')
            
            if not items and not f_note: continue
            
            cat_html = f"<b>{cat}</b>"
            if f_note: cat_html += f"<br><br><span style='color:#0066cc; font-size:0.9em;'>[주간고정]<br>{f_note}</span>"
            
            content_html = ""
            if not items: content_html = "특이사항 없음"
            else:
                for it in items:
                    p_str = f" ({it['prog']}%)" if it['type'] in ['daily', 'issue'] else ""
                    content_html += f"<div style='margin-bottom:8px;'>{it['content']}{p_str}</div>"
            
            cat_prog = int(sum(it['prog'] for it in items) / len(items)) if items else 100
            xls_hr += f"<tr><td style='vertical-align: top; {font_css}'>{cat_html}</td><td style='vertical-align: top; {font_css}'>{content_html}</td><td style='text-align:center; vertical-align: top; {font_css}'><b>{cat_prog}%</b></td><td style='{font_css}'></td></tr>"
            
        th = f"<tr><th style='background:#e0f7fa; padding:8px; width:15%; {font_css}'>업무분류</th><th style='background:#e0f7fa; padding:8px; width:50%; {font_css}'>지난주진행업무 ({s_w_str} ~ {e_w_str})</th><th style='background:#e0f7fa; padding:8px; width:10%; {font_css}'>진행률</th><th style='background:#e0f7fa; padding:8px; width:25%; {font_css}'>금주예정업무</th></tr>"
        xls_html = f"<html><head><meta charset='utf-8'></head><body><h2 style='text-align:center; {font_css}'>◎경영지원부 주간 보고◎</h2><table border='1' style='border-collapse:collapse; width:100%; border: 1px solid #ccc;'>{th}{xls_hr}</table></body></html>"
        
        st.download_button("💾 주간 보고서 (Excel) 다운로드", xls_html.encode('utf-8-sig'), f"[주간보고]경영지원부_{s_w}_{e_w}.xls")
