import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑 V17", layout="wide")

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

# --- [보안 및 로그인] ---
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

# --- [데이터 할당 및 💡 정렬 처리] ---
all_daily, proj_data, sub_data, routine_data, kpi_config, user_data, cat_data = load_db_data()

# 프로젝트 정렬: 생성된 순서(id)대로 고정 (요청 4번)
proj_data = sorted(proj_data, key=lambda x: int(x.get('id', 0)))
sub_data = sorted(sub_data, key=lambda x: int(x.get('id', 0)))

cat_list = sorted(list(set([str(c.get('분류명', '')) for c in cat_data if pd.notna(c.get('분류명')) and str(c.get('분류명')).strip() != ""])))
if not cat_list: cat_list = ["경영관리", "재무업무", "기타"]

my_kpi_opts = [str(k.get('KPI명', '')) for k in kpi_config if pd.notna(k.get('KPI명')) and str(k.get('구분', '공통')).strip() in ['공통', u_name] and str(k.get('KPI명')).strip() != ""]

# --- [사이드바] ---
with st.sidebar:
    st.subheader(f"👤 {u_name}님 ({u_role})")
    if st.button("🔄 최신 데이터 불러오기", use_container_width=True, type="primary"): apply_changes()
    st.divider()
    if st.button("🚪 로그아웃", use_container_width=True): 
        st.session_state['logged_in'] = False
        st.rerun()

st.title("🚀 NOWSYSTEM 통합 업무 관리")
t_date = st.date_input("📅 업무 기준일 선택", datetime.date.today(), key="main_date")
t_str = t_date.strftime("%Y-%m-%d")

is_locked = False

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
    
    # 💡 폼 초기화 적용 (Enter 키 지원 및 입력 후 공백화 - 요청 2, 3번)
    task_type = st.radio("업무 종류 선택", ["일반/데일리 업무", "프로젝트 연동 업무"], horizontal=True)
    if task_type == "일반/데일리 업무":
        with st.form("add_daily_normal_form", clear_on_submit=True):
            my_routines = [r.get('업무명') for r in routine_data if r.get('담당자') == u_name]
            sel_opt = st.selectbox("업무명", ["✏️ 직접 입력"] + my_routines)
            n_task = st.text_input("새 업무명 (직접 입력)") 
                
            c1, c2 = st.columns(2)
            n_cat = c1.selectbox("분류", cat_list) 
            n_kpi = c2.selectbox("연관 KPI", dropdown_opts + ["기타"])
            
            if st.form_submit_button("업무 추가", type="primary"):
                final_task = n_task if sel_opt == "✏️ 직접 입력" else sel_opt
                if final_task:
                    supabase.table('daily').insert({"날짜": t_str, "업무명": final_task, "진행률": 0, "프로젝트연동": "FALSE", "분류": n_cat, "연결프로젝트": "", "KPI": n_kpi, "담당자": u_name, "보고서제외": False}).execute()
                    apply_changes()
    else: 
        with st.form("add_daily_proj_form", clear_on_submit=True):
            my_projs = [p.get('프로젝트명') for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name) and not (str(p.get('보관함이동')).upper() == "TRUE" or p.get('보관함이동') == True)]
            sel_p = st.selectbox("진행 중인 프로젝트", my_projs) if my_projs else None
            my_subs = [s.get('세부업무명') for s in sub_data if s.get('프로젝트명') == sel_p] if sel_p else []
            sel_s = st.selectbox("프로젝트 하위 세부업무", my_subs) if my_subs else None
            
            if st.form_submit_button("프로젝트 업무 당겨오기", type="primary"):
                if sel_p and sel_s:
                    p_info = next((p for p in proj_data if p.get('프로젝트명') == sel_p), {})
                    supabase.table('daily').insert({"날짜": t_str, "업무명": sel_s, "진행률": 0, "프로젝트연동": "TRUE", "분류": p_info.get('분류', '프로젝트'), "연결프로젝트": f"{sel_p}::{sel_s}", "KPI": p_info.get('KPI', '기타'), "담당자": u_name, "보고서제외": False}).execute()
                    apply_changes()

    st.divider()
    
    # 일과 목록
    for i, row in enumerate(filtered_daily):
        r_id = row.get('id', f"temp_d_{i}") 
        col1, col2, col3, col4 = st.columns([4, 4, 1, 1])
        disp_name = str(row.get('업무명','')).replace('\n', '<br>')
        badge = f" <small style='color:blue;'>[{row.get('담당자','')}]</small>" if u_role == "마스터" else ""
        is_proj = str(row.get('프로젝트연동', 'FALSE')).upper() == "TRUE"
        
        if is_proj: col1.markdown(f"**[{row.get('분류', '프로젝트')}]** <span style='color:#555;'>{str(row.get('연결프로젝트', '')).replace('::', ' > ')}</span>{badge}", unsafe_allow_html=True)
        else: col1.markdown(f"**[{row.get('분류', '기타')}]** {disp_name}{badge}", unsafe_allow_html=True)
            
        cur_p = int(row.get('진행률', 0) if str(row.get('진행률', 0)).isdigit() else 0)
        new_p = col2.slider("진행", 0, 100, cur_p, 10, key=f"d_sld_{r_id}", label_visibility="collapsed")
        
        if new_p != cur_p:
            if isinstance(r_id, int) or str(r_id).isdigit():
                supabase.table('daily').update({"진행률": new_p}).eq('id', r_id).execute()
                if is_proj:
                    p_info = str(row.get('연결프로젝트', ''))
                    if "::" in p_info:
                        p_n, s_n = p_info.split("::", 1)
                        s_id_match = next((s.get('id') for s in sub_data if s.get('프로젝트명') == p_n and s.get('세부업무명') == s_n), None)
                        if s_id_match: supabase.table('sub_tasks').update({"진행률": new_p}).eq('id', s_id_match).execute()
                apply_changes()
        
        # 💡 보고서 제외 토글 (요청 7번)
        is_ex = bool(row.get('보고서제외', False))
        new_ex = col3.checkbox("🚫제외", value=is_ex, key=f"d_ex_{r_id}", help="체크 시 보고서에 출력되지 않습니다.")
        if new_ex != is_ex:
            supabase.table('daily').update({"보고서제외": new_ex}).eq('id', r_id).execute()
            apply_changes()

        if col4.button("🗑️", key=f"d_del_{r_id}"):
            if isinstance(r_id, int) or str(r_id).isdigit():
                supabase.table('daily').delete().eq('id', r_id).execute()
                apply_changes()

    st.write("---")
    # 💡 고정 업무 가시성 확대 (요청 6번)
    st.subheader("📌 나의 데일리 고정 업무 (루틴)")
    c_r1, c_r2 = st.columns([1, 1])
    with c_r1:
        with st.form("add_routine_form", clear_on_submit=True):
            r_task = st.text_input("새 데일리 업무명 등록")
            r_cat = st.selectbox("분류", cat_list) 
            r_kpi = st.selectbox("연관 KPI", dropdown_opts + ["기타"])
            if st.form_submit_button("루틴 목록에 추가"):
                if r_task:
                    supabase.table('routines').insert({"업무명": r_task, "분류": r_cat, "KPI": r_kpi, "담당자": u_name}).execute()
                    apply_changes()
    with c_r2:
        for i, r in enumerate(routine_data):
            if r.get('담당자') == u_name:
                r_id = r.get('id', f"temp_r_{i}")
                rr1, rr2 = st.columns([4, 1])
                rr1.write(f"· [{r.get('분류')}] {r.get('업무명')}")
                if rr2.button("삭제", key=f"rdel_{r_id}"):
                    if isinstance(r_id, int) or str(r_id).isdigit():
                        supabase.table('routines').delete().eq('id', r_id).execute()
                        apply_changes()

# ==========================================
# 탭 2: 프로젝트 관리 
# ==========================================
with tabs[1]:
    st.header("📁 프로젝트 현황")
    with st.expander("✨ 신규 프로젝트 등록", expanded=False):
        # 💡 폼 초기화 적용 (요청 3번)
        with st.form("new_proj_form", clear_on_submit=True):
            pc1, pc2 = st.columns(2)
            p_name = pc1.text_input("프로젝트명")
            p_cat = pc2.selectbox("분류", cat_list) 
            p_start = pc1.date_input("시작일")
            p_kpi = pc2.selectbox("연관 KPI", dropdown_opts + ["기타"])
            
            if st.form_submit_button("프로젝트 저장", type="primary"):
                if p_name:
                    supabase.table('projects').insert({"프로젝트명": p_name, "시작일": str(p_start), "분류": p_cat, "KPI": p_kpi, "담당자": u_name}).execute()
                    st.success(f"[{p_name}] 저장 완료!")
                    apply_changes()

    for i, p in enumerate(proj_data):
        r_id = p.get('id', f"temp_p_{i}")
        is_archived_bool = True if str(p.get("보관함이동")).upper() == "TRUE" or p.get("보관함이동") == True else False
        if (u_role != "마스터" and p.get('담당자') != u_name) or is_archived_bool: continue
            
        pn = p.get("프로젝트명", "")
        owner = f" ({p.get('담당자','')})" if u_role == "마스터" else ""
        my_s_list = sub_dict.get(pn, [])
        
        # 💡 프로젝트 전체 진행률 계산 (요청 1번)
        total_p = sum(int(s.get('진행률', 0) if str(s.get('진행률',0)).isdigit() else 0) for s in my_s_list)
        avg_p = int(total_p / len(my_s_list)) if len(my_s_list) > 0 else 0
        
        with st.expander(f"📂 {pn} [{p.get('분류')}] - 📊 전체 진행률: {avg_p}% {owner}"):
            # 💡 엔터키 및 폼 초기화 적용 (요청 2, 3번)
            with st.form(key=f"sub_form_{r_id}", clear_on_submit=True):
                sc1, sc2 = st.columns([4,1])
                new_sub = sc1.text_input("세부 업무명 (입력 후 Enter)")
                if sc2.form_submit_button("하위 업무 추가"):
                    if new_sub:
                        supabase.table('sub_tasks').insert({"프로젝트명": pn, "세부업무명": new_sub, "진행률": 0, "담당자": u_name, "보고서제외": False}).execute()
                        apply_changes()
                        
            for j, s in enumerate(my_s_list):
                s_id = s.get('id', f"temp_s_{j}")
                sl1, sl2, sl3, sl4, sl5 = st.columns([4, 3, 1, 1, 1])
                sl1.markdown(f"· {str(s.get('세부업무명','')).replace('\n','<br>')}", unsafe_allow_html=True)
                
                cur_sp = int(s.get('진행률',0) if str(s.get('진행률',0)).isdigit() else 0)
                sp = sl2.slider("진행", 0, 100, cur_sp, 10, key=f"s_sld_{s_id}", label_visibility="collapsed")
                if sp != cur_sp:
                    if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').update({"진행률": sp}).eq('id', s_id).execute()
                        apply_changes()
                
                # 💡 완료 버튼 처리 (요청 1번)
                if sl3.button("✅완료", key=f"sdone_{s_id}"):
                     if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').update({"진행률": 100}).eq('id', s_id).execute()
                        apply_changes()

                # 💡 보고서 제외 (요청 7번)
                s_ex = bool(s.get('보고서제외', False))
                s_new_ex = sl4.checkbox("🚫제외", value=s_ex, key=f"s_ex_{s_id}")
                if s_new_ex != s_ex:
                    supabase.table('sub_tasks').update({"보고서제외": s_new_ex}).eq('id', s_id).execute()
                    apply_changes()
                
                if sl5.button("🗑️", key=f"sdel_{s_id}"):
                    if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').delete().eq('id', s_id).execute()
                        apply_changes()
            
            st.write("---")
            ac1, ac2, ac3 = st.columns([2,1,1])
            if ac2.button("📦 보관함 이동", key=f"arc_{r_id}"):
                if isinstance(r_id, int) or str(r_id).isdigit():
                    supabase.table('projects').update({"보관함이동": True}).eq('id', r_id).execute()
                    apply_changes()
            if ac3.button("🗑️ 프로젝트 삭제", key=f"pdel_{r_id}"):
                if isinstance(r_id, int) or str(r_id).isdigit():
                    supabase.table('projects').delete().eq('id', r_id).execute()
                    apply_changes()

# ==========================================
# 탭 3 & 4: 설정 탭 
# ==========================================
if u_role == "마스터":
    with tab_set1:
        st.header("⚙️ 설정 1 (사내 계정 및 KPI 관리)")
        # (이전과 동일 로직)
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
    st.header("📈 전사 통합 KPI" if u_role == "마스터" else f"📈 {u_name}님 전용 KPI 현황")
    stats = {}
    for d in all_daily:
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
# 탭 5: 데이터/보고서 (💡 전체 내용 편집 기능 추가)
# ==========================================
with tab_rep:
    st.header("📊 데이터 및 보고서 관리")
    
    # 💡 [요청 5번] 입력된 모든 업무 내용 일괄 수정 모드 추가
    with st.expander("🛠️ 등록된 전체 업무 일괄 수정 (오타/내용 변경)"):
        st.info("표 안의 글씨를 더블클릭하여 바로 수정하고, 맨 아래 [일괄 수정 저장] 버튼을 누르세요.")
        t1, t2, t3 = st.tabs(["📝 일일 업무", "📁 프로젝트", "📋 하위 세부업무"])
        
        with t1:
            e_d_df = st.data_editor(pd.DataFrame(all_daily), key="ed_d", use_container_width=True)
        with t2:
            e_p_df = st.data_editor(pd.DataFrame(proj_data), key="ed_p", use_container_width=True)
        with t3:
            e_s_df = st.data_editor(pd.DataFrame(sub_data), key="ed_s", use_container_width=True)
            
        if st.button("💾 전체 데이터 일괄 수정 저장", type="primary"):
            for r in e_d_df.to_dict('records'): supabase.table('daily').upsert(r).execute()
            for r in e_p_df.to_dict('records'): supabase.table('projects').upsert(r).execute()
            for r in e_s_df.to_dict('records'): supabase.table('sub_tasks').upsert(r).execute()
            st.success("모든 수정 사항이 DB에 저장되었습니다!")
            apply_changes()

    st.divider()
    
    st.subheader("🖨️ 맞춤형 보고서 출력")
    r_type = st.radio("보고서 종류 선택", ["일일(HTML)", "기간별(Excel)"], horizontal=True)
    if r_type == "일일(HTML)":
        r_d = st.date_input("보고 날짜", t_date, key="r_date_p")
        r_s = r_d.strftime("%Y-%m-%d")
        
        # 💡 [요청 7번] 보고서 제외(True)인 항목은 보고서 배열에서 아예 뺍니다.
        rep_daily = [d for d in all_daily if str(d.get('날짜')) == r_s and (u_role == "마스터" or d.get('담당자') == u_name) and not bool(d.get('보고서제외', False))]
        rep_proj = [p for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name)]
        
        h_d_html = "".join([f"<li style='margin-bottom:8px;'><b>[{t.get('분류','기타')}]</b> {'(완료)' if int(t.get('진행률',0))==100 else f'({t.get('진행률',0)}%)'} {str(t.get('업무명','')).replace(chr(10), '<br>')} <span style='color:#1976D2; font-size:0.9em;'>[{t.get('담당자','')}]</span></li>" for t in rep_daily])
        
        h_p_html = ""
        for p in rep_proj:
            if bool(p.get('보관함이동', False)) or str(p.get('보관함이동')).upper() == "TRUE": continue
            pn = p.get('프로젝트명', '')
            st_txt = "(완료)" if str(p.get('완료여부')).upper() == "TRUE" else ""
            
            # 하위 업무 중 '보고서 제외'가 아닌 것만 출력
            valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue # 표시할 하위업무가 없으면 프로젝트 타이틀도 패스
            
            h_p_html += f"<div style='margin-top:15px;'><h4 style='margin-bottom:5px;'>■ [{p.get('분류','기타')}] {pn} <span style='color:#555; font-size:0.9em;'>({p.get('담당자','')})</span> <span style='color:#2e7d32;'>{st_txt}</span></h4><ul style='margin-top:0;'>"
            for s in valid_subs:
                prog = int(s.get('진행률',0))
                icon = "✓" if prog==100 else ("▶" if prog>0 else "□")
                h_p_html += f"<li style='margin-bottom:5px;'>{icon} {str(s.get('세부업무명','')).replace(chr(10), '<br>')} ({prog}%) <span style='color:#1976D2; font-size:0.9em;'>[{s.get('담당자','')}]</span></li>"
            h_p_html += "</ul></div>"
            
        title_txt = "전사 업무 보고서" if u_role == "마스터" else f"{u_name} 업무 보고서"
        full_html = f"<html><body style='font-family:sans-serif;'><h2>[{r_s}] {title_txt}</h2><h3>■ 일일 업무</h3><ul style='line-height:1.5;'>{h_d_html}</ul><hr><h3>■ 프로젝트 현황</h3>{h_p_html}</body></html>"
        st.components.v1.html(full_html, height=300, scrolling=True)
        c_btn1, c_btn2 = st.columns([1, 3])
        with c_btn1: st.download_button("📥 HTML 다운로드", full_html.encode('utf-8'), f"Report_{r_s}.html")

    elif r_type == "기간별(Excel)":
        c_ds1, c_ds2 = st.columns(2)
        s_w = c_ds1.date_input("시작일", t_date - datetime.timedelta(days=7), key="ws")
        e_w = c_ds2.date_input("종료일", t_date, key="we")
        rep_proj = [p for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name)]
        xls_hr = ""
        for p in rep_proj:
            if bool(p.get('보관함이동', False)) or str(p.get('보관함이동')).upper() == "TRUE": continue
            pn = p.get('프로젝트명', '')
            cat = p.get('분류', '기타')
            ph = f"<b>[{pn}]</b><br>"
            
            valid_subs = [s for s in sub_dict.get(pn, []) if not bool(s.get('보고서제외', False))]
            if not valid_subs: continue
            
            total_p = 0
            for s in valid_subs:
                prog = int(s.get('진행률',0))
                total_p += prog
                owner_s = f" [{s.get('담당자','')}]" if u_role == "마스터" and s.get('담당자') else ""
                ph += f"- {str(s.get('세부업무명','')).replace(chr(10), '<br>')} ({prog}%){owner_s}<br>"
            avg_p = int(total_p / len(valid_subs)) if valid_subs else 0
            xls_hr += f"<tr><td><b>{cat}</b></td><td>{ph}</td><td style='text-align:center;'><b>{avg_p}%</b></td><td></td></tr>"
            
        th = "<tr><th style='background:#e0f7fa; padding:8px;'>분류</th><th style='background:#e0f7fa; padding:8px;'>업무내역</th><th style='background:#e0f7fa; padding:8px;'>진행률</th><th style='background:#e0f7fa; padding:8px;'>예정사항</th></tr>"
        title_txt = "전사" if u_role == "마스터" else u_name
        xls_html = f"<html><meta charset='utf-8'><style>td {{border: 1px solid #ccc; padding: 8px; vertical-align: top; line-height:1.5;}}</style><body><h2>{title_txt} 업무 보고서 ({s_w} ~ {e_w})</h2><table style='border-collapse:collapse; width:100%; border: 1px solid #ccc;'>{th}{xls_hr}</table></body></html>"
        st.download_button("💾 Excel 다운로드 (.xls)", xls_html.encode('utf-8-sig'), f"Report_{s_w}_{e_w}.xls")
