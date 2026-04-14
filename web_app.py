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
        st.warning(f"일부 데이터를 불러오지 못했습니다: {e}")
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

# --- [데이터 할당] ---
all_daily, proj_data, sub_data, routine_data, kpi_config, user_data, cat_data = load_db_data()

cat_list = sorted(list(set([str(c.get('분류명', '')) for c in cat_data if pd.notna(c.get('분류명')) and str(c.get('분류명')).strip() != ""])))
if not cat_list:
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
                        p_info = next((p for p in proj_data if p.get('프로젝트명') == sel_p), {})
                        p_cat = p_info.get('분류', '프로젝트')
                        p_kpi = p_info.get('KPI', '기타')
                        supabase.table('daily').insert({"날짜": t_str, "업무명": sel_s, "진행률": 0, "프로젝트연동": "TRUE", "분류": p_cat, "연결프로젝트": f"{sel_p}::{sel_s}", "KPI": p_kpi, "담당자": u_name}).execute()
                        apply_changes()

    st.divider()
    
    for i, row in enumerate(filtered_daily):
        r_id = row.get('id', f"temp_d_{i}") 
        col1, col2, col3 = st.columns([4, 5, 1])
        disp_name = str(row.get('업무명','')).replace('\n', '<br>')
        badge = f" <small style='color:blue;'>[{row.get('담당자','')}]</small>" if u_role == "마스터" else ""
        is_proj = str(row.get('프로젝트연동', 'FALSE')).upper() == "TRUE"
        
        if is_proj:
            p_info = str(row.get('연결프로젝트', ''))
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

    with st.expander("⚙️ 나의 데일리 업무(루틴) 목록 관리", expanded=False):
        with st.form("add_routine_form"):
            rc1, rc2, rc3 = st.columns([2,1,1])
            r_task = rc1.text_input("새 데일리 업무명")
            r_cat = rc2.selectbox("분류", cat_list) 
            r_kpi = rc3.selectbox("연관 KPI", dropdown_opts + ["기타"])
            if st.form_submit_button("목록에 추가하기"):
                if r_task:
                    supabase.table('routines').insert({"업무명": r_task, "분류": r_cat, "KPI": r_kpi, "담당자": u_name}).execute()
                    apply_changes()
        for i, r in enumerate(routine_data):
            if r.get('담당자') == u_name:
                r_id = r.get('id', f"temp_r_{i}")
                rc1, rc2 = st.columns([4, 1])
                rc1.write(f"· **[{r.get('분류')}]** {r.get('업무명')}")
                if rc2.button("목록에서 삭제", key=f"rdel_{r_id}"):
                    if isinstance(r_id, int) or str(r_id).isdigit():
                        supabase.table('routines').delete().eq('id', r_id).execute()
                        apply_changes()

# ==========================================
# 탭 2: 프로젝트 관리 
# ==========================================
with tabs[1]:
    st.header("📁 프로젝트 현황")
    with st.expander("✨ 신규 프로젝트 등록", expanded=False):
        pc1, pc2 = st.columns(2)
        p_name = pc1.text_input("프로젝트명", disabled=is_locked)
        p_cat = pc2.selectbox("분류", cat_list, key="p_c", disabled=is_locked) 
        p_start = pc1.date_input("시작일", key="p_s", disabled=is_locked)
        p_kpi = pc2.selectbox("연관 KPI", dropdown_opts + ["기타"], key="p_k", disabled=is_locked)
        if st.button("프로젝트 저장", disabled=is_locked):
            if p_name:
                supabase.table('projects').insert({"프로젝트명": p_name, "시작일": str(p_start), "완료일": "", "완료여부": "FALSE", "보류여부": "FALSE", "보관함이동": "FALSE", "비고": "", "분류": p_cat, "KPI": p_kpi, "담당자": u_name}).execute()
                apply_changes()

    for i, p in enumerate(proj_data):
        r_id = p.get('id', f"temp_p_{i}")
        if (u_role != "마스터" and p.get('담당자') != u_name) or str(p.get("보관함이동", "FALSE")).upper() == "TRUE": continue
            
        pn = p.get("프로젝트명", "")
        owner = f" ({p.get('담당자','')})" if u_role == "마스터" else ""
        
        with st.expander(f"📂 {pn} [{p.get('분류')}]{owner}"):
            with st.form(key=f"sub_form_{r_id}"):
                sc1, sc2 = st.columns([3,1])
                new_sub = sc1.text_input("세부 업무명")
                if sc2.form_submit_button("하위 업무 추가", disabled=is_locked):
                    if new_sub:
                        supabase.table('sub_tasks').insert({"프로젝트명": pn, "세부업무명": new_sub, "진행률": 0, "사용여부": "TRUE", "담당자": u_name}).execute()
                        apply_changes()
            for j, s in enumerate(sub_dict.get(pn, [])):
                s_id = s.get('id', f"temp_s_{j}")
                sl1, sl2 = st.columns([6, 4])
                sl1.markdown(f"· {str(s.get('세부업무명','')).replace('\n','<br>')}", unsafe_allow_html=True)
                cur_sp = int(s.get('진행률',0) if str(s.get('진행률',0)).isdigit() else 0)
                sp = sl2.slider("진행", 0, 100, cur_sp, 10, key=f"s_sld_{s_id}", disabled=is_locked, label_visibility="collapsed")
                if not is_locked and sp != cur_sp:
                    if isinstance(s_id, int) or str(s_id).isdigit():
                        supabase.table('sub_tasks').update({"진행률": sp}).eq('id', s_id).execute()
                        st.toast(f"✅ 진행률 {sp}% 저장 완료!")
                        apply_changes()
            
            st.write("---")
            ac1, ac2, ac3 = st.columns([2,1,1])
            if ac2.button("📦 보관함 이동", key=f"arc_{r_id}", disabled=is_locked):
                if isinstance(r_id, int) or str(r_id).isdigit():
                    supabase.table('projects').update({"보관함이동": "TRUE"}).eq('id', r_id).execute()
                    apply_changes()
            if ac3.button("🗑️ 프로젝트 삭제", key=f"pdel_{r_id}", disabled=is_locked):
                if isinstance(r_id, int) or str(r_id).isdigit():
                    supabase.table('projects').delete().eq('id', r_id).execute()
                    apply_changes()

# ==========================================
# 탭 3: 마스터 전용 설정 1 (KPI & 계정)
# ==========================================
if u_role == "마스터":
    with tab_set1:
        st.header("⚙️ 설정 1 (사내 계정 및 KPI 관리)")
        c1, c2 = st.columns(2)
        
        with c1:
            st.subheader("👥 사내 계정 목록")
            u_df = pd.DataFrame(user_data)
            e_u_df = st.data_editor(u_df, num_rows="dynamic", use_container_width=True)
            if st.button("계정 정보 저장"):
                orig_u_ids = set(u_df['id'].dropna()) if 'id' in u_df.columns else set()
                new_u_ids = set(e_u_df['id'].dropna()) if 'id' in e_u_df.columns else set()
                for did in orig_u_ids - new_u_ids: supabase.table('users').delete().eq('id', did).execute()
                for r in e_u_df.to_dict('records'):
                    if pd.isna(r.get('id')): r.pop('id', None) 
                    supabase.table('users').upsert(r).execute()
                st.success("계정 정보 저장 완료!")
                apply_changes()
        
        with c2:
            st.subheader("🎯 개인별 & 공통 KPI 지표")
            k_df = pd.DataFrame(kpi_config)
            if 'KPI명' not in k_df.columns: k_df['KPI명'] = ""
            if '구분' not in k_df.columns: k_df['구분'] = '공통'
            disp_cols = [col for col in k_df.columns if col not in ['분류명']]
            k_df = k_df[disp_cols]
            
            e_k_df = st.data_editor(k_df, num_rows="dynamic", use_container_width=True)
            if st.button("KPI 지표 저장"):
                orig_k_ids = set(k_df['id'].dropna()) if 'id' in k_df.columns else set()
                new_k_ids = set(e_k_df['id'].dropna()) if 'id' in e_k_df.columns else set()
                for did in orig_k_ids - new_k_ids: supabase.table('settings').delete().eq('id', did).execute()
                for r in e_k_df.to_dict('records'):
                    if pd.isna(r.get('id')): r.pop('id', None)
                    if r.get('KPI명'): supabase.table('settings').upsert(r).execute()
                st.success("KPI 지표 저장 완료!")
                apply_changes()

# ==========================================
# 탭 4: 마스터 전용 설정 2 (업무 분류)
# ==========================================
if u_role == "마스터":
    with tab_set2:
        st.header("⚙️ 설정 2 (업무 분류 전용 관리)")
        st.info("💡 카테고리를 독립된 DB 테이블에서 안전하게 관리합니다. 이곳에서 변경한 내용은 즉시 전사 시스템에 반영됩니다.")
        
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
            st.success("새로운 업무 분류가 안전하게 저장되었습니다!")
            apply_changes()

# ==========================================
# KPI 현황 탭 & 보고서 탭
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
            
    if not stats: st.info("현재 표시할 KPI 실적이 없습니다.")
    else:
        for k_name, data in stats.items():
            avg = int(data["sum"] / data["count"]) if data["count"] > 0 else 0
            st.write(f"**{k_name}** (총 {data['count']}건)")
            st.progress(avg / 100, text=f"평균 달성률: {avg}%")

with tab_rep:
    st.header("📊 데이터 및 보고서 관리")
    with st.expander("📥 과거 업무 일괄 업로드 (CSV)"):
        temp_df = pd.DataFrame(columns=["날짜", "업무명", "진행률", "프로젝트연동", "분류", "연결프로젝트", "KPI", "담당자"])
        st.download_button("📄 업로드 양식 받기", temp_df.to_csv(index=False).encode('utf-8-sig'), "template.csv")
        up_file = st.file_uploader("CSV 파일 선택", type="csv")
        if up_file and st.button("시트로 일괄 전송"):
            df_up = pd.read_csv(up_file).fillna("")
            if u_role != "마스터": df_up['담당자'] = u_name 
            new_rows = df_up.to_dict('records')
            for row in new_rows: supabase.table('daily').insert(row).execute()
            st.success("업로드 완료!")
            apply_changes()
    st.divider()
    
    st.subheader("🖨️ 맞춤형 보고서 출력")
    r_type = st.radio("보고서 종류 선택", ["일일(HTML)", "기간별(Excel)"], horizontal=True)
    if r_type == "일일(HTML)":
        r_d = st.date_input("보고 날짜", t_date, key="r_date_p")
        r_s = r_d.strftime("%Y-%m-%d")
        rep_daily = [d for d in all_daily if str(d.get('날짜')) == r_s and (u_role == "마스터" or d.get('담당자') == u_name)]
        rep_proj = [p for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name)]
        
        h_d_html = "".join([f"<li style='margin-bottom:8px;'><b>[{t.get('분류','기타')}]</b> {'(완료)' if int(t.get('진행률',0))==100 else f'({t.get('진행률',0)}%)'} {str(t.get('업무명','')).replace(chr(10), '<br>')} <span style='color:#1976D2; font-size:0.9em;'>[{t.get('담당자','')}]</span></li>" for t in rep_daily])
        
        h_p_html = ""
        for p in rep_proj:
            if str(p.get("보관함이동", "FALSE")).upper() == "TRUE": continue
            pn = p.get('프로젝트명', '')
            st_txt = "(완료)" if str(p.get('완료여부','FALSE')).upper() == "TRUE" else ""
            h_p_html += f"<div style='margin-top:15px;'><h4 style='margin-bottom:5px;'>■ [{p.get('분류','기타')}] {pn} <span style='color:#555; font-size:0.9em;'>({p.get('담당자','')})</span> <span style='color:#2e7d32;'>{st_txt}</span></h4><ul style='margin-top:0;'>"
            for s in sub_dict.get(pn, []):
                prog = int(s.get('진행률',0))
                icon = "✓" if prog==100 else ("▶" if prog>0 else "□")
                h_p_html += f"<li style='margin-bottom:5px;'>{icon} {str(s.get('세부업무명','')).replace(chr(10), '<br>')} ({prog}%) <span style='color:#1976D2; font-size:0.9em;'>[{s.get('담당자','')}]</span></li>"
            h_p_html += "</ul></div>"
            
        title_txt = "전사 업무 보고서" if u_role == "마스터" else f"{u_name} 업무 보고서"
        full_html = f"<html><body style='font-family:sans-serif;'><h2>[{r_s}] {title_txt}</h2><h3>■ 일일 업무</h3><ul style='line-height:1.5;'>{h_d_html}</ul><hr><h3>■ 프로젝트 현황</h3>{h_p_html}</body></html>"
        st.components.v1.html(full_html, height=300, scrolling=True)
        c_btn1, c_btn2 = st.columns([1, 3])
        with c_btn1: st.download_button("📥 HTML 다운로드", full_html.encode('utf-8'), f"Report_{r_s}.html")
        with st.expander("📋 HTML 코드 복사"): st.code(full_html, language="html")

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
                owner_s = f" [{s.get('담당자','')}]" if u_role == "마스터" and s.get('담당자') else ""
                ph += f"- {str(s.get('세부업무명','')).replace(chr(10), '<br>')} ({prog}%){owner_s}<br>"
            avg_p = int(total_p / len(my_s)) if my_s else 0
            xls_hr += f"<tr><td><b>{cat}</b></td><td>{ph}</td><td style='text-align:center;'><b>{avg_p}%</b></td><td></td></tr>"
            
        th = "<tr><th style='background:#e0f7fa; padding:8px;'>분류</th><th style='background:#e0f7fa; padding:8px;'>업무내역</th><th style='background:#e0f7fa; padding:8px;'>진행률</th><th style='background:#e0f7fa; padding:8px;'>예정사항</th></tr>"
        title_txt = "전사" if u_role == "마스터" else u_name
        xls_html = f"<html><meta charset='utf-8'><style>td {{border: 1px solid #ccc; padding: 8px; vertical-align: top; line-height:1.5;}}</style><body><h2>{title_txt} 업무 보고서 ({s_w} ~ {e_w})</h2><table style='border-collapse:collapse; width:100%; border: 1px solid #ccc;'>{th}{xls_hr}</table></body></html>"
        st.download_button("💾 Excel 다운로드 (.xls)", xls_html.encode('utf-8-sig'), f"Report_{s_w}_{e_w}.xls")
