import streamlit as st
import gspread
import datetime
import pandas as pd
import json

# 1. 웹페이지 설정
st.set_page_config(page_title="NOWSYSTEM 관제탑", layout="wide")

# 2. 구글 시트 연결
@st.cache_resource
def get_gspread_client():
    try:
        creds_json = st.secrets["google_credentials"]
        credentials = json.loads(creds_json)
        return gspread.service_account_from_dict(credentials)
    except:
        return gspread.service_account(filename="now_secret.json")

try:
    gc = get_gspread_client()
    sh = gc.open("업무관리_DB")
    w_d = sh.worksheet("일일업무")
    w_p = sh.worksheet("프로젝트")
    w_s = sh.worksheet("세부업무")
    w_st = sh.worksheet("설정")
    w_u = sh.worksheet("계정관리")
    
    sheet_list = [s.title for s in sh.worksheets()]
    if "루틴업무" not in sheet_list:
        w_r = sh.add_worksheet(title="루틴업무", rows="100", cols="10")
        w_r.append_row(["업무명", "분류", "KPI", "담당자"])
    else:
        w_r = sh.worksheet("루틴업무")
except Exception as e:
    st.error(f"구글 시트 연결 실패: {e}")
    st.stop()

# 💡 고유 행 번호(_r)를 부여하여 슬라이더 실종(오류) 완벽 해결
def parse_batch_data(sheet_values):
    if not sheet_values or len(sheet_values) < 2: return []
    headers = [str(h).strip() for h in sheet_values[0]]
    records = []
    for idx, row in enumerate(sheet_values[1:]):
        padded_row = row + [''] * (len(headers) - len(row))
        record = dict(zip(headers, padded_row[:len(headers)]))
        record['_r'] = idx + 2 # 구글 시트의 실제 행 번호 (1번은 제목)
        records.append(record)
    return records

@st.cache_data(ttl=60)
def load_db_data():
    try:
        ranges = ['일일업무', '프로젝트', '세부업무', '루틴업무', '설정', '계정관리']
        batch = sh.values_batch_get(ranges)
        v_ranges = batch.get('valueRanges', [])
        return (
            parse_batch_data(v_ranges[0].get('values', [])) if len(v_ranges) > 0 else [],
            parse_batch_data(v_ranges[1].get('values', [])) if len(v_ranges) > 1 else [],
            parse_batch_data(v_ranges[2].get('values', [])) if len(v_ranges) > 2 else [],
            parse_batch_data(v_ranges[3].get('values', [])) if len(v_ranges) > 3 else [],
            parse_batch_data(v_ranges[4].get('values', [])) if len(v_ranges) > 4 else [],
            parse_batch_data(v_ranges[5].get('values', [])) if len(v_ranges) > 5 else []
        )
    except:
        st.error("⏳ 구글 시트 동시 접속 대기 중입니다. 1분 뒤 새로고침(F5) 해주세요.")
        st.stop()

def apply_changes():
    load_db_data.clear() 
    st.rerun()

# --- [보안 및 로그인] ---
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_info' not in st.session_state: st.session_state['user_info'] = {}

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

# --- [사이드바] ---
with st.sidebar:
    st.subheader(f"👤 {u_name}님 ({u_role})")
    if st.button("🔄 최신 데이터 불러오기", use_container_width=True, type="primary"):
        apply_changes()
    st.divider()
    if st.button("🚪 로그아웃", use_container_width=True): 
        st.session_state['logged_in'] = False
        st.rerun()

# --- [데이터 로드] ---
all_daily, proj_data, sub_data, routine_data, kpi_config, user_data = load_db_data()

for k in kpi_config:
    if '구분' not in k: k['구분'] = '공통'
my_kpi_opts = [str(k.get('KPI명', '')) for k in kpi_config if str(k.get('구분', '공통')).strip() in ['공통', u_name]]

st.title("🚀 NOWSYSTEM 통합 업무 관리")
t_date = st.date_input("📅 업무 기준일 선택", datetime.date.today(), key="main_date")
t_str = t_date.strftime("%Y-%m-%d")

if f"lock_{t_str}" not in st.session_state: st.session_state[f"lock_{t_str}"] = False
is_locked = st.session_state[f"lock_{t_str}"]

if u_role == "마스터":
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str]
    tabs = st.tabs(["📝 전사 일과 관리", "📁 프로젝트 관리", "⚙️ 마스터 설정", "📈 통합 KPI", "📊 데이터/보고서"])
    tab_kpi = tabs[3]; tab_rep = tabs[4]
else:
    filtered_daily = [d for d in all_daily if str(d.get('날짜')) == t_str and d.get('담당자') == u_name]
    tabs = st.tabs(["📝 나의 일과", "📁 나의 프로젝트", "📈 나의 KPI", "📊 데이터/보고서"])
    tab_kpi = tabs[2]; tab_rep = tabs[3]

sub_dict = {p.get("프로젝트명", ""): [] for p in proj_data}
for s in sub_data:
    if u_role == "마스터" or s.get('담당자') == u_name:
        pn = s.get("프로젝트명", "")
        if pn in sub_dict: sub_dict[pn].append(s)

# ==========================================
# 탭 1: 일과 관리 (💡 일회성/프로젝트/데일리 완벽 통합)
# ==========================================
with tabs[0]:
    st.header(f"📝 {t_str} 업무 리스트")
    dropdown_opts = [k.get('KPI명', '') for k in kpi_config] if u_role == "마스터" else my_kpi_opts
    
    with st.expander("➕ 오늘의 업무 추가", expanded=not is_locked):
        task_type = st.radio("업무 종류 선택", ["일반/데일리 업무", "프로젝트 연동 업무"], horizontal=True)
        
        if task_type == "일반/데일리 업무":
            my_routines = [r.get('업무명') for r in routine_data if r.get('담당자') == u_name]
            # 데일리 업무를 선택하거나 직접 입력하도록 통합
            sel_opt = st.selectbox("업무명", ["✏️ 직접 입력"] + my_routines)
            if sel_opt == "✏️ 직접 입력":
                n_task = st.text_input("새 업무명 (직접 입력)", disabled=is_locked)
            else:
                n_task = sel_opt
                
            c1, c2 = st.columns(2)
            n_cat = c1.selectbox("분류", ["경영관리", "재무업무", "입찰업무", "조달업무", "현장업무", "기타"], disabled=is_locked)
            n_kpi = c2.selectbox("연관 KPI", dropdown_opts + ["기타"], disabled=is_locked)
            
            if st.button("업무 추가", disabled=is_locked, type="primary"):
                if n_task:
                    w_d.append_row([t_str, n_task, 0, "FALSE", n_cat, "", n_kpi, u_name])
                    apply_changes()
                    
        else: # 프로젝트 연동 업무 선택 시
            my_projs = [p.get('프로젝트명') for p in proj_data if (u_role == "마스터" or p.get('담당자') == u_name) and str(p.get("보관함이동", "FALSE")).upper() != "TRUE"]
            sel_p = st.selectbox("진행 중인 프로젝트", my_projs) if my_projs else None
            
            if sel_p:
                my_subs = [s.get('세부업무명') for s in sub_data if s.get('프로젝트명') == sel_p]
                sel_s = st.selectbox("프로젝트 하위 세부업무", my_subs) if my_subs else None
                
                if st.button("프로젝트 업무 당겨오기", disabled=is_locked, type="primary"):
                    if sel_p and sel_s:
                        p_kpi = next((p.get('KPI') for p in proj_data if p.get('프로젝트명') == sel_p), "기타")
                        # 프로젝트명과 세부업무명을 "::"로 묶어서 저장 (나중에 슬라이더 동기화용)
                        w_d.append_row([t_str, sel_s, 0, "TRUE", "프로젝트", f"{sel_p}::{sel_s}", p_kpi, u_name])
                        apply_changes()
            else:
                st.info("현재 할당된 진행 중 프로젝트가 없습니다.")

    st.divider()
    
    # 💡 고유 행 번호(_r)를 사용하여 슬라이더 충돌 방지
    for row in filtered_daily:
        r_idx = row['_r']
        col1, col2, col3 = st.columns([4, 5, 1])
        disp_name = str(row.get('업무명','')).replace('\n', '<br>')
        badge = f" <small style='color:blue;'>[{row.get('담당자','')}]</small>" if u_role == "마스터" else ""
        is_proj = str(row.get('프로젝트연동', 'FALSE')).upper() == "TRUE"
        
        if is_proj:
            p_info = str(row.get('연결프로젝트', ''))
            col1.markdown(f"**[📂프로젝트]** <span style='color:#555;'>{p_info.replace('::', ' > ')}</span>{badge}", unsafe_allow_html=True)
        else:
            col1.markdown(f"**[{row.get('분류', '기타')}]** {disp_name}{badge}", unsafe_allow_html=True)
            
        cur_p = int(row.get('진행률', 0) if str(row.get('진행률', 0)).isdigit() else 0)
        new_p = col2.slider("진행", 0, 100, cur_p, 10, key=f"d_sld_{r_idx}", disabled=is_locked, label_visibility="collapsed")
        
        if not is_locked and new_p != cur_p:
            try:
                # 1. 일일 업무 시트 진행률 업데이트
                w_d.update_cell(r_idx, 3, new_p) 
                
                # 💡 [핵심] 프로젝트 연동 업무라면, '세부업무' 시트의 진행률도 동시에 올립니다!
                if is_proj:
                    p_info = str(row.get('연결프로젝트', ''))
                    if "::" in p_info:
                        p_n, s_n = p_info.split("::", 1)
                        for s in sub_data:
                            if s.get('프로젝트명') == p_n and s.get('세부업무명') == s_n:
                                w_s.update_cell(s['_r'], 3, new_p)
                                break
                st.toast(f"✅ 진행률 {new_p}% 저장 완료!")
            except:
                st.warning("⏳ 구글 시트 접속 지연. 잠시 후 다시 조절해주세요.")
        
        if col3.button("🗑️", key=f"d_del_{r_idx}", disabled=is_locked):
            w_d.delete_rows(r_idx)
            apply_changes()

    # (데일리 업무 목록을 미리 관리해두는 셋팅 창)
    with st.expander("⚙️ 나의 데일리 업무(루틴) 목록 관리", expanded=False):
        with st.form("add_routine_form"):
            rc1, rc2, rc3 = st.columns([2,1,1])
            r_task = rc1.text_input("새 데일리 업무명")
            r_cat = rc2.selectbox("분류", ["경영관리", "재무업무", "입찰업무", "조달업무", "현장업무", "기타"])
            r_kpi = rc3.selectbox("연관 KPI", dropdown_opts + ["기타"])
            if st.form_submit_button("목록에 추가하기"):
                if r_task:
                    w_r.append_row([r_task, r_cat, r_kpi, u_name])
                    apply_changes()
        for r in routine_data:
            if r.get('담당자') == u_name:
                r_idx = r['_r'] 
                rc1, rc2 = st.columns([4, 1])
                rc1.write(f"· **[{r.get('분류')}]** {r.get('업무명')}")
                if rc2.button("목록에서 삭제", key=f"rdel_{r_idx}"):
                    w_r.delete_rows(r_idx)
                    apply_changes()

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
                apply_changes()

    for p in proj_data:
        r_idx = p['_r']
        if (u_role != "마스터" and p.get('담당자') != u_name) or str(p.get("보관함이동", "FALSE")).upper() == "TRUE": continue
            
        pn = p.get("프로젝트명", "")
        owner = f" ({p.get('담당자','')})" if u_role == "마스터" else ""
        
        with st.expander(f"📂 {pn} [{p.get('분류')}]{owner}"):
            with st.form(key=f"sub_form_{r_idx}"):
                sc1, sc2 = st.columns([3,1])
                new_sub = sc1.text_input("세부 업무명")
                if sc2.form_submit_button("하위 업무 추가", disabled=is_locked):
                    if new_sub:
                        w_s.append_row([pn, new_sub, 0, "TRUE", u_name])
                        apply_changes()
            for s in sub_dict.get(pn, []):
                sl1, sl2 = st.columns([6, 4])
                sl1.markdown(f"· {str(s.get('세부업무명','')).replace('\n','<br>')}", unsafe_allow_html=True)
                cur_sp = int(s.get('진행률',0) if str(s.get('진행률',0)).isdigit() else 0)
                sp = sl2.slider("진행", 0, 100, cur_sp, 10, key=f"s_sld_{s['_r']}", disabled=is_locked, label_visibility="collapsed")
                
                if not is_locked and sp != cur_sp:
                    try:
                        w_s.update_cell(s['_r'], 3, sp)
                        st.toast(f"✅ 진행률 {sp}% 저장 완료!")
                    except:
                        pass
            
            st.write("---")
            ac1, ac2, ac3 = st.columns([2,1,1])
            if ac2.button("📦 보관함 이동", key=f"arc_{r_idx}", disabled=is_locked):
                w_p.update_cell(r_idx, 6, "TRUE")
                apply_changes()
            if ac3.button("🗑️ 프로젝트 삭제", key=f"pdel_{r_idx}", disabled=is_locked):
                w_p.delete_rows(r_idx)
                apply_changes()

# ==========================================
# 탭 3: 마스터 전용 설정
# ==========================================
if u_role == "마스터":
    with tabs[2]:
        st.header("⚙️ 마스터 전용 설정")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("👥 사내 계정 목록")
            u_df = pd.DataFrame(user_data)
            e_u_df = st.data_editor(u_df, num_rows="dynamic", use_container_width=True)
            if st.button("계정 정보 저장"):
                w_u.clear(); w_u.update([e_u_df.columns.values.tolist()] + e_u_df.values.tolist())
                st.success("저장 완료!")
                apply_changes()
        
        with c2:
            st.subheader("🎯 개인별 & 공통 KPI 관리")
            k_df = pd.DataFrame(kpi_config)
            if '구분' not in k_df.columns: k_df['구분'] = '공통'
            cols = k_df.columns.tolist()
            if '구분' in cols:
                cols.insert(0, cols.pop(cols.index('구분')))
                k_df = k_df[cols]
            e_k_df = st.data_editor(k_df, num_rows="dynamic", use_container_width=True)
            if st.button("KPI 지표 저장"):
                w_st.clear(); w_st.update([e_k_df.columns.values.tolist()] + e_k_df.values.tolist())
                st.success("저장 완료!")
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
            new_rows = df_up.values.tolist()
            w_d.append_rows(new_rows)
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
