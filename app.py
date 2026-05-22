import streamlit as st
import pandas as pd
import zipfile
import io
import re
import random

st.set_page_config(page_title="출결신고서 통합 HWPX 생성기", layout="wide")

st.title("📄 출결신고서 통합 HWPX 자동 생성기")
st.markdown("나이스 엑셀을 업로드하면 기존 매크로 로직으로 가공하여 **완벽한 서식의 단일 HWPX 파일**로 합쳐서 생성합니다.")

# 1. 파일 업로드 UI
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 데이터 및 학급 설정")
    grade = st.number_input("학년", value=3)
    class_num = st.number_input("반", value=3)
    excel_file = st.file_uploader("나이스 출결 엑셀/CSV 업로드", type=['xlsx', 'csv'])

with col2:
    st.subheader("2. 한글 양식 업로드")
    st.info("🚨 템플릿 HWPX 내에 {{성명}}, {{사유}} 등의 태그가 정상적으로 존재해야 합니다.")
    template_file = st.file_uploader("HWPX 템플릿 파일 업로드", type=['hwpx'])

# 2. VBA 로직 이식 데이터 가공 함수
def process_data_vba(df, g, c):
    processed = []
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        num_raw = str(row.iloc[1]).replace('.0', '').strip() if pd.notna(row.iloc[1]) else ""
        name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        raw_type = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        periods = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
        reason = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""

        if not raw_type or raw_type.startswith("미인정"): continue

        kind, reason_type = "", ""
        for k in ["결석", "지각", "조퇴", "결과"]:
            if raw_type.endswith(k):
                kind = k
                reason_type = raw_type[:-len(k)]
                break
        
        if not reason_type or not num_raw or not name: continue

        # 번호 정렬용 정수 변환
        num_int = int(num_raw) if num_raw.isdigit() else 99

        # 날짜 파싱
        clean_date = date_str
        while clean_date.endswith('.'): clean_date = clean_date[:-1]
        clean_date = clean_date.replace('-', '.')
        parts = clean_date.split('.')
        y, m, d = 0, 0, 0
        if len(parts) >= 3 and parts[0].isdigit() and parts[1].isdigit() and parts[2].isdigit():
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])

        # 교시 파싱
        p_arr = [int(x) for x in re.findall(r'\d+', periods) if "교시" in periods]
        period_note = ""
        if kind == "결석": period_note = "1일간"
        elif kind == "지각": period_note = f"~{max(p_arr)}교시까지" if p_arr else ""
        elif kind == "조퇴": period_note = f"{min(p_arr)}교시부터" if p_arr else ""
        elif kind == "결과": period_note = ",".join([f"{p}교시" for p in p_arr]) if p_arr else ""

        formatted_date = f"{y}년 {m}월 {d}일" if y > 0 else ""

        processed.append({
            "{{학년}}": str(g),
            "{{반}}": str(c),
            "{{번호}}": str(num_raw),
            "{{성명}}": name,
            "{{사유유형}}": reason_type,
            "{{신고종류}}": kind,
            "{{출결구분_원본}}": raw_type,
            "{{시작일_연}}": str(y),
            "{{시작일_월}}": str(m),
            "{{시작일_일}}": str(d),
            "{{종료일_연}}": str(y),
            "{{종료일_월}}": str(m),
            "{{종료일_일}}": str(d),
            "{{시작일}}": formatted_date,
            "{{종료일}}": formatted_date,
            "{{결시정보}}": period_note,
            "{{사유}}": reason,
            "{{작성일_연}}": str(y),
            "{{작성일_월}}": str(m),
            "{{작성일_일}}": str(d),
            "_sort_num": num_int,
            "_sort_m": m,
            "_sort_d": d,
            "selected": True
        })
    
    # VBA와 동일하게 번호 -> 월 -> 일 순으로 정렬
    processed.sort(key=lambda x: (x["_sort_num"], x["_sort_m"], x["_sort_d"]))
    return processed

# 3. 깨짐 방지용 HWPX 무한 복사 및 결합 함수
def build_integrated_hwpx(template_bytes, data_list):
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        zip_files = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    
    # 본문 데이터 로드
    section_path = 'Contents/section0.xml'
    if section_path not in zip_files:
        raise ValueError("올바른 HWPX 서식이 아닙니다.")
        
    origin_xml = zip_files[section_path].decode('utf-8')
    
    # 한글 문서 내부 섹션 틀 분리 구문 자동 파싱
    sec_match = re.search(r'(<hs:sec[^>]*>)(.*?)(</hs:sec>)', origin_xml, re.DOTALL)
    if not sec_match:
        raise ValueError("HWPX 내부 구문을 분석할 수 없습니다.")
        
    prefix = origin_xml[:sec_match.start(2)]
    body_template = sec_match.group(2)
    suffix = origin_xml[sec_match.end(2):]
    
    merged_bodies = []
    
    for idx, student in enumerate(data_list):
        page_xml = body_template
        
        # 태그 치환 진행 (VBA 헤더 규칙 기반)
        for tag, value in student.items():
            if tag.startswith('{{'):
                page_xml = page_xml.replace(tag, value)
        
        # 내부 엘리먼트 고유 ID 난수화하여 한글 프로그램 내 충돌 및 파일 깨짐 원천 차단
        def rand_id(m):
            return f' id="{random.randint(100000, 9999999)}"'
        page_xml = re.sub(r' id="\d+"', rand_id, page_xml)
        
        # 페이지 간의 자동 구분을 위해 매 페이지 끝에 한글 전용 구획 추가
        if idx < len(data_list) - 1:
            # 마지막 페이지가 아닐 때만 한글 인쇄용 강제 쪽나누기 태그 삽입
            if "</hp:p>" in page_xml:
                page_xml = page_xml.rstrip().replace("</hp:p>", "</hp:p><hp:p><hp:run><hp:ctrl><hc:colBr/></hp:ctrl></hp:run></hp:p>", 1)
                
        merged_bodies.append(page_xml)
        
    # 재조립 진행
    final_xml = prefix + "".join(merged_bodies) + suffix
    zip_files[section_path] = final_xml.encode('utf-8')
    
    # ZIP 포맷 재압축 및 바이너리 출력
    out_io = io.BytesIO()
    with zipfile.ZipFile(out_io, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fpath, fdata in zip_files.items():
            zout.writestr(fpath, fdata)
            
    return out_io.getvalue()

# 4. 실행 프로세스
if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file) if excel_file.name.endswith('.xlsx') else pd.read_csv(excel_file)
        if 'attendance_records' not in st.session_state:
            st.session_state.attendance_records = process_data_vba(df, grade, class_num)
    except Exception as e:
        st.error(f"데이터 파일 분석 중 오류: {e}")

    if 'attendance_records' in st.session_state and st.session_state.attendance_records:
        st.divider()
        st.subheader("⚙️ 출결 가공 완료 대상자 명단")
        
        display_df = pd.DataFrame(st.session_state.attendance_records)
        view_cols = ['selected', '{{시작일_월}}', '{{시작일_일}}', '{{번호}}', '{{성명}}', '{{출결구분_원본}}', '{{사유}}']
        edited_df = st.data_editor(display_df[view_cols], hide_index=True)
        
        # 최종 선택된 학생 필터링
        selected_students = [st.session_state.attendance_records[i] for i, r in edited_df.iterrows() if r['selected']]
        
        if st.button("🔥 1개의 HWPX 파일로 최종 결과물 통합 생성하기", type="primary"):
            if not selected_students:
                st.warning("선택된 학생이 없습니다.")
            else:
                with st.spinner("선생님의 완벽한 한글 서식으로 조립 중..."):
                    try:
                        final_hwpx = build_integrated_hwpx(template_file.read(), selected_students)
                        first_m = str(selected_students[0]['{{시작일_월}}']).zfill(2)
                        
                        st.success(f"🎉 총 {len(selected_students)}명의 결석신고서가 포함된 단일 한글 파일이 완성되었습니다!")
                        st.download_button(
                            label="📥 완성된 결석신고서_통합본.hwpx 다운로드",
                            data=final_hwpx,
                            file_name=f"결석신고서_통합본_2026년{first_m}월.hwpx",
                            mime="application/octet-stream"
                        )
                    except Exception as e:
                        st.error(f"한글 파일 조립 중 실패: {e}\n양식 내의 태그 상태를 확인해 주세요.")
