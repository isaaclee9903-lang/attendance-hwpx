import streamlit as st
import pandas as pd
import zipfile
import io
import re
import random

st.set_page_config(page_title="출결신고서 HWPX 병합 생성기", layout="wide")

st.title("📄 출결신고서 HWPX 자동 병합 생성기")
st.markdown("선택한 모든 학생의 결석신고서를 **하나의 HWPX 파일(여러 페이지)**로 합쳐서 만들어줍니다.")

# 1. 파일 업로드 UI
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 데이터 업로드")
    grade = st.number_input("학년", value=3)
    class_num = st.number_input("반", value=3)
    excel_file = st.file_uploader("나이스 출결 엑셀/CSV 업로드", type=['xlsx', 'csv'])

with col2:
    st.subheader("2. 양식 업로드")
    st.info("🚨 템플릿 맨 마지막에 반드시 **[Ctrl+Enter]로 쪽 나누기**를 추가한 파일을 올려주세요!")
    template_file = st.file_uploader("HWPX 템플릿 파일 업로드", type=['hwpx'])

# 2. 데이터 가공 함수 (VBA 로직 유지)
def process_data(df, g, c):
    processed = []
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
        num = str(row.iloc[1]).replace('.0', '') if pd.notna(row.iloc[1]) else ""
        name = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        raw_type = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""
        periods = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""
        reason = str(row.iloc[5]) if pd.notna(row.iloc[5]) else ""

        if not raw_type or str(raw_type).startswith("미인정"): continue

        kind, reason_type = "", ""
        for k in ["결석", "지각", "조퇴", "결과"]:
            if raw_type.endswith(k):
                kind = k
                reason_type = raw_type[:-len(k)]
                break
        
        if not reason_type or not num or not name: continue

        clean_date = date_str.replace('-', '.').rstrip('.')
        parts = clean_date.split('.')
        y, m, d = "", "", ""
        if len(parts) >= 3:
            y, m, d = parts[0], parts[1].zfill(2), parts[2].zfill(2)

        p_arr = [int(x) for x in re.findall(r'\d+', periods)]
        period_note = ""
        if kind == "결석": period_note = "1일간"
        elif kind == "지각": period_note = f"~{max(p_arr)}교시까지" if p_arr else ""
        elif kind == "조퇴": period_note = f"{min(p_arr)}교시부터" if p_arr else ""
        elif kind == "결과": period_note = ", ".join([f"{p}교시" for p in p_arr]) if p_arr else ""

        period_str = f"{y}년 {m}월 {d}일 ({period_note})"

        processed.append({
            "{{학년}}": str(g),
            "{{반}}": str(c),
            "{{번호}}": str(num),
            "{{성명}}": name.strip(),
            "{{결석종류}}": raw_type,
            "{{결석기간}}": period_str,
            "{{결석사유}}": reason.strip(),
            "{{년}}": y,
            "{{월}}": m,
            "{{일}}": d,
            "selected": True
        })
    return processed

# 3. HWPX 내부 XML 무한 복사 및 병합 함수 (핵심)
def generate_merged_hwpx(template_bytes, data_list):
    # HWPX 압축 풀기
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        zip_data = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    
    content = zip_data['Contents/section0.xml'].decode('utf-8')
    
    # <hs:sec> 내부의 본문(1페이지 분량) 전체 덩어리 추출
    match = re.search(r'(<hs:sec[^>]*>)(.*?)(</hs:sec>)', content, re.DOTALL)
    if not match:
        raise ValueError("HWPX 구조를 인식할 수 없습니다.")
    
    start_tag = match.group(1)
    inner_xml = match.group(2)
    end_tag = match.group(3)
    
    merged_inner_xml = ""
    
    # 선택된 학생 수만큼 본문 덩어리 복사 및 치환
    for data_dict in data_list:
        student_xml = inner_xml
        
        # 1. 빈칸(태그) 치환
        for tag, val in data_dict.items():
            if tag.startswith('{{') and tag.endswith('}}'):
                student_xml = student_xml.replace(tag, val)
                
        # 2. HWPX 문서 충돌을 막기 위해 복사본들의 내부 ID값을 랜덤으로 변경
        def repl_id(m):
            return f' id="{random.randint(1000000000, 4200000000)}"'
        
        student_xml = re.sub(r' id="\d+"', repl_id, student_xml)
        
        # 병합 데이터에 이어붙이기
        merged_inner_xml += student_xml
        
    # 복사된 덩어리들을 원본 껍데기(XML)에 다시 집어넣기
    final_xml = start_tag + merged_inner_xml + end_tag
    zip_data['Contents/section0.xml'] = final_xml.encode('utf-8')
    
    # 다시 HWPX 파일로 압축해서 내보내기
    out_io = io.BytesIO()
    with zipfile.ZipFile(out_io, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname, fdata in zip_data.items():
            zout.writestr(fname, fdata)
            
    return out_io.getvalue()

# 4. 메인 실행 로직
if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file) if excel_file.name.endswith('.xlsx') else pd.read_csv(excel_file)
        if 'processed_data' not in st.session_state:
            st.session_state.processed_data = process_data(df, grade, class_num)
    except Exception as e:
        st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {e}")

    if 'processed_data' in st.session_state and len(st.session_state.processed_data) > 0:
        st.divider()
        st.subheader("3. 대상자 선택 및 통합 파일 생성")
        
        display_df = pd.DataFrame(st.session_state.processed_data)
        view_df = display_df[['selected', '{{월}}', '{{일}}', '{{번호}}', '{{성명}}', '{{결석종류}}', '{{결석사유}}']].copy()
        edited_df = st.data_editor(view_df, hide_index=True)

        if st.button("🖨️ 체크된 인원 HWPX (단일 파일) 일괄 생성"):
            selected_data = [st.session_state.processed_data[idx] for idx, row in edited_df.iterrows() if row['selected']]
            
            if not selected_data:
                st.warning("선택된 학생이 없습니다.")
            else:
                with st.spinner("하나의 HWPX 파일로 병합 중입니다... (3초 소요)"):
                    try:
                        merged_hwpx_bytes = generate_merged_hwpx(template_file.read(), selected_data)
                        
                        st.success(f"✅ 총 {len(selected_data)}명 분량의 결석신고서가 1개의 HWPX 파일로 완성되었습니다!")
                        st.download_button(
                            label="📥 결석신고서_통합본.hwpx 다운로드",
                            data=merged_hwpx_bytes,
                            file_name="결석신고서_통합본.hwpx",
                            mime="application/octet-stream"
                        )
                    except Exception as e:
                        st.error(f"문서 병합 중 오류가 발생했습니다: {e}")
