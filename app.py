import streamlit as st
import pandas as pd
import zipfile
import io
import re

st.set_page_config(page_title="출결신고서 HWPX 메일머지", layout="wide")

st.title("📄 출결신고서 HWPX 자동 생성기")
st.markdown("나이스 출결 엑셀과 HWPX 템플릿을 업로드하면, 서식 훼손 없이 내용만 치환하여 생성합니다.")

# 1. 파일 업로드 UI
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 데이터 업로드")
    grade = st.number_input("학년", value=3)
    class_num = st.number_input("반", value=3)
    excel_file = st.file_uploader("나이스 출결 엑셀/CSV 업로드", type=['xlsx', 'csv'])

with col2:
    st.subheader("2. 양식 업로드")
    st.info("반드시 태그(예: {{성명}})가 삽입된 .hwpx 파일을 올려주세요.")
    template_file = st.file_uploader("HWPX 템플릿 파일 업로드", type=['hwpx'])

# 2. 데이터 가공 함수 (VBA 로직 완벽 이식)
def process_data(df, g, c):
    processed = []
    # 첫 줄이 헤더이므로 skip하지 않고 pandas가 자동으로 헤더로 잡았다고 가정
    for _, row in df.iterrows():
        # 데이터 클렌징
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

        # 날짜 파싱
        clean_date = date_str.replace('-', '.').rstrip('.')
        parts = clean_date.split('.')
        y, m, d = "", "", ""
        if len(parts) >= 3:
            y, m, d = parts[0], parts[1].zfill(2), parts[2].zfill(2)

        # 교시 파싱
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
            "_filename": f"{m}월{d}일_{num}번_{name}_{kind}.hwpx", # 파일 저장용 (치환용 아님)
            "selected": True
        })
    return processed

# 3. HWPX 내부 XML 치환 함수
def generate_hwpx(template_bytes, data_dict):
    # HWPX는 사실 ZIP 파일입니다. 메모리에서 압축을 풉니다.
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        zip_data = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    
    # 본문이 담긴 section0.xml을 꺼내서 태그를 교체합니다.
    content = zip_data['Contents/section0.xml'].decode('utf-8')
    for tag, val in data_dict.items():
        if not tag.startswith('_') and tag != 'selected':
            content = content.replace(tag, val)
    zip_data['Contents/section0.xml'] = content.encode('utf-8')

    # 다시 ZIP(HWPX)으로 묶어줍니다.
    out_io = io.BytesIO()
    with zipfile.ZipFile(out_io, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname, fdata in zip_data.items():
            zout.writestr(fname, fdata)
    
    return out_io.getvalue()


# 메인 로직
if excel_file and template_file:
    # 엑셀 읽기
    try:
        df = pd.read_excel(excel_file) if excel_file.name.endswith('.xlsx') else pd.read_csv(excel_file)
        if 'processed_data' not in st.session_state:
            st.session_state.processed_data = process_data(df, grade, class_num)
    except Exception as e:
        st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {e}")

    if 'processed_data' in st.session_state and len(st.session_state.processed_data) > 0:
        st.divider()
        st.subheader("3. 대상자 선택 및 생성")
        
        # 화면에 표 출력 (체크박스 편집 가능)
        display_df = pd.DataFrame(st.session_state.processed_data)
        # 화면에 보여줄 컬럼만 필터링
        view_df = display_df[['selected', '{{월}}', '{{일}}', '{{번호}}', '{{성명}}', '{{결석종류}}', '{{결석사유}}']].copy()
        
        edited_df = st.data_editor(view_df, hide_index=True)

        if st.button("체크된 인원 HWPX 일괄 생성 (ZIP)"):
            template_bytes = template_file.read()
            
            # 생성된 HWPX들을 담을 가상의 ZIP 파일
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for idx, row in edited_df.iterrows():
                    if row['selected']:
                        target_data = st.session_state.processed_data[idx]
                        # HWPX 1개 생성
                        hwpx_bytes = generate_hwpx(template_bytes, target_data)
                        # ZIP 안에 파일 추가
                        zip_file.writestr(target_data['_filename'], hwpx_bytes)
            
            st.success("✅ 문서 생성이 완료되었습니다!")
            st.download_button(
                label="📥 완성된 결석신고서 모음 다운로드 (.zip)",
                data=zip_buffer.getvalue(),
                file_name="결석신고서_자동생성.zip",
                mime="application/zip"
            )
