import openpyxl
import pyperclip
import os
import streamlit as st

# 웹 페이지 제목
st.title("엑셀 데이터 복사")
st.caption(":rainbow[지정된 키워드 바로 아래 행부터 전체 내용이 클립보드에 복사됩니다.]")

# 기본 폴더 및 키워드 설정
default_directory_path = "C:/Users/jaguar/Downloads"  # 기본 경로
default_keywords = ["중간_CNS", "zh-hans", "CNS", "zh_CN", "Simplified Chinese"]  # 기본 키워드

# 폴더 경로 입력 (사용자가 수정 가능)
directory_path = st.text_input("📂 파일이 있는 폴더 경로", value=default_directory_path)

# 폴더 내 엑셀 파일 자동 탐색 및 선택
xlsx_files = [f for f in os.listdir(directory_path) if f.endswith(".xlsx")] if os.path.exists(directory_path) else []
if not xlsx_files:
    st.warning("⚠️ 해당 폴더에 `.xlsx` 파일이 없습니다. 경로를 확인하세요.")
file_name = st.selectbox("📄 파일 선택", xlsx_files) if xlsx_files else None

# 키워드 입력 (사용자가 수정 가능)
keywords_input = st.text_area("🔍 찾을 키워드(언어열 이름)", value=", ".join(default_keywords))
keywords = [keyword.strip() for keyword in keywords_input.split(",")]

# 실행 버튼
if st.button("실행"):
    if not file_name:
        st.error("❌ 파일을 선택하세요!")
    else:
        file_path = os.path.join(directory_path, file_name)
        if os.path.exists(file_path):
            # 엑셀 파일 열기
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active

            # 특정 키워드가 있는 행과 열 찾기
            target_row, target_column = None, None
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True), start=1):
                for col_idx, cell_value in enumerate(row, start=1):
                    if cell_value in keywords:  # 키워드가 포함된 행 찾기
                        target_row, target_column = row_idx, col_idx
                        break
                if target_row:
                    break  # 첫 번째 일치하는 키워드만 찾음

            # 키워드 발견 여부 확인
            if target_row is None:
                st.error("❌ 키워드를 찾을 수 없습니다.")
            else:
                # 키워드 아래 행부터 끝까지 해당 열의 데이터 가져오기
                values = [
                    str(ws.cell(row=i, column=target_column).value).replace("\n", "\n")  # 줄바꿈 유지
                    for i in ra
