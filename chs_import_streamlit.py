import streamlit as st
import openpyxl
import os

st.title('엑셀 데이터 복사')
st.caption(":rainbow[지정된 키워드 바로 아래 행부터 전체 내용이 클립보드에 복사됩니다.]")

# 파일 업로드 기능 추가
uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])

# 키워드 입력
default_keywords = ["중간_CNS", "zh-hans", "CNS", "zh_CN", "Simplified Chinese"]
keywords_input = st.text_area("찾을 키워드", value=", ".join(default_keywords))

formatted_text = ""  # 복사할 텍스트를 저장할 변수

# 실행 버튼
if st.button("실행"):
    if uploaded_file is None:
        st.error("❌ 파일을 업로드해주세요!")
    else:
        # 업로드된 파일을 openpyxl로 불러오기
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active

        keywords = [k.strip() for k in keywords_input.split(",")]

        # 키워드 찾기
        target_row, target_column = None, None
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            for col_idx, cell_value in enumerate(row, start=1):
                if cell_value in keywords:
                    target_row, target_column = row_idx, col_idx
                    break
            if target_row:
                break

        if target_row is None:
            st.error("❌ 키워드를 찾을 수 없습니다.")
        else:
            values = [
                str(ws.cell(row=i, column=target_column).value).replace("\n", "\r\n")
                for i in range(target_row + 1, ws.max_row + 1)
                if ws.cell(row=i, column=target_column).value is not None
            ]
            
            if values:
                formatted_text = "\r\n".join(f'"{value}"' for value in values)
                st.success("✅ 데이터 추출 완료!")
            else:
                st.warning("⚠️ 복사할 데이터가 없습니다.")

        wb.close()

# 📋 "복사하기" 버튼 추가
if formatted_text:
    st.text_area("복사할 내용", formatted_text, height=200)
    st.write("👉 아래 버튼을 클릭 후 **Ctrl + C** 또는 **Cmd + C**로 복사하세요!")
    st.button("📋 복사하
