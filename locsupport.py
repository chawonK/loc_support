import streamlit as st
import openpyxl
import pandas as pd
import io
import fitz  
import re
import tempfile
import zipfile
from io import BytesIO
from docx import Document
from pptx import Presentation

# 페이지 설정
st.set_page_config(page_title="엑셀 도구 모음", layout="centered")

# 사이드바 메뉴
st.sidebar.title("엑셀 도구 모음")
page = st.sidebar.radio(" ", ("엑셀 데이터 복사", "엑셀 파일 미리보기", "엑셀 시트 분할", "단어수 카운터(파일)","단어수 카운터(웹)","월간 보고 데이터"))

# 1. 엑셀 데이터 복사
if page == "엑셀 데이터 복사":
    st.title('📄엑셀 데이터 복사')
    st.write(":rainbow[지정된 키워드 바로 아래 행부터 전체 내용이 복사됩니다.]")

    uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx", "xls"])
    default_keywords = ["중간_CNS", "zh-hans", "CNS", "zh_CN", "Simplified Chinese", "CNS (중국어 간체)"]
    keywords_input = st.text_area("찾을 키워드(언어열 이름)", value=", ".join(default_keywords))

    if uploaded_file:
        keywords = [keyword.strip() for keyword in keywords_input.split(',')]
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws = wb.active

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
            values = [str(ws.cell(row=i, column=target_column).value or "") for i in range(target_row + 1, ws.max_row + 1)]
            formatted_text = "\n".join(f'"{value}"' for value in values)

            if values:
                st.success("✅ 데이터 준비 완료! 아래 박스에서 복사해서 사용하세요.")
                st.text_area("복사된 내용", formatted_text, height=200)
                st.download_button("📥 텍스트 다운로드", formatted_text, "data.txt")
            else:
                st.warning("⚠️ 복사할 데이터가 없습니다.")

        wb.close()

# 2. 엑셀 시트 분할
elif page == "엑셀 시트 분할":
    st.title("✂️ 엑셀 시트 분할")
    st.caption("※ 엑셀 파일의 각 시트를 새로운 파일로 분할하여 ZIP으로 다운로드합니다.")
    
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        file_name = uploaded_file.name
        excel_file = pd.ExcelFile(uploaded_file)
    
        if not excel_file.sheet_names:
            st.error("❌ 해당 엑셀 파일에 시트가 없습니다.")
        else:
            # ZIP 파일을 메모리에 생성
            zip_buffer = BytesIO()
            zip_file = zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED)
    
            for sheet_name in excel_file.sheet_names:
                df = excel_file.parse(sheet_name)
                
                # 각 시트를 엑셀 파일로 변환
                sheet_io = BytesIO()
                with pd.ExcelWriter(sheet_io, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                
                # ZIP에 추가
                zip_file.writestr(f"{sheet_name}.xlsx", sheet_io.getvalue())
    
            zip_file.close()  # ZIP 파일 닫기
    
            # ZIP 파일 다운로드
            zip_buffer.seek(0)
            st.download_button(
                label="📥 ZIP 파일 다운로드",
                data=zip_buffer,
                file_name=f"{file_name}_시트분할.zip",
                mime="application/zip",
            )

# 3. 단어수 카운터
elif page == "단어수 카운터(파일)":
    st.title("🔢 단어수 카운터(파일)")
    st.caption("※ 여러 언어가 섞인 파일은 단어수가 부정확하게 나옵니다. 참고용으로만 이용해주세요.")
    uploaded_file = st.file_uploader("파일 업로드 (Word, PPTX, Excel, PDF, TXT)", type=["pptx", "docx", "xlsx", "pdf", "txt"])

    def count_words(text):
        return len(text.split()) if text else 0

    if uploaded_file:
        file_name = uploaded_file.name.lower()
        word_count, file_preview = 0, ""

        if file_name.endswith(".docx"):
            doc = Document(uploaded_file)
            file_preview = "\n".join([para.text[:300] for para in doc.paragraphs])
            word_count = count_words(file_preview)
        elif file_name.endswith(".pptx"):
            prs = Presentation(uploaded_file)
            file_preview = "\n".join([shape.text[:300] for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
            word_count = count_words(file_preview)
        elif file_name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(uploaded_file)
            file_preview = "\n".join([str(cell.value)[:300] for ws in wb.worksheets for row in ws.iter_rows() for cell in row if cell.value])
            word_count = count_words(file_preview)
        elif file_name.endswith(".pdf"):
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            file_preview = "\n".join([page.get_text()[:300] for page in doc])
            word_count = count_words(file_preview)
        elif file_name.endswith(".txt"):
            file_preview = uploaded_file.read().decode("utf-8")[:300]
            word_count = count_words(file_preview)

        st.markdown(f"### 단어수: <span style='color: #4CAF50; font-size: 24px;'>{word_count}</span>", unsafe_allow_html=True)
        st.text_area("파일 내용 미리보기:", file_preview, height=200)

# 4. 월간 보고 데이터
elif page == "월간 보고 데이터":
    st.title("📊 Jira CSV 데이터 추출기")
    st.caption("Jira에서 요청 필터 페이지에서 우측 상단의 '보기' > 'CSV (모든 필드)'로 다운로드 받은 월별 전체 요청 파일을 업로드 하면 월별 프로젝트별 요청수, 단어수 합계를 볼 수 있습니다.")
    uploaded_file = st.file_uploader("CSV 파일 업로드", type=["csv"])

    if uploaded_file:
        df = pd.read_csv(uploaded_file)
        df_filtered = df[['프로젝트 이름', '요약', '기한', '생성일']].copy()
        df_filtered[['단어수', '기준 언어']] = df_filtered['요약'].str.extract(r'\[(\d+)\s*([A-Za-z]+)\]')
        df_filtered['단어수'] = pd.to_numeric(df_filtered['단어수'], errors='coerce')
        project_summary_df = df_filtered.groupby('프로젝트 이름').agg(요청수=('요약', 'count'), 단어수_합계=('단어수', 'sum')).reset_index()
        total_row = pd.DataFrame({'프로젝트 이름': ['합계'], '요청수': [project_summary_df['요청수'].sum()], '단어수_합계': [project_summary_df['단어수_합계'].sum()]})
        project_summary_df = pd.concat([project_summary_df, total_row], ignore_index=True)

        # C2 셀에서 시트명을 동적으로 생성
        c2_value = df_filtered['기한'].iloc[0]  # C2 내용 (기한 열의 첫 번째 값)
        sheet_name_prefix = c2_value[2:4] + c2_value[5:7]  # 2, 3, 5, 6번째 글자 추출
        original_sheet_name = sheet_name_prefix  # '원본 데이터' 시트명
        summary_sheet_name = sheet_name_prefix + " 월별 통계"  # '프로젝트별 요약' 시트명

        # 동적으로 파일명 생성
        file_name = f"{sheet_name_prefix}_project_summary.xlsx"

        # Excel 파일로 저장
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name=original_sheet_name)
            project_summary_df.to_excel(writer, index=False, sheet_name=summary_sheet_name)
        output.seek(0)

        # 다운로드 버튼
        st.download_button(
            label="📥 엑셀 파일 다운로드",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 5. 단어수 카운터
elif page == "단어수 카운터(웹)":
    st.title("🔢 단어수 카운터(웹)")
    st.write("텍스트를 입력하면 띄어쓰기 기준으로 단어 수를 계산합니다.")

    def count_words(text):
        words = text.split()
        return len(words)

    # 초기 상태 설정
    if 'text_input' not in st.session_state:
        st.session_state.text_input = ""
    
    if 'word_count' not in st.session_state:
        st.session_state.word_count = 0
    
    # 단어 수 업데이트 함수
    def update_word_count():
        st.session_state.word_count = count_words(st.session_state.text_input)

    st.subheader(f"단어 수: {st.session_state.word_count}")
    
    # 텍스트 입력 필드
    text_input = st.text_area("텍스트 입력", height=200, key="text_input")
    
    # 버튼 클릭 시 단어 수 업데이트
    if st.button("단어 수 계산"):
        update_word_count()
    
    

# 6. 엑셀 파일 미리보기
elif page == "엑셀 파일 미리보기":
    st.title("🔍엑셀 파일 미리보기")
    st.write("파일을 업로드 하면 내용을 미리 볼 수 있습니다.")

    # 파일 업로드
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # 엑셀 파일 로드
        xls = pd.ExcelFile(uploaded_file)
        
        # 데이터프레임 로드 (첫 번째 시트 자동 선택)
        df = pd.read_excel(xls, sheet_name=0)
        
        # 데이터 미리보기
        st.write("### 데이터 미리보기")
        st.dataframe(df.head(20))  # 상위 20개 행 미리보기
