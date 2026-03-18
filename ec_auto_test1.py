import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="EC_AUTO 데이터 대조 도구", layout="wide")

st.title("📂 EC_AUTO 데이터 대조 자동화")
st.info("동료분께: 드롭박스에서 받은 파일을 아래에 각각 넣어주세요.")

# 1. 파일 업로드 섹션
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ 기준 엑셀 파일 (1Project)")
    project_file = st.file_uploader("1Project 폴더의 엑셀 파일을 선택하세요", type=['xlsx', 'xls'])

with col2:
    st.subheader("2️⃣ 대조 대상 파일들 (Daejo_excel)")
    daejo_files = st.file_uploader("Daejo_excel 폴더의 파일들을 모두 선택하세요", accept_multiple_files=True)

# 2. 데이터 처리 버튼
if st.button("🚀 데이터 대조 시작"):
    if project_file and daejo_files:
        try:
            # 엑셀 읽기
            df = pd.read_excel(project_file)
            col_a = df.columns[0] # 첫 번째 열 기준
            
            # 업로드된 파일들의 이름 리스트 생성 (확장자 제거)
            daejo_names = [f.name.split('.')[0] for f in daejo_files]
            
            # 대조 작업
            matched_df = df[df[col_a].astype(str).str.strip().isin(daejo_names)]
            
            if not matched_df.empty:
                st.success(f"✅ 대조 완료! 총 {len(matched_df)}건이 일치합니다.")
                
                # 결과물을 엑셀로 변환 (메모리 상에서 처리)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    matched_df.to_excel(writer, index=False)
                
                # 다운로드 버튼 생성
                st.download_button(
                    label="📥 결과 파일 다운로드 (Result_Excel)",
                    data=output.getvalue(),
                    file_name="Final_Result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("⚠️ 일치하는 데이터가 없습니다.")
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
    else:
        st.error("❗ 모든 파일을 업로드해야 시작할 수 있습니다.")
