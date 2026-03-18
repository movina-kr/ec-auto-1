import os
import pandas as pd
import sys

def run_process():
    # 1. 드롭박스 또는 로컬 경로 설정 (실행 파일 위치 기준 상위 폴더 탐색)
    # 프로그램이 실행되는 위치를 기준으로 경로를 잡습니다.
    base_path = os.getcwd() 
    
    project_dir = os.path.join(base_path, '1Project')
    daejo_dir = os.path.join(base_path, 'Daejo_excel')
    result_dir = os.path.join(base_path, 'Result_Excel')

    # 폴더 존재 확인
    for path in [project_dir, daejo_dir, result_dir]:
        if not os.path.exists(path):
            print(f"오류: '{path}' 폴더를 찾을 수 없습니다.")
            input("엔터를 눌러 종료하세요...")
            return

    try:
        # 2. Daejo_excel 폴더 내 파일명 목록 (확장자 제외)
        daejo_files = [os.path.splitext(f)[0] for f in os.listdir(daejo_dir) 
                       if os.path.isfile(os.path.join(daejo_dir, f))]
        
        # 3. 1Project 폴더 내 첫 번째 엑셀 파일 읽기
        project_files = [f for f in os.listdir(project_dir) if f.endswith(('.xlsx', '.xls'))]
        
        if not project_files:
            print("1Project 폴더에 엑셀 파일이 없습니다.")
        else:
            target_file = os.path.join(project_dir, project_files[0])
            df = pd.read_excel(target_file)
            
            # A컬럼(첫 번째 열) 추출 및 대조
            col_a = df.columns[0]
            # 파일명과 A컬럼 데이터 비교 (양끝 공백 제거 후 비교)
            matched_df = df[df[col_a].astype(str).str.strip().isin(daejo_files)]
            
            # 4. 결과 저장
            if not matched_df.empty:
                output_path = os.path.join(result_dir, 'Matched_Result.xlsx')
                matched_df.to_excel(output_path, index=False)
                print(f"성공! {len(matched_df)}개의 행을 추출하여 Result_Excel에 저장했습니다.")
            else:
                print("일치하는 데이터가 없습니다.")

    except Exception as e:
        print(f"실행 중 오류 발생: {e}")
    
    input("\n작업 완료! 엔터를 누르면 창이 닫힙니다...")

if __name__ == "__main__":
    run_process()