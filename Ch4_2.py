from tkinter import filedialog
import os
import re
import zipfile
import shutil

# 결과를 저장할 폴더와 임시 저장 폴더 이름을 변수로 지정
result_f = r'C:\Data\Book\MinimizedPython\Ch3\Ch3_3\Result' # 사용자 환경에 따라 폴더 경로 수정할 것
temp_f = 'Temp'

# Office 파일별 확장자명 & 중간 폴더명 선언하기
ext = ['ppt', 'xls', 'doc']
middir = ['ppt', 'xl', 'word']

# 작업할 폴더 선택하기
TPath = filedialog.askdirectory()

# 파일 확장자명에 따라 반복문 실행
for idx, x in enumerate(ext):
    # 확장자명 x에 따른 파일 리스트 가져오기
    files = [f for f in os.listdir(TPath) if re.match(r'.*[.]' + x, f)]
    
    # 파일별 반복문
    for file in files:

        # 저장할 파일 경로 생성
        newpath = os.path.join(result_f, middir[idx], file[:file.find('.')]) 
        if not os.path.exists(newpath):
            os.makedirs(newpath)

        # 임시 폴더에 MS Office 파일 압축 해제
        with zipfile.ZipFile(os.path.join(TPath,file), 'r') as zip_ref:
            zip_ref.extractall(temp_f)

        # 임시 폴더에서 이미지 파일을 저장할 파일 경로에 복사
        shutil.copytree(os.path.join(temp_f,middir[idx],'media'), newpath, 
                        dirs_exist_ok=True)

        # 임시 폴더 삭제
        shutil.rmtree(temp_f)