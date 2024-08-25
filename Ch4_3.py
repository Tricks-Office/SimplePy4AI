import tkinter as tk
from tkinter import filedialog
import os
import re
import zipfile
import shutil

# 이미지 가져오기 기능 함수 만들기
def getOfficeImage():
    # 결과를 저장할 폴더와 임시 저장 폴더 이름을 변수로 지정
    result_f = txtFolder2.get("1.0", "end-1c")
    temp_f = 'Temp'

    # Office 파일별 확장자명 & 중간 폴더명 선언하기
    ext = ['ppt', 'xls', 'doc']
    middir = ['ppt', 'xl', 'word']

    # 작업할 폴더 선택하기
    TPath = txtFolder1.get("1.0", "end-1c")

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

# 작업 대상 폴더 경로 변경 버튼에 이벤트 추가
def selectTargetFolder():
    TPath = filedialog.askdirectory()
    txtFolder1.delete(1.0, tk.END)
    txtFolder1.insert(tk.INSERT, chars=TPath)

# 결과 폴더 경로 변경 버튼에 이벤트 추가
def selectResultFolder():
    RPath = filedialog.askdirectory()
    txtFolder2.delete(1.0, tk.END)
    txtFolder2.insert(tk.INSERT, chars=RPath)

# GUI 창 생성하고 기본정보 설정
win = tk.Tk()
win.geometry("680x130")
win.title('Office 파일 이미지 가져오기')

# Label Frame 설정하고 기본 정보 설정
frameFolder = tk.LabelFrame(win, text = "폴더 경로")
frameFolder.pack()
frameBtn = tk.Frame(win)
frameBtn.pack()

# 라벨 만들고 창에 올리기
lbFolder1 = tk.Label(frameFolder, text = '작업 대상', width = 20)
lbFolder2 = tk.Label(frameFolder, text = '결과', width = 20)

# 텍스트 상자 만들고 창에 올리기
txtFolder1 = tk.Text(frameFolder, width = 40, height = 1, padx=5)
txtFolder2 = tk.Text(frameFolder, width = 40, height = 1, padx=5)
txtFolder1.insert(tk.INSERT, chars="작업 대상 폴더 경로를 입력하세요.")
txtFolder2.insert(tk.INSERT, chars="결과 폴더 경로를 입력하세요.")  

# 폴더 경로 변경 버튼 만들고 창에 올리기
btnFolder1 = tk.Button(frameFolder, text = '경로 선택', width = 20, 
                       command=selectTargetFolder)
btnFolder2 = tk.Button(frameFolder, text = '경로 선택', width = 20, 
                       command=selectResultFolder)

# 폴더 GUI grid로 배치하기
lbFolder1.grid(row=0, column=0, padx=10)
lbFolder2.grid(row=1, column=0, padx=10)
txtFolder1.grid(row=0, column=1, padx=10, pady=5)
txtFolder2.grid(row=1, column=1, padx=10, pady=10)
btnFolder1.grid(row=0, column=2, padx=10)
btnFolder2.grid(row=1, column=2, padx=10)

# 실행 버튼 만들고 창에 올리기
btnRun = tk.Button(frameBtn, text = '이미지 가져오기', width = 30,
                     command=getOfficeImage)
btnRun.pack(pady=10)

# GUI 반복
win.mainloop()