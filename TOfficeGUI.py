import tkinter as tk
from tkinter import ttk
from tkinter import font
import pandas as pd
from tkinter import filedialog
import TOffice as tof

### 기준정보로 등록해야 하는 정보
## GUI 타이틀
win_title = "Tricks-Office Automation"

## GUI 크기
Gsize = "680x580"

## 폴더 / 파일 리스트
# 각 행의 라벨 정보 List로 정리
label_f = ['Target', 'Result','Format','Dictionary']
# 각 행에서 다루는 값이 폴더일때는 0, 파일일때는 1로 구분자
fileyn = [0, 0, 1, 1]
# 폴더 / 파일 리스트 기준정보 불러오기 (기준정보 관리 파일명 : GUI_Master.csv)
# 기준정보 파일이 없을 경우 초기화
try : 
    df_FileFolder= pd.read_csv("TOffice_Path_Master.csv")
except :
    d = {'Item' : ['Please select file / folder path'] * len(label_f)}
    df_FileFolder = pd.DataFrame(data=d)

## 파라미터 텍스트 상자 리스트
label_para = ['Eng Font', 'Korean Font']
# 폴더 / 파일 리스트 기준정보 불러오기 (기준정보 관리 파일명 : Para.csv)
# 기준정보 파일이 없을 경우 초기화
try : 
    dfP= pd.read_csv("TOffice_Font_Master.csv")
except :
    d = {'Item' : ['Arial'] * len(label_para)}
    dfP = pd.DataFrame(data=d)


### GUI용 함수
# GUIMaster Data 업데이트 : 
def update_Master(idx, var):
    df_FileFolder.Item[idx] = var
    df_FileFolder.to_csv('TOffice_Path_Master.csv', index=False)

# Para Data 업데이트 :
def update_Para(idx, var):
    dfP.Item[idx] = var
    dfP.to_csv('TOffice_Font_Master.csv', index=False)


# 폴더/파일 경로 바꾸는 버튼을 눌렀을때 업데이트
def onClick(i, fileYN):
    # 폴더 경로 바꾸는 로직 (fineYN = 0 일때)
    if fileYN == 0:
        folder_selected = filedialog.askdirectory()
        var = folder_selected
    # 파일 경로 바꾸는 로직 (fineYN = 0 이 아닐때)
    else:
        folder_selected = filedialog.askopenfile()
        var = folder_selected.name

    txtPath[i].delete('1.0', tk.END)
    txtPath[i].insert(tk.INSERT, chars=var)
    update_Master(i,var)
    
## Main Code
# GUI 구성
win = tk.Tk()
win.geometry(Gsize)
win.title(win_title)

# Frame 설정하기 (frame1 : 폴더, frame2 : 파일, frame3 : 글꼴, frame4 : 암호)
frame1 = tk.LabelFrame(win, text = '폴더', padx = 5, pady=5, width = 100)
frame2 = tk.LabelFrame(win, text = '파일', padx = 5, pady=5, width = 100)
frame3 = tk.LabelFrame(win, text = '글꼴', padx = 5, pady=5)
frame4 = tk.Frame(win)
frame5 = tk.Frame(win)
frame1.pack(padx = 10, pady=10)
frame2.pack(padx = 10, pady=10)
frame3.pack(padx = 10, pady=10)
frame4.pack(padx = 10, pady=10)
frame5.pack(padx = 10, pady=10)

# 파일 / 폴더 경로 설정 GUI
lbName = []
txtPath = []
btnPath =[]

for i,x in enumerate(label_f):
    match i:
        case 0 | 1:
            lbName.append(tk.Label(frame1, text=x, width=15,padx =5, pady = 5))
            txtPath.append(tk.Text(frame1, width = 50, height = 1, padx =5, pady = 5, background='lightgrey'))
            btnPath.append(tk.Button(frame1, text="Change Path", width=10, padx =5, pady = 5, command=lambda i=i: onClick(i,fileyn[i])))
        case 2 | 3:
            lbName.append(tk.Label(frame2, text=x, width=15,padx =5, pady = 5))
            txtPath.append(tk.Text(frame2, width = 50, height = 1, padx =5, pady = 5, background='lightgrey'))
            btnPath.append(tk.Button(frame2, text="Change Path", width=10, padx =5, pady = 5, command=lambda i=i: onClick(i,fileyn[i])))

    # 폴더/파일 이름 초기값 넣기
    txtPath[i].insert(tk.INSERT, chars=df_FileFolder.Item[i])

    lbName[i].grid(row=i, column=0, padx =5, sticky=tk.W)
    txtPath[i].grid(row=i, column=1, padx =5, sticky=tk.W)
    btnPath[i].grid(row=i, column=2, padx =5, sticky=tk.W)

# 글꼴 구역의 각 구성 요소 생성하기
lbEFont = tk.Label(frame3, text = '영문 글꼴', width = 10, padx =5, pady =5)
cbEFont = ttk.Combobox(frame3, width = 25)
lbHFont = tk.Label(frame3, text = '한글 글꼴', width = 10, padx =5, pady =5)
cbHFont = ttk.Combobox(frame3, width = 25)

# 폰트 정보 콤보 박스에 리스트 추가하고 readonly 속성주기
fonts = list(font.families())
fonts.sort()
cbEFont['values'] = fonts
cbHFont['values'] = fonts
cbEFont['state'] = 'readonly'
cbHFont['state'] = 'readonly'

# 글꼴 초깃값 불러오기
cbEFont.current(fonts.index(dfP['Item'][0]))
cbHFont.current(fonts.index(dfP['Item'][1]))

# 글꼴 변경시 master data 저장하기
cbEFont.bind('<<ComboboxSelected>>', lambda _ : update_Para(0, cbEFont.get()))
cbHFont.bind('<<ComboboxSelected>>', lambda _ : update_Para(1, cbHFont.get()))

# 글꼴 구역의 각 구성 요소 배치하기
lbEFont.grid(row=0, column=0, padx =5, pady =5)
cbEFont.grid(row=0, column=1, padx =5, pady =5)
lbHFont.grid(row=0, column=2, padx =5, pady =5)
cbHFont.grid(row=0, column=3, padx =5, pady =5)


# 암호 구역의 각 구성 요소 생성하기
lbpw = tk.Label(frame4, text = '암호', width = 10, padx =5, pady =5)
txtpw = tk.Text(frame4, width = 55, height = 1, padx = 5, pady=5)

# 암호 구역의 각 구성 요소 배치하기
lbpw.grid(row=0, column=0, padx =5, pady =5)
txtpw.grid(row=0, column=1, padx =5, pady =5)

# 암호 초기값 입력
txtpw.insert(tk.INSERT, chars='Pas$W0rd')

# 버튼 생성하기
btnSaveImg = tk.Button(frame5, text = '문서 파일 이미지 저장', width = 25, padx=5, pady=5, 
                        command = lambda : tof.D_Img(df_FileFolder['Item'][0], df_FileFolder['Item'][1]))
btnMergeXl = tk.Button(frame5, text = '같은 형식 엑셀 병합', width = 25, padx=5, pady=5, 
                        command = lambda : tof.C_xl(df_FileFolder['Item'][0], df_FileFolder['Item'][1]))
btnMergeXl2 = tk.Button(frame5, text = '유사 형식 엑셀 병합', width = 25, padx=5, pady=5, 
                        command = lambda : tof.M_xl(df_FileFolder['Item'][0], df_FileFolder['Item'][1], df_FileFolder['Item'][2]))
btnPptFont = tk.Button(frame5, text = 'PPT 폰트 통일 (글꼴)', width = 25, padx=5, pady=5, 
                        command = lambda : tof.PptFont(df_FileFolder['Item'][0], df_FileFolder['Item'][1], dfP['Item'][0],
                        dfP['Item'][1]))
btnPptDic = tk.Button(frame5, text = 'PPT 단어 일괄 변경', width = 25, padx=5, pady=5, 
                        command = lambda : tof.PPT_Dic(df_FileFolder['Item'][0], df_FileFolder['Item'][1], 
                        df_FileFolder['Item'][3]))
btnWordDic = tk.Button(frame5, text = '워드 단어 일괄 변경', width = 25, padx=5, pady=5, 
                        command = lambda : tof.Doc_Dic(df_FileFolder['Item'][0], df_FileFolder['Item'][1],
                        df_FileFolder['Item'][3]))
btnDocPDF = tk.Button(frame5, text = '오피스 문서 PDF 전환', width = 25, padx=5, pady=5, 
                        command = lambda : tof.D_PDF(df_FileFolder['Item'][0], df_FileFolder['Item'][1]))
btnImgPDF = tk.Button(frame5, text = '이미지 PDF 전환', width = 25, padx=5, pady=5, 
                        command = lambda : tof.Img_PDF(df_FileFolder['Item'][0], df_FileFolder['Item'][1]))
btnPDFpw = tk.Button(frame5, text = 'PDF 암호 설정 (암호)', width = 25, padx=5, pady=5, 
                        command = lambda : tof.pw_PDF(df_FileFolder['Item'][0], df_FileFolder['Item'][1], 
                        txtpw.get('1.0','end-1c')))


# 버튼 배치하기
btnSaveImg.grid(row=0, column=0, padx =5, pady =5)
btnMergeXl.grid(row=0, column=1, padx =5, pady =5)
btnMergeXl2.grid(row=0, column=2, padx =5, pady =5)
btnPptFont.grid(row=1, column=0, padx =5, pady =5)
btnPptDic.grid(row=1, column=1, padx =5, pady =5)
btnWordDic.grid(row=1, column=2, padx =5, pady =5)
btnDocPDF.grid(row=2, column=0, padx =5, pady =5)
btnImgPDF.grid(row=2, column=1, padx =5, pady =5)
btnPDFpw.grid(row=2, column=2, padx =5, pady =5)


win.mainloop()