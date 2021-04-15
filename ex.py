from io import BytesIO
import win32clipboard  # pywin32
from PIL import Image

import pyautogui as pag
import keyboard
import time
from datetime import datetime

import cv2 # opencv-python
import pytesseract

import tkinter as tk
import threading

# EXTENSION
# while True : 무한반복 = 무한대기
# if keyboard.is_pressed("q") : 특정 키가 눌렸을 때 실행되게 하기위해
# pag.position() : 마우스 위치 좌표 뽑아내기

filepath = "./captureimg.png"

def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def setKeybtn():
    print("setKey")

    img = tk.PhotoImage(file = filepath)
    # leftlabel.configure(image=img)
    # leftlabel.image = img
    # https://python-forum.io/Thread-Tkinter-How-do-I-change-an-image-dynamically

def setColorbtn(color):
    print(color)
    rightpanel.configure(bg=color)

def startbtn(): # 시작 버튼 눌렀을 때, "캡쳐 이미지 -> OCR Text로 변경" 실행
    print("start")
    while True:
        fail = 0
        while True:
            if keyboard.is_pressed("q"):
                start_t = datetime.now()  # 키를 눌렀을 때의 시간
                start = pag.position()  # 시작점 저장
                print(start)
                time.sleep(0.2)
                break

        while True:
            now_t = datetime.now()
            if (now_t - start_t).total_seconds() >= 5:  # 일정 시간이 지났을 때 캡쳐 아닌 것으로 판단
                fail = 1
                break
            if keyboard.is_pressed("q"):
                end = pag.position()  # 끝 점 저장
                print(end)
                time.sleep(0.2)
                break

        if (fail == 0) \
                and (start[0] < end[0]) \
                and (start[1] < end[1]):
            xx = start[0]
            yy = start[1]
            ww = end[0] - start[0]
            hh = end[1] - start[1]
            pag.screenshot(filepath, region=(xx, yy, ww, hh))

            # 이미지 저장
            image = Image.open(filepath)
            output = BytesIO()
            image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()

            # 클립보드에 저장
            send_to_clipboard(win32clipboard.CF_DIB, data)

            # 이미지 OCR 분석
            # 1. Github
            # https://tesseract-ocr.github.io/
            # https://github.com/UB-Mannheim/tesseract/wiki - Binaries, Windows - Tesseract at UB Mannheim

            # 2. Github에서 exe 파일 다운 방법
            # https://junyoung-jamong.github.io/computer/vision,/ocr/2019/01/30/Python%EC%97%90%EC%84%9C-Tesseract%EB%A5%BC-%EC%9D%B4%EC%9A%A9%ED%95%B4-OCR-%EC%88%98%ED%96%89%ED%95%98%EA%B8%B0.html

            # 3. 출처
            # https://niceman.tistory.com/155
            # - 텍스트 추출 정확도는 이미지에 따라서 크게 차이날 수 있다.
            # - 텍스트가 잘 정리되고 나열된 상태의 이미지라면 더 정확한 인식이 가능하다.

            # 4. 사용법
            # tesseract -c preserve_interword_spaces=1 kor.png stdout -l kor > kor.txt
            # tesseract -c preserve_interword_spaces=1 koeng.png stdout -l kor+eng > koeng.txt 등

            # OCR 정확도 개선,
            # https://junyoung-jamong.github.io/computer/vision,/ocr/2019/01/30/Python%EC%97%90%EC%84%9C-Tesseract%EB%A5%BC-%EC%9D%B4%EC%9A%A9%ED%95%B4-OCR-%EC%88%98%ED%96%89%ED%95%98%EA%B8%B0.html
            image = cv2.imread(filepath)
            grayImg = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            grayImg = cv2.threshold(grayImg, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
            cv2.imwrite(filepath, grayImg)

            # OCR, Image -> Text 변경
            pytesseract.pytesseract.tesseract_cmd = r'.\tresseract\tesseract.exe'
            text = pytesseract.image_to_string(Image.open(filepath), lang='eng+kor')

            # 프로그램 왼쪽에 이미지 띄워주기
            img = tk.PhotoImage(file=filepath)
            leftpanel.configure(image=img)
            leftpanel.image = img
            # https://stackoverflow.com/questions/3482081/how-to-update-the-image-of-a-tkinter-label-widget

            # 오른쪽에 텍스트 띄워주기
            # label:
            #   rightpanel.configure(text=text)
            #   rightpanel.text = text
            # text:
            rightpanel.delete(0, tk.END)
            rightpanel.insert(tk.INSERT, text)

            # pyinstaller라는 라이브러리 활용해서 exe 파일로 만들 수 있습니다.
            # icon 형태로 실행 파일을 만들어서 더블 클릭하면 계속 이 프로그램이 실행되게끔 할 수 있습니다.
            # 그런 상태에서 q, q를 이용해서 캡쳐를 하고
            # ctrl + V를 하게 되면
            # 우리가 쉽게 print screen capture를 할 수 있는 프로그램을 쓸 수 있는 것입니다.

if __name__ == '__main__':
    # GUI
    app = tk.Tk()
    app.title("OCR")
    app.geometry("1000x400+100+100") # 윈도우 창의 크기, x축, y축
    app.resizable(True, True) # 최대화 가능 범위(width, height)

    menubar = tk.Menu(app)

    # Sub Menuqq
    submenu = tk.Menu(menubar, tearoff=0)
    submenu.add_command(label="Key", command=setKeybtn)

    # Color sub menu
    colormenu = tk.Menu(submenu, tearoff=0)
    colormenu.add_command(label='gold', command=lambda:setColorbtn('gold'))
    colormenu.add_command(label='blue', command=lambda:setColorbtn('skyblue'))
    colormenu.add_command(label='pink', command=lambda:setColorbtn('pink'))
    colormenu.add_command(label='gray', command=lambda:setColorbtn('lightgray'))
    colormenu.add_command(label='white', command=lambda:setColorbtn('white'))
    submenu.add_cascade(label="Color", menu=colormenu)

    # Menu
    menubar.add_cascade(label="Start", command=threading.Thread(target=startbtn).start())
    menubar.add_cascade(label="Setting", menu=submenu)

    app.config(menu=menubar)

    # Left label - Capture Image
    leftpanel = tk.Label(app, width=400, bg="black",
                         text="단축키를 통해\n텍스트로 번역할 이미지를\n시작과 끝점을 이용하여 캡쳐 해주세요.\n(기본 단축키는 q입니다.)", fg="white",
                         font=("Times", "16"))
    leftpanel.pack(side="left", anchor="nw", fill="y", expand="yes")

    # Right label - OCR text
    rightpanel = tk.Entry(app, width=600, bg="gold", justify=tk.CENTER, font=("Times", "16"))
    rightpanel.pack(side="left", anchor="ne", fill="y", expand="yes")

    app.mainloop()

# 기록
# auto-py-to-exe
# https://stackoverflow.com/questions/54993317/how-to-compile-python-code-with-pytesseract-to-exe

# 기능 추가
# 1. 단축키
# 2. font
# 3. Machine learning
# https://niceman.tistory.com/157?category=1009824
