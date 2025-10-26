import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import keyboard

"""
Description:
- 이 코드는 Google Sheets API를 사용하여 퀴즈 데이터를 가져옵니다.
- 퀴즈 데이터는 Google Sheets의 특정 시트에서 가져옵니다.
- GUI는 tkinter를 사용하여 전체 화면 모드로 구현됩니다.
- 사용자는 한글 단어에 대한 영어 정답을 입력해야 합니다.

Quiz data format in Google Sheets:
| 한글 단어 | 영어 정답 |
=> https://docs.google.com/spreadsheets/d/1BHkAT3j75_jq5qM5p1AZ73NaR4JhcxP7uBeWZRE0CD8/edit?usp=sharing

exe 배포 : 
pyinstaller --onefile --windowed --add-data "quizapp-credentials.json;." quizapp.py
"""
# quizapp 실행시 Windows 키 차단
def block_windows_key():
    keyboard.block_key('left windows')
    keyboard.block_key('right windows')

def unblock_windows_key():
    keyboard.unblock_key('left windows')
    keyboard.unblock_key('right windows')

def on_closing():
    unblock_windows_key()
    root.destroy()

# F1 키로 버전 정보 보기
def show_version():
    messagebox.showinfo("버전 정보", "QuizApp v0.3.0\n2025-10-26")
    # QuizApp v0.3.0 : Google Sheets에서 오늘 날짜 시트를 불러오도록 수정

# Google Sheets API 인증 설정
def fetch_quiz_data():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    # creds = ServiceAccountCredentials.from_json_keyfile_name("quizapp-credentials.json", scope)    
    def resource_path(relative_path):
        """PyInstaller 환경에서도 파일 경로를 안전하게 가져오기"""
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    creds = ServiceAccountCredentials.from_json_keyfile_name(resource_path("quizapp-credentials.json"), scope)
    
    client = gspread.authorize(creds)
    # Google Sheets 문서 이름과 시트 이름
    # 오늘 날짜 기반 시트 이름
    today = datetime.today().strftime("%Y-%m-%d")
    try:
        sheet = client.open("Shooting").worksheet(today)
    except gspread.exceptions.WorksheetNotFound:
        messagebox.showerror("시트 없음", f"{today} 날짜의 퀴즈 시트가 존재하지 않습니다.")
        exit()

    data = sheet.get_all_records()

    # 퀴즈 데이터: [(한글, 영어)]
    return [(row['한글 단어'], row['영어 정답']) for row in data]

# 퀴즈 데이터 불러오기
quiz_data = fetch_quiz_data()
current_index = 0

# 정답 확인 함수
def check_answer():
    global current_index
    user_input = entry.get().strip().lower()
    correct_answer = quiz_data[current_index][1].lower()
    if user_input == correct_answer:
        current_index += 1
        if current_index >= len(quiz_data):
            messagebox.showinfo("성공!", "모든 문제를 맞췄습니다!")
            root.destroy()
        else:
            update_question()
    else:
        messagebox.showerror("틀렸습니다", "정답이 아닙니다. 다시 시도하세요.")

# 문제 업데이트
def update_question():
    entry.delete(0, tk.END)
    korean_word = quiz_data[current_index][0]
    label.config(text=f"다음 영어 단어를 맞춰보세요:\n\n'{korean_word}'")

# 전체 화면 GUI 설정
root = tk.Tk()
root.title("한글 → 영어 단어 퀴즈")
root.attributes('-fullscreen', True)
# root.geometry("800x600")
root.protocol("WM_DELETE_WINDOW", on_closing)
block_windows_key()
root.configure(bg='black')

# 예: F1 키로 버전 정보 보기
root.bind("<F1>", lambda event: show_version())

label = tk.Label(root, text="", font=("Arial", 28), fg="white", bg="black")
label.pack(pady=80)

# 엔터키로도 정답 제출 할 수 있도록...
entry = tk.Entry(root, font=("Arial", 24))
entry.pack()
entry.bind('<Return>', lambda event: check_answer())

# 버튼을 클릭해서 정답 제출 할 수 있도록...
button = tk.Button(root, text="정답 제출", font=("Arial", 20), command=check_answer)
button.pack(pady=30)

# Alt+F4 방지
def disable_event():
    pass
root.protocol("WM_DELETE_WINDOW", disable_event)

update_question()
root.mainloop()
