import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import keyboard
import psutil
import subprocess
import time
import logging

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
# F1 키로 버전 정보 보기
def show_version():
    messagebox.showinfo("버전 정보", "QuizApp v0.4.0\n2025-10-26")
    # QuizApp v0.4.0 : 프로세스 종료 로그 추가
    # QuizApp v0.3.0 : Google Sheets에서 오늘 날짜 시트를 불러오도록 수정

# 로그 설정
# 로그 디렉터리 확인 및 생성
log_dir = r"C:\temp"
os.makedirs(log_dir, exist_ok=True)  # 폴더가 없으면 자동 생성

# 로그 설정
logging.basicConfig(
    filename=os.path.join(log_dir, "log.txt"),
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

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

# quizapp 실행 전의 프로세스를 모두 종료하고 quizapp을 실행하는 함수
SAFE_PROCESSES = [
    "quizapp.exe"
    , "explorer.exe"
    , "winlogon.exe"
    , "svchost.exe"
    , "system"
    , "services.exe"
    , "lsass.exe"
    , "Registry"
    , "smss.exe"
    , "csrss.exe"
    , "wininit.exe"
    , "wmiprvse.exe"
    , "vmms.exe"
    , "sihost.exe"
    , "aggregatorhost.exe"
    , "usbservice64.exe"
    , "securityhealthservice.exe"
    , "code.exe"
]
def launch_quizapp_and_close_others(app_path=r"C:\apps\quizapp\quizapp.exe", wait_time=3):
    """
    quizapp 실행 전의 프로세스를 모두 종료하고 quizapp을 실행합니다.
    
    Parameters:
    - app_path: 실행할 quizapp 경로
    - wait_time: quizapp 실행 후 대기 시간 (초)
    """
    # 1. 현재 실행 중인 프로세스 목록 저장
    before = {p.pid: p.info for p in psutil.process_iter(['name', 'create_time'])}

    # 2. quizapp 실행
    subprocess.Popen(app_path)

    # 3. 앱이 안정적으로 실행될 수 있도록 대기
    time.sleep(wait_time)

    # 4. quizapp 이전에 실행된 프로세스 종료
    logging.info("###프로세스 종료 작업 시작")
    for pid, info in before.items():
        try:
            proc = psutil.Process(pid)
            name = proc.name().lower()
            if name not in SAFE_PROCESSES:
                proc.terminate()
                logging.info(f"종료됨: {name} (PID: {pid})")
            else:
                logging.info(f"유지됨: {name} (PID: {pid})")
        except Exception as e:
            logging.warning(f"종료 실패: PID {pid}, 오류: {e}")
    logging.info("###프로세스 종료 작업 완료")
launch_quizapp_and_close_others()

update_question()
root.mainloop()
