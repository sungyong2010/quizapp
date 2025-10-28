import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import keyboard
import psutil
import pygetwindow as gw
import subprocess
import time
import logging
import win32gui
import win32process
import threading
import signal

"""
Description:
- 이 코드는 Google Sheets API를 사용하여 퀴즈 데이터를 가져옵니다.
- 퀴즈 데이터는 Google Sheets의 특정 시트에서 가져옵니다.
- GUI는 tkinter를 사용하여 전체 화면 모드로 구현됩니다.
- 사용자는 한글 단어에 대한 영어 정답을 입력해야 합니다.

Quiz data format in Google Sheets:
| 한글 단어 | 영어 정답 | 힌트 |
=> https://docs.google.com/spreadsheets/d/1BHkAT3j75_jq5qM5p1AZ73NaR4JhcxP7uBeWZRE0CD8/edit?usp=sharing

exe 배포 : 
python -O -m PyInstaller --onefile --windowed --add-data "quizapp-credentials.json;." quizapp.py
"""

# F1 키로 버전 정보 보기
def show_version():
    messagebox.showinfo("버전 정보", "QuizApp v1.1.0\n2025-10-28")
    # QuizApp v1.1.0 : Google Sheets 메시지 템플릿 기능 추가 개선
    # QuizApp v1.0.0 : Google Sheets 메시지 템플릿 기능 추가
    # QuizApp v0.8.0 : hidden code to exit program added
    # QuizApp v0.7.0 : 로블록스 프로세스 종료 기능 추가
    # QuizApp v0.6.0 : cmd.exe 에 대한 예외 처리 추가
    # QuizApp v0.5.0 : foreground 프로세스 종료 시 로그 기록 추가
    # QuizApp v0.4.0 : 프로세스 종료 로그 추가
    # QuizApp v0.3.0 : Google Sheets에서 오늘 날짜 시트를 불러오도록 수정


# 로그 설정
# 로그 디렉터리 확인 및 생성
log_dir = r"C:\temp"
os.makedirs(log_dir, exist_ok=True)  # 폴더가 없으면 자동 생성

# 기존 로그 파일 삭제 (존재하는 경우)
log_file_path = os.path.join(log_dir, "log.txt")
if os.path.exists(log_file_path):
    try:
        os.remove(log_file_path)
    except OSError:
        pass  # 파일 삭제 실패 시 무시

# 로그 설정
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filemode="w",  # 쓰기 모드로 새 파일 생성
)


# quizapp 실행시 Windows 키 차단
def block_windows_key():
    keyboard.block_key("left windows")
    keyboard.block_key("right windows")


def unblock_windows_key():
    keyboard.unblock_key("left windows")
    keyboard.unblock_key("right windows")


def on_closing():
    # 프로세스 모니터링 중지
    process_monitor.stop_monitoring()
    unblock_windows_key()
    root.destroy()


# Google Sheets API 인증 설정
# 퀴즈 데이터와 메시지 템플릿을 한 번에 가져오기
def fetch_quiz_and_message():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    def resource_path(relative_path):
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        resource_path("quizapp-credentials.json"), scope
    )
    client = gspread.authorize(creds)

    today = datetime.today().strftime("%Y-%m-%d")
    try:
        sheet = client.open("Shooting").worksheet(today)
    except gspread.exceptions.WorksheetNotFound:
        messagebox.showerror(
            "시트 없음", f"{today} 날짜의 퀴즈 시트가 존재하지 않습니다."
        )
        exit()

    data = sheet.get_all_records()
    quiz_data = [
        (
            row.get("한글 단어", ""),
            row.get("영어 정답", ""),
            row.get("힌트", "")
        )
        for row in data
    ]

    # 메시지 템플릿도 같이 가져오기
    COMMON_MSG = (
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "다음 영어 단어를 맞춰보세요({current_num}/{total_num}):\n\n"
        "'{korean_word}'"
    )
    try:
        msg_sheet = client.open("Shooting").worksheet("msg")
        msg = msg_sheet.cell(1, 1).value
        if msg and msg.strip():
            message_template = msg.strip() + "\n\n" + COMMON_MSG
        else:
            raise Exception
    except Exception:
        message_template = (
            "우리 준기가 오늘 외운 영어 단어로 언젠가\n"
            "외국 친구들과 웃으며 이야기하는 모습을 상상해봐.\n\n"
            "그 순간을 위해 지금 우리가 함께\n"
            "외국 친구들과 웃으며 이야기하는 모습을 상상해봐.\n\n"
            "그 순간을 위해 지금 우리가 함께 노력하고 있는 거야.\n"
            "힘들어도 아빠가 끝까지 함께 할게\n\n"
            + COMMON_MSG
        )

    return quiz_data, message_template

# 퀴즈 데이터 불러오기
quiz_data, quiz_message_template = fetch_quiz_and_message()
current_index = 0

# 정답 확인 함수
def check_answer():
    global current_index
    user_input = entry.get().strip()
    
    # 숨겨진 종료 코드 확인 (대소문자 구분 없음)
    if user_input == "dkfmaekdns":
        logging.info("숨겨진 종료 코드 입력됨 - 프로그램 종료")
        messagebox.showinfo("종료", "프로그램을 종료합니다.")
        process_monitor.stop_monitoring()
        # 종료 직전에 explorer.exe 복구
        # subprocess.Popen("explorer.exe")
        unblock_windows_key()
        root.destroy()
        sys.exit()
        return
    
    # 일반적인 정답 확인 (소문자로 변환)e
    user_input_lower = user_input.lower()
    correct_answer = quiz_data[current_index][1].lower()
    
    if user_input_lower == correct_answer:
        current_index += 1
        if current_index >= len(quiz_data):
            messagebox.showinfo("성공!", "모든 문제를 맞췄습니다!")
            process_monitor.stop_monitoring()
            # 종료 직전에 explorer.exe 복구
            # subprocess.Popen("explorer.exe")
            unblock_windows_key()
            root.destroy()
            sys.exit()
        else:
            update_question()
    else:
        messagebox.showerror("틀렸습니다", "정답이 아닙니다. 다시 시도하세요.")
        # 오답 후 입력 필드 초기화
        entry.delete(0, tk.END)


# 문제 업데이트
def update_question():
    entry.delete(0, tk.END)
    korean_word = quiz_data[current_index][0]
    hint = quiz_data[current_index][2]  # 힌트 컬럼

    # 현재 문제 번호와 전체 문제 수
    current_num = current_index + 1
    total_num = len(quiz_data)

    message = quiz_message_template.format(
        current_num=current_num,
        total_num=total_num,
        korean_word=korean_word
    )

    label.config(text=message)

    # 힌트 라벨 업데이트 (없으면 생성)
    hint_text = f"({hint})" if hint else ""
    if not hasattr(update_question, "hint_label"):
        update_question.hint_label = tk.Label(
            root,
            text=hint_text,
            font=("Arial", int(28 * 0.7)),  # 70% 크기
            fg="yellow",
            bg="black"
        )
        update_question.hint_label.pack()
    else:
        update_question.hint_label.config(text=hint_text)

    # 포커스 강제 설정 (지연 후 재적용 및 독점 포커스)
    entry.focus_set()
    entry.grab_set()  # 입력 필드가 포커스를 독점
    root.after(100, lambda: entry.focus_set())


# 디버그 모드 설정 (C언어의 #ifdef DEBUG와 유사)
# __debug__는 python -O로 실행시 False가 됨
DEBUG_MODE = __debug__

# 조기 프로세스 정리 (임포트 완료 즉시 실행)
def early_process_cleanup():
    """프로그램 로딩 중 조기 프로세스 정리"""
    try:
        # 간단한 프로세스 정리 (빠른 실행을 위해 최소화)
        unsafe_processes = ["cmd.exe"
                            , "notepad.exe"
                            # , "explorer.exe"
                            ]
        
        # DEBUG 모드가 아닌 경우에만 브라우저도 종료
        if not DEBUG_MODE:
            unsafe_processes.extend(["chrome.exe"
                                     , "firefox.exe"
                                     , "msedge.exe"
                                     , "powershell.exe"])
        
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'].lower() in unsafe_processes:
                    proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    except Exception:
        pass  # 에러 발생 시 무시하고 계속 진행

# 임포트 완료 즉시 조기 프로세스 정리 실행
early_process_cleanup()

# 포그라운드 프로세스 종료 함수 (먼저 정의)
def terminate_foreground_processes(safe_processes=None):
    if safe_processes is None:
        safe_processes = [
            "quizapp.exe"
            , "code.exe"
            , "windowsterminal.exe"
            , "wt.exe"
            , "openonsole.exe"
            # , "explorer.exe"
        ]
        
        # DEBUG 모드에서만 chrome.exe 허용 (C언어 #ifdef DEBUG와 유사)
        if DEBUG_MODE:
            safe_processes.append("chrome.exe")
            safe_processes.append("vsclient.exe")
            safe_processes.append("powershell.exe")

    # 로블록스 프로세스는 무조건 종료 대상
    BLOCKED_PROCESSES = [
        "robloxplayerbeta.exe", "roblox.exe", "robloxstudio.exe"
    ]

    # 현재 실행 중인 프로세스 이름도 보호
    current_process_name = psutil.Process(os.getpid()).name().lower()
    if current_process_name not in safe_processes:
        safe_processes.append(current_process_name)

    logging.info("### 포그라운드 프로세스 종료 시작")

    def enum_window_callback(hwnd, pid_list):
        if win32gui.IsWindowVisible(hwnd):
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                pid_list.add(pid)
            except Exception:
                pass

    visible_pids = set()
    win32gui.EnumWindows(enum_window_callback, visible_pids)

    for pid in visible_pids:
        try:
            proc = psutil.Process(pid)
            name = proc.name().lower()
            if name in BLOCKED_PROCESSES:
                proc.terminate()
                logging.info(f"로블록스 종료됨: {name} (PID: {pid})")
            if name not in safe_processes:
                proc.terminate()
                logging.info(f"종료됨: {name} (PID: {pid})")
            else:
                logging.info(f"유지됨: {name} (PID: {pid})")
        except Exception as e:
            logging.warning(f"종료 실패: PID {pid}, 오류: {e}")

    logging.info("### 포그라운드 프로세스 종료 완료")

# 백그라운드 프로세스 모니터링
class ProcessMonitor:
    def __init__(self):
        self.running = True
        self.monitor_thread = None
        
    def start_monitoring(self):
        """백그라운드에서 프로세스 모니터링 시작"""
        if not DEBUG_MODE:  # 릴리즈 모드에서만 모니터링
            self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
            self.monitor_thread.start()
            
    def stop_monitoring(self):
        """프로세스 모니터링 중지"""
        self.running = False
        
    def _monitor_loop(self):
        """프로세스 모니터링 루프"""
        unsafe_processes = ["cmd.exe"
                            , "notepad.exe"
                            , "chrome.exe"
                            , "firefox.exe"
                            , "msedge.exe"
                            # , "explorer.exe"
                            ]
        # 로블록스 프로세스는 무조건 종료 대상
        BLOCKED_PROCESSES = [
            "robloxplayerbeta.exe", "roblox.exe", "robloxstudio.exe"
        ]

        while self.running:
            try:
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        if proc.info['name'].lower() in BLOCKED_PROCESSES:
                            proc.terminate()
                            logging.info(f"모니터링: {proc.info['name']} 종료 (PID: {proc.info['pid']})")
                        if proc.info['name'].lower() in unsafe_processes:
                            proc.terminate()
                            logging.info(f"모니터링: {proc.info['name']} 종료 (PID: {proc.info['pid']})")
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                time.sleep(2)  # 2초마다 체크
            except Exception as e:
                logging.warning(f"모니터링 오류: {e}")
                time.sleep(5)

# 프로세스 모니터 생성
process_monitor = ProcessMonitor()

# 전체 화면 GUI 설정
root = tk.Tk()
root.title("한글 → 영어 단어 퀴즈")
# 전체 화면 대신 최대화로 변경 (포커스 호환성 향상)
root.state('zoomed')  # 'zoomed'는 최대화 모드
root.overrideredirect(True)  # 타이틀바 및 최소/최대/닫기 버튼 제거
root.protocol("WM_DELETE_WINDOW", on_closing)
block_windows_key()
root.configure(bg="black")

# ✅ 포커스 강제 설정
root.focus_force()

# 프로그램 시작 즉시 프로세스 종료 (보안 강화)
logging.info("프로그램 시작 - 즉시 프로세스 정리 실행")
terminate_foreground_processes()
# 프로세스 종료 후 포커스 재설정
root.after(500, lambda: entry.focus_set())  # 500ms 지연 후 포커스

# 백그라운드 모니터링 시작
process_monitor.start_monitoring()
logging.info("백그라운드 프로세스 모니터링 시작")

# 예: F1 키로 버전 정보 보기
root.bind("<F1>", lambda event: show_version())

# 상단 프레임 (X 버튼용)
top_frame = tk.Frame(root, bg="black")
top_frame.pack(fill="x", side="top")

# X 버튼 (우상단) - DEBUG 모드에서만 표시
def close_app():
    # 프로세스 모니터링 중지
    process_monitor.stop_monitoring()
    on_closing()

# DEBUG 모드에서만 X 버튼 생성
if DEBUG_MODE:
    close_button = tk.Button(top_frame, text="✕", font=("Arial", 20), 
                            fg="white", bg="red", activebackground="darkred",
                            command=close_app, width=3, height=1)
    close_button.pack(side="right", padx=10, pady=5)

label = tk.Label(root, text="", font=("Arial", 28), fg="white", bg="black")
label.pack(pady=80)

# 엔터키로도 정답 제출 할 수 있도록...
entry = tk.Entry(root, font=("Arial", 24))
entry.pack()
entry.bind("<Return>", lambda event: check_answer())

# 버튼을 클릭해서 정답 제출 할 수 있도록...
button = tk.Button(root, text="정답 제출", font=("Arial", 20), command=check_answer)
button.pack(pady=30)

# Entry 필드에 자동 포커스 설정
entry.focus_set()
root.after(200, lambda: entry.focus_set())  # 추가 지연 포커스

# Alt+F4 방지 (릴리즈 빌드에만 적용)
def disable_event():
    pass

# DEBUG 모드가 아닌 경우에만 Alt+F4 방지 적용
if not DEBUG_MODE:
    root.protocol("WM_DELETE_WINDOW", disable_event)
else:
    # DEBUG 모드에서는 정상적으로 창 닫기 허용
    root.protocol("WM_DELETE_WINDOW", on_closing)

if __name__ == "__main__":
    update_question()
    root.mainloop()
