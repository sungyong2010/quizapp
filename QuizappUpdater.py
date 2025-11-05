"""
<exe 배포>
작업 스케줄러에서 SYSTEM 계정으로 실행할 예정이면 
UAC 프롬프트 회피를 위해 일반(manifest 기본 asInvoker)로 빌드:
==========================================================
python -O -m PyInstaller --onefile `
    --noconsole `
    --name QuizappUpdater `
    quizappupdater.py
"""
import os
import sys
import json
import time
import shutil
import ctypes
import logging
import tempfile
import subprocess
from datetime import datetime

try:
    import requests
    import psutil
    import win32security
    import ntsecuritycon
except ImportError:
    print("Required modules missing. Please install requests and psutil.")
    sys.exit(1)

REMOTE_VERSION_URL = "https://raw.githubusercontent.com/sungyong2010/anything-to-share/main/Quizapp/version.json"
REMOTE_EXE_URL = "https://raw.githubusercontent.com/sungyong2010/anything-to-share/main/Quizapp/Quizapp.exe"
TARGET_EXE_PATH = r"c:\Apps\quizapp\Quizapp.exe"
LOCAL_VERSION_JSON = r"c:\Apps\quizapp\version.json"  # Store local version for comparison
BACKUP_DIR = r"c:\Apps\quizapp\backup"
LOG_PATH = r"c:\temp\quizappupdater.log"
SCHEDULED_TASK_NAME = r"\QuizApp 자동 실행"

# 기존 로그 파일이 있으면 삭제
if os.path.exists(LOG_PATH):
    try:
        os.remove(LOG_PATH)
    except Exception as e:
        print(f"Failed to remove old log file: {e}")

# 로그 디렉토리 생성 및 로깅 설정
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)


def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def show_popup(message: str, title: str = "QuizApp Updater", timeout_ms: int = 2500, flags: int = 0x00000040):
    """Show a transient informational popup.

    Attempts to use MessageBoxTimeoutW (undocumented but widely available) so it auto-closes.
    Falls back to MessageBoxW if timeout variant unavailable (user must dismiss then).

    flags default adds MB_ICONINFORMATION (0x40). Caller can pass other MB_* flags OR them in.
    """
    try:
        user32 = ctypes.windll.user32
        # Try timeout version: int MessageBoxTimeoutW(HWND, LPCWSTR, LPCWSTR, UINT, WORD, DWORD)
        try:
            MessageBoxTimeoutW = user32.MessageBoxTimeoutW
            MessageBoxTimeoutW(None, message, title, flags, 0, timeout_ms)
            return
        except AttributeError:
            pass
        # Fallback standard message box (modal)
        user32.MessageBoxW(None, message, title, flags)
    except Exception as e:
        logging.warning(f"Popup failed: {e} | message={message}")


def read_local_version() -> str:
    if not os.path.exists(LOCAL_VERSION_JSON):
        return "0.0.0"
    try:
        with open(LOCAL_VERSION_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("latest_version", "0.0.0")
    except Exception as e:
        logging.warning(f"Failed reading local version: {e}")
        return "0.0.0"


def fetch_remote_version():
    try:
        resp = requests.get(REMOTE_VERSION_URL, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        return data.get("latest_version"), data.get("download_url", REMOTE_EXE_URL)
    except Exception as e:
        logging.error(f"Remote version fetch failed: {e}")
        return None, None


def version_tuple(v: str):
    return tuple(int(part) for part in v.split('.') if part.isdigit())


def needs_update(local_v: str, remote_v: str) -> bool:
    try:
        return version_tuple(remote_v) > version_tuple(local_v)
    except Exception:
        return remote_v != local_v


def terminate_running_exe(exe_path: str, process_name: str = "quizapp.exe", timeout: int = 10):
    logging.info("Checking for running quizapp.exe processes to terminate")
    end_time = time.time() + timeout
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            name = (proc.info.get('name') or '').lower()
            exe = (proc.info.get('exe') or '').lower()
            if name == process_name.lower() or exe == exe_path.lower():
                logging.info(f"Terminating PID {proc.pid} ({name})")
                proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            logging.warning(f"Process termination issue: {e}")
    # Wait for termination
    while time.time() < end_time:
        alive = False
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                name = (proc.info.get('name') or '').lower()
                exe = (proc.info.get('exe') or '').lower()
                if name == process_name.lower() or exe == exe_path.lower():
                    alive = True
                    break
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        if not alive:
            logging.info("All quizapp.exe processes terminated")
            return True
        time.sleep(0.5)
    logging.error("Timeout waiting for quizapp.exe to terminate")
    return False


def backup_existing(exe_path: str):
    if not os.path.exists(exe_path):
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"quizapp_{timestamp}.exe")
    try:
        shutil.copy2(exe_path, dest)
        logging.info(f"Backup created: {dest}")
    except Exception as e:
        logging.warning(f"Backup failed: {e}")


def download_new_exe(url: str) -> str:
    logging.info(f"Downloading new quizapp.exe from {url}")
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, "quizapp_new.exe")
        with open(temp_path, 'wb') as f:
            f.write(resp.content)
        logging.info(f"Downloaded to {temp_path}")
        return temp_path
    except Exception as e:
        logging.error(f"Download failed: {e}")
        return ""


def replace_exe(new_path: str, target_path: str) -> bool:
    if not new_path or not os.path.exists(new_path):
        logging.error("New exe path invalid")
        return False
    
    try:
        # Remove target if exists with retry
        for _ in range(10):
            if os.path.exists(target_path):
                try:
                    os.remove(target_path)
                    break
                except PermissionError:
                    time.sleep(0.5)
            else:
                break
        
        # 파일 복사
        shutil.move(new_path, target_path)
        
        # 새로운 exe 파일에 권한 설정
        if set_permissions(target_path):
            logging.info(f"Replaced {target_path} and set permissions successfully")
            return True
        else:
            logging.error(f"Failed to set permissions on {target_path}")
            return False
    except Exception as e:
        logging.error(f"Replace failed: {e}")
        return False


def set_permissions(path):
    """Sets appropriate permissions on the specified file or directory:
    - SYSTEM and Administrators get full control
    - Everyone gets only read and execute permissions
    """
    try:
        import win32security
        import ntsecuritycon as con
        
        # Get SIDs for different security principals
        everyone = win32security.ConvertStringSidToSid("S-1-1-0")  # Everyone
        system = win32security.ConvertStringSidToSid("S-1-5-18")   # SYSTEM
        admins = win32security.ConvertStringSidToSid("S-1-5-32-544")  # Administrators
        
        # Get current security descriptor
        sd = win32security.GetFileSecurity(path, win32security.DACL_SECURITY_INFORMATION)
        dacl = win32security.ACL()
        
        # Add full control for SYSTEM and Administrators
        dacl.AddAccessAllowedAce(win32security.ACL_REVISION, con.FILE_ALL_ACCESS, system)
        dacl.AddAccessAllowedAce(win32security.ACL_REVISION, con.FILE_ALL_ACCESS, admins)
        
        # Add only read and execute permissions for Everyone
        everyone_permissions = (
            # Basic read permissions
            con.GENERIC_READ |
            # Execute permissions
            con.GENERIC_EXECUTE |
            # Additional permissions
            con.SYNCHRONIZE |
            # Specific file permissions
            con.FILE_READ_DATA |
            con.FILE_READ_ATTRIBUTES |
            con.FILE_READ_EA |
            con.FILE_EXECUTE
        )
        dacl.AddAccessAllowedAce(win32security.ACL_REVISION, everyone_permissions, everyone)
        
        # Set the new DACL
        sd.SetSecurityDescriptorDacl(1, dacl, 0)
        win32security.SetFileSecurity(path, win32security.DACL_SECURITY_INFORMATION, sd)
        
        logging.info(f"Successfully set permissions for {path}")
        return True
    except Exception as e:
        logging.error(f"Failed to set permissions for {path}: {e}")
        return False

def write_local_version(version: str):
    try:
        data = {"latest_version": version, "updated_at": datetime.utcnow().isoformat()}
        with open(LOCAL_VERSION_JSON, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logging.info(f"Local version updated to {version}")
    except Exception as e:
        logging.error(f"Writing local version failed: {e}")


def launch_app():
    # Try scheduled task trigger first
    try:
        result = subprocess.run(
            ["schtasks", "/run", "/TN", SCHEDULED_TASK_NAME],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            logging.info(f"Scheduled task triggered: {SCHEDULED_TASK_NAME}")
            return
        else:
            logging.warning(f"schtasks failed (rc={result.returncode}): {result.stderr.strip()}")
    except Exception as e:
        logging.warning(f"schtasks exception: {e}")
    # Fallback: direct start (권한 상승된 상태로 실행될 수 있음)
    try:
        os.startfile(TARGET_EXE_PATH)
        logging.info("Fallback direct launch executed")
    except Exception as e:
        logging.error(f"Fallback launch failed: {e}")


def main():
    logging.info("Updater started")
    # show_popup("업데이트 확인 중...", timeout_ms=1800)
    if not is_admin():
        logging.error("Updater not running with admin rights. Exiting.")
        show_popup("관리자 권한 아님 - 종료", timeout_ms=2500, flags=0x00000010)  # icon hand
        return
    local_v = read_local_version()
    remote_v, download_url = fetch_remote_version()
    if not remote_v:
        logging.info("Remote version unavailable. Launching existing app.")
        show_popup("원격 버전 정보 없음. 앱 실행", timeout_ms=2500, flags=0x00000030)  # warning icon
        launch_app()
        return
    if needs_update(local_v, remote_v):
        logging.info(f"Update needed. Local={local_v} Remote={remote_v}")
        show_popup(f"새 버전 {remote_v} 업데이트 진행", timeout_ms=2200)
        if not terminate_running_exe(TARGET_EXE_PATH):
            logging.error("Cannot terminate running instance. Abort update.")
            show_popup("실행 중 프로세스 종료 실패 - 기존 앱 실행", timeout_ms=3000, flags=0x00000030)
            launch_app()
            return
        backup_existing(TARGET_EXE_PATH)
        show_popup("다운로드 중...", timeout_ms=2000)
        new_path = download_new_exe(download_url)
        if replace_exe(new_path, TARGET_EXE_PATH):
            write_local_version(remote_v)
            show_popup("업데이트 완료! 앱 실행", timeout_ms=2500)
        else:
            logging.error("Update failed during replace phase")
            show_popup("업데이트 실패 - 기존 앱 실행", timeout_ms=3000, flags=0x00000030)
    else:
        logging.info(f"No update needed. Local={local_v} Remote={remote_v}")
        # show_popup("최신 버전입니다. 앱 실행", timeout_ms=2000)
    launch_app()


if __name__ == "__main__":
    main()