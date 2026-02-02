import subprocess
import pyautogui
import time as tm
import cv2
import numpy as np
import mss
import pandas as pd
import pyperclip
import os
import json
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from screeninfo import get_monitors

# Windows 레지스트리 접근을 위한 winreg (Windows 전용)
try:
    import winreg
except ImportError:
    winreg = None  # Windows가 아닌 경우

# MSS 모듈 설치 및 임포트 확인
try:
    import mss
except ModuleNotFoundError:
    subprocess.check_call(["pip", "install", "mss"])
    import mss

# screeninfo 모듈 설치 및 임포트
try:
    from screeninfo import get_monitors
except ModuleNotFoundError:
    subprocess.check_call(["pip", "install", "screeninfo"])
    from screeninfo import get_monitors

# 카카오톡 경로 자동 탐색 함수
def find_kakao_path():
    """
    카카오톡 실행 파일 경로를 자동으로 찾는 함수
    여러 경로를 순차적으로 확인하여 존재하는 경로를 반환
    """
    # 확인할 경로 목록
    possible_paths = [
        # 일반적인 설치 경로
        'C:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe',
        'C:/Program Files/Kakao/KakaoTalk/KakaoTalk.exe',
        'D:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe',
        'D:/Program Files/Kakao/KakaoTalk/KakaoTalk.exe',
        # 사용자별 AppData 경로
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Kakao/KakaoTalk/KakaoTalk.exe'),
        os.path.join(os.environ.get('APPDATA', ''), 'Kakao/KakaoTalk/KakaoTalk.exe'),
    ]
    
    # 레지스트리에서 찾기 (Windows)
    if winreg is not None:
        try:
            # HKEY_LOCAL_MACHINE에서 찾기
            reg_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            ]
            
            for hkey, reg_path in reg_paths:
                try:
                    key = winreg.OpenKey(hkey, reg_path)
                    for i in range(winreg.QueryInfoKey(key)[0]):
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            subkey = winreg.OpenKey(key, subkey_name)
                            try:
                                display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                if "KakaoTalk" in display_name or "카카오톡" in display_name:
                                    install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                    if install_location:
                                        kakao_exe = os.path.join(install_location, "KakaoTalk.exe")
                                        if os.path.exists(kakao_exe):
                                            winreg.CloseKey(subkey)
                                            winreg.CloseKey(key)
                                            print(f"[DEBUG] 레지스트리에서 카카오톡 경로 발견: {kakao_exe}")
                                            return kakao_exe
                            except (FileNotFoundError, OSError):
                                pass
                            finally:
                                winreg.CloseKey(subkey)
                        except (OSError, WindowsError):
                            continue
                    winreg.CloseKey(key)
                except (OSError, WindowsError):
                    continue
        except Exception as e:
            print(f"[DEBUG] 레지스트리 검색 중 오류 (무시): {e}")
    
    # 일반 경로들 확인
    for path in possible_paths:
        if path and os.path.exists(path):
            print(f"[DEBUG] 카카오톡 경로 발견: {path}")
            return path
    
    # 시작 메뉴에서 찾기
    try:
        start_menu_paths = [
            os.path.join(os.environ.get('APPDATA', ''), r'Microsoft\Windows\Start Menu\Programs'),
            os.path.join(os.environ.get('PROGRAMDATA', ''), r'Microsoft\Windows\Start Menu\Programs'),
        ]
        
        for start_menu in start_menu_paths:
            if os.path.exists(start_menu):
                for root, dirs, files in os.walk(start_menu):
                    for file in files:
                        if file.lower() == 'kakaotalk.lnk' or 'kakaotalk' in file.lower():
                            # .lnk 파일이므로 실제 경로를 찾아야 함
                            # 간단하게 KakaoTalk 폴더를 찾아서 확인
                            pass
    except Exception as e:
        print(f"[DEBUG] 시작 메뉴 검색 중 오류 (무시): {e}")
    
    return None

# 경로 및 변수 설정
try:
    with open('kakao_path.txt', 'r') as f:
        KAKAO_PATH = f.read().strip()
        if not os.path.exists(KAKAO_PATH):
            print(f"[DEBUG] kakao_path.txt에 지정된 경로가 존재하지 않습니다: {KAKAO_PATH}")
            print("[DEBUG] 자동으로 카카오톡 경로를 찾는 중...")
            found_path = find_kakao_path()
            if found_path:
                KAKAO_PATH = found_path
                print(f"[DEBUG] 자동으로 찾은 경로 사용: {KAKAO_PATH}")
            else:
                print("[WARNING] 카카오톡 경로를 자동으로 찾을 수 없습니다. 기본 경로를 사용합니다.")
                KAKAO_PATH = 'C:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe'
except FileNotFoundError:
    print("[DEBUG] kakao_path.txt 파일을 찾을 수 없습니다. 자동으로 카카오톡 경로를 찾는 중...")
    found_path = find_kakao_path()
    if found_path:
        KAKAO_PATH = found_path
        print(f"[DEBUG] 자동으로 찾은 경로 사용: {KAKAO_PATH}")
    else:
        KAKAO_PATH = 'C:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe'
        print("[WARNING] 카카오톡 경로를 자동으로 찾을 수 없습니다. 기본 경로를 사용합니다.")

# 작업 디렉토리 변경 전에 기본 경로 저장
BASE_DIR = os.getcwd()
EXCEL_PATH = os.path.join(BASE_DIR, "chat_nickname_list.xlsx")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.json")

# 템플릿 저장 함수
def save_template(user_id, password, chat_type, chatroom_type, text_file_path):
    """템플릿을 JSON 파일로 저장하는 함수"""
    template = {
        "user_id": user_id,
        "password": password,
        "chat_type": chat_type,
        "chatroom_type": chatroom_type,
        "text_file_path": text_file_path
    }
    try:
        with open(TEMPLATE_PATH, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        print(f"[DEBUG] 템플릿 저장 완료: {TEMPLATE_PATH}")
        return True
    except Exception as e:
        print(f"[ERROR] 템플릿 저장 실패: {e}")
        return False

# 템플릿 로드 함수
def load_template():
    """템플릿을 JSON 파일에서 로드하는 함수"""
    try:
        if not os.path.exists(TEMPLATE_PATH):
            print(f"[DEBUG] 템플릿 파일이 없습니다: {TEMPLATE_PATH}")
            return None
        with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
            template = json.load(f)
        print(f"[DEBUG] 템플릿 로드 완료: {template}")
        return template
    except Exception as e:
        print(f"[ERROR] 템플릿 로드 실패: {e}")
        return None

# 기본 템플릿 생성 (처음 실행 시)
def create_default_template():
    """기본 템플릿을 생성하는 함수"""
    default_template = {
        "user_id": "01021456993",
        "password": "@rhdi120",
        "chat_type": "오픈채팅",
        "chatroom_type": "히딩크_체험단(임시)",
        "text_file_path": r"C:\Users\darli\Desktop\github\macro_kakao\모집공고\팔도_이천비락식혜.txt"
    }
    if not os.path.exists(TEMPLATE_PATH):
        save_template(
            default_template["user_id"],
            default_template["password"],
            default_template["chat_type"],
            default_template["chatroom_type"],
            default_template["text_file_path"]
        )

# 엑셀 파일에서 시트명 목록을 읽어오는 함수
def get_excel_sheet_names(excel_path):
    """
    엑셀 파일의 시트명 목록을 반환하는 함수
    """
    try:
        if not os.path.exists(excel_path):
            print(f"[WARNING] 엑셀 파일을 찾을 수 없습니다: {excel_path}")
            return []
        excel_file = pd.ExcelFile(excel_path)
        sheet_names = excel_file.sheet_names
        print(f"[DEBUG] 엑셀 시트명 목록: {sheet_names}")
        return sheet_names
    except Exception as e:
        print(f"[ERROR] 엑셀 파일 읽기 오류: {e}")
        return []

# 작업 디렉토리 설정
os.chdir("좌표")
print("Current working directory:", os.getcwd())

# 이미지 파일 경로 설정
COORDS = {
    "login_arr": "loginarr.png",
    "id_bt01": "id_bt01.png",
    "id_bt02": "id_bt02.png",
    "id_bt03": "id_bt03.png",
    "loginscs_bt": "loginscs_bt.png",
    "login_emt_bt": "login_emt_bt.png",
    "login_stanby_bt": "login_stanby_bt.png",
    "logout_popup": "logout_popup.png",
    "login_push": "login_push.png",
    "leftprofile_bt01": "leftprofile_bt01.png",
    "leftchat_bt01": "leftchat_bt01.png",
    "topchat02": "topchat02.png",
    "topopenchat_on": "topopenchat_on.png",
    "topopenchat_off": "topopenchat_off.png",
    "chatscsbt": "chatscsbt.png",
    "chatscs2bt": "chatscs2bt.png",
    "chatclose_bt": "chatclose_bt.png"
}

# 모니터 설정을 자동으로 감지하는 함수
def get_monitor_configs():
    monitors = get_monitors()
    monitor_configs = []
    
    print("\n=== 모니터 감지 결과 ===")
    print(f"감지된 모니터 수: {len(monitors)}")
    
    for i, monitor in enumerate(monitors, 1):
        monitor_config = {
            'top': 0,
            'left': monitor.x,
            'width': monitor.width,
            'height': monitor.height
        }
        monitor_configs.append(monitor_config)
        
        print(f"\n모니터 {i}:")
        print(f"  - 해상도: {monitor.width}x{monitor.height}")
        print(f"  - 위치: x={monitor.x}, y={monitor.y}")
        print(f"  - 기본 모니터: {'예' if monitor.is_primary else '아니오'}")
    
    print("\n모니터 감지 완료")
    print("==================\n")
    
    return monitor_configs

# 모니터 설정 자동 감지
MONITOR_CONFIGS = get_monitor_configs()

def capture_screen(region):
    with mss.mss() as sct:
        screenshot = sct.grab(region)
        img = np.array(screenshot)
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        return img

def match_template(screen, template_path):
    # 디버깅: 파일 경로 확인
    abs_path = os.path.abspath(template_path)
    print(f"[DEBUG] 이미지 파일 찾는 중: {template_path}")
    print(f"[DEBUG] 절대 경로: {abs_path}")
    print(f"[DEBUG] 파일 존재 여부: {os.path.exists(abs_path)}")
    
    if not os.path.exists(abs_path):
        print(f"[ERROR] 파일이 존재하지 않습니다: {abs_path}")
        print(f"[DEBUG] 현재 작업 디렉토리: {os.getcwd()}")
        raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {abs_path}")
    
    template = cv2.imread(template_path, cv2.IMREAD_COLOR)
    if template is None:
        print(f"[ERROR] cv2.imread가 None을 반환했습니다: {abs_path}")
        raise ValueError(f"이미지를 읽을 수 없습니다: {template_path}")

    result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(result)
    if max_val >= 0.8:  # 일치율이 80% 이상인 경우
        return (max_loc[0] + template.shape[1] // 2, max_loc[1] + template.shape[0] // 2)
    return None

def locate_image_on_monitor(image_path, monitor):
    screen = capture_screen(monitor)
    location = match_template(screen, image_path)
    if location:
        location = (location[0] + monitor['left'], location[1] + monitor['top'])
    return location

def locate_image_on_monitors(image_path):
    print(f"[DEBUG] 이미지 검색 시작: {image_path}")
    for i, monitor in enumerate(MONITOR_CONFIGS):
        print(f"[DEBUG] 모니터 {i+1}에서 검색 중...")
        location = locate_image_on_monitor(image_path, monitor)
        if location:
            print(f"[DEBUG] 이미지 찾음: {location}")
            return location
    print(f"[DEBUG] 모든 모니터에서 이미지를 찾지 못함: {image_path}")
    return None

def locate_and_click(image_path):
    print(f"{image_path}을(를) 찾고 클릭합니다.")
    location = locate_image_on_monitors(image_path)
    if location:
        print(f"{image_path}의 위치를 찾았습니다: {location}, 클릭합니다.")
        pyautogui.click(location)
        tm.sleep(0.5)
        return True
    else:
        print(f"{image_path}의 위치를 찾을 수 없습니다.")
        return False

def locate_and_click_offset(image_path, offset_x=0, offset_y=0):
    print(f"{image_path}을(를) 찾고 오프셋으로 클릭합니다: Offset: ({offset_x}, {offset_y})")
    location = locate_image_on_monitors(image_path)
    if location:
        location = (location[0] + offset_x, location[1] + offset_y)
        print(f"{image_path}의 위치를 찾았습니다: {location}, 클릭합니다.")
        pyautogui.click(location)
        tm.sleep(0.5)
        return True
    else:
        print(f"{image_path}의 위치를 찾을 수 없습니다.")
        return False

def run_kakao_macro(combo_id, entry_pw, chat_type, chatroom_type, text_content):
    try:
        # 디버깅: 경로 정보 출력
        print(f"\n[DEBUG] ========== 실행 시작 ==========")
        print(f"[DEBUG] EXCEL_PATH: {EXCEL_PATH}")
        print(f"[DEBUG] EXCEL_PATH 존재 여부: {os.path.exists(EXCEL_PATH)}")
        print(f"[DEBUG] KAKAO_PATH: {KAKAO_PATH}")
        print(f"[DEBUG] KAKAO_PATH 존재 여부: {os.path.exists(KAKAO_PATH)}")
        print(f"[DEBUG] 현재 작업 디렉토리: {os.getcwd()}")
        print(f"[DEBUG] BASE_DIR: {BASE_DIR}")
        print(f"[DEBUG] ================================\n")
        
        # 카카오톡 실행
        print("카카오톡 실행 중...")
        if not os.path.exists(KAKAO_PATH):
            raise FileNotFoundError(f"카카오톡 실행 파일을 찾을 수 없습니다: {KAKAO_PATH}")
        subprocess.Popen([KAKAO_PATH])
        
        # 카카오톡 실행 후 상태 확인
        print("\n카카오톡 실행 후 상태 확인 중...")
        
        # 2초 간격으로 60회 체크하여 현재 상태 확인
        state = None  # "logged_out", "logged_in", None
        for i in range(60):
            # 이미 로그아웃된 상태인지 확인 (login_stanby_bt.png)
            if locate_image_on_monitors(COORDS["login_stanby_bt"]):
                print("이미 로그아웃된 상태 감지 (login_stanby_bt)")
                state = "logged_out"
                break
            # 로그인된 상태인지 확인 (loginscs_bt.png 또는 login_emt_bt.png)
            if locate_image_on_monitors(COORDS["loginscs_bt"]):
                print("로그인된 상태 감지 (loginscs_bt)")
                state = "logged_in"
                break
            if locate_image_on_monitors(COORDS["login_emt_bt"]):
                print("로그인된 상태 감지 (login_emt_bt)")
                state = "logged_in"
                break
            print(f"상태 확인 중... ({i+1}/60)")
            tm.sleep(2)
        
        # 상태에 따라 분기 처리
        if state == "logged_out":
            # 이미 로그아웃된 상태이므로 바로 로그인 수행
            print("이미 로그아웃된 상태입니다. 바로 로그인을 수행합니다.")
        elif state == "logged_in":
            # 로그인된 상태이므로 로그아웃 후 로그인 화면 대기
            print("로그인된 상태입니다. 로그아웃을 수행합니다.")
            # 로그인 성공 확인용 버튼 클릭 (실패해도 계속 진행)
            if not locate_and_click(COORDS["loginscs_bt"]):
                print(f"[WARNING] 로그인 성공 확인 버튼({COORDS['loginscs_bt']})을 찾을 수 없지만 계속 진행합니다.")
            tm.sleep(5)
            
            # 로그아웃 수행
            print("로그아웃 중...")
            pyautogui.hotkey('alt', 'n')
            
            # 로그인 화면이 나타날 때까지 대기
            print("로그인 화면 대기 중...")
            max_wait = 60  # 최대 대기 시간 (초)
            interval = 1   # 확인 간격 (초)
            start_time = tm.time()
            
            while tm.time() - start_time < max_wait:
                if locate_image_on_monitors(COORDS["login_arr"]):
                    print("로그인 화면 감지됨")
                    break
                print("로그인 화면 대기 중...")
                tm.sleep(interval)
            else:
                raise TimeoutError("로그인 화면이 나타나지 않았습니다.")
        else:
            # 아무것도 감지되지 않았으므로 로그인 화면이 나타날 때까지 대기
            print("상태를 확인할 수 없습니다. 로그인 화면 대기 중...")
            max_wait = 60  # 최대 대기 시간 (초)
            interval = 1   # 확인 간격 (초)
            start_time = tm.time()
            
            while tm.time() - start_time < max_wait:
                if locate_image_on_monitors(COORDS["login_arr"]):
                    print("로그인 화면 감지됨")
                    break
                if locate_image_on_monitors(COORDS["login_stanby_bt"]):
                    print("로그인 화면 감지됨 (login_stanby_bt)")
                    break
                print("로그인 화면 대기 중...")
                tm.sleep(interval)
            else:
                raise TimeoutError("로그인 화면이 나타나지 않았습니다.")

        # 로그인 수행
        if not locate_and_click(COORDS["login_arr"]):
            # 실패 시 logout_popup.png 팝업이 떠있는지 확인
            print(f"[DEBUG] {COORDS['login_arr']} 클릭 실패. logout_popup.png 팝업 확인 중...")
            logout_popup_location = locate_image_on_monitors(COORDS["logout_popup"])
            if logout_popup_location:
                print(f"[DEBUG] logout_popup.png 팝업 감지됨. 팝업을 클릭하고 엔터를 누릅니다.")
                pyautogui.click(logout_popup_location)
                tm.sleep(0.5)
                pyautogui.press('enter')
                tm.sleep(2)  # 팝업 처리 후 대기
                print(f"[DEBUG] logout_popup.png 처리 완료. login_arr 재시도 중...")
                # 재시도
                if not locate_and_click(COORDS["login_arr"]):
                    raise RuntimeError(f"로그인 버튼({COORDS['login_arr']})을 찾을 수 없습니다. 로그인 화면이 표시되지 않았을 수 있습니다.")
            else:
                raise RuntimeError(f"로그인 버튼({COORDS['login_arr']})을 찾을 수 없습니다. 로그인 화면이 표시되지 않았을 수 있습니다.")
        tm.sleep(0.3)
        user_id_image = {
            "01021456993": COORDS["id_bt01"],
            "bnam91@naver.com": COORDS["id_bt02"],
            "jisu1021104@gmail.com": COORDS["id_bt03"]
        }
        if not locate_and_click(user_id_image[combo_id]):
            raise RuntimeError(f"아이디 선택 버튼({user_id_image[combo_id]})을 찾을 수 없습니다.")
        tm.sleep(1)
        pyperclip.copy(entry_pw)
        pyautogui.hotkey('ctrl', 'v')
        tm.sleep(0.3)
        pyautogui.press('enter')
        # loginscs_bt.png 또는 login_push.png 중 하나가 나타날 때까지 동적으로 체크
        print(f"[DEBUG] 로그인 결과 확인 중... (loginscs_bt.png 또는 login_push.png 대기)")
        max_wait = 120  # 최대 대기 시간 (초)
        interval = 2   # 확인 간격 (초)
        start_time = tm.time()
        login_success = False
        login_push_processed = False  # login_push.png 처리 여부
        
        while tm.time() - start_time < max_wait:
            # loginscs_bt.png 확인 (로그인 성공)
            if locate_image_on_monitors(COORDS["loginscs_bt"]):
                print(f"[DEBUG] loginscs_bt.png 감지됨. 로그인 성공 확인.")
                login_success = True
                break
            
            # login_push.png 확인 (다른 PC에서 로그인 중)
            if not login_push_processed:  # 아직 처리하지 않은 경우에만 확인
                login_push_location = locate_image_on_monitors(COORDS["login_push"])
                if login_push_location:
                    print(f"[DEBUG] login_push.png 팝업 감지됨. 팝업을 클릭하고 엔터를 누릅니다.")
                    pyautogui.click(login_push_location)
                    tm.sleep(0.5)
                    pyautogui.press('enter')
                    tm.sleep(2)  # 팝업 처리 후 초기 대기
                    print(f"[DEBUG] login_push.png 처리 완료. loginscs_bt.png 동적 체크 시작...")
                    login_push_processed = True
                    # login_push 처리 후에는 loginscs_bt.png만 계속 확인
            
            if login_push_processed:
                # login_push 처리 후 loginscs_bt.png 동적 체크
                if locate_image_on_monitors(COORDS["loginscs_bt"]):
                    print(f"[DEBUG] login_push 처리 후 loginscs_bt.png 감지됨. 로그인 성공 확인.")
                    login_success = True
                    break
                else:
                    print(f"[DEBUG] login_push 처리 후 loginscs_bt.png 대기 중... ({int(tm.time() - start_time)}초 경과)")
            
            if not login_push_processed:
                print(f"[DEBUG] 로그인 결과 확인 중... ({int(tm.time() - start_time)}초 경과)")
            
            tm.sleep(interval)
        
        if not login_success:
            messagebox.showerror("로그인 실패", f"로그인에 실패했습니다. ({max_wait}초 내에 loginscs_bt.png를 찾지 못했습니다.)")
            return
        
        # 좌측 채팅 버튼 클릭
        if not locate_and_click(COORDS["leftchat_bt01"]):
            raise RuntimeError(f"좌측 채팅 버튼({COORDS['leftchat_bt01']})을 찾을 수 없습니다.")
        tm.sleep(0.5)

        # 채팅 설정
        if chat_type == "친구(신규)":
            if not locate_and_click(COORDS["leftprofile_bt01"]):
                raise RuntimeError(f"프로필 버튼({COORDS['leftprofile_bt01']})을 찾을 수 없습니다.")
        elif chat_type == "개인채팅":
            if locate_image_on_monitors(COORDS["topopenchat_on"]):
                if not locate_and_click_offset(COORDS["topopenchat_off"], -86, 0):  # 개인채팅 탭 클릭
                    raise RuntimeError(f"개인채팅 탭 버튼({COORDS['topopenchat_off']})을 찾을 수 없습니다.")
        elif chat_type == "오픈채팅":
            if locate_image_on_monitors(COORDS["topopenchat_off"]):
                if not locate_and_click(COORDS["topopenchat_on"]):  # 오픈채팅 탭 클릭
                    raise RuntimeError(f"오픈채팅 탭 버튼({COORDS['topopenchat_on']})을 찾을 수 없습니다.")

        # 채팅방 찾기 및 열기
        print(f"[DEBUG] 엑셀 파일 읽기 시도: {EXCEL_PATH}")
        print(f"[DEBUG] 시트 이름: {chatroom_type}")
        if not os.path.exists(EXCEL_PATH):
            raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {EXCEL_PATH}")
        sheet = pd.read_excel(EXCEL_PATH, sheet_name=chatroom_type)
        print(f"[DEBUG] 엑셀 파일 읽기 성공, 행 수: {len(sheet)}")
        for chatroom_name in sheet.iloc[:, 2]:
            pyautogui.hotkey('ctrl', 'f')
            tm.sleep(0.5)
            for _ in range(20):
                pyautogui.press('backspace')
            pyperclip.copy(chatroom_name)
            pyautogui.hotkey('ctrl', 'v')
            tm.sleep(0.5)
            pyautogui.press('enter')
            tm.sleep(3)
            if not locate_image_on_monitors(COORDS["chatscsbt"]):
                print(f"{chatroom_name} 채팅방을 열지 못했습니다.")
                continue

            # 채팅 입력 및 전송
            pyperclip.copy(text_content)
            pyautogui.hotkey('ctrl', 'v')
            tm.sleep(2)  # 추가된 대기 시간
            if locate_image_on_monitors(COORDS["chatscs2bt"]):
                print(f"{chatroom_name} 채팅방에 메시지 전송 성공.")
            else:
                print(f"{chatroom_name} 채팅방에 메시지 전송 실패.")
                continue
            
            # pyautogui.press('enter') # DEV
            tm.sleep(1)
            pyautogui.press('esc')  # ESC 키를 눌러 채팅창 닫기
            tm.sleep(1)

    except FileNotFoundError as e:
        error_msg = f"파일을 찾을 수 없습니다: {e}"
        print(f"[ERROR] {error_msg}")
        print(f"[DEBUG] 현재 작업 디렉토리: {os.getcwd()}")
        print(f"[DEBUG] BASE_DIR: {BASE_DIR}")
        messagebox.showerror("오류", error_msg)
    except Exception as e:
        error_msg = f"작업 중 오류가 발생했습니다: {e}"
        print(f"[ERROR] {error_msg}")
        print(f"[DEBUG] 예외 타입: {type(e).__name__}")
        import traceback
        print(f"[DEBUG] 상세 오류:\n{traceback.format_exc()}")
        messagebox.showerror("오류", error_msg)

class KakaoMacroApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KakaoTalk 예약 프로그램")
        self.geometry("800x600")
        
        # 아이디별 비밀번호 매핑 (create_widgets 호출 전에 정의)
        self.password_map = {
            "01021456993": "@rhdi120",
            "bnam91@naver.com": "",
            "jisu1021104@gmail.com": ""
        }
        
        self.text_filename = ""
        self.text_content = ""
        
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="아이디").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        tk.Label(self, text="패스워드").grid(row=1, column=0, padx=10, pady=5, sticky='w')

        self.combo_id = ttk.Combobox(self, values=["01021456993", "bnam91@naver.com", "jisu1021104@gmail.com"])
        self.combo_id.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        self.combo_id.current(0)
        self.combo_id.bind("<<ComboboxSelected>>", self.on_id_changed)

        self.entry_pw = tk.Entry(self, show="*")
        self.entry_pw.grid(row=1, column=1, padx=10, pady=5, sticky='ew')
        
        # 초기 비밀번호 설정
        self.on_id_changed()

        tk.Label(self, text="채팅탭").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.combo_chat_type = ttk.Combobox(self, values=["오픈채팅", "개인채팅", "친구(신규)"])
        self.combo_chat_type.grid(row=2, column=1, padx=10, pady=5, sticky='ew')
        self.combo_chat_type.current(0)

        tk.Label(self, text="채팅방유형").grid(row=3, column=0, padx=10, pady=5, sticky='w')
        # 엑셀 파일에서 시트명을 동적으로 읽어옴
        sheet_names = get_excel_sheet_names(EXCEL_PATH)
        if not sheet_names:
            # 엑셀 파일을 읽을 수 없는 경우 기본값 사용
            sheet_names = ["히딩크_체험단", "히딩크_무한", "실배송_박항서", "홍명보_가구매", "이강인_침투", "김민재_부업"]
            print("[WARNING] 엑셀 파일에서 시트명을 읽을 수 없어 기본값을 사용합니다.")
        self.combo_chatroom_type = ttk.Combobox(self, values=sheet_names)
        self.combo_chatroom_type.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        if sheet_names:
            self.combo_chatroom_type.current(0)

        self.btn_browse = tk.Button(self, text="내용 파일 불러오기", command=self.load_text_file)
        self.btn_browse.grid(row=4, column=1, padx=10, pady=5, sticky='ew')

        self.label_filename = tk.Label(self, text="")
        self.label_filename.grid(row=4, column=2, padx=10, pady=5, sticky='w')

        tk.Label(self, text="메모").grid(row=5, column=0, padx=10, pady=5, sticky='w')
        self.entry_memo = tk.Entry(self)
        self.entry_memo.grid(row=5, column=1, padx=10, pady=5, sticky='ew')

        self.btn_add = tk.Button(self, text="예약 추가", command=self.add_reservation)
        self.btn_add.grid(row=6, column=1, padx=10, pady=5, sticky='ew')

        self.btn_execute = tk.Button(self, text="즉시 실행", command=self.execute_now)
        self.btn_execute.grid(row=6, column=0, padx=10, pady=5, sticky='ew')

        self.btn_stop = tk.Button(self, text="작업 중지", command=self.stop_action)
        self.btn_stop.grid(row=6, column=2, padx=10, pady=5, sticky='ew')

        columns = ("채팅탭", "채팅방유형", "내용", "메모", "상태", "삭제하기")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        self.tree.grid(row=9, column=0, columnspan=3, sticky='nsew')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        self.tree.bind("<Double-1>", self.delete_reservation)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(9, weight=1)

    def on_id_changed(self, event=None):
        """아이디가 변경될 때 비밀번호 필드를 자동으로 채우는 함수"""
        selected_id = self.combo_id.get()
        if selected_id in self.password_map:
            password = self.password_map[selected_id]
            self.entry_pw.delete(0, tk.END)
            self.entry_pw.insert(0, password)

    def load_text_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            self.text_filename = file_path.split('/')[-1]
            self.label_filename.config(text=self.text_filename)
            with open(file_path, "r", encoding="utf-8") as file:
                self.text_content = file.read()

    def add_reservation(self):
        reservation = (
            self.combo_chat_type.get(),
            self.combo_chatroom_type.get(),
            self.text_filename,
            self.entry_memo.get(),
            "예약대기",
            "삭제"
        )

        if messagebox.askyesno("확인", "예약을 추가하시겠습니까?"):
            for item in self.tree.get_children():
                if self.tree.item(item, 'values')[:4] == reservation[:4]:
                    messagebox.showerror("중복 예약", "동일한 예약이 이미 존재합니다.")
                    return
            self.tree.insert("", "end", values=reservation)

    def delete_reservation(self, event):
        selected_item = self.tree.selection()[0]
        self.tree.delete(selected_item)

    def execute_now(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("경고", "실행할 예약을 선택하세요.")
            return

        if messagebox.askyesno("확인", "즉시 실행하시겠습니까?"):
            reservation = self.tree.item(selected_item, 'values')
            self.perform_action(reservation)

    def perform_action(self, reservation):
        chat_type, chatroom_type, text_file, memo, status, delete = reservation
        messagebox.showinfo("즉시 실행", f"예약된 작업을 즉시 실행합니다:\n\n{reservation}")
        run_kakao_macro(self.combo_id.get(), self.entry_pw.get(), chat_type, chatroom_type, self.text_content)
        self.tree.set(self.tree.selection()[0], column="상태", value="완료")

    def stop_action(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("경고", "작업 중지할 예약을 선택하세요.")
            return

        for item in selected_item:
            self.tree.set(item, column="상태", value="작업중지")

def run_template():
    """템플릿으로 즉시 실행하는 함수"""
    template = load_template()
    if not template:
        print("[ERROR] 템플릿 파일을 찾을 수 없습니다.")
        return False
    
    try:
        if os.path.exists(template['text_file_path']):
            with open(template['text_file_path'], "r", encoding="utf-8") as f:
                text_content = f.read()
        else:
            print(f"[ERROR] 내용 파일을 찾을 수 없습니다: {template['text_file_path']}")
            return False
        
        if text_content:
            print(f"[INFO] 템플릿으로 즉시 실행합니다.")
            print(f"[INFO] 아이디: {template['user_id']}")
            print(f"[INFO] 채팅탭: {template['chat_type']}")
            print(f"[INFO] 채팅방유형: {template['chatroom_type']}")
            print(f"[INFO] 내용파일: {os.path.basename(template['text_file_path'])}")
            run_kakao_macro(
                template['user_id'],
                template['password'],
                template['chat_type'],
                template['chatroom_type'],
                text_content
            )
            return True
        else:
            print("[ERROR] 내용 파일이 비어있습니다.")
            return False
    except Exception as e:
        print(f"[ERROR] 템플릿 실행 중 오류가 발생했습니다: {e}")
        import traceback
        print(traceback.format_exc())
        return False

if __name__ == "__main__":
    # 명령줄 인자 파싱
    parser = argparse.ArgumentParser(description="KakaoTalk 매크로 프로그램")
    parser.add_argument(
        "-t", "--template",
        action="store_true",
        help="템플릿으로 즉시 실행 (CLI 모드)"
    )
    args = parser.parse_args()
    
    # 기본 템플릿 생성 (없는 경우)
    create_default_template()
    
    # CLI 모드: 템플릿으로 즉시 실행
    if args.template:
        run_template()
    else:
        # GUI 모드: 템플릿 사용 여부 확인
        template = load_template()
        if template:
            # 템플릿 사용 여부를 묻는 다이얼로그
            root = tk.Tk()
            root.withdraw()  # 메인 윈도우 숨기기
            
            response = messagebox.askyesnocancel(
                "템플릿 사용",
                f"자주 사용하는 템플릿을 사용하시겠습니까?\n\n"
                f"아이디: {template['user_id']}\n"
                f"채팅탭: {template['chat_type']}\n"
                f"채팅방유형: {template['chatroom_type']}\n"
                f"내용파일: {os.path.basename(template['text_file_path'])}\n\n"
                f"예: 템플릿으로 즉시 실행\n"
                f"아니오: 직접 선택\n"
                f"취소: 프로그램 종료"
            )
            
            root.destroy()
            
            if response is True:  # 템플릿 사용
                # 템플릿 파일 내용 읽기
                try:
                    if os.path.exists(template['text_file_path']):
                        with open(template['text_file_path'], "r", encoding="utf-8") as f:
                            text_content = f.read()
                    else:
                        messagebox.showerror("오류", f"내용 파일을 찾을 수 없습니다:\n{template['text_file_path']}")
                        text_content = ""
                    
                    if text_content:
                        # 템플릿으로 즉시 실행
                        if messagebox.askyesno("즉시 실행", "템플릿 설정으로 즉시 실행하시겠습니까?"):
                            run_kakao_macro(
                                template['user_id'],
                                template['password'],
                                template['chat_type'],
                                template['chatroom_type'],
                                text_content
                            )
                except Exception as e:
                    messagebox.showerror("오류", f"템플릿 실행 중 오류가 발생했습니다:\n{e}")
            elif response is False:  # 직접 선택
                app = KakaoMacroApp()
                app.mainloop()
            # response is None (취소)인 경우 프로그램 종료
        else:
            # 템플릿이 없으면 일반 GUI 실행
            app = KakaoMacroApp()
            app.mainloop()
