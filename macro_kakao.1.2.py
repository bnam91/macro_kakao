import subprocess
import pyautogui
import time as tm
import cv2
import numpy as np
import mss
import pandas as pd
import pyperclip
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime
from screeninfo import get_monitors

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

# 경로 및 변수 설정
try:
    with open('kakao_path.txt', 'r') as f:
        KAKAO_PATH = f.read().strip()
except FileNotFoundError:
    KAKAO_PATH = 'C:/Program Files (x86)/Kakao/KakaoTalk/KakaoTalk.exe'
    print("kakao_path.txt 파일을 찾을 수 없습니다. 기본 경로를 사용합니다.")

EXCEL_PATH = "chat_nickname_list.xlsx"

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
    "leftprofile_bt01": "leftprofile_bt01.png",
    "leftchat_bt01": "leftchat_bt01.png",
    "topchat02": "topchat02.png",
    "topopenchat_on": "topopenchat_on.png",
    "topopenchat_off": "topopenchat_off.png",
    "chatscsbt": "chatscsbt.png",
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
        print(f"  • 해상도: {monitor.width}x{monitor.height}")
        print(f"  • 위치: x={monitor.x}, y={monitor.y}")
        print(f"  • 기본 모니터: {'예' if monitor.is_primary else '아니오'}")
    
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
    template = cv2.imread(template_path, cv2.IMREAD_COLOR)
    if template is None:
        raise ValueError(f"이미지를 찾을 수 없습니다: {template_path}")

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
    for monitor in MONITOR_CONFIGS:
        location = locate_image_on_monitor(image_path, monitor)
        if location:
            return location
    return None

def locate_and_click(image_path):
    print(f"{image_path}을(를) 찾고 클릭합니다.")
    location = locate_image_on_monitors(image_path)
    if location:
        print(f"{image_path}의 위치를 찾았습니다: {location}, 클릭합니다.")
        pyautogui.click(location)
        tm.sleep(0.5)
    else:
        print(f"{image_path}의 위치를 찾을 수 없습니다.")

def locate_and_click_offset(image_path, offset_x=0, offset_y=0):
    print(f"{image_path}을(를) 찾고 오프셋으로 클릭합니다: Offset: ({offset_x}, {offset_y})")
    location = locate_image_on_monitors(image_path)
    if location:
        location = (location[0] + offset_x, location[1] + offset_y)
        print(f"{image_path}의 위치를 찾았습니다: {location}, 클릭합니다.")
        pyautogui.click(location)
        tm.sleep(0.5)
    else:
        print(f"{image_path}의 위치를 찾을 수 없습니다.")

def run_kakao_macro(combo_id, entry_pw, chat_type, chatroom_type, text_content):
    try:
        # 카카오톡 실행
        print("카카오톡 실행 중...")
        subprocess.Popen([KAKAO_PATH])
        
        # 최대 60초까지 대기하면서 2초마다 로그인 버튼 확인
        max_wait = 60  # 최대 대기 시간 (초)
        interval = 2   # 확인 간격 (초)
        start_time = tm.time()
        
        while tm.time() - start_time < max_wait:
            if locate_image_on_monitors(COORDS["loginscs_bt"]):
                print("카카오톡 실행 완료")
                break
            print("카카오톡 실행 대기 중...")
            tm.sleep(interval)
        else:
            raise TimeoutError("카카오톡 실행 시간이 초과되었습니다.")

        # 로그인 성공 확인용 버튼 클릭
        locate_and_click(COORDS["loginscs_bt"])
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

        # 로그인 수행
        locate_and_click(COORDS["login_arr"])
        tm.sleep(0.3)
        user_id_image = {
            "01021456993": COORDS["id_bt01"],
            "bnam91@naver.com": COORDS["id_bt02"],
            "jisu1021104@gmail.com": COORDS["id_bt03"]
        }
        locate_and_click(user_id_image[combo_id])
        tm.sleep(1)
        pyperclip.copy(entry_pw)
        pyautogui.hotkey('ctrl', 'v')
        tm.sleep(0.3)
        pyautogui.press('enter')
        tm.sleep(15)
        if not locate_image_on_monitors(COORDS["loginscs_bt"]):
            messagebox.showerror("로그인 실패", "로그인에 실패했습니다.")
            return
        
        # 좌측 채팅 버튼 클릭
        locate_and_click(COORDS["leftchat_bt01"])
        tm.sleep(0.5)

        # 채팅 설정
        if chat_type == "친구(신규)":
            locate_and_click(COORDS["leftprofile_bt01"])
        elif chat_type == "개인채팅":
            if locate_image_on_monitors(COORDS["topopenchat_on"]):
                locate_and_click_offset(COORDS["topopenchat_off"], -86, 0)  # 개인채팅 탭 클릭
        elif chat_type == "오픈채팅":
            if locate_image_on_monitors(COORDS["topopenchat_off"]):
                locate_and_click(COORDS["topopenchat_on"])  # 오픈채팅 탭 클릭

        # 채팅방 찾기 및 열기
        sheet = pd.read_excel(EXCEL_PATH, sheet_name=chatroom_type)
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
            if locate_image_on_monitors("chatscs2bt.png"):
                print(f"{chatroom_name} 채팅방에 메시지 전송 성공.")
            else:
                print(f"{chatroom_name} 채팅방에 메시지 전송 실패.")
                continue
            
            # pyautogui.press('enter')
            tm.sleep(1)
            pyautogui.press('esc')  # ESC 키를 눌러 채팅창 닫기
            tm.sleep(1)

    except Exception as e:
        messagebox.showerror("오류", f"작업 중 오류가 발생했습니다: {e}")

class KakaoMacroApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("KakaoTalk 예약 프로그램")
        self.geometry("800x600")
        
        self.create_widgets()
        self.text_filename = ""
        self.text_content = ""

    def create_widgets(self):
        tk.Label(self, text="아이디").grid(row=0, column=0, padx=10, pady=5, sticky='w')
        tk.Label(self, text="패스워드").grid(row=1, column=0, padx=10, pady=5, sticky='w')

        self.combo_id = ttk.Combobox(self, values=["01021456993", "bnam91@naver.com", "jisu1021104@gmail.com"])
        self.combo_id.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        self.combo_id.current(0)

        self.entry_pw = tk.Entry(self, show="*")
        self.entry_pw.grid(row=1, column=1, padx=10, pady=5, sticky='ew')

        tk.Label(self, text="채팅탭").grid(row=2, column=0, padx=10, pady=5, sticky='w')
        self.combo_chat_type = ttk.Combobox(self, values=["오픈채팅", "개인채팅", "친구(신규)"])
        self.combo_chat_type.grid(row=2, column=1, padx=10, pady=5, sticky='ew')
        self.combo_chat_type.current(0)

        tk.Label(self, text="채팅방유형").grid(row=3, column=0, padx=10, pady=5, sticky='w')
        self.combo_chatroom_type = ttk.Combobox(self, values=["히딩크_체험단", "히딩크_무한", "실배송_박항서", "홍명보_가구매", "이강인_침투", "김민재_부업"])
        self.combo_chatroom_type.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        self.combo_chatroom_type.current(0)

        self.btn_browse = tk.Button(self, text="내용 파일 불러오기", command=self.load_text_file)
        self.btn_browse.grid(row=4, column=1, padx=10, pady=5, sticky='ew')

        self.label_filename = tk.Label(self, text="")
        self.label_filename.grid(row=4, column=2, padx=10, pady=5, sticky='w')

        tk.Label(self, text="예약일").grid(row=5, column=0, padx=10, pady=5, sticky='w')
        self.entry_date = DateEntry(self, date_pattern='yyyy-mm-dd')
        self.entry_date.grid(row=5, column=1, padx=10, pady=5, sticky='ew')

        tk.Label(self, text="시간").grid(row=6, column=0, padx=10, pady=5, sticky='w')
        self.combo_time = ttk.Combobox(self, values=[f"{hour:02d}:{minute:02d}" for hour in range(24) for minute in [0, 15, 30, 45]])
        self.combo_time.grid(row=6, column=1, padx=10, pady=5, sticky='ew')
        self.combo_time.current(0)

        tk.Label(self, text="메모").grid(row=7, column=0, padx=10, pady=5, sticky='w')
        self.entry_memo = tk.Entry(self)
        self.entry_memo.grid(row=7, column=1, padx=10, pady=5, sticky='ew')

        self.btn_add = tk.Button(self, text="예약 추가", command=self.add_reservation)
        self.btn_add.grid(row=8, column=1, padx=10, pady=5, sticky='ew')

        self.btn_execute = tk.Button(self, text="즉시 실행", command=self.execute_now)
        self.btn_execute.grid(row=8, column=0, padx=10, pady=5, sticky='ew')

        self.btn_stop = tk.Button(self, text="작업 중지", command=self.stop_action)
        self.btn_stop.grid(row=8, column=2, padx=10, pady=5, sticky='ew')

        columns = ("채팅탭", "채팅방유형", "내용", "예약일", "시간", "메모", "상태", "삭제하기")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        self.tree.grid(row=9, column=0, columnspan=3, sticky='nsew')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        self.tree.bind("<Double-1>", self.delete_reservation)

        self.btn_load_previous = tk.Button(self, text="이전 예약내역 불러오기", command=self.load_previous_reservations)
        self.btn_load_previous.grid(row=10, column=0, columnspan=3, padx=10, pady=5, sticky='ew')

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(9, weight=1)

    def load_text_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if file_path:
            self.text_filename = file_path.split('/')[-1]
            self.label_filename.config(text=self.text_filename)
            with open(file_path, "r", encoding="utf-8") as file:
                self.text_content = file.read()

    def add_reservation(self):
        # 현재 시간과 입력된 예약 시간을 비교하여 경고창을 표시
        selected_date = self.entry_date.get_date()
        selected_time = self.combo_time.get()
        selected_datetime_str = f"{selected_date} {selected_time}"
        selected_datetime = datetime.strptime(selected_datetime_str, "%Y-%m-%d %H:%M")

        if selected_datetime < datetime.now():
            messagebox.showwarning("경고", "예약 시간이 이미 지났습니다.")
            return

        reservation = (
            self.combo_chat_type.get(),
            self.combo_chatroom_type.get(),
            self.text_filename,
            self.entry_date.get(),
            self.combo_time.get(),
            self.entry_memo.get(),
            "예약대기",
            "삭제"
        )

        if messagebox.askyesno("확인", "해당 시간으로 예약을 추가하시겠습니까?"):
            for item in self.tree.get_children():
                if self.tree.item(item, 'values')[:6] == reservation[:6]:
                    messagebox.showerror("중복 예약", "동일한 예약이 이미 존재합니다.")
                    return
            self.tree.insert("", "end", values=reservation)

    def delete_reservation(self, event):
        selected_item = self.tree.selection()[0]
        self.tree.delete(selected_item)

    def load_previous_reservations(self):
        pass

    def execute_now(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("경고", "실행할 예약을 선택하세요.")
            return

        if messagebox.askyesno("확인", "즉시 실행하시겠습니까?"):
            reservation = self.tree.item(selected_item, 'values')
            self.perform_action(reservation)

    def perform_action(self, reservation):
        chat_type, chatroom_type, text_file, date, time, memo, status, delete = reservation
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

if __name__ == "__main__":
    app = KakaoMacroApp()
    app.mainloop()
