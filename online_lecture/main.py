import tkinter as tk
from tkinter import filedialog, ttk
import os
import json
import sys
import threading
from utils import load_workbook, paste_school, paste_student

CONFIG_FILE = "config/config.json"

# stdout을 Text 창으로 리디렉션
class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass

# 처리 함수
def execute_process(dir_path, source_path="temp.xlsx", target_path="main.xlsx"):
    print(f"소스: {source_path}")
    print(f"타겟: {target_path}")
    print("처리 시작...")
    wb_source = load_workbook(source_path)
    ws_source = wb_source.active
    
    wb_target = load_workbook(target_path)
    ws_target_school = wb_target["2.교육(일지)"]
    ws_target_student = wb_target["3.교육(명부)"]

    row = 2
    cnt = 0
    while ws_source[f"A{row}"].value is not None:
        cnt += 1
        row += 1

    target_counter_school = 8
    target_counter_student = 5
    progress_bar["value"] = 0
    for source_pointer in range(2, cnt + 2, 1):
        target_counter_school = paste_school(ws_source, ws_target_school, source_pointer, target_counter_school)
        target_counter_student = paste_student(dir_path, ws_source, ws_target_student, source_pointer, target_counter_student)
        rate = (source_pointer-1) * 100 // cnt
        progress_bar["value"] = rate
        progress_label.config(text=f"{rate}% 완료")
        window.update_idletasks()

    wb_target.save(target_path)
    print("처리 완료!")

def select_schooldir():
    global dir_path, last_dir
    folder = filedialog.askdirectory(title="신청서 폴더 선택", initialdir=last_dir)
    if folder:
        dir_label.config(text=folder)
        dir_path = folder
        last_dir = folder
        save_config()

def select_source():
    global source_path, last_dir
    filepath = filedialog.askopenfilename(title="신청서 정보 파일 선택", initialdir=last_dir)
    if filepath:
        source_label.config(text=filepath)
        source_path = filepath
        last_dir = os.path.dirname(filepath)
        save_config()

def select_target():
    global target_path, last_dir
    filepath = filedialog.askopenfilename(title="양식 파일 선택", initialdir=last_dir)
    if filepath:
        target_label.config(text=filepath)
        target_path = filepath
        last_dir = os.path.dirname(filepath)
        save_config()

# 처리 시작 (버튼 클릭 시)
def start_process():
    if not source_path or not target_path:
        print("신청서 정보 파일과 양식 파일을 모두 선택하세요.")
        return
    threading.Thread(target=execute_process, args=(dir_path, source_path, target_path), daemon=True).start()

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_config():
    config = {
        "last_dir": last_dir,
        "source_path": source_path,
        "target_path": target_path,
        "dir_path": dir_path
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f)

if __name__ == "__main__":
    # 윈도우 생성
    window = tk.Tk()
    window.title("2025년 시청자미디어센터 온라인특강 운영실적")
    window.geometry("700x500")

    config = load_config()
    last_dir = config.get("last_dir", ".")
    source_path = config.get("source_path", "")
    target_path = config.get("target_path", "")
    dir_path = config.get("dir_path", "")

    # 상단: 소스/타겟 선택 버튼 + 경로 라벨
    frame_top = tk.Frame(window)
    frame_top.pack(pady=10)

    # 폴더 선택
    frame_dir = tk.Frame(frame_top)
    frame_dir.pack(side=tk.LEFT, padx=20)
    btn_dir = tk.Button(frame_dir, text="신청서 폴더 선택", command=select_schooldir)
    btn_dir.pack()
    dir_label = tk.Label(frame_dir, text=dir_path or "(폴더 경로 없음)", wraplength=250)
    dir_label.pack()

    # 소스 선택
    frame_source = tk.Frame(frame_top)
    frame_source.pack(side=tk.LEFT, padx=20)
    btn_source = tk.Button(frame_source, text="신청서 정보 파일 선택", command=select_source)
    btn_source.pack()
    source_label = tk.Label(frame_source, text=source_path or "(경로 없음)", wraplength=250)
    source_label.pack()

    # 타겟 선택
    frame_target = tk.Frame(frame_top)
    frame_target.pack(side=tk.LEFT, padx=20)
    btn_target = tk.Button(frame_target, text="양식 파일 선택", command=select_target)
    btn_target.pack()
    target_label = tk.Label(frame_target, text=target_path or "(경로 없음)", wraplength=250)
    target_label.pack()

    # 처리 시작 버튼
    btn_start = tk.Button(window, text="처리 시작", command=start_process)
    btn_start.pack(pady=10)

    
    # 진행 바
    progress_bar = ttk.Progressbar(window, length=600, mode='determinate')
    progress_bar.pack(pady=10)

    # 진행률 라벨
    progress_label = tk.Label(window, text="0% 완료")
    progress_label.pack()
    

    # 출력창 (stdout 리디렉션)
    output_text = tk.Text(window, height=10)
    output_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # stdout 연결
    sys.stdout = StdoutRedirector(output_text)
    sys.stderr = sys.stdout

    # 실행
    window.mainloop()
