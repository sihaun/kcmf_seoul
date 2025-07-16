import tkinter as tk
from tkinter import filedialog, ttk
import sys
import os
import threading
from utils import load_workbook, paste_student, return_class, perfectcopy

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
def execute_process(list_path, source_path="temp.xlsx", target_path="main.xlsx"):
    print(f"소스: {source_path}")
    print(f"타겟: {target_path}")
    print("처리 시작...")
    wb_source = load_workbook(source_path)
    ws_source = wb_source.active
    
    wb_target = load_workbook(target_path)
    ws_target_student = wb_target["3.교육(명부)"]

    row = 2
    cnt = 0
    while ws_source[f"A{row}"].value is not None:
        cnt += 1
        row += 1

    target_counter_student = 5
    progress_bar["value"] = 0
    for source_pointer in range(2, cnt + 2, 1):
        target_counter_student = paste_student(list_path, ws_source, ws_target_student, source_pointer, target_counter_student)
        rate = (source_pointer-1) * 100 // cnt
        progress_bar["value"] = rate
        progress_label.config(text=f"{rate}% 완료")
        window.update_idletasks()

    wb_target.save(target_path)
    print("처리 완료!")

def paste_student(dir_path, source_sheet, target_sheet, source_pointer: int=2, target_counter : int=4):

    ws_source = source_sheet

    ws_target = target_sheet

    list_counter = 5
    try:
        name_list, gender_list = listing_student(dir_path, ws_source[f"A{source_pointer}"].value)

        for i in range(ws_source[f"G{source_pointer}"].value): # source의 인구수만큼 반복
            school_name = ws_source[f"B{source_pointer}"].value
            class_name = ws_source[f"C{source_pointer}"].value
            name = name_list[i]
            gender = gender_list[i]
            address = ws_source[f"H{source_pointer}"].value

            if gender == None:
                print(f"school : {school_name}, index : {list_counter}. 인원과 명단 수가 맞지 않음.")
                raise IndexError(f"index : {list_counter}. 인원과 명단 수가 맞지 않음.")
            # 80
            for i in range(65, 80, 1):
                column = chr(i)
                if column == "D":
                    value = return_class(class_name, school_name)
                elif column == "G":
                    value = name
                elif column == "J":
                    value = "100%"
                elif column == "L":
                    value = gender
                elif column == "O":
                    value = address
                else : 
                    value = None
                perfectcopy(ws_target[f"{column}{target_counter}"], ws_target[f"{column}2"], value)

            target_counter += 1
            list_counter += 1

    except:
        pass

    return target_counter

def listing_student(file_path : str):
    wb = load_workbook(file_path)
    ws = wb.active

    # 값과 일부 서식 복사 (수동으로)
    name_list = []
    gender_list = []
    pointer = 5
    for row in range(pointer, ws.max_row + 1):
        name_list.append(ws[f"B{row}"].value)
        gender_list.append(ws[f"D{row}"].value)

    name_count = {}
    result = []

    for name in name_list:
        # 이름 카운트 증가
        name_count[name] = name_count.get(name, 0) + 1

        # 두 자리 숫자로 포맷
        suffix = f"{name_count[name]:02d}"

        # 이름 + 순번 조합
        result.append(f"{name}{suffix}")

    result_name = [name[:-2] if name.endswith("01") else name for name in result]

    return result_name, gender_list

# 명단 파일 선택
def select_list():
    filepath = filedialog.askopenfilename(title="신청서 정보 파일 선택")
    if filepath:
        list_label.config(text=filepath)
        global list_path
        list_path = filepath


# 소스 파일 선택
def select_source():
    filepath = filedialog.askopenfilename(title="신청서 정보 파일 선택")
    if filepath:
        source_label.config(text=filepath)
        global source_path
        source_path = filepath

# 타겟 파일 선택
def select_target():
    filepath = filedialog.askopenfilename(title="양식 파일 선택")
    if filepath:
        target_label.config(text=filepath)
        global target_path
        target_path = filepath

# 처리 시작 (버튼 클릭 시)
def start_process():
    if not source_path or not target_path:
        print("신청서 정보 파일과 양식 파일을 모두 선택하세요.")
        return
    threading.Thread(target=execute_process, args=(list_path, source_path, target_path), daemon=True).start()

if __name__ == "__main__":
    # 윈도우 생성
    window = tk.Tk()
    window.title("운영실적 에러 수정")
    window.geometry("700x500")

    # 초기 경로 변수
    source_path = ""
    target_path = ""

    # 상단: 소스/타겟 선택 버튼 + 경로 라벨
    frame_top = tk.Frame(window)
    frame_top.pack(pady=10)

    # 파일 선택
    frame_list = tk.Frame(frame_top)
    frame_list.pack(side=tk.LEFT, padx=20)
    btn_list = tk.Button(frame_list, text="명단 파일 선택", command=select_list)
    btn_list.pack()
    list_label = tk.Label(frame_list, text="(파일 경로 없음)", wraplength=250)
    list_label.pack()

    # 소스 선택
    frame_source = tk.Frame(frame_top)
    frame_source.pack(side=tk.LEFT, padx=20)
    btn_source = tk.Button(frame_source, text="신청서 정보 파일 선택", command=select_source)
    btn_source.pack()
    source_label = tk.Label(frame_source, text="(경로 없음)", wraplength=250)
    source_label.pack()

    # 타겟 선택
    frame_target = tk.Frame(frame_top)
    frame_target.pack(side=tk.LEFT, padx=20)
    btn_target = tk.Button(frame_target, text="양식 파일 선택", command=select_target)
    btn_target.pack()
    target_label = tk.Label(frame_target, text="(경로 없음)", wraplength=250)
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
