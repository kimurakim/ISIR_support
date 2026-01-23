import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os, re

class MAKEroot_label:
    def __init__(self, root_name):
        self.root = root_name
        self.labels = {}

    def create_label(self, label_name, label_text, x, y, font_default=("Arial",11), font_color="black"):
        make_label = tk.Label(self.root, text=label_text, font=font_default, justify="left", fg=font_color)
        make_label.place(x=x, y=y)
        self.labels[label_name] = make_label

    def update_label(self, name, new_text):
        if name in self.labels:
            self.labels[name].configure(text=new_text)

class MAKEroot_entry:
    def __init__(self, objectname, x, y, focus=0, on_enter=None):
        self.var = tk.StringVar()
        self.make_entry = tk.Entry(objectname, textvariable=self.var)
        self.make_entry.place(x=x, y=y)
        if focus:
            self.make_entry.focus_set()
        if on_enter:
            self.make_entry.bind("<Return>", lambda event: on_enter())

class MAKEroot_btn:
    def __init__(self, root_name):
        self.root = root_name

    def create_btn(self, btn_text, x, y, command, font_default=("Arial",11), font_color="black"):
        make_btn = tk.Button(self.root, text=btn_text, font=font_default, fg=font_color, command=command)
        make_btn.place(x=x, y=y)
        return make_btn

class GUI_APP:
    def __init__(self, root, root_title):
        self.root = root
        self.root.title(root_title)
        self.root.geometry("700x500+100+100")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        self.tab1 = tk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="SMD 저항 확인")
        self.setup_tab1(self.tab1)

        self.tab2 = tk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="기타 소자 확인")
        self.setup_tab2(self.tab2)

    def setup_tab1(self, frame):
        self.make_label = MAKEroot_label(frame)
        self.make_btn = MAKEroot_btn(frame)
        self.data_pairs = []
        self.temp_pair = []

        self.make_label.create_label("entry_label", "저항 이름을 입력해 주세요(예 : R1일 경우에는 1 입력)", x=30, y=30)
        self.entry = MAKEroot_entry(frame, x=30, y=60, focus=1, on_enter=self.handle_enter)
        self.make_label.create_label("register_list", "입력 list", x=200, y=120)

        self.make_btn.create_btn("완료", x=30, y=90, command=self.finish_input)

    def handle_enter(self): #엔터 눌렀을 때
        value = self.entry.var.get().strip()
        if value:
            self.temp_pair.append(value) #값 추가(value에)
            self.entry.var.set("") #값 초기화
            if len(self.temp_pair) == 1:
                self.make_label.update_label("entry_label", "저항 값을 입력해 주세요")
            elif len(self.temp_pair) == 2:
                self.data_pairs.append(self.temp_pair)
                print(f"저장됨: {self.temp_pair}")
                self.temp_pair = []
                self.make_label.update_label("register_list", self.data_pairs)
                self.make_label.update_label("entry_label", "저항 이름을 입력해 주세요(예 : R1일 경우에는 1 입력)")
                self.root.bell()  #비프음 출력

    def finish_input(self):
        df = pd.DataFrame(self.data_pairs, columns=["열2", "열1"])
        df_grouped = df.groupby("열1", as_index=False)["열2"].agg(','.join)
        
        import re
        def natural_sort_key(s):
            return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', s)]
        
        df_grouped["열2"] = df_grouped["열2"].apply(
            lambda x: ','.join(sorted(
                set(['R' + item if not item.startswith('R') else item for item in x.split(',')]),
                key=natural_sort_key
            ))
    )


        folder_path = filedialog.askdirectory(title="결과 파일을 저장할 폴더를 선택하세요")
        if folder_path:
            save_path = os.path.join(folder_path, "result.xlsx")
            df_grouped.to_excel(save_path, index=False)
            print(df_grouped)
            print(f"결과가 저장되었습니다: {save_path}")
            self.root.destroy()
        else:
            print("저장이 취소되었습니다.")

    def setup_tab2(self, frame):
        self.list_input = []

        # Entry1: 접두어 입력 (예: D)
        tk.Label(frame, text="접두어 입력 (예: D):").place(x=30, y=20)
        self.entry_prefix_var = tk.StringVar()
        self.entry_prefix = tk.Entry(frame, textvariable=self.entry_prefix_var)
        self.entry_prefix.place(x=180, y=20)

        # Entry2: 값 입력
        tk.Label(frame, text="값 입력:").place(x=30, y=60)
        self.entry2_var = tk.StringVar()
        self.entry2 = tk.Entry(frame, textvariable=self.entry2_var)
        self.entry2.place(x=100, y=60)
        self.entry2.bind("<Return>", lambda e: self.add_to_list())

        #btn1 : 클립보드 값 따오기
        self.clip_btn = tk.Button(frame, text="클립보드 붙여넣기", command=self.load_clipboard)
        self.clip_btn.place(x=200, y=90)
        
        # 결과 표시 라벨
        self.result_label = tk.Label(frame, text="", fg="blue", justify="left")
        self.result_label.place(x=30, y=130)
        
        self.result_extra_label = tk.Label(frame, text="", fg="blue", justify="left")
        self.result_extra_label.place(x=230, y=130)

        # 결과 버튼
        self.result_btn = tk.Button(frame, text="결과", command=self.show_unique_values)
        self.result_btn.place(x=100, y=90)

    def add_to_list(self):
        prefix = self.entry_prefix_var.get().strip().upper()
        if not prefix or not re.fullmatch(r"[A-Z]+", prefix):
            messagebox.showwarning("잘못된 접두어", "⚠ 접두어는 영어 알파벳만 입력해 주세요 (예: D, R 등)")
            return

        value = self.entry2_var.get().strip()
        if value:
            self.list_input.append(value)
            print(f"입력값 추가됨: {value}")
        self.entry2_var.set("")

    def show_unique_values(self): #접두어 추가해서 입력값 중복 제거 후 결과 반환 (mapped_vals)
        prefix = self.entry_prefix_var.get().strip().upper()
        if not prefix or not re.fullmatch(r"[A-Z]+", prefix):
            messagebox.showwarning("잘못된 접두어", "⚠ 접두어는 영어 알파벳만 입력해 주세요 (예: D, R 등)")
            return

        unique_vals = sorted(set(self.list_input))
        mapped_vals = [f"{prefix}{val}" for val in unique_vals]

        # new_items (클립보드에서 붙여넣은 원본 값)과 mapped_vals 비교
        if hasattr(self, 'new_clipboard_items'):
            clipboard_set = set(self.new_clipboard_items)
            mapped_set = set(mapped_vals)
            # new_items에 있지만 mapped_vals에는 없는 값
            missing_items = sorted(clipboard_set - mapped_set)
            extra_items = sorted(mapped_set - clipboard_set)
            extra_items = [item for item in extra_items if not item.startswith(f"{prefix}{prefix}")]
            
            if missing_items:
                print(f"⛔ 매핑되지 않은 항목: {missing_items}")
                self.result_label.config(text=f"\n❗미 발견 소자 항목:\n" + "\n".join(missing_items))
                if extra_items:
                    self.result_extra_label.config(text=f"\n❗PARTLIST에 없는 항목:\n" + "\n".join(extra_items))
            else:
                self.result_label.config(text="✅ 모든 항목이 매핑되었습니다.")
        else:
            self.result_label.config(text="고유 값:\n" + "\n".join(mapped_vals))
        
        
        ### 엑셀 파일로 결과 저장 부분
        df_missing = pd.DataFrame(missing_items, columns=["미발견 소자 (missing_items)"])
        df_extra = pd.DataFrame(extra_items, columns=["PARTLIST에 없음 (extra_items)"])

        # 사용자에게 저장할 폴더 선택 요청
        folder_path = filedialog.askdirectory(title="결과 파일을 저장할 폴더를 선택하세요")
        if folder_path:
            save_path = os.path.join(folder_path, "result2.xlsx")
            
            # 하나의 Excel 파일에 두 시트로 저장 (선택사항)
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df_missing.to_excel(writer, index=False, sheet_name="찾지 못한 소자")
                df_extra.to_excel(writer, index=False, sheet_name="PARTLIST에 존재하지 않음")

            messagebox.showinfo("저장 완료", f"결과 파일이 저장되었습니다:\n{save_path}")
            print(f"✅ result2.xlsx 저장 완료: {save_path}")
            self.root.destroy()
        else:
            print("❌ 저장 경로 선택이 취소되었습니다.")
        
    
    
    def load_clipboard(self): #클립보드 가져오기
        new_items = get_clipboard_data_as_list()
        if not new_items:
            messagebox.showwarning("붙여넣기 실패", "클립보드에서 유효한 데이터를 찾을 수 없습니다.")
            return
        
    # 접두어 자동 추출 후 entry1에 지정
        auto_prefix = extract_common_prefix(new_items)
        if auto_prefix:
            self.entry_prefix_var.set(auto_prefix)
        print(f"🔍 접두어 자동 지정: {auto_prefix}")

        self.new_clipboard_items = new_items  # 마지막 붙여넣기 내용 저장
        self.list_input.extend(new_items)
        print(f"클립보드에서 추가된 항목: {new_items}")
        self.result_label.config(text=f"{len(new_items)}개 항목이 추가되었습니다.\nLIST : {new_items}")
        
        


def get_clipboard_data_as_list(): #접두어(R,L 등) 엑셀에서 클립보드 따오기 위해서 제작
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()  # GUI 창 숨김
    try:
        raw_data = root.clipboard_get()
        normalized = raw_data.replace("\r", "").replace("\n", ",").replace('"', '')
        return [item.strip() for item in normalized.split(",") if item.strip()]
    except Exception as e:
        print(f"클립보드 읽기 실패: {e}")
        return []

def extract_common_prefix(items): #접두어 자동 추출
    if not items:
        return ""

    # 숫자 시작 전까지 문자만 추출 (예: "DCP1" → "DCP")
    prefixes = [re.match(r'^[A-Za-z]+', item) for item in items]
    prefixes = [m.group() for m in prefixes if m]

    if not prefixes:
        return ""

    # 공통 접두어 구하기 (예: DCP1, DCP2, DCP3 → DCP)
    prefix = os.path.commonprefix(prefixes)
    return prefix


def main():
    root = tk.Tk()
    app = GUI_APP(root, "GUI 앱")
    root.mainloop()

main()
