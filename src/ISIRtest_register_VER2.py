import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import re


# ══════════════════════════════════════════════════════════════
#  WidgetFactory : Label / Entry / Button 통합 관리
# ══════════════════════════════════════════════════════════════
class WidgetFactory:
    """Label, Entry, Button을 통합 관리하는 팩토리 클래스"""

    def __init__(self, root):
        self.root = root
        self._widgets = {}

    # ── 내부 헬퍼 ──────────────────────────────────────────
    def _store(self, name, widget):
        if name:
            self._widgets[name] = widget
        return widget

    def get(self, name):
        return self._widgets.get(name)

    # ── Label ──────────────────────────────────────────────
    def label(self, x, y, text="", name=None,
              font=("Arial", 11), color="black"):
        w = tk.Label(self.root, text=text, font=font,
                     justify="left", fg=color)
        w.place(x=x, y=y)
        return self._store(name, w)

    def update_label(self, name, new_text):
        w = self.get(name)
        if w:
            w.configure(text=new_text)

    # ── Entry ──────────────────────────────────────────────
    def entry(self, x, y, name=None, width=20,
              focus=False, on_enter=None):
        var = tk.StringVar()
        w = tk.Entry(self.root, textvariable=var, width=width)
        w.place(x=x, y=y)
        w._var = var

        if focus: # 처음 실행시 마우스 올리는 상태로 할지
            w.focus_set()
        if on_enter:
            w.bind("<Return>", lambda e: on_enter())

        return self._store(name, w)

    def get_value(self, name):
        w = self.get(name)
        return w._var.get() if w else ""

    def set_value(self, name, text):
        w = self.get(name)
        if w:
            w._var.set(text)

    # ── Button ─────────────────────────────────────────────
    def button(self, x, y, text="", command=None, name=None,
               font=("Arial", 11), color="black"):
        w = tk.Button(self.root, text=text, font=font,
                      fg=color, command=command)
        w.place(x=x, y=y)
        return self._store(name, w)


# ══════════════════════════════════════════════════════════════
#  유틸 함수
# ══════════════════════════════════════════════════════════════
def get_clipboard_data_as_list():
    # 엑셀에서 복사한 클립보드 데이터를 리스트로 반환
    # 클립보드만 쓰는 root를 만들고, 작업 종료 시 destory로 창 닫음
    root = tk.Tk()
    root.withdraw()
    try:
        raw_data = root.clipboard_get()
        normalized = raw_data.replace("\r", "").replace("\n", ",").replace('"', '')
        return [item.strip() for item in normalized.split(",") if item.strip()]
    except Exception as e:
        print(f"클립보드 읽기 실패: {e}")
        return []
    finally:
        root.destroy()


def extract_common_prefix(items):
    """항목 리스트에서 공통 접두어 추출 (예: DCP1, DCP2 → DCP)"""
    if not items:
        return ""
    prefixes = [re.match(r'^[A-Za-z]+', item) for item in items]
    prefixes = [m.group() for m in prefixes if m]
    if not prefixes:
        return ""
    return os.path.commonprefix(prefixes)


def natural_sort_key(s):
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', s)]


# ══════════════════════════════════════════════════════════════
#  탭1 : SMD 저항 확인
# ══════════════════════════════════════════════════════════════
class Tab_SMD(tk.Frame):
    """저항 이름 → 저항 값 순서로 입력받아 쌍으로 저장 후 엑셀 출력"""

    INPUT_STEPS = [
        "저항 이름을 입력해 주세요 (예: R1 → 1 입력)",
        "저항 값을 입력해 주세요",
    ]

    def __init__(self, parent):
        super().__init__(parent)
        self.ui = WidgetFactory(self)
        self.data_pairs: list[list[str]] = []
        self.temp_pair:  list[str]       = []
        self._build()

    def _build(self):
        self.ui.label(30, 30, text=self.INPUT_STEPS[0], name="guide")
        self.ui.entry(30, 60, name="inp", focus=True, on_enter=self._on_enter)
        self.ui.button(30, 90, text="완료", command=self._finish)

        self.ui.label(200, 120, text="입력 목록", name="list_title")

        self.listbox = tk.Listbox(self, width=35, height=15)
        self.listbox.place(x=200, y=145)

    # ── 입력 처리 ───────────────────────────────────────────
    def _on_enter(self):
        value = self.ui.get_value("inp").strip()
        if not value:
            return

        self.temp_pair.append(value)
        self.ui.set_value("inp", "")

        step = len(self.temp_pair)
        if step < len(self.INPUT_STEPS):
            self.ui.update_label("guide", self.INPUT_STEPS[step])
        else:
            self._save_pair()

    def _save_pair(self):
        pair = self.temp_pair[:]
        self.data_pairs.append(pair)
        name, value = pair
        self.listbox.insert(tk.END, f"R{name}  →  {value}")
        print(f"저장됨: {pair}")

        self.temp_pair = []
        self.ui.update_label("guide", self.INPUT_STEPS[0])
        self.winfo_toplevel().bell()

    # ── 완료 : DataFrame 가공 후 엑셀 저장 ─────────────────
    def _finish(self):
        if not self.data_pairs:
            messagebox.showwarning("데이터 없음", "저장된 데이터가 없습니다.")
            return

        df = pd.DataFrame(self.data_pairs, columns=["열2", "열1"])
        df_grouped = df.groupby("열1", as_index=False)["열2"].agg(','.join)

        df_grouped["열2"] = df_grouped["열2"].apply(
            lambda x: ','.join(sorted(
                set(['R' + item if not item.startswith('R') else item
                     for item in x.split(',')]),
                key=natural_sort_key
            ))
        )

        folder_path = filedialog.askdirectory(title="결과 파일을 저장할 폴더를 선택하세요")
        if folder_path:
            save_path = os.path.join(folder_path, "result.xlsx")
            df_grouped.to_excel(save_path, index=False)
            print(df_grouped)
            print(f"결과가 저장되었습니다: {save_path}")
            self.winfo_toplevel().destroy()
        else:
            print("저장이 취소되었습니다.")


# ══════════════════════════════════════════════════════════════
#  탭2 : 기타 소자 확인
# ══════════════════════════════════════════════════════════════
class Tab_Other(tk.Frame):
    """접두어 + 값 입력 또는 클립보드로 소자 목록 비교 후 엑셀 저장"""

    def __init__(self, parent):
        super().__init__(parent)
        self.ui = WidgetFactory(self)
        self.list_input: list[str] = []
        self.new_clipboard_items: list[str] = []
        self._build()

    def _build(self):
        # 접두어 입력
        self.ui.label(30, 20,  text="접두어 입력 (예: D):")
        self.ui.entry(180, 20, name="prefix")

        # 값 입력
        self.ui.label(30, 60,  text="값 입력:")
        self.ui.entry(100, 60, name="value", on_enter=self._add_to_list)

        # 버튼
        self.ui.button(100, 90, text="결과",            command=self._show_unique_values)
        self.ui.button(200, 90, text="클립보드 붙여넣기", command=self._load_clipboard)

        # 결과 레이블
        self.ui.label(30,  130, name="result_main",  color="blue")
        self.ui.label(230, 130, name="result_extra", color="blue")

    # ── 값 추가 ─────────────────────────────────────────────
    def _add_to_list(self):
        prefix = self.ui.get_value("prefix").strip().upper()
        if not prefix or not re.fullmatch(r"[A-Z]+", prefix):
            messagebox.showwarning("잘못된 접두어",
                                   "⚠ 접두어는 영어 알파벳만 입력해 주세요 (예: D, R 등)")
            return

        value = self.ui.get_value("value").strip()
        if value:
            self.list_input.append(value)
            print(f"입력값 추가됨: {value}")
        self.ui.set_value("value", "")

    # ── 결과 표시 및 엑셀 저장 ──────────────────────────────
    def _show_unique_values(self):
        prefix = self.ui.get_value("prefix").strip().upper()
        if not prefix or not re.fullmatch(r"[A-Z]+", prefix):
            messagebox.showwarning("잘못된 접두어",
                                   "⚠ 접두어는 영어 알파벳만 입력해 주세요 (예: D, R 등)")
            return

        unique_vals = sorted(set(self.list_input))
        mapped_vals = [f"{prefix}{val}" for val in unique_vals]

        missing_items: list[str] = []
        extra_items:   list[str] = []

        if self.new_clipboard_items:
            clipboard_set = set(self.new_clipboard_items)
            mapped_set    = set(mapped_vals)
            missing_items = sorted(clipboard_set - mapped_set)
            extra_items   = sorted(
                [i for i in (mapped_set - clipboard_set)
                 if not i.startswith(f"{prefix}{prefix}")]
            )

            if missing_items:
                print(f"⛔ 매핑되지 않은 항목: {missing_items}")
                self.ui.update_label("result_main",
                                     "❗미 발견 소자 항목:\n" + "\n".join(missing_items))
                if extra_items:
                    self.ui.update_label("result_extra",
                                         "❗PARTLIST에 없는 항목:\n" + "\n".join(extra_items))
            else:
                self.ui.update_label("result_main", "✅ 모든 항목이 매핑되었습니다.")
        else:
            self.ui.update_label("result_main", "고유 값:\n" + "\n".join(mapped_vals))

        # 엑셀 저장
        self._save_result(missing_items, extra_items)

    def _save_result(self, missing_items: list, extra_items: list):
        df_missing = pd.DataFrame(missing_items, columns=["미발견 소자 (missing_items)"])
        df_extra   = pd.DataFrame(extra_items,   columns=["PARTLIST에 없음 (extra_items)"])

        folder_path = filedialog.askdirectory(title="결과 파일을 저장할 폴더를 선택하세요")
        if folder_path:
            save_path = os.path.join(folder_path, "result2.xlsx")
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                df_missing.to_excel(writer, index=False, sheet_name="찾지 못한 소자")
                df_extra.to_excel(writer, index=False, sheet_name="PARTLIST에 존재하지 않음")

            messagebox.showinfo("저장 완료", f"결과 파일이 저장되었습니다:\n{save_path}")
            print(f"✅ result2.xlsx 저장 완료: {save_path}")
            self.winfo_toplevel().destroy()
        else:
            print("❌ 저장 경로 선택이 취소되었습니다.")

    # ── 클립보드 로드 ────────────────────────────────────────
    def _load_clipboard(self):
        new_items = get_clipboard_data_as_list()
        if not new_items:
            messagebox.showwarning("붙여넣기 실패",
                                   "클립보드에서 유효한 데이터를 찾을 수 없습니다.")
            return

        auto_prefix = extract_common_prefix(new_items)
        if auto_prefix:
            self.ui.set_value("prefix", auto_prefix)
        print(f"🔍 접두어 자동 지정: {auto_prefix}")

        self.new_clipboard_items = new_items
        self.list_input.extend(new_items)
        print(f"클립보드에서 추가된 항목: {new_items}")
        self.ui.update_label("result_main",
                             f"{len(new_items)}개 항목이 추가되었습니다.\nLIST : {new_items}")

# ══════════════════════════════════════════════════════════════
#  탭 3 : test
# ══════════════════════════════════════════════════════════════

class Tab_test(tk.Frame):
    def __init(self, parent):
        super().__init__(parent)
        self.ui = WidgetFactory(self)
        self._build()


# ══════════════════════════════════════════════════════════════
#  메인 앱 : 탭 조립만 담당
# ══════════════════════════════════════════════════════════════
class GUI_APP:
    TABS = [
        ("SMD 저항 확인",   Tab_SMD),
        ("기타 소자 확인",  Tab_Other),
    ]

    def __init__(self, root, title="부품 확인기"):
        self.root = root
        self.root.title(title)
        self.root.geometry("700x500+100+100")

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)

        for label, TabClass in self.TABS:
            tab = TabClass(nb)
            nb.add(tab, text=label)


# ══════════════════════════════════════════════════════════════
#  실행
# ══════════════════════════════════════════════════════════════
def main():
    root = tk.Tk()
    GUI_APP(root, "부품 확인기")
    root.mainloop()


if __name__ == "__main__":
    main()