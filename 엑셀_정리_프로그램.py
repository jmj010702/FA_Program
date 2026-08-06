# -*- coding: utf-8 -*-
"""
시트 정리 프로그램 - 화면 (GUI)

RMA 시트의 특정 열을 골라서 새 시트/기존 시트로 정리해주는 tkinter 화면.
실제 엑셀 읽기/쓰기 로직은 core.py 에 있고, 이 파일은 화면과 사용자 입력만 담당한다.

프로그램 소개, 사용법, 버전 히스토리는 README.md 참고.

필요한 라이브러리: openpyxl
설치: pip install openpyxl
"""

import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from core import get_sheet_names, read_raw_sheet, compute_result_rows, run_conversion


# ------------------------------------------------------------------
# 화면 (GUI)
# ------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("시트 정리 프로그램")
        root.geometry("760x700")
        root.minsize(700, 550)

        self.filepath = tk.StringVar()
        self.target_mode = tk.StringVar(value="new")
        self._raw_row_data = {}
        self._preview_row_data = {}

        # ---- 파일 선택 ----
        frame_file = tk.Frame(root)
        frame_file.pack(fill="x", padx=10, pady=(10, 5))
        tk.Label(frame_file, text="엑셀 파일:").pack(side="left")
        tk.Entry(frame_file, textvariable=self.filepath, width=55).pack(side="left", padx=5)
        tk.Button(frame_file, text="찾아보기", command=self.browse_file).pack(side="left")

        self.sheet_list_label = tk.Label(
            root, text="파일을 선택하면 이 파일에 있는 시트 목록이 여기에 표시됩니다.",
            fg="gray30", anchor="w", justify="left", wraplength=720
        )
        self.sheet_list_label.pack(fill="x", padx=12, pady=(0, 5))

        # ---- 원본 시트 옵션 ----
        frame_opt = tk.LabelFrame(root, text="원본 시트 설정")
        frame_opt.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_opt, text="가져올 시트 이름:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.rma_sheet_name = tk.StringVar(value="RMA")
        self.combo_rma_sheet = ttk.Combobox(frame_opt, textvariable=self.rma_sheet_name, width=15)
        self.combo_rma_sheet.grid(row=0, column=1, padx=5)
        self.combo_rma_sheet.bind("<<ComboboxSelected>>", self.on_rma_sheet_selected)

        tk.Label(frame_opt, text="헤더(열 제목) 행 번호:").grid(row=0, column=2, sticky="w", padx=(20, 0))
        self.rma_header_row = tk.StringVar(value="2")
        tk.Entry(frame_opt, textvariable=self.rma_header_row, width=5).grid(row=0, column=3, padx=5)

        tk.Label(
            frame_opt,
            text="열 제목이 있는 줄 번호예요. (예: 2행에 제목이 있으면 2 → 데이터는 3행부터 자동으로 읽음)",
            fg="gray30", justify="left"
        ).grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=(0, 5))

        tk.Label(frame_opt, text="가져올 열 (비우면 기본 15개 열 사용):").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.column_spec = tk.StringVar(value="")
        tk.Entry(frame_opt, textvariable=self.column_spec, width=50).grid(row=2, column=1, columnspan=3, sticky="w", padx=5)

        tk.Label(
            frame_opt,
            text="입력 예시: 열이름1,열이름2  또는  1,3,5  또는  A,C,E",
            fg="gray30", justify="left"
        ).grid(row=3, column=0, columnspan=4, sticky="w", padx=5, pady=(0, 5))

        # ---- 대상 시트 옵션 ----
        frame_target = tk.LabelFrame(root, text="결과를 저장할 위치")
        frame_target.pack(fill="x", padx=10, pady=5)

        tk.Radiobutton(frame_target, text="새 시트 만들기", variable=self.target_mode, value="new",
                        command=self.update_target_widgets).grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.new_sheet_name = tk.StringVar(value=datetime.datetime.now().strftime("SYSTEM_%Y%m%d"))
        self.entry_new_name = tk.Entry(frame_target, textvariable=self.new_sheet_name, width=25)
        self.entry_new_name.grid(row=0, column=1, sticky="w", padx=5)

        tk.Radiobutton(frame_target, text="기존 시트에 이어서 추가", variable=self.target_mode, value="existing",
                        command=self.update_target_widgets).grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.existing_sheet_name = tk.StringVar()
        self.combo_existing = ttk.Combobox(frame_target, textvariable=self.existing_sheet_name, width=23, state="disabled")
        self.combo_existing.grid(row=1, column=1, sticky="w", padx=5)

        # ---- 복사 옵션 ----
        frame_copy_opt = tk.LabelFrame(root, text="복사 옵션")
        frame_copy_opt.pack(fill="x", padx=10, pady=5)

        self.copy_style = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_copy_opt, text="서식(글자색/배경색/테두리 등)도 함께 복사",
                        variable=self.copy_style).grid(row=0, column=0, sticky="w", padx=5, pady=3)

        self.copy_formula = tk.BooleanVar(value=True)
        tk.Checkbutton(frame_copy_opt, text="원본이 수식이면 값 대신 수식 자체를 복사",
                        variable=self.copy_formula).grid(row=1, column=0, sticky="w", padx=5, pady=3)

        # ---- 미리보기 / 실행 버튼 ----
        frame_buttons = tk.Frame(root)
        frame_buttons.pack(fill="x", padx=10, pady=10)

        tk.Button(
            frame_buttons, text="미리보기 (저장 안 함)", command=self.on_preview,
            bg="#4a5568", fg="white", font=("Malgun Gothic", 11, "bold"), height=2
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))

        tk.Button(
            frame_buttons, text="실행 (실제 저장)", command=self.on_run,
            bg="#2b6cb0", fg="white", font=("Malgun Gothic", 11, "bold"), height=2
        ).pack(side="left", fill="x", expand=True, padx=(5, 0))

        # ---- 결과를 보여줄 탭 (원본 시트 보기 / 미리보기 표 / 처리 로그) ----
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # 원본 시트 보기 탭 (파일에 실제로 뭐가 들어있는지 그대로 보여줌)
        frame_raw = tk.Frame(notebook)
        notebook.add(frame_raw, text="원본 시트 보기")

        raw_top = tk.Frame(frame_raw)
        raw_top.pack(fill="x", padx=5, pady=5)
        tk.Label(raw_top, text="볼 시트:").pack(side="left")
        self.raw_view_sheet = tk.StringVar()
        self.combo_raw_sheet = ttk.Combobox(raw_top, textvariable=self.raw_view_sheet, width=20, state="readonly")
        self.combo_raw_sheet.pack(side="left", padx=5)
        self.combo_raw_sheet.bind("<<ComboboxSelected>>", self.on_raw_sheet_selected)
        self.raw_info = tk.Label(raw_top, text="", fg="gray30")
        self.raw_info.pack(side="left", padx=10)

        raw_tree_frame = tk.Frame(frame_raw)
        raw_tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.raw_tree = ttk.Treeview(raw_tree_frame, show="headings")
        raw_vsb = ttk.Scrollbar(raw_tree_frame, orient="vertical", command=self.raw_tree.yview)
        raw_hsb = ttk.Scrollbar(raw_tree_frame, orient="horizontal", command=self.raw_tree.xview)
        self.raw_tree.configure(yscrollcommand=raw_vsb.set, xscrollcommand=raw_hsb.set)
        self.raw_tree.grid(row=0, column=0, sticky="nsew")
        raw_vsb.grid(row=0, column=1, sticky="ns")
        raw_hsb.grid(row=1, column=0, sticky="ew")
        raw_tree_frame.grid_rowconfigure(0, weight=1)
        raw_tree_frame.grid_columnconfigure(0, weight=1)
        # Shift + 마우스 휠로도 좌우 스크롤 가능하게 (창이 좁아져 열이 안 보일 때 편리)
        self.raw_tree.bind("<Shift-MouseWheel>", lambda e: self.raw_tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.raw_tree.bind("<<TreeviewSelect>>", self.on_raw_row_selected)

        self.raw_detail = tk.Label(
            frame_raw, text="※ 표에서 행을 클릭하면 그 행 전체 내용이 여기 표시됩니다. (열이 잘려도 확인 가능)",
            fg="gray30", justify="left", anchor="w",
            wraplength=720, padx=8, pady=6
        )
        self.raw_detail.pack(fill="x", padx=5, pady=(0, 5))

        # 미리보기 탭 (매칭 결과)
        frame_preview = tk.Frame(notebook)
        notebook.add(frame_preview, text="결과 미리보기")

        self.preview_info = tk.Label(frame_preview, text="아직 미리보기를 실행하지 않았습니다.", fg="gray30", anchor="w")
        self.preview_info.pack(fill="x", padx=5, pady=(5, 0))

        tree_frame = tk.Frame(frame_preview)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.tree.bind("<Shift-MouseWheel>", lambda e: self.tree.xview_scroll(int(-1 * (e.delta / 120)), "units"))
        self.tree.bind("<<TreeviewSelect>>", self.on_preview_row_selected)

        self.preview_detail = tk.Label(
            frame_preview, text="※ 표에서 행을 클릭하면 그 행 전체 내용이 여기 표시됩니다. (열이 잘려도 확인 가능)",
            fg="gray30", justify="left", anchor="w",
            wraplength=720, padx=8, pady=6
        )
        self.preview_detail.pack(fill="x", padx=5, pady=(0, 5))

        # 로그 탭
        frame_log = tk.Frame(notebook)
        notebook.add(frame_log, text="처리 로그")
        self.log_box = scrolledtext.ScrolledText(
            frame_log, height=16
        )
        self.log_box.pack(fill="both", expand=True, padx=5, pady=5)

        self.notebook = notebook
        self.frame_preview = frame_preview
        self.frame_raw = frame_raw
        self.frame_log = frame_log

    def update_target_widgets(self):
        if self.target_mode.get() == "new":
            self.entry_new_name.config(state="normal")
            self.combo_existing.config(state="disabled")
        else:
            self.entry_new_name.config(state="disabled")
            self.combo_existing.config(state="readonly")

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="RMA 시트가 들어있는 엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        if path:
            self.filepath.set(path)
            try:
                names = get_sheet_names(path)
                self.sheet_list_label.config(
                    text=f"이 파일의 시트 목록 ({len(names)}개): " + ", ".join(names)
                )
                self.combo_existing["values"] = names
                if names:
                    self.existing_sheet_name.set(names[0])
                self.combo_rma_sheet["values"] = names
                if names:
                    # 이름에 'RMA'가 들어간 시트가 있으면 그걸 기본 선택, 없으면 첫번째 시트
                    guess = next((n for n in names if "RMA" in n.upper()), names[0])
                    self.rma_sheet_name.set(guess)
                self.combo_raw_sheet["values"] = names
                if names:
                    self.raw_view_sheet.set(guess)
                    self.load_raw_sheet(path, guess)
            except Exception as e:
                messagebox.showwarning("알림", f"시트 목록을 읽는 중 문제가 발생했습니다: {e}")

    def log(self, message):
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.root.update_idletasks()

    def fill_tree(self, headers, rows, filled_names=None):
        filled_names = set(filled_names or [])
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = headers
        self._preview_row_data = {}  # 행 iid -> {열이름: 값} (상세보기용)
        first_filled_idx = None
        for i, h in enumerate(headers):
            is_filled = h in filled_names
            label = f"★ {h}" if is_filled else h
            self.tree.heading(h, text=label)
            # stretch=False: 창을 좁혀도 글씨가 뭉개지지 않고, 대신 가로 스크롤로 보게 함
            self.tree.column(h, width=max(80, min(160, len(str(h)) * 12 + 40)), anchor="w", stretch=False)
            if is_filled and first_filled_idx is None:
                first_filled_idx = i
        for row in rows:
            display_row = ["" if v is None else v for v in row]
            iid = self.tree.insert("", "end", values=display_row)
            self._preview_row_data[iid] = dict(zip(headers, display_row))

        # 이번에 실제로 채워지는 열이 화면 오른쪽에 있어서 안 보이는 경우, 그 열이 보이게 자동 스크롤
        if first_filled_idx is not None and headers:
            self.root.update_idletasks()
            fraction = max(0, (first_filled_idx - 1)) / max(1, len(headers))
            self.tree.xview_moveto(fraction)
        self.preview_detail.config(text="※ 표에서 행을 클릭하면 그 행 전체 내용이 여기 표시됩니다. (열이 잘려도 확인 가능)")

    def on_preview_row_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        row_data = self._preview_row_data.get(sel[0], {})
        text = "   |   ".join(f"{k}: {v}" for k, v in row_data.items() if k)
        self.preview_detail.config(text=text or "(빈 행)")

    def load_raw_sheet(self, path, sheet_name):
        try:
            col_letters, rows, max_row = read_raw_sheet(path, sheet_name, max_rows=300)
        except Exception as e:
            messagebox.showerror("오류", f"시트 내용을 읽는 중 문제가 발생했습니다: {e}")
            return

        self.raw_tree.delete(*self.raw_tree.get_children())
        self.raw_tree["columns"] = col_letters
        self._raw_row_data = {}  # 행 iid -> {열문자: 값} (상세보기용)
        for h in col_letters:
            self.raw_tree.heading(h, text=h)
            width = 45 if h == "행" else 100
            # stretch=False: 창을 좁혀도 글씨가 뭉개지지 않고, 대신 가로 스크롤로 보게 함
            self.raw_tree.column(h, width=width, anchor="w", stretch=False)
        for row in rows:
            display_row = ["" if v is None else v for v in row]
            iid = self.raw_tree.insert("", "end", values=display_row)
            self._raw_row_data[iid] = dict(zip(col_letters, display_row))

        info = f"'{sheet_name}' 시트: 전체 {max_row}행 중 {len(rows)}행 표시"
        if max_row > 300:
            info += " (300행까지만 미리보기, 실제 데이터는 전체가 있습니다)"
        self.raw_info.config(text=info)
        self.raw_detail.config(text="※ 표에서 행을 클릭하면 그 행 전체 내용이 여기 표시됩니다. (열이 잘려도 확인 가능)")

    def on_raw_row_selected(self, event=None):
        sel = self.raw_tree.selection()
        if not sel:
            return
        row_data = self._raw_row_data.get(sel[0], {})
        text = "   |   ".join(f"{k}: {v}" for k, v in row_data.items() if k and k != "행")
        self.raw_detail.config(text=text or "(빈 행)")

    def on_raw_sheet_selected(self, event=None):
        """'볼 시트' 드롭다운에서 시트를 고르면 바로 내용을 보여주고, RMA 시트 이름도 같이 맞춘다."""
        sheet = self.raw_view_sheet.get().strip()
        if not sheet:
            return
        self.rma_sheet_name.set(sheet)  # RMA 시트 이름 콤보박스와 동기화
        path = self.filepath.get().strip()
        if path and os.path.exists(path):
            self.load_raw_sheet(path, sheet)

    def on_rma_sheet_selected(self, event=None):
        """'RMA 시트 이름' 드롭다운에서 시트를 고르면, '볼 시트'도 같이 맞추고 내용을 보여준다."""
        sheet = self.rma_sheet_name.get().strip()
        if not sheet:
            return
        self.raw_view_sheet.set(sheet)  # 볼 시트 콤보박스와 동기화
        path = self.filepath.get().strip()
        if path and os.path.exists(path):
            self.load_raw_sheet(path, sheet)

    def refresh_sheet_lists(self, path, prefer_sheet=None):
        """파일에 시트가 추가/변경된 뒤 각종 드롭다운과 안내 문구를 다시 채운다."""
        try:
            names = get_sheet_names(path)
        except Exception as e:
            self.log(f"시트 목록을 다시 읽는 중 문제가 발생했습니다: {e}")
            return

        self.sheet_list_label.config(text=f"이 파일의 시트 목록 ({len(names)}개): " + ", ".join(names))
        self.combo_existing["values"] = names
        self.combo_rma_sheet["values"] = names
        self.combo_raw_sheet["values"] = names

        if prefer_sheet and prefer_sheet in names:
            self.existing_sheet_name.set(prefer_sheet)
        elif names and self.existing_sheet_name.get() not in names:
            self.existing_sheet_name.set(names[0])

        # 지금 '원본 시트 보기' 탭에서 보고 있던 시트가 여전히 있으면 그대로 새로고침
        if self.raw_view_sheet.get() in names:
            self.load_raw_sheet(path, self.raw_view_sheet.get())

    def _validate_common_inputs(self):
        path = self.filepath.get().strip()
        if not path:
            messagebox.showwarning("알림", "먼저 엑셀 파일을 선택해주세요.")
            return None
        if not os.path.exists(path):
            messagebox.showerror("오류", "파일을 찾을 수 없습니다.")
            return None
        try:
            header_row = int(self.rma_header_row.get().strip())
        except ValueError:
            messagebox.showerror("오류", "헤더 행 번호는 숫자로 입력해주세요.")
            return None
        mode = self.target_mode.get()
        if mode == "existing" and not self.existing_sheet_name.get().strip():
            messagebox.showwarning("알림", "이어서 추가할 기존 시트를 선택해주세요.")
            return None
        return path, header_row, mode

    def on_preview(self):
        validated = self._validate_common_inputs()
        if not validated:
            return
        path, header_row, mode = validated

        self.log_box.delete("1.0", "end")
        try:
            headers, rows, missing_cols, filled_names = compute_result_rows(
                filepath=path,
                rma_sheet_name=self.rma_sheet_name.get().strip(),
                rma_header_row=header_row,
                column_spec_text=self.column_spec.get().strip(),
                target_mode=mode,
                existing_sheet_name=self.existing_sheet_name.get().strip(),
                log=self.log,
                max_preview_rows=200,
            )
            self.fill_tree(headers, rows, filled_names)
            info = (
                f"미리보기: {len(rows)}개 행 (실제 실행 시 원본 파일에 저장됩니다. 아직 저장되지 않았습니다.)  "
                f"★표시된 열 = 이번에 실제로 채워지는 열: {', '.join(filled_names)}"
            )
            if len(rows) >= 200:
                info += "  ※ 화면에는 최대 200행까지만 보여줍니다. 실제 실행 시엔 전체가 처리됩니다."
            self.preview_info.config(text=info)
            self.notebook.select(self.frame_preview)
        except Exception as e:
            self.log(f"오류 발생: {e}")
            self.notebook.select(self.frame_log)  # 로그 탭
            messagebox.showerror("오류", str(e))

    def on_run(self):
        validated = self._validate_common_inputs()
        if not validated:
            return
        path, header_row, mode = validated

        self.log_box.delete("1.0", "end")
        self.notebook.select(self.frame_log)  # 실행 시작하면 로그 탭으로 전환
        try:
            sheet = run_conversion(
                filepath=path,
                rma_sheet_name=self.rma_sheet_name.get().strip(),
                rma_header_row=header_row,
                column_spec_text=self.column_spec.get().strip(),
                target_mode=mode,
                new_sheet_name=self.new_sheet_name.get().strip(),
                existing_sheet_name=self.existing_sheet_name.get().strip(),
                copy_style=self.copy_style.get(),
                copy_formula=self.copy_formula.get(),
                log=self.log,
            )
            # 새로 생기거나 갱신된 시트가 바로 드롭다운/목록/원본 보기에 반영되도록 새로고침
            self.refresh_sheet_lists(path, prefer_sheet=sheet)
            messagebox.showinfo("완료", f"'{sheet}' 시트에 정리가 끝났습니다. (목록이 자동으로 새로고침되었습니다)")
        except Exception as e:
            self.log(f"오류 발생: {e}")
            messagebox.showerror("오류", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
