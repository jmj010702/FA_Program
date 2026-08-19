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

from core import get_sheet_names, get_sheet_headers, read_raw_sheet, compute_result_rows, run_conversion


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

        # "가져올 열" 빌더 UI 상태: 추가한 열들을 순서대로 담아두고,
        # 여기서 core.py에 넘길 column_spec 문자열을 자동으로 만들어낸다.
        self.column_entries = []
        self.column_spec = tk.StringVar(value="")

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

        tk.Label(frame_opt, text="가져올 열 (비워두면 기본 15개 열 사용):").grid(row=0, column=0, sticky="w", padx=5, pady=(5, 3))

        self._sheet_header_rows = {}  # {시트이름: 마지막으로 사용한 헤더 행 번호} 자동 기억용

        frame_col_add = tk.Frame(frame_opt)
        frame_col_add.grid(row=1, column=0, columnspan=4, sticky="w", padx=5, pady=2)

        # 1번째 줄: 시트 / 헤더 행 / 열 / 새로고침
        tk.Label(frame_col_add, text="시트:").grid(row=0, column=0, sticky="w")
        self.col_sheet_source = tk.StringVar(value="RMA")
        self.combo_col_sheet = ttk.Combobox(frame_col_add, textvariable=self.col_sheet_source, width=14, state="readonly")
        self.combo_col_sheet.grid(row=0, column=1, sticky="w", padx=(2, 10))
        self.combo_col_sheet.bind("<<ComboboxSelected>>", self.on_col_sheet_changed)

        tk.Label(frame_col_add, text="헤더 행:").grid(row=0, column=2, sticky="w")
        self.rma_header_row = tk.StringVar(value="2")
        tk.Entry(frame_col_add, textvariable=self.rma_header_row, width=4).grid(row=0, column=3, sticky="w", padx=(2, 10))

        tk.Label(frame_col_add, text="열:").grid(row=0, column=4, sticky="w")
        self.col_ref_selected = tk.StringVar(value="")
        self.combo_col_ref = ttk.Combobox(frame_col_add, textvariable=self.col_ref_selected, width=16, state="readonly")
        self.combo_col_ref.grid(row=0, column=5, sticky="w", padx=(2, 4))

        tk.Button(frame_col_add, text="새로고침", command=self.refresh_col_ref_options).grid(row=0, column=6, sticky="w")

        # 2번째 줄: 출력 이름 / 추가 / 공백 열 추가
        tk.Label(frame_col_add, text="출력 이름(선택):").grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.col_output_name = tk.StringVar(value="")
        tk.Entry(frame_col_add, textvariable=self.col_output_name, width=20).grid(
            row=1, column=2, columnspan=2, sticky="w", padx=(2, 10), pady=(6, 0)
        )

        tk.Button(frame_col_add, text="추가", command=self.add_column_entry).grid(
            row=1, column=4, sticky="w", padx=(0, 5), pady=(6, 0)
        )
        tk.Button(frame_col_add, text="공백 열 추가", command=self.add_blank_entry).grid(
            row=1, column=5, columnspan=2, sticky="w", pady=(6, 0)
        )

        frame_col_list = tk.Frame(frame_opt)
        frame_col_list.grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=(8, 3))

        self.column_listbox = tk.Listbox(frame_col_list, width=62, height=5, exportselection=False)
        self.column_listbox.pack(side="left")

        frame_col_list_buttons = tk.Frame(frame_col_list)
        frame_col_list_buttons.pack(side="left", padx=(8, 0), fill="y")
        tk.Button(frame_col_list_buttons, text="▲ 위로", width=10,
                  command=lambda: self.move_selected_column_entry(-1)).pack(fill="x")
        tk.Button(frame_col_list_buttons, text="▼ 아래로", width=10,
                  command=lambda: self.move_selected_column_entry(1)).pack(fill="x", pady=(2, 0))
        tk.Button(frame_col_list_buttons, text="선택 삭제", width=10,
                  command=self.remove_selected_column_entry).pack(fill="x", pady=(8, 0))
        tk.Button(frame_col_list_buttons, text="전체 지우기", width=10,
                  command=self.clear_column_entries).pack(fill="x", pady=(2, 0))

        self.spec_preview_label = tk.Label(
            frame_opt, text="만들어지는 입력값: (없음 → 기본 15개 열 사용)",
            fg="gray30", justify="left", anchor="w", wraplength=700
        )
        self.spec_preview_label.grid(row=3, column=0, columnspan=4, sticky="w", padx=5, pady=(0, 5))

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
        self.combo_existing.bind("<<ComboboxSelected>>", self.on_existing_sheet_selected)

        tk.Label(frame_target, text="새 열 삽입 위치 (선택, 비우면 맨 뒤에 추가):").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.insert_ref_column = tk.StringVar(value="")
        self.combo_insert_ref = ttk.Combobox(frame_target, textvariable=self.insert_ref_column, width=20, state="disabled")
        self.combo_insert_ref.grid(row=2, column=1, sticky="w", padx=5)

        self.insert_position = tk.StringVar(value="after")
        frame_insert_pos = tk.Frame(frame_target)
        frame_insert_pos.grid(row=2, column=2, columnspan=2, sticky="w")
        self.radio_insert_before = tk.Radiobutton(frame_insert_pos, text="선택한 열 앞에", variable=self.insert_position, value="before")
        self.radio_insert_before.pack(side="left")
        self.radio_insert_after = tk.Radiobutton(frame_insert_pos, text="선택한 열 뒤에", variable=self.insert_position, value="after")
        self.radio_insert_after.pack(side="left")

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
            self.combo_insert_ref.config(state="disabled")
        else:
            self.entry_new_name.config(state="disabled")
            self.combo_existing.config(state="readonly")
            self.combo_insert_ref.config(state="readonly")

    # ---- "가져올 열" 빌더 ----

    def refresh_col_ref_options(self):
        """'시트' 콤보박스에서 고른 시트의 실제 헤더 이름들로 '열' 콤보박스를 채운다."""
        path = self.filepath.get().strip()
        sheet = self.col_sheet_source.get().strip()
        try:
            header_row = int(self.rma_header_row.get().strip())
        except ValueError:
            header_row = 1

        headers = []
        if path and sheet and os.path.exists(path):
            try:
                headers = get_sheet_headers(path, sheet, header_row)
            except Exception:
                headers = []

        self.combo_col_ref["values"] = headers
        self.col_ref_selected.set(headers[0] if headers else "")

    def on_col_sheet_changed(self, event=None):
        """'시트' 콤보박스를 바꾸면, 그 시트에서 마지막으로 썼던 헤더 행 번호를 자동으로 불러온다."""
        sheet = self.col_sheet_source.get().strip()
        if sheet in self._sheet_header_rows:
            self.rma_header_row.set(str(self._sheet_header_rows[sheet]))
        self.refresh_col_ref_options()

    def _get_primary_sheet_name(self):
        """실제 처리(행 순회)의 기준이 되는 시트 이름.
        이미 추가한 열이 있으면 그중 첫 번째 '실제' 열(빈 열 제외)의 시트를 기준으로 삼고,
        아직 추가한 열이 없으면 지금 '시트' 콤보박스에서 선택 중인 값을 기준으로 삼는다."""
        first_real = next((e for e in self.column_entries if not e["is_blank"]), None)
        if first_real:
            return first_real["sheet_name"]
        return self.col_sheet_source.get().strip()

    def _get_primary_header_row(self):
        """기준 시트의 헤더 행 번호. 규칙은 _get_primary_sheet_name과 동일하게 첫 번째 실제 열 기준."""
        first_real = next((e for e in self.column_entries if not e["is_blank"]), None)
        if first_real:
            return first_real["header_row"]
        return int(self.rma_header_row.get().strip())

    def _get_sheet_header_rows(self):
        """{시트이름: 헤더 행 번호} - 지금까지 추가한 열들이 사용한 시트별 헤더 행을 core.py에 그대로 전달."""
        return {
            e["sheet_name"]: e["header_row"]
            for e in self.column_entries
            if not e["is_blank"]
        }

    def add_column_entry(self):
        sheet = self.col_sheet_source.get().strip()
        ref = self.col_ref_selected.get().strip()
        if not sheet or not ref:
            messagebox.showwarning("알림", "먼저 시트와 열을 선택해주세요.")
            return
        try:
            header_row = int(self.rma_header_row.get().strip())
        except ValueError:
            messagebox.showerror("오류", "헤더 행 번호는 숫자로 입력해주세요.")
            return
        output_name = self.col_output_name.get().strip() or ref
        self.column_entries.append(
            {"output_name": output_name, "sheet_name": sheet, "ref": ref, "is_blank": False, "header_row": header_row}
        )
        self._sheet_header_rows[sheet] = header_row
        self.col_output_name.set("")
        self._refresh_column_listbox()

    def add_blank_entry(self):
        output_name = self.col_output_name.get().strip()
        self.column_entries.append(
            {"output_name": output_name, "sheet_name": None, "ref": None, "is_blank": True, "header_row": None}
        )
        self.col_output_name.set("")
        self._refresh_column_listbox()

    def remove_selected_column_entry(self):
        sel = self.column_listbox.curselection()
        if not sel:
            return
        del self.column_entries[sel[0]]
        self._refresh_column_listbox()

    def move_selected_column_entry(self, direction):
        sel = self.column_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        new_idx = idx + direction
        if 0 <= new_idx < len(self.column_entries):
            self.column_entries[idx], self.column_entries[new_idx] = (
                self.column_entries[new_idx],
                self.column_entries[idx],
            )
            self._refresh_column_listbox()
            self.column_listbox.selection_set(new_idx)

    def clear_column_entries(self):
        self.column_entries = []
        self._refresh_column_listbox()

    def _entry_label(self, entry):
        if entry["is_blank"]:
            return f"(빈 열: {entry['output_name']})" if entry["output_name"] else "(빈 열)"
        primary = self._get_primary_sheet_name()
        src = entry["ref"] if entry["sheet_name"] == primary else f"{entry['sheet_name']}!{entry['ref']}"
        return f"{entry['output_name']}  ←  {src}"

    def _refresh_column_listbox(self):
        self.column_listbox.delete(0, "end")
        for entry in self.column_entries:
            self.column_listbox.insert("end", self._entry_label(entry))
        self._refresh_spec_preview()

    def _build_column_spec_text(self):
        primary = self._get_primary_sheet_name()
        tokens = []
        for entry in self.column_entries:
            if entry["is_blank"]:
                tokens.append(f"{entry['output_name']}=" if entry["output_name"] else "")
                continue
            sheet_prefix = f"{entry['sheet_name']}!" if entry["sheet_name"] and entry["sheet_name"] != primary else ""
            tokens.append(f"{entry['output_name']}={sheet_prefix}{entry['ref']}")
        return ",".join(tokens)

    def _refresh_spec_preview(self):
        spec_text = self._build_column_spec_text()
        self.column_spec.set(spec_text)
        primary = self._get_primary_sheet_name()
        primary_desc = f"기준 시트: {primary}" if primary else "기준 시트: (선택 안 됨)"
        self.spec_preview_label.config(
            text=f"{primary_desc}   |   만들어지는 입력값: " + (spec_text or "(없음 → 기본 15개 열 사용)")
        )

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
                    self.refresh_insert_ref_options(path, names[0])
                # 이름에 'RMA'가 들어간 시트가 있으면 그걸 기본 선택, 없으면 첫번째 시트
                guess = next((n for n in names if "RMA" in n.upper()), names[0]) if names else ""
                self.combo_raw_sheet["values"] = names
                if names:
                    self.raw_view_sheet.set(guess)
                    self.load_raw_sheet(path, guess)

                self.combo_col_sheet["values"] = names
                self.column_entries = []
                self._sheet_header_rows = {}
                if names:
                    self.col_sheet_source.set(guess)
                    self.refresh_col_ref_options()
                self._refresh_column_listbox()
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
        """'볼 시트' 드롭다운에서 시트를 고르면 바로 내용을 보여준다."""
        sheet = self.raw_view_sheet.get().strip()
        if not sheet:
            return
        path = self.filepath.get().strip()
        if path and os.path.exists(path):
            self.load_raw_sheet(path, sheet)

    def on_existing_sheet_selected(self, event=None):
        """'기존 시트에 이어서 추가' 대상 시트를 고르면, 삽입 위치 콤보박스도 그 시트의 헤더로 갱신."""
        path = self.filepath.get().strip()
        sheet = self.existing_sheet_name.get().strip()
        if path and sheet and os.path.exists(path):
            self.refresh_insert_ref_options(path, sheet)

    def refresh_insert_ref_options(self, path, sheet_name):
        """삽입 위치 콤보박스의 목록을 지정한 시트의 현재 헤더 이름들로 채운다."""
        try:
            headers = get_sheet_headers(path, sheet_name)
        except Exception:
            headers = []
        self.combo_insert_ref["values"] = headers
        if self.insert_ref_column.get() not in headers:
            self.insert_ref_column.set("")

    def refresh_sheet_lists(self, path, prefer_sheet=None):
        """파일에 시트가 추가/변경된 뒤 각종 드롭다운과 안내 문구를 다시 채운다."""
        try:
            names = get_sheet_names(path)
        except Exception as e:
            self.log(f"시트 목록을 다시 읽는 중 문제가 발생했습니다: {e}")
            return

        self.sheet_list_label.config(text=f"이 파일의 시트 목록 ({len(names)}개): " + ", ".join(names))
        self.combo_existing["values"] = names
        self.combo_raw_sheet["values"] = names
        self.combo_col_sheet["values"] = names

        if prefer_sheet and prefer_sheet in names:
            self.existing_sheet_name.set(prefer_sheet)
        elif names and self.existing_sheet_name.get() not in names:
            self.existing_sheet_name.set(names[0])

        if self.existing_sheet_name.get() in names:
            self.refresh_insert_ref_options(path, self.existing_sheet_name.get())

        if self.col_sheet_source.get() in names:
            self.refresh_col_ref_options()
        self._refresh_column_listbox()

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
            header_row = self._get_primary_header_row()
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
                rma_sheet_name=self._get_primary_sheet_name(),
                rma_header_row=header_row,
                column_spec_text=self.column_spec.get().strip(),
                target_mode=mode,
                existing_sheet_name=self.existing_sheet_name.get().strip(),
                log=self.log,
                insert_ref_column=self.insert_ref_column.get().strip(),
                insert_position=self.insert_position.get(),
                max_preview_rows=200,
                sheet_header_rows=self._get_sheet_header_rows(),
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
                rma_sheet_name=self._get_primary_sheet_name(),
                rma_header_row=header_row,
                column_spec_text=self.column_spec.get().strip(),
                target_mode=mode,
                new_sheet_name=self.new_sheet_name.get().strip(),
                existing_sheet_name=self.existing_sheet_name.get().strip(),
                copy_style=self.copy_style.get(),
                copy_formula=self.copy_formula.get(),
                log=self.log,
                insert_ref_column=self.insert_ref_column.get().strip(),
                insert_position=self.insert_position.get(),
                sheet_header_rows=self._get_sheet_header_rows(),
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
