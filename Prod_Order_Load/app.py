# -*- coding: utf-8 -*-
"""
공정발주 통합 유저폼(GUI).
시작 시 SQLite만 불러와 목록 표시(B). 통합은 필요할 때만 실행.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from pathlib import Path
from threading import Thread
import sys
import traceback

from config import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_OUTPUT_FILENAME,
    SOURCE_PATHS_FILE,
    DEFAULT_SOURCE_FOLDER_PATHS,
)
from folder_create import (
    load_folder_create_settings,
    save_folder_create_settings,
    create_folder_structure,
    create_folder_structure_grouped,
    copy_template_and_fill_cover,
    create_blank_workbook,
)
from consolidation import process_folders
from sqlite_query import (
    query_consolidated,
    get_last_exported_at,
    fetch_distinct_column,
)
from version_info import get_version

_FILTER_SPEC = [
    ("작업지시번호", "작업지시번호"),
    ("고객사", "고객사"),
    ("사업명", "사업명"),
    ("품명", "품명"),
    ("품번", "품번"),
    ("발주사양", "발주사양"),
    ("날짜 시작", "날짜"),
    ("날짜 종료", "날짜"),
]


class App:
    def __init__(self):
        self.root = tk.Tk()
        self._app_version = get_version()
        self.root.title(f"공정발주 데이터 통합 v{self._app_version}")
        self.root.minsize(1180, 760)
        self.root.geometry("1320x920")

        self.output_dir = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.output_filename = tk.StringVar(value=DEFAULT_OUTPUT_FILENAME)
        self.save_excel_var = tk.BooleanVar(value=False)
        self.running = False
        self._last_output_path: Path | None = None
        self.sqlite_path_var = tk.StringVar(
            value=str((DEFAULT_OUTPUT_DIR / DEFAULT_OUTPUT_FILENAME).with_suffix(".sqlite"))
        )

        self.filter_job = tk.StringVar(value="(전체)")
        self.filter_customer = tk.StringVar(value="(전체)")
        self.filter_business = tk.StringVar(value="(전체)")
        self.filter_product = tk.StringVar(value="(전체)")
        self.filter_code = tk.StringVar(value="(전체)")
        self.filter_spec = tk.StringVar(value="(전체)")
        self.filter_date_from = tk.StringVar(value="(전체)")
        self.filter_date_to = tk.StringVar(value="(전체)")

        self._filter_vars = [
            self.filter_job,
            self.filter_customer,
            self.filter_business,
            self.filter_product,
            self.filter_code,
            self.filter_spec,
            self.filter_date_from,
            self.filter_date_to,
        ]

        self._tv_cols: list[str] = []
        self._tv_by_iid: dict[str, tuple] = {}

        self.cb_data: list[ttk.Combobox] = []

        self.hint_var = tk.StringVar(value="")
        self.last_merge_var = tk.StringVar(value="마지막 통합: (없음)")

        _fc = load_folder_create_settings()
        self.folder_create_base_var = tk.StringVar(value=_fc.get("base_path") or str(DEFAULT_OUTPUT_DIR))
        self.folder_template_var = tk.StringVar(value=_fc.get("template_xlsx_path") or "")
        self.folder_cover_sheet_var = tk.StringVar(value=_fc.get("cover_sheet_name") or "표지")

        self._build_ui()
        self._startup_load_db()

    def _log(self, msg: str) -> None:
        """백그라운드 스레드에서도 안전하게 한 줄 안내 갱신."""

        def _set() -> None:
            self.hint_var.set(str(msg)[:300])

        try:
            self.root.after(0, _set)
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        frm = ttk.Frame(self.root, padding=8)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        nb = ttk.Notebook(frm)
        nb.grid(row=0, column=0, sticky=tk.NSEW)

        tab_data = ttk.Frame(nb, padding=8)
        nb.add(tab_data, text="데이터")

        self._build_tab_data(tab_data)

    def _add_filter_block(
        self, parent: ttk.Frame, start_row: int, into_list: list[ttk.Combobox]
    ) -> None:
        flt = ttk.LabelFrame(parent, text="목록 필터 (콤보 — DB 고유값, 직접 입력도 가능)")
        flt.grid(row=start_row, column=0, sticky=tk.EW, pady=(0, 6))
        for c in (1, 3, 5, 7):
            flt.columnconfigure(c, weight=1)

        labels = [x[0] for x in _FILTER_SPEC]
        # 2줄(세로)로 압축: 8개 필터를 4개씩 배치
        grid_pos = [(0, 0), (0, 2), (0, 4), (0, 6), (1, 0), (1, 2), (1, 4), (1, 6)]
        for label, var, (gr, gc) in zip(labels, self._filter_vars, grid_pos):
            ttk.Label(flt, text=label).grid(row=gr, column=gc, sticky=tk.W, padx=4, pady=2)
            cb = ttk.Combobox(
                flt,
                textvariable=var,
                width=24,
                values=["(전체)"],
                state="normal",
            )
            cb.grid(row=gr, column=gc + 1, sticky=tk.EW, padx=4, pady=2)
            into_list.append(cb)

    def _build_tab_data(self, tab_data: ttk.Frame) -> None:
        tab_data.rowconfigure(6, weight=3)
        tab_data.rowconfigure(8, weight=1)
        tab_data.columnconfigure(0, weight=1)

        # 작업(통합 실행) + 설정관리
        job = ttk.LabelFrame(tab_data, text="작업")
        job.grid(row=0, column=0, sticky=tk.EW, pady=(0, 8))
        for c in (1, 3, 5):
            job.columnconfigure(c, weight=1)

        ttk.Label(job, text="저장 경로:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(job, textvariable=self.output_dir).grid(row=0, column=1, sticky=tk.EW, padx=4, pady=2)
        ttk.Button(job, text="찾아보기...", command=self._browse_dir).grid(row=0, column=2, padx=4, pady=2)

        ttk.Label(job, text="파일명(.xlsx):").grid(row=0, column=3, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(job, textvariable=self.output_filename).grid(row=0, column=4, sticky=tk.EW, padx=4, pady=2)
        ttk.Checkbutton(job, text="엑셀도 저장", variable=self.save_excel_var).grid(row=0, column=5, sticky=tk.W)

        btns = ttk.Frame(job)
        btns.grid(row=1, column=0, columnspan=6, sticky=tk.W, padx=4, pady=(4, 2))
        self.btn_run = ttk.Button(btns, text="통합 실행 (SQLite 갱신)", command=self._run)
        self.btn_run.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="설정 관리...", command=self._open_settings_manager).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="SQLite 열기", command=self._open_sqlite_file).pack(side=tk.LEFT, padx=(0, 8))
        self.btn_open_result = ttk.Button(btns, text="결과 엑셀 열기", command=self._open_result_file, state=tk.DISABLED)
        self.btn_open_result.pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="상태: 대기 중")
        ttk.Label(job, textvariable=self.status_var, foreground="#333").grid(
            row=2, column=0, columnspan=6, sticky=tk.W, padx=4, pady=(2, 2)
        )
        ttk.Label(job, textvariable=self.last_merge_var, foreground="#333").grid(
            row=3, column=0, columnspan=6, sticky=tk.W, padx=4, pady=(0, 2)
        )
        self.progress = ttk.Progressbar(job, mode="indeterminate", length=520)
        self.progress.grid(row=4, column=0, columnspan=6, sticky=tk.EW, padx=4, pady=(2, 2))
        ttk.Label(job, textvariable=self.hint_var, wraplength=1180, foreground="#555").grid(
            row=5, column=0, columnspan=6, sticky=tk.EW, padx=4, pady=(2, 2)
        )

        # DB 경로
        db_row = ttk.Frame(tab_data)
        db_row.grid(row=1, column=0, sticky=tk.EW, pady=(0, 8))
        db_row.columnconfigure(1, weight=1)
        ttk.Label(db_row, text="SQLite 파일:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(db_row, textvariable=self.sqlite_path_var).grid(row=0, column=1, sticky=tk.EW)
        ttk.Button(db_row, text="찾아보기...", command=self._browse_sqlite).grid(row=0, column=2, padx=(8, 0))

        self._add_filter_block(tab_data, 2, self.cb_data)
        btn_row = ttk.Frame(tab_data)
        btn_row.grid(row=3, column=0, sticky=tk.W, pady=(0, 8))
        ttk.Button(btn_row, text="목록 불러오기", command=self._load_tree_from_db).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="필터 초기화", command=self._clear_filters_data_tab).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="콤보 목록 새로고침", command=self._reload_combos_from_current_db).pack(
            side=tk.LEFT
        )

        fc = ttk.LabelFrame(tab_data, text="폴더 생성 (선택 행 기준: 기본경로/고객사_사업명/폴더명/하위폴더…)")
        fc.grid(row=3, column=0, sticky=tk.EW, padx=(520, 0))
        fc.columnconfigure(1, weight=1)

        ttk.Label(fc, text="기본 경로:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        ttk.Entry(fc, textvariable=self.folder_create_base_var, width=48).grid(
            row=0, column=1, sticky=tk.EW, padx=4, pady=2
        )
        ttk.Button(fc, text="찾아보기...", command=self._browse_folder_base).grid(row=0, column=2, padx=4, pady=2)

        ttk.Button(fc, text="기본경로·하위폴더 목록 편집...", command=self._edit_folder_create_settings).grid(
            row=1, column=0, columnspan=3, sticky=tk.W, padx=4, pady=(2, 4)
        )

        act = ttk.Frame(fc)
        act.grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=4, pady=(2, 4))
        ttk.Button(act, text="선택 행만", command=self._run_folder_create_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(act, text="현재 목록 전체", command=self._run_folder_create_all_filtered).pack(side=tk.LEFT)

        tree_frame = ttk.Frame(tab_data)
        tree_frame.grid(row=6, column=0, sticky=tk.NSEW, pady=(0, 4))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            tree_frame,
            show="headings",
            selectmode="extended",
        )
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        ttk.Label(tab_data, text="선택 행 상세 (위 표는 DB의 모든 열을 가로 스크롤로 확인):").grid(
            row=7, column=0, sticky=tk.W, pady=(8, 4)
        )
        self.detail_text = scrolledtext.ScrolledText(tab_data, height=7, width=120, wrap=tk.WORD)
        self.detail_text.grid(row=8, column=0, sticky=tk.NSEW, pady=(0, 4))

    def _open_settings_manager(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("설정 관리")
        win.geometry("860x640")
        win.transient(self.root)

        nb = ttk.Notebook(win)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tab_src = ttk.Frame(nb, padding=10)
        tab_folder = ttk.Frame(nb, padding=10)
        nb.add(tab_src, text="소스 경로")
        nb.add(tab_folder, text="폴더/템플릿")

        # --- 소스 경로 ---
        ttk.Label(tab_src, text="소스 폴더 경로 (한 줄에 하나씩):").pack(anchor=tk.W)
        txt_src = scrolledtext.ScrolledText(tab_src, height=18, width=110, wrap=tk.NONE)
        txt_src.pack(fill=tk.BOTH, expand=True, pady=(6, 8))

        paths: list[str] = []
        try:
            if SOURCE_PATHS_FILE.exists():
                import json

                with SOURCE_PATHS_FILE.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                    paths = data.get("folders") or data.get("paths") or []
            else:
                paths = list(DEFAULT_SOURCE_FOLDER_PATHS)
        except Exception as e:
            messagebox.showwarning("로드 경고", str(e))
            paths = list(DEFAULT_SOURCE_FOLDER_PATHS)

        if paths:
            txt_src.insert(tk.END, "\n".join(str(p) for p in paths))

        def save_src() -> None:
            import json

            raw = txt_src.get("1.0", tk.END)
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            data = {"folders": lines}
            try:
                SOURCE_PATHS_FILE.parent.mkdir(parents=True, exist_ok=True)
                with SOURCE_PATHS_FILE.open("w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.hint_var.set("소스 경로 설정을 저장했습니다.")
            except Exception as e:
                messagebox.showerror("저장 오류", str(e))

        btn_src = ttk.Frame(tab_src)
        btn_src.pack(fill=tk.X)
        ttk.Button(btn_src, text="저장", command=save_src).pack(side=tk.RIGHT, padx=(6, 0))

        # --- 폴더/템플릿 ---
        s = load_folder_create_settings()
        frm = ttk.Frame(tab_folder)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="기본 경로:").grid(row=0, column=0, sticky=tk.W, pady=4)
        base_var = tk.StringVar(value=s.get("base_path") or "")
        ttk.Entry(frm, textvariable=base_var).grid(row=0, column=1, sticky=tk.EW, pady=4, padx=(8, 0))
        ttk.Button(
            frm,
            text="찾아보기...",
            command=lambda: self._browse_dir_into_var(base_var, title="폴더 생성 기본 경로"),
        ).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(frm, text="템플릿 엑셀(.xlsx):").grid(row=1, column=0, sticky=tk.W, pady=4)
        tmpl_var = tk.StringVar(value=s.get("template_xlsx_path") or "")
        ttk.Entry(frm, textvariable=tmpl_var).grid(row=1, column=1, sticky=tk.EW, pady=4, padx=(8, 0))
        ttk.Button(frm, text="찾아보기...", command=lambda: self._browse_template_into_var(tmpl_var)).grid(
            row=1, column=2, padx=(8, 0), pady=4
        )

        ttk.Label(frm, text="표지 시트명:").grid(row=2, column=0, sticky=tk.W, pady=4)
        cover_var = tk.StringVar(value=s.get("cover_sheet_name") or "표지")
        ttk.Entry(frm, textvariable=cover_var, width=24).grid(row=2, column=1, sticky=tk.W, pady=4, padx=(8, 0))

        ttk.Label(frm, text="하위 폴더 이름(한 줄에 하나):").grid(row=3, column=0, sticky=tk.NW, pady=(8, 4))
        txt_sub = scrolledtext.ScrolledText(frm, height=14, width=60, wrap=tk.NONE)
        txt_sub.grid(row=3, column=1, columnspan=2, sticky=tk.NSEW, pady=(8, 4), padx=(8, 0))
        frm.rowconfigure(3, weight=1)
        subs = s.get("subfolders") or []
        if subs:
            txt_sub.insert(tk.END, "\n".join(str(x) for x in subs))

        def save_folder() -> None:
            raw = txt_sub.get("1.0", tk.END)
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            bp = (base_var.get() or "").strip()
            try:
                save_folder_create_settings(
                    bp,
                    lines,
                    template_xlsx_path=(tmpl_var.get() or "").strip(),
                    cover_sheet_name=(cover_var.get() or "").strip() or "표지",
                )
                self.folder_create_base_var.set(bp or str(DEFAULT_OUTPUT_DIR))
                self.folder_template_var.set((tmpl_var.get() or "").strip())
                self.folder_cover_sheet_var.set((cover_var.get() or "").strip() or "표지")
                self.hint_var.set("폴더/템플릿 설정을 저장했습니다.")
            except OSError as e:
                messagebox.showerror("저장 오류", str(e))

        btn_folder = ttk.Frame(tab_folder)
        btn_folder.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_folder, text="저장", command=save_folder).pack(side=tk.RIGHT, padx=(6, 0))

        # 하단 닫기
        bottom = ttk.Frame(win)
        bottom.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(bottom, text="닫기", command=win.destroy).pack(side=tk.RIGHT)

    def _browse_dir_into_var(self, var: tk.StringVar, *, title: str) -> None:
        path = filedialog.askdirectory(title=title, initialdir=var.get() or str(Path.cwd()))
        if path:
            var.set(path)

    def _make_group_name(self, customer: str, business: str) -> str:
        c = (customer or "").strip()
        b = (business or "").strip()
        if c and b:
            return f"{c}_{b}"
        return c or b or "미지정"

    def _browse_folder_base(self) -> None:
        path = filedialog.askdirectory(
            title="폴더 생성 기본 경로",
            initialdir=self.folder_create_base_var.get() or str(Path.cwd()),
        )
        if path:
            self.folder_create_base_var.set(path)

    def _browse_template_xlsx(self) -> None:
        path = filedialog.askopenfilename(
            title="복사할 템플릿 엑셀 선택",
            filetypes=[("Excel", "*.xlsx"), ("모든 파일", "*.*")],
            initialdir=str(Path(self.folder_template_var.get() or ".").parent),
        )
        if path:
            self.folder_template_var.set(path)

    def _edit_folder_create_settings(self) -> None:
        win = tk.Toplevel(self.root)
        win.title("폴더 생성 설정")
        win.geometry("760x520")
        win.transient(self.root)

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.columnconfigure(1, weight=1)

        s = load_folder_create_settings()
        ttk.Label(frm, text="기본 경로 (한 줄):").grid(row=0, column=0, sticky=tk.NW, pady=(0, 8))
        base_var = tk.StringVar(value=s.get("base_path") or "")
        ttk.Entry(frm, textvariable=base_var, width=70).grid(row=0, column=1, sticky=tk.EW, pady=(0, 8))

        ttk.Label(frm, text="복사할 템플릿 엑셀(.xlsx) 경로:").grid(row=1, column=0, sticky=tk.NW, pady=(0, 8))
        tmpl_var = tk.StringVar(value=s.get("template_xlsx_path") or "")
        tmpl_row = ttk.Frame(frm)
        tmpl_row.grid(row=1, column=1, sticky=tk.EW, pady=(0, 8))
        tmpl_row.columnconfigure(0, weight=1)
        ttk.Entry(tmpl_row, textvariable=tmpl_var).grid(row=0, column=0, sticky=tk.EW, padx=(0, 8))
        ttk.Button(tmpl_row, text="찾아보기...", command=lambda: self._browse_template_into_var(tmpl_var)).grid(
            row=0, column=1
        )

        ttk.Label(frm, text="표지 시트명(없으면 첫 시트 사용):").grid(row=2, column=0, sticky=tk.NW, pady=(0, 8))
        cover_var = tk.StringVar(value=s.get("cover_sheet_name") or "표지")
        ttk.Entry(frm, textvariable=cover_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=(0, 8))

        ttk.Label(frm, text="하위 폴더 이름 (한 줄에 하나, 비우면 상위 폴더만 생성):").grid(
            row=3, column=0, sticky=tk.NW, pady=(4, 4)
        )
        txt = scrolledtext.ScrolledText(frm, height=12, width=60, wrap=tk.NONE)
        txt.grid(row=3, column=1, sticky=tk.NSEW, pady=(4, 8))
        frm.rowconfigure(3, weight=1)
        subs = s.get("subfolders") or []
        if subs:
            txt.insert(tk.END, "\n".join(str(x) for x in subs))

        def save_fc() -> None:
            raw = txt.get("1.0", tk.END)
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            bp = (base_var.get() or "").strip()
            try:
                save_folder_create_settings(
                    bp,
                    lines,
                    template_xlsx_path=(tmpl_var.get() or "").strip(),
                    cover_sheet_name=(cover_var.get() or "").strip() or "표지",
                )
                self.folder_create_base_var.set(bp or str(DEFAULT_OUTPUT_DIR))
                self.folder_template_var.set((tmpl_var.get() or "").strip())
                self.folder_cover_sheet_var.set((cover_var.get() or "").strip() or "표지")
                win.destroy()
                self.hint_var.set("폴더 생성 설정을 저장했습니다.")
            except OSError as e:
                messagebox.showerror("저장 오류", str(e))

        br = ttk.Frame(frm)
        br.grid(row=4, column=0, columnspan=2, sticky=tk.E, pady=(8, 0))
        ttk.Button(br, text="저장", command=save_fc).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(br, text="취소", command=win.destroy).pack(side=tk.RIGHT)

    def _browse_template_into_var(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="복사할 템플릿 엑셀 선택",
            filetypes=[("Excel", "*.xlsx"), ("모든 파일", "*.*")],
            initialdir=str(Path(var.get() or ".").parent),
        )
        if path:
            var.set(path)

    def _collect_folder_items_from_rows(self, cols: list[str], rows: list[tuple]) -> list[dict]:
        # 폴더 생성 + 파일 복사/표지입력에 필요한 컬럼
        required_cols = (
            "폴더명",
            "고객사",
            "사업명",
            "BOM파일명",
            "작업지시번호",
            "품명",
            "품번",
            "공정",
            "고객사납품",
            "자재입고수량",
            "발주사양",
        )
        for required in required_cols:
            if required not in cols:
                messagebox.showwarning(
                    "열 없음",
                    f"DB에 「{required}」 열이 없습니다.\n통합 후 데이터 탭에서 목록을 다시 불러오세요.",
                )
                return []

        i_folder = cols.index("폴더명")
        i_cust = cols.index("고객사")
        i_biz = cols.index("사업명")
        i_bom = cols.index("BOM파일명")
        i_job = cols.index("작업지시번호")
        i_prod = cols.index("품명")
        i_code = cols.index("품번")
        i_proc = cols.index("공정")
        i_ship = cols.index("고객사납품")
        i_in = cols.index("자재입고수량")
        i_spec = cols.index("발주사양")

        out: list[dict] = []
        for row in rows:
            folder = row[i_folder] if i_folder < len(row) else None
            cust = row[i_cust] if i_cust < len(row) else None
            biz = row[i_biz] if i_biz < len(row) else None
            if folder is None or str(folder).strip() == "":
                continue
            group = self._make_group_name("" if cust is None else str(cust), "" if biz is None else str(biz))
            bom_name = row[i_bom] if i_bom < len(row) else None
            p_values = [
                "" if row[i_job] is None else str(row[i_job]),
                "" if cust is None else str(cust),
                "" if biz is None else str(biz),
                "" if row[i_prod] is None else str(row[i_prod]),
                "" if row[i_code] is None else str(row[i_code]),
                "" if row[i_proc] is None else str(row[i_proc]),
                "" if row[i_ship] is None else str(row[i_ship]),
                "" if row[i_in] is None else str(row[i_in]),
                "" if row[i_spec] is None else str(row[i_spec]),
            ]
            out.append(
                {
                    "group": group,
                    "folder": str(folder).strip(),
                    "bom_filename": "" if bom_name is None else str(bom_name).strip(),
                    "p_values": [v.strip() for v in p_values],
                }
            )
        return out

    def _run_folder_create_selected(self) -> None:
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showwarning(
                "선택 없음",
                "「데이터」 탭 표에서 행을 하나 이상 선택한 뒤 다시 시도하세요.\n(Ctrl·Shift로 다중 선택)",
            )
            return
        rows = [self._tv_by_iid[iid] for iid in sel if iid in self._tv_by_iid]
        items = self._collect_folder_items_from_rows(self._tv_cols, rows)
        if not items:
            return
        self._folder_create_thread(items)

    def _run_folder_create_all_filtered(self) -> None:
        iids = self.tree.get_children()
        if not iids:
            messagebox.showwarning(
                "목록 없음",
                "「데이터」 탭에서 「목록 불러오기」로 표를 먼저 채워 주세요.",
            )
            return
        rows = [self._tv_by_iid[iid] for iid in iids if iid in self._tv_by_iid]
        items = self._collect_folder_items_from_rows(self._tv_cols, rows)
        if not items:
            return
        self._folder_create_thread(items)

    def _folder_create_thread(self, items: list[dict]) -> None:
        base_str = (self.folder_create_base_var.get() or "").strip()
        if not base_str:
            messagebox.showwarning("경로 없음", "기본 경로를 입력하거나 찾아보기로 선택하세요.")
            return

        st = load_folder_create_settings()
        subs = list(st.get("subfolders") or [])
        template_path = Path((st.get("template_xlsx_path") or "").strip())
        cover_sheet = (st.get("cover_sheet_name") or "표지").strip() or "표지"

        def work() -> None:
            err_tail = ""
            try:
                save_folder_create_settings(
                    base_str,
                    subs,
                    template_xlsx_path=str(template_path) if str(template_path).strip() else "",
                    cover_sheet_name=cover_sheet,
                )
            except OSError as e:
                err_tail = str(e)
            ok_p, sk, errs = 0, 0, []
            if not err_tail:
                # 1) 폴더 구조: base/group/폴더명/(sub...)
                pairs = [(it["group"], it["folder"]) for it in items]
                ok_p, sk, errs = create_folder_structure_grouped(
                    Path(base_str),
                    pairs,
                    subs,
                    log=self._log,
                )
                # 2) 파일 생성/복사: base/group 에 생성
                base = Path(base_str)
                for it in items:
                    group_dir = base / it["group"]
                    group_dir.mkdir(parents=True, exist_ok=True)

                    # A) 새 워크북 생성: {폴더명}-표.xlsx
                    blank_name = f'{it["folder"]}-좌표.xlsx'
                    ok1, e1 = create_blank_workbook(dest_xlsx=group_dir / blank_name, log=self._log)
                    if (not ok1) and e1:
                        errs.append(f'{it["group"]}/{blank_name}: {e1}')

                    # B) 템플릿 복사 + 표지 입력: {BOM파일명}.xlsx
                    bom_base = it.get("bom_filename") or ""
                    if bom_base:
                        dest = group_dir / f"{bom_base}.xlsx"
                        ok2, e2 = copy_template_and_fill_cover(
                            template_xlsx=template_path,
                            dest_xlsx=dest,
                            cover_sheet_name=cover_sheet,
                            p1_to_p9_values=list(it.get("p_values") or []),
                            log=self._log,
                        )
                        if (not ok2) and e2:
                            errs.append(f'{it["group"]}/{dest.name}: {e2}')
                    else:
                        errs.append(f'{it["group"]}/{it["folder"]}: BOM파일명이 비어 있어 템플릿 복사를 건너뜀')

            def done() -> None:
                if err_tail:
                    messagebox.showerror("설정 저장 실패", err_tail)
                    return
                msg = f"처리: {ok_p}개\n건너뜀: {sk}개\n오류: {len(errs)}건"
                if errs:
                    msg += "\n\n" + "\n".join(errs[:15])
                    if len(errs) > 15:
                        msg += "\n…"
                messagebox.showinfo("폴더 생성", msg)
                self.hint_var.set("폴더 생성 작업을 마쳤습니다.")

            self.root.after(0, done)

        self.hint_var.set("폴더 생성 중…")
        Thread(target=work, daemon=True).start()

    def _collect_folder_names_from_rows(self, cols: list[str], rows: list[tuple]) -> list[str]:
        # (레거시) 예전 방식 호환을 위해 남겨둠
        if "폴더명" not in cols:
            messagebox.showwarning(
                "열 없음",
                "DB에 「폴더명」 열이 없습니다.\n통합 후 데이터 탭에서 목록을 다시 불러오세요.",
            )
            return []
        icol = cols.index("폴더명")
        names: list[str] = []
        for row in rows:
            v = row[icol] if icol < len(row) else None
            if v is None or str(v).strip() == "":
                continue
            names.append(str(v).strip())
        return names

    def _startup_load_db(self) -> None:
        self._sync_sqlite_path_from_output()
        p = Path(self.sqlite_path_var.get().strip())
        self._update_last_merge_label(p)
        if p.is_file():
            self._reload_filter_options(p)
            self._load_tree_from_db()
            self.hint_var.set("저장된 DB를 불러왔습니다. 필요할 때만 「통합 실행」하세요.")
        else:
            self.hint_var.set("DB 파일이 없습니다. 「통합 실행」으로 처음 생성할 수 있습니다.")

    def _update_last_merge_label(self, db_path: Path) -> None:
        ts = get_last_exported_at(db_path)
        if ts:
            self.last_merge_var.set(f"마지막 통합: {ts}")
        else:
            self.last_merge_var.set("마지막 통합: (DB 없음 또는 데이터 없음)")

    def _reload_filter_options(self, db_path: Path) -> None:
        if not db_path.is_file():
            return

        def _apply_data(idx: int, vals: list[str]) -> None:
            cb = self.cb_data[idx]
            cur = (cb.get() or "").strip()
            cb["values"] = vals
            if cur in vals:
                cb.set(cur)
            else:
                cb.set("(전체)")

        for i in range(6):
            col = _FILTER_SPEC[i][1]
            _apply_data(i, ["(전체)"] + fetch_distinct_column(db_path, col, limit=400))
        date_vals = ["(전체)"] + fetch_distinct_column(db_path, "날짜", limit=500)
        _apply_data(6, date_vals)
        _apply_data(7, date_vals)

    def _reload_combos_from_current_db(self) -> None:
        p = Path(self.sqlite_path_var.get().strip())
        if not p.is_file():
            messagebox.showwarning("파일 없음", "SQLite 경로를 확인하세요.")
            return
        self._reload_filter_options(p)
        self.hint_var.set("콤보 목록을 DB 기준으로 갱신했습니다.")

    def _sync_sqlite_path_from_output(self) -> None:
        out_dir = (self.output_dir.get() or "").strip()
        out_name = (self.output_filename.get() or "").strip() or DEFAULT_OUTPUT_FILENAME
        stem = Path(out_name).stem or "공정발주내역"
        self.sqlite_path_var.set(str(Path(out_dir) / f"{stem}.sqlite"))

    def _browse_dir(self) -> None:
        path = filedialog.askdirectory(
            title="저장할 폴더 선택",
            initialdir=self.output_dir.get() or str(Path.cwd()),
        )
        if path:
            self.output_dir.set(path)
            self._sync_sqlite_path_from_output()

    def _browse_sqlite(self) -> None:
        path = filedialog.askopenfilename(
            title="SQLite 파일 선택",
            filetypes=[("SQLite", "*.sqlite"), ("모든 파일", "*.*")],
            initialdir=str(Path(self.sqlite_path_var.get() or ".").parent),
        )
        if path:
            self.sqlite_path_var.set(path)
            p = Path(path)
            self._update_last_merge_label(p)
            self._reload_filter_options(p)
            self._load_tree_from_db()
            self.hint_var.set(f"DB 경로 변경: {p.name}")

    def _reset_defaults(self) -> None:
        self.output_dir.set(str(DEFAULT_OUTPUT_DIR))
        self.output_filename.set(DEFAULT_OUTPUT_FILENAME)
        self._sync_sqlite_path_from_output()
        self.hint_var.set("기본 경로/파일명으로 복원했습니다.")

    def _open_result_file(self) -> None:
        p = self._last_output_path
        if not p or not p.is_file():
            messagebox.showwarning(
                "파일 없음",
                "엑셀을 저장한 실행만 열 수 있습니다.\n「엑셀 파일도 함께 저장」을 켜고 통합해 주세요.",
            )
            return
        try:
            os.startfile(str(p.resolve()))
        except OSError as e:
            messagebox.showerror("열기 실패", str(e))

    def _open_sqlite_file(self) -> None:
        p = Path(self.sqlite_path_var.get().strip())
        if not p.is_file():
            messagebox.showwarning("파일 없음", f"SQLite 파일이 없습니다.\n{p}")
            return
        try:
            os.startfile(str(p.resolve()))
        except OSError as e:
            messagebox.showerror("열기 실패", str(e))

    def _clear_filters(self) -> None:
        for v in self._filter_vars:
            v.set("(전체)")

    def _clear_filters_data_tab(self) -> None:
        self._clear_filters()
        self._load_tree_from_db()

    def _load_tree_from_db(self) -> None:
        db_path = Path(self.sqlite_path_var.get().strip())
        if not db_path.is_file():
            messagebox.showwarning("파일 없음", f"SQLite 파일이 없습니다.\n{db_path}")
            return

        cols, rows = query_consolidated(
            db_path,
            job_contains=self.filter_job.get(),
            customer_contains=self.filter_customer.get(),
            business_contains=self.filter_business.get(),
            product_contains=self.filter_product.get(),
            code_contains=self.filter_code.get(),
            spec_contains=self.filter_spec.get(),
            date_from=self.filter_date_from.get(),
            date_to=self.filter_date_to.get(),
        )
        for c in self.tree.get_children():
            self.tree.delete(c)
        self.detail_text.delete("1.0", tk.END)

        if not cols:
            self.hint_var.set("테이블이 없거나 비어 있습니다.")
            return

        self._tv_cols = list(cols)
        self._tv_by_iid.clear()
        shown = list(cols)
        idxs = list(range(len(cols)))
        self.tree["columns"] = shown
        for cname in shown:
            self.tree.heading(cname, text=cname)
            w = min(220, max(72, min(len(cname) * 9 + 24, 180)))
            self.tree.column(cname, width=w, minwidth=56, stretch=tk.YES)

        for row in rows:
            rid = row[0]
            if rid is None:
                continue
            iid = str(rid)
            self._tv_by_iid[iid] = row
            vals = tuple("" if row[j] is None else str(row[j]) for j in idxs)
            self.tree.insert("", tk.END, iid=iid, values=vals)

        self.hint_var.set(f"[데이터 탭] {len(rows)}행, 열 {len(shown)}개 표시 (가로 스크롤)")

    def _on_tree_select(self, _evt=None) -> None:
        sel = self.tree.selection()
        self.detail_text.delete("1.0", tk.END)
        if not sel or not self._tv_cols:
            return
        iid = sel[0]
        row = self._tv_by_iid.get(iid)
        if not row:
            return
        lines = [f"{name}: {row[i] if i < len(row) else ''}" for i, name in enumerate(self._tv_cols)]
        self.detail_text.insert(tk.END, "\n".join(lines))

    def _edit_source_paths(self) -> None:
        # 레거시: 기존 버튼을 없애면서도 외부 호출이 있으면 설정관리로 연결
        self._open_settings_manager()

    def _run(self) -> None:
        if self.running:
            return
        out_dir = (self.output_dir.get() or "").strip()
        out_name = (self.output_filename.get() or "").strip()
        if not out_name:
            messagebox.showwarning("입력 오류", "파일명을 입력해 주세요.")
            return
        if not out_name.lower().endswith(".xlsx"):
            out_name += ".xlsx"
        self.output_filename.set(out_name)
        output_path = Path(out_dir) / out_name
        self._sync_sqlite_path_from_output()
        save_xlsx = bool(self.save_excel_var.get())

        self.running = True
        self.btn_run.configure(state=tk.DISABLED)
        self.btn_open_result.configure(state=tk.DISABLED)
        self._last_output_path = None
        self.status_var.set("상태: 통합 실행 중...")
        self.hint_var.set("통합 중…")
        self.root.title(f"공정발주 데이터 통합 v{self._app_version} (실행 중)")
        self.progress.start(50)

        def work() -> None:
            ok = False
            err_tb = ""
            try:
                ok = process_folders(output_path, self._log, save_excel=save_xlsx)
            except Exception as e:
                err_tb = f"{e}\n{traceback.format_exc()}"
            finally:
                sqlite_path = output_path.with_suffix(".sqlite")
                excel_ok = save_xlsx and ok and output_path.is_file()
                sqlite_ok = ok and sqlite_path.is_file()

                def _done() -> None:
                    self.progress.stop()
                    self.running = False
                    self.status_var.set("상태: 완료" if ok else "상태: 실패")
                    self.btn_run.configure(state=tk.NORMAL)
                    if excel_ok:
                        self._last_output_path = output_path
                        self.btn_open_result.configure(state=tk.NORMAL)
                    else:
                        self._last_output_path = None
                        self.btn_open_result.configure(state=tk.DISABLED)
                    if err_tb:
                        self.hint_var.set("통합 중 오류가 발생했습니다.")
                        messagebox.showerror("통합 오류", err_tb[:4000])
                    elif ok:
                        self.hint_var.set("통합이 완료되었습니다.")
                    else:
                        self.hint_var.set("통합이 실패했거나 저장에 문제가 있을 수 있습니다.")
                    if sqlite_ok:
                        self.sqlite_path_var.set(str(sqlite_path))
                        self._update_last_merge_label(sqlite_path)
                        self._reload_filter_options(sqlite_path)
                        self._load_tree_from_db()
                    self.root.title(f"공정발주 데이터 통합 v{self._app_version}")

                self.root.after(0, _done)

        Thread(target=work, daemon=True).start()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = App()
    app.run()


if __name__ == "__main__":
    main()
    sys.exit(0)
