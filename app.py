# -*- coding: utf-8 -*-
"""
공정발주 통합 유저폼(GUI).
기본 저장 경로·파일명을 사용하며, 경로 선택/파일명 입력으로 변경할 수 있다.
"""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from pathlib import Path
from threading import Thread
import queue
import sys

from config import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_OUTPUT_FILENAME,
    SOURCE_PATHS_FILE,
    DEFAULT_SOURCE_FOLDER_PATHS,
)
from consolidation import process_folders


def _default_output_path() -> Path:
    return DEFAULT_OUTPUT_DIR / DEFAULT_OUTPUT_FILENAME


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("공정발주 데이터 통합")
        self.root.minsize(480, 420)
        self.root.geometry("560x500")

        self.output_dir = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.output_filename = tk.StringVar(value=DEFAULT_OUTPUT_FILENAME)
        self.log_queue: queue.Queue = queue.Queue()
        self.running = False

        self._build_ui()
        self._poll_log()

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # 저장 경로
        ttk.Label(frm, text="저장 경로:").grid(row=0, column=0, sticky=tk.W, pady=(0, 4))
        path_row = ttk.Frame(frm)
        path_row.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        self.root.columnconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        e_path = ttk.Entry(path_row, textvariable=self.output_dir, width=50)
        e_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(path_row, text="찾아보기...", command=self._browse_dir).pack(side=tk.RIGHT)

        # 파일명
        ttk.Label(frm, text="파일명:").grid(row=2, column=0, sticky=tk.W, pady=(0, 4))
        ttk.Entry(frm, textvariable=self.output_filename, width=40).grid(
            row=3, column=0, sticky=tk.EW, pady=(0, 12)
        )

        # 기본값 복원
        ttk.Button(
            frm,
            text="기본 경로/파일명으로 복원",
            command=self._reset_defaults,
        ).grid(row=4, column=0, sticky=tk.W, pady=(0, 16))

        # 실행
        run_row = ttk.Frame(frm)
        run_row.grid(row=5, column=0, sticky=tk.W, pady=(0, 8))
        btn_run = ttk.Button(run_row, text="통합 실행", command=self._run)
        btn_run.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_run = btn_run
        self.status_var = tk.StringVar(value="상태: 대기 중")
        ttk.Label(run_row, textvariable=self.status_var).pack(side=tk.LEFT)

        # 소스 경로 관리 버튼
        ttk.Button(
            frm,
            text="소스 경로 관리...",
            command=self._edit_source_paths,
        ).grid(row=5, column=1, sticky=tk.W, padx=(12, 0))

        # 진행 표시 (실행 중일 때만 애니메이션)
        self.progress = ttk.Progressbar(frm, mode="indeterminate", length=300)
        self.progress.grid(row=6, column=0, sticky=tk.EW, pady=(0, 8))

        # 로그
        ttk.Label(frm, text="로그:").grid(row=7, column=0, sticky=tk.W, pady=(0, 4))
        self.log_text = scrolledtext.ScrolledText(
            frm, height=14, width=70, state=tk.DISABLED, wrap=tk.WORD
        )
        self.log_text.grid(row=8, column=0, sticky=tk.NSEW, pady=(0, 8))
        frm.rowconfigure(8, weight=1)
        frm.columnconfigure(0, weight=1)

    def _log(self, msg: str):
        self.log_queue.put(msg)

    def _poll_log(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        if self.running:
            self.root.after(200, self._poll_log)
        else:
            self.root.after(100, self._poll_log)

    def _browse_dir(self):
        path = filedialog.askdirectory(
            title="저장할 폴더 선택",
            initialdir=self.output_dir.get() or str(Path.cwd()),
        )
        if path:
            self.output_dir.set(path)

    def _reset_defaults(self):
        self.output_dir.set(str(DEFAULT_OUTPUT_DIR))
        self.output_filename.set(DEFAULT_OUTPUT_FILENAME)
        self._log("기본 경로/파일명으로 복원했습니다.")

    def _edit_source_paths(self):
        """소스 데이터 경로 목록을 추가/수정/삭제하는 간단한 다이얼로그."""
        win = tk.Toplevel(self.root)
        win.title("소스 데이터 경로 관리")
        win.geometry("640x360")
        win.transient(self.root)

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="소스 폴더 경로 (한 줄에 하나씩):").pack(anchor=tk.W)
        txt = scrolledtext.ScrolledText(frm, height=12, width=80, wrap=tk.NONE)
        txt.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        # 현재 설정 로드
        paths = []
        try:
            if SOURCE_PATHS_FILE.exists():
                import json

                with SOURCE_PATHS_FILE.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                    paths = data.get("folders") or data.get("paths") or []
            else:
                paths = DEFAULT_SOURCE_FOLDER_PATHS
        except Exception as e:
            self._log(f"[경고] 소스 경로 설정 로드 실패: {e}")
            paths = DEFAULT_SOURCE_FOLDER_PATHS

        if paths:
            txt.insert(tk.END, "\n".join(str(p) for p in paths))

        def save_paths():
            import json

            raw = txt.get("1.0", tk.END)
            lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            data = {"folders": lines}
            try:
                with SOURCE_PATHS_FILE.open("w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self._log("소스 경로 설정을 저장했습니다.")
                win.destroy()
            except Exception as e:
                messagebox.showerror("저장 오류", f"소스 경로 설정 저장 중 오류가 발생했습니다.\n{e}")

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_row, text="저장", command=save_paths).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_row, text="취소", command=win.destroy).pack(side=tk.RIGHT)

    def _run(self):
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
        self.running = True
        self.btn_run.configure(state=tk.DISABLED)
        self.status_var.set("상태: 통합 실행 중...")
        self.root.title("공정발주 데이터 통합 (실행 중)")
        self.progress.start(50)
        self._log("통합을 시작합니다...")
        self._log(f"출력: {output_path}")

        def work():
            try:
                ok = process_folders(output_path, self._log)
                self.log_queue.put("[완료]" if ok else "[실패] 저장 단계에서 오류가 났을 수 있습니다.")
            except Exception as e:
                self.log_queue.put(f"[오류] {e}")
                import traceback
                self.log_queue.put(traceback.format_exc())
            finally:
                self.log_queue.put("---")
                self.running = False

                def _done():
                    self.progress.stop()
                    self.status_var.set("상태: 완료")
                    self.btn_run.configure(state=tk.NORMAL)
                    self.root.title("공정발주 데이터 통합")

                self.root.after(0, _done)

        Thread(target=work, daemon=True).start()

    def run(self):
        self.root.mainloop()


def main():
    app = App()
    app.run()


if __name__ == "__main__":
    main()
    sys.exit(0)
