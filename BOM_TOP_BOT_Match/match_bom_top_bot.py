from __future__ import annotations

import argparse
import json
import math
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Tuple

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.utils.cell import range_boundaries


Key = Tuple[str, str]  # (coord, material)


@dataclass
class SheetData:
    label: str
    sheet_name: str
    key_to_qty: Dict[Key, float] = field(default_factory=dict)
    key_to_rows: Dict[Key, List[int]] = field(default_factory=dict)
    coord_to_count: Dict[str, int] = field(default_factory=dict)
    coord_to_rows: Dict[str, List[int]] = field(default_factory=dict)
    coord_to_material_counts: Dict[str, Dict[str, int]] = field(default_factory=dict)


def norm_text(v: object, *, case_insensitive: bool) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if case_insensitive:
        s = s.lower()
    return s


def norm_qty(v: object) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
        return float(v)
    s = str(v).strip()
    if not s:
        return 0.0
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def split_coords(coord_text: object) -> List[str]:
    if coord_text is None:
        return []
    s = str(coord_text).strip()
    if not s:
        return []
    parts = [p.strip() for p in s.split(",")]
    return [p for p in parts if p]


def parse_single_col_range(addr: str) -> Tuple[int, int, int, int]:
    """
    Returns (min_col, min_row, max_col, max_row). Validates single-column range.
    """
    min_col, min_row, max_col, max_row = range_boundaries(addr)
    if min_col != max_col:
        raise ValueError(f"Range must be single-column: {addr}")
    if max_row < min_row:
        raise ValueError(f"Invalid range rows: {addr}")
    return min_col, min_row, max_col, max_row


def read_col_range_values(ws, addr: str) -> Tuple[List[object], List[int]]:
    min_col, min_row, _max_col, max_row = parse_single_col_range(addr)
    values: List[object] = []
    rows: List[int] = []
    for r in range(min_row, max_row + 1):
        values.append(ws.cell(row=r, column=min_col).value)
        rows.append(r)
    return values, rows


def bump(d: Dict[str, int], k: str, inc: int = 1) -> None:
    d[k] = d.get(k, 0) + inc


def add_row_list(d: Dict, k, row_num: int) -> None:
    if k not in d:
        d[k] = [row_num]
    else:
        d[k].append(row_num)


def build_sheet_data(
    wb,
    *,
    label: str,
    sheet_name: str,
    material_range: str,
    coord_range: str,
    qty_range: str,
    case_insensitive: bool,
) -> SheetData:
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Sheet not found: {sheet_name}")
    ws = wb[sheet_name]

    mats, mat_rows = read_col_range_values(ws, material_range)
    coords, _coord_rows = read_col_range_values(ws, coord_range)
    qtys, _qty_rows = read_col_range_values(ws, qty_range)

    if not (len(mats) == len(coords) == len(qtys)):
        raise ValueError(
            f"{label}: ranges must have same number of rows "
            f"(mat={len(mats)}, coord={len(coords)}, qty={len(qtys)})"
        )

    sd = SheetData(label=label, sheet_name=sheet_name)

    for i in range(len(mats)):
        mat = norm_text(mats[i], case_insensitive=case_insensitive)
        coord_text = coords[i]
        qty = norm_qty(qtys[i])
        row_num = mat_rows[i]

        if not mat and (coord_text is None or str(coord_text).strip() == ""):
            continue

        for coord_raw in split_coords(coord_text):
            coord = norm_text(coord_raw, case_insensitive=case_insensitive)
            if coord:
                bump(sd.coord_to_count, coord, 1)
                add_row_list(sd.coord_to_rows, coord, row_num)
                if coord not in sd.coord_to_material_counts:
                    sd.coord_to_material_counts[coord] = {}
                mcounts = sd.coord_to_material_counts[coord]
                mcounts[mat] = mcounts.get(mat, 0) + 1

            if coord and mat:
                key: Key = (coord, mat)
                sd.key_to_qty[key] = sd.key_to_qty.get(key, 0.0) + qty
                add_row_list(sd.key_to_rows, key, row_num)

    return sd


def compute_status(in_bom: bool, in_board: bool, bom_qty: float, board_qty: float) -> str:
    if in_bom and in_board:
        if abs(bom_qty - board_qty) < 1e-9:
            return "OK"
        return "QTY_MISMATCH"
    if in_bom and not in_board:
        return "MISSING_ON_TOPBOT"
    if (not in_bom) and in_board:
        return "EXTRA_ON_TOPBOT"
    return "UNKNOWN"


def ensure_fresh_sheet(wb, name: str):
    if name in wb.sheetnames:
        ws_old = wb[name]
        wb.remove(ws_old)
    return wb.create_sheet(title=name)


def write_table(ws, headers: List[str], rows: Iterable[Iterable[object]]) -> None:
    bold = Font(bold=True)
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = bold

    for row in rows:
        ws.append(list(row))

    ws.auto_filter.ref = ws.dimensions
    autofit_columns(ws, max_scan_rows=5000)


def autofit_columns(ws, *, max_scan_rows: int) -> None:
    # openpyxl doesn't have true AutoFit; approximate via max string length.
    dims: Dict[int, int] = {}
    max_row = min(ws.max_row, max_scan_rows)
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            s = str(v)
            dims[c] = max(dims.get(c, 0), len(s))
    for c, ln in dims.items():
        ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = min(max(10, ln + 2), 60)


def materials_summary(sd: SheetData, coord: str) -> str:
    m = sd.coord_to_material_counts.get(coord, {})
    parts = [f"{k}({v})" for k, v in m.items()]
    return ", ".join(parts)


def rows_summary(rows: List[int]) -> str:
    return ",".join(str(r) for r in rows)


def _load_workbook(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    ext = os.path.splitext(path)[1].lower()
    return load_workbook(path, keep_vba=(ext == ".xlsm"))


def run_match(config: dict) -> str:
    """단일 파일(excel_path) 또는 3개 파일(bom_file, top_file, bot_file) 모드 지원."""
    opts = config.get("options", {})
    inplace = bool(opts.get("inplace", False))
    case_insensitive = bool(opts.get("case_insensitive", False))
    output_path = config.get("output_path") or ""

    bom_cfg = config["bom"]
    top_cfg = config["top"]
    bot_cfg = config["bot"]

    # 3개 파일 모드
    if "bom_file" in config:
        bom_path = config["bom_file"]
        top_path = config["top_file"]
        bot_path = config["bot_file"]
        wb_bom = _load_workbook(bom_path)
        wb_top = _load_workbook(top_path)
        wb_bot = _load_workbook(bot_path)
        bom = build_sheet_data(
            wb_bom, label="BOM", sheet_name=bom_cfg["sheet"],
            material_range=bom_cfg["material_range"],
            coord_range=bom_cfg["coord_range"],
            qty_range=bom_cfg["qty_range"],
            case_insensitive=case_insensitive,
        )
        top = build_sheet_data(
            wb_top, label="TOP", sheet_name=top_cfg["sheet"],
            material_range=top_cfg["material_range"],
            coord_range=top_cfg["coord_range"],
            qty_range=top_cfg["qty_range"],
            case_insensitive=case_insensitive,
        )
        bot = build_sheet_data(
            wb_bot, label="BOT", sheet_name=bot_cfg["sheet"],
            material_range=bot_cfg["material_range"],
            coord_range=bot_cfg["coord_range"],
            qty_range=bot_cfg["qty_range"],
            case_insensitive=case_insensitive,
        )
        wb = Workbook()
        default_sheet = wb.active
        if default_sheet:
            wb.remove(default_sheet)
        base_path = bom_path
    else:
        # 단일 파일 모드
        excel_path = config["excel_path"]
        if not os.path.exists(excel_path):
            raise FileNotFoundError(excel_path)
        wb = _load_workbook(excel_path)
        base_path = excel_path
        bom = build_sheet_data(
            wb, label="BOM", sheet_name=bom_cfg["sheet"],
            material_range=bom_cfg["material_range"],
            coord_range=bom_cfg["coord_range"],
            qty_range=bom_cfg["qty_range"],
            case_insensitive=case_insensitive,
        )
        top = build_sheet_data(
            wb, label="TOP", sheet_name=top_cfg["sheet"],
            material_range=top_cfg["material_range"],
            coord_range=top_cfg["coord_range"],
            qty_range=top_cfg["qty_range"],
            case_insensitive=case_insensitive,
        )
        bot = build_sheet_data(
            wb, label="BOT", sheet_name=bot_cfg["sheet"],
            material_range=bot_cfg["material_range"],
            coord_range=bot_cfg["coord_range"],
            qty_range=bot_cfg["qty_range"],
            case_insensitive=case_insensitive,
        )

    all_keys = set(bom.key_to_qty) | set(top.key_to_qty) | set(bot.key_to_qty)

    match_rows = []
    unmatched_top_rows = []
    unmatched_bot_rows = []
    ok_count = 0
    unmatched_top_count = 0
    unmatched_bot_count = 0

    for coord, mat in sorted(all_keys):
        in_b = (coord, mat) in bom.key_to_qty
        in_t = (coord, mat) in top.key_to_qty
        in_bo = (coord, mat) in bot.key_to_qty

        q_b = bom.key_to_qty.get((coord, mat), 0.0)
        q_t = top.key_to_qty.get((coord, mat), 0.0)
        q_bo = bot.key_to_qty.get((coord, mat), 0.0)
        q_tb = q_t + q_bo

        status = compute_status(in_b, (in_t or in_bo), q_b, q_tb)
        status_top = compute_status(in_b, in_t, q_b, q_t)
        status_bot = compute_status(in_b, in_bo, q_b, q_bo)
        key_str = f"{coord}|{mat}"

        row = [
            key_str, coord, mat, q_b, q_t, q_bo, q_tb,
            "Y" if in_b else "", "Y" if in_t else "", "Y" if in_bo else "",
            status,
        ]
        match_rows.append(row)

        if status == "OK":
            ok_count += 1
        if status_top != "OK" and (in_b or in_t):
            unmatched_top_count += 1
            unmatched_top_rows.append([
                key_str, coord, mat, q_b, q_t,
                "Y" if in_b else "", "Y" if in_t else "", status_top,
            ])
        if status_bot != "OK" and (in_b or in_bo):
            unmatched_bot_count += 1
            unmatched_bot_rows.append([
                key_str, coord, mat, q_b, q_bo,
                "Y" if in_b else "", "Y" if in_bo else "", status_bot,
            ])

    ws_match = ensure_fresh_sheet(wb, "Match_Result")
    write_table(
        ws_match,
        headers=[
            "Key(coord|material)", "Coord", "Material",
            "BOM_Qty", "TOP_Qty", "BOT_Qty", "TOP+BOT_Qty",
            "In_BOM", "In_TOP", "In_BOT", "Status",
        ],
        rows=match_rows,
    )

    unmatched_rows = [r[:7] + [r[10]] for r in match_rows if r[10] != "OK"]
    ws_un = ensure_fresh_sheet(wb, "Unmatched")
    write_table(
        ws_un,
        headers=[
            "Key(coord|material)", "Coord", "Material",
            "BOM_Qty", "TOP_Qty", "BOT_Qty", "TOP+BOT_Qty", "Status",
        ],
        rows=unmatched_rows,
    )

    ws_un_top = ensure_fresh_sheet(wb, "Unmatched_TOP")
    write_table(
        ws_un_top,
        headers=[
            "Key(coord|material)", "Coord", "Material",
            "BOM_Qty", "TOP_Qty", "In_BOM", "In_TOP", "Status",
        ],
        rows=unmatched_top_rows,
    )

    ws_un_bot = ensure_fresh_sheet(wb, "Unmatched_BOT")
    write_table(
        ws_un_bot,
        headers=[
            "Key(coord|material)", "Coord", "Material",
            "BOM_Qty", "BOT_Qty", "In_BOM", "In_BOT", "Status",
        ],
        rows=unmatched_bot_rows,
    )

    dup_rows = []
    dup_count_bom = dup_count_top = dup_count_bot = 0
    for sd in (bom, top, bot):
        for coord, cnt in sorted(sd.coord_to_count.items()):
            if cnt > 1:
                if sd.label == "BOM":
                    dup_count_bom += 1
                elif sd.label == "TOP":
                    dup_count_top += 1
                else:
                    dup_count_bot += 1
                dup_rows.append([
                    sd.label, sd.sheet_name, coord, cnt,
                    materials_summary(sd, coord),
                    rows_summary(sd.coord_to_rows.get(coord, [])),
                ])
    ws_dup = ensure_fresh_sheet(wb, "Coord_Duplicates")
    write_table(
        ws_dup,
        headers=["SheetLabel", "SheetName", "Coord", "Count", "Materials(count)", "Rows"],
        rows=dup_rows,
    )

    ws_summary = ensure_fresh_sheet(wb, "Summary")
    write_table(
        ws_summary,
        headers=[
            "구분", "매칭됨(OK)", "불일치(TOP)", "불일치(BOT)",
            "중복좌표_BOM", "중복좌표_TOP", "중복좌표_BOT",
        ],
        rows=[
            [
                "BOM 기준",
                ok_count,
                unmatched_top_count,
                unmatched_bot_count,
                dup_count_bom,
                dup_count_top,
                dup_count_bot,
            ]
        ],
    )

    if "bom_file" in config:
        out = output_path or (os.path.splitext(bom_path)[0] + "_matched.xlsx")
    elif inplace:
        out = base_path
    else:
        if output_path:
            out = output_path
        else:
            root, ext = os.path.splitext(base_path)
            out = f"{root}_matched{ext}" if ext.lower() in (".xlsx", ".xlsm") else f"{root}_matched.xlsx"

    wb.save(out)
    return out


def _get_sheet_names(file_path: str) -> List[str]:
    """엑셀 파일에서 시트 이름 목록 반환. 실패 시 빈 리스트."""
    if not file_path or not os.path.exists(file_path):
        return []
    try:
        wb = load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names
    except Exception:
        return []


def run_gui() -> None:
    """tkinter 폼: BOM/TOP/BOT 파일·시트·범위 선택 후 매칭 실행."""
    root = tk.Tk()
    root.title("BOM / TOP / BOT 매칭")
    root.geometry("520x480")
    root.resizable(True, True)

    main_frame = ttk.Frame(root, padding=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="BOM 파일:").grid(row=0, column=0, sticky=tk.W, pady=2)
    bom_path_var = tk.StringVar()
    ttk.Entry(main_frame, textvariable=bom_path_var, width=42).grid(row=0, column=1, padx=4, pady=2)

    def browse_bom():
        path = filedialog.askopenfilename(
            title="BOM 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if path:
            bom_path_var.set(path)
            names = _get_sheet_names(path)
            bom_sheet_combo["values"] = names
            if names:
                bom_sheet_combo.set(names[0])

    ttk.Button(main_frame, text="찾아보기", command=browse_bom).grid(row=0, column=2, pady=2)

    ttk.Label(main_frame, text="TOP 파일:").grid(row=1, column=0, sticky=tk.W, pady=2)
    top_path_var = tk.StringVar()
    ttk.Entry(main_frame, textvariable=top_path_var, width=42).grid(row=1, column=1, padx=4, pady=2)

    def browse_top():
        path = filedialog.askopenfilename(
            title="TOP 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if path:
            top_path_var.set(path)
            names = _get_sheet_names(path)
            top_sheet_combo["values"] = names
            if names:
                top_sheet_combo.set(names[0])

    ttk.Button(main_frame, text="찾아보기", command=browse_top).grid(row=1, column=2, pady=2)

    ttk.Label(main_frame, text="BOT 파일:").grid(row=2, column=0, sticky=tk.W, pady=2)
    bot_path_var = tk.StringVar()
    ttk.Entry(main_frame, textvariable=bot_path_var, width=42).grid(row=2, column=1, padx=4, pady=2)

    def browse_bot():
        path = filedialog.askopenfilename(
            title="BOT 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if path:
            bot_path_var.set(path)
            names = _get_sheet_names(path)
            bot_sheet_combo["values"] = names
            if names:
                bot_sheet_combo.set(names[0])

    ttk.Button(main_frame, text="찾아보기", command=browse_bot).grid(row=2, column=2, pady=2)

    ttk.Label(main_frame, text="저장 경로 (비우면 BOM과 같은 폴더에 _matched.xlsx):").grid(
        row=3, column=0, sticky=tk.W, pady=2
    )
    output_path_var = tk.StringVar()
    ttk.Entry(main_frame, textvariable=output_path_var, width=42).grid(row=3, column=1, padx=4, pady=2)

    def browse_output():
        path = filedialog.asksaveasfilename(
            title="결과 저장 위치",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("모든 파일", "*.*")],
        )
        if path:
            output_path_var.set(path)

    ttk.Button(main_frame, text="찾아보기", command=browse_output).grid(row=3, column=2, pady=2)

    ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=10)
    ttk.Label(main_frame, text="시트 및 범위 (각 1열, 예: B2:B200):").grid(
        row=5, column=0, columnspan=3, sticky=tk.W
    )

    bom_sheet_var = tk.StringVar()
    top_sheet_var = tk.StringVar()
    bot_sheet_var = tk.StringVar()
    bom_mat_var = tk.StringVar(value="B2:B200")
    bom_coord_var = tk.StringVar(value="C2:C200")
    bom_qty_var = tk.StringVar(value="D2:D200")
    top_mat_var = tk.StringVar(value="B2:B200")
    top_coord_var = tk.StringVar(value="C2:C200")
    top_qty_var = tk.StringVar(value="D2:D200")
    bot_mat_var = tk.StringVar(value="B2:B200")
    bot_coord_var = tk.StringVar(value="C2:C200")
    bot_qty_var = tk.StringVar(value="D2:D200")

    def _grid_section(r, label, combo_widget, mat_var, coord_var, qty_var):
        ttk.Label(main_frame, text=label).grid(row=r, column=0, sticky=tk.W, pady=2)
        combo_widget.grid(row=r, column=1, padx=4, pady=2)
        ttk.Label(main_frame, text="자재범위").grid(row=r + 1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=mat_var, width=22).grid(row=r + 1, column=1, padx=4, pady=2)
        ttk.Label(main_frame, text="좌표범위").grid(row=r + 2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=coord_var, width=22).grid(row=r + 2, column=1, padx=4, pady=2)
        ttk.Label(main_frame, text="수량범위").grid(row=r + 3, column=0, sticky=tk.W, pady=2)
        ttk.Entry(main_frame, textvariable=qty_var, width=22).grid(row=r + 3, column=1, padx=4, pady=2)

    bom_sheet_combo = ttk.Combobox(main_frame, textvariable=bom_sheet_var, width=22)
    top_sheet_combo = ttk.Combobox(main_frame, textvariable=top_sheet_var, width=22)
    bot_sheet_combo = ttk.Combobox(main_frame, textvariable=bot_sheet_var, width=22)

    _grid_section(6, "BOM 시트", bom_sheet_combo, bom_mat_var, bom_coord_var, bom_qty_var)
    _grid_section(10, "TOP 시트", top_sheet_combo, top_mat_var, top_coord_var, top_qty_var)
    _grid_section(14, "BOT 시트", bot_sheet_combo, bot_mat_var, bot_coord_var, bot_qty_var)

    case_insensitive_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(
        main_frame, text="좌표/자재 대소문자 무시", variable=case_insensitive_var
    ).grid(row=18, column=0, columnspan=2, sticky=tk.W, pady=8)

    def run_match_from_gui() -> None:
        bom_path = bom_path_var.get().strip()
        top_path = top_path_var.get().strip()
        bot_path = bot_path_var.get().strip()
        if not bom_path or not top_path or not bot_path:
            messagebox.showerror("입력 오류", "BOM, TOP, BOT 파일을 모두 선택해 주세요.")
            return
        if not bom_sheet_var.get().strip():
            messagebox.showerror("입력 오류", "BOM 시트를 선택해 주세요.")
            return
        if not top_sheet_var.get().strip():
            messagebox.showerror("입력 오류", "TOP 시트를 선택해 주세요.")
            return
        if not bot_sheet_var.get().strip():
            messagebox.showerror("입력 오류", "BOT 시트를 선택해 주세요.")
            return
        cfg = {
            "bom_file": bom_path,
            "top_file": top_path,
            "bot_file": bot_path,
            "output_path": output_path_var.get().strip() or None,
            "bom": {
                "sheet": bom_sheet_var.get().strip(),
                "material_range": bom_mat_var.get().strip(),
                "coord_range": bom_coord_var.get().strip(),
                "qty_range": bom_qty_var.get().strip(),
            },
            "top": {
                "sheet": top_sheet_var.get().strip(),
                "material_range": top_mat_var.get().strip(),
                "coord_range": top_coord_var.get().strip(),
                "qty_range": top_qty_var.get().strip(),
            },
            "bot": {
                "sheet": bot_sheet_var.get().strip(),
                "material_range": bot_mat_var.get().strip(),
                "coord_range": bot_coord_var.get().strip(),
                "qty_range": bot_qty_var.get().strip(),
            },
            "options": {
                "inplace": False,
                "case_insensitive": case_insensitive_var.get(),
            },
        }
        try:
            out = run_match(cfg)
            messagebox.showinfo("완료", f"결과 저장:\n{out}")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    ttk.Button(main_frame, text="매칭 실행", command=run_match_from_gui).grid(
        row=19, column=1, pady=12
    )

    root.mainloop()


def load_config(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=False, help="Path to config JSON (없으면 GUI 실행)")
    args = ap.parse_args()

    if args.config:
        cfg = load_config(args.config)
        out = run_match(cfg)
        print(out)
    else:
        run_gui()


if __name__ == "__main__":
    main()

