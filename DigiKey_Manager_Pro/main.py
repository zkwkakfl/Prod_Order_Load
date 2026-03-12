"""
디지키 API를 사용한 파트넘버 조회 애플리케이션
엑셀 파일에서 시트를 선택하고, 파트넘버를 더블클릭하여 조회하는 GUI 프로그램
버전 1.2.6 - 유사 자재 목록 선택, 조회 실패 시 Google 웹 검색 기능
"""

import difflib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import webbrowser
from excel_handler import ExcelHandler
from digikey_api import DigikeyAPIClient, RateLimitExceeded
from database import PartDatabase

# 유사 자재: 최소 유사도(0~1), 목록 최대 개수
SIMILAR_PARTS_MIN_RATIO = 0.6
SIMILAR_PARTS_MAX_COUNT = 10

# 컬럼 선택 다이얼로그: 표 형태 미리보기에 표시할 최대 행 수
COLUMN_PREVIEW_MAX_ROWS = 20


class DigikeyViewerApp:
    """메인 애플리케이션 클래스"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("디지키 파트넘버 조회 프로그램 v1.2.6")
        self.root.geometry("1200x700")
        
        # API 일일 호출 제한 (디지키 Product Information API)
        self.api_daily_limit = 1000
        
        # 데이터 저장 변수
        self.excel_handler = ExcelHandler()
        self.digikey_api = DigikeyAPIClient()
        self.part_db = PartDatabase()  # 파트넘버 데이터베이스
        self.current_df = None  # 현재 로드된 엑셀 데이터
        self.query_results = []  # 조회 결과 저장
        self.config_file = "config.txt"  # 설정 파일 경로
        
        # config 파일에서 API 키 로드
        self.load_config()
        
        # 디지키 API 설정 확인
        self.check_api_config()
        
        # GUI 초기화 (먼저 UI를 생성해야 함)
        self.init_ui()
        
        # 시작 시 파일 및 시트 선택 유저폼 표시 (약간의 지연 후)
        self.root.after(100, self.show_initial_setup)
    
    def init_ui(self):
        """UI 초기화"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 윈도우 종료 프로토콜 설정
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 메뉴바
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="엑셀 파일 열기", command=self.load_excel_file)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.on_closing)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="설정", menu=settings_menu)
        settings_menu.add_command(label="디지키 API 설정", command=self.show_api_settings)
        
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도구", menu=tools_menu)
        tools_menu.add_command(label="데이터베이스 통계", command=self.show_db_stats)
        tools_menu.add_command(label="API 호출 통계", command=self.show_api_stats)
        
        # 상단 도구바
        toolbar = ttk.Frame(main_frame)
        toolbar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(toolbar, text="엑셀 파일 열기", command=self.load_excel_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="시트 선택", command=self.select_sheet).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="디지키 API 설정", command=self.show_api_settings).pack(side=tk.LEFT, padx=5)
        
        self.sheet_label = ttk.Label(toolbar, text="선택된 시트: 없음")
        self.sheet_label.pack(side=tk.LEFT, padx=10)
        
        # API 호출 통계 라벨
        self.api_stats_label = ttk.Label(toolbar, text="")
        self.api_stats_label.pack(side=tk.RIGHT, padx=10)
        self.update_api_stats_label()
        
        # 탭 위젯 생성
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 탭 1: 시트 데이터 리스트뷰
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="시트 데이터")
        self.setup_tab1()
        
        # 탭 2: 조회 목록 및 상세정보
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="조회 결과")
        self.setup_tab2()
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
    
    def setup_tab1(self):
        """탭 1 설정: 시트 데이터 표시"""
        # 프레임 설정
        frame = ttk.Frame(self.tab1, padding="5")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 트리뷰 (리스트뷰 역할) 및 스크롤바
        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
        
        self.tree1 = ttk.Treeview(frame, yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.config(command=self.tree1.yview)
        scrollbar_x.config(command=self.tree1.xview)
        
        # 그리드 배치
        self.tree1.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        
        # 더블클릭 이벤트 바인딩
        self.tree1.bind("<Double-1>", self.on_part_double_click)
    
    def setup_tab2(self):
        """탭 2 설정: 조회 결과 및 상세정보"""
        # 좌우 분할 프레임
        paned = ttk.PanedWindow(self.tab2, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 왼쪽: 조회 결과 리스트뷰
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=3)
        
        scrollbar_y2 = ttk.Scrollbar(left_frame, orient=tk.VERTICAL)
        scrollbar_x2 = ttk.Scrollbar(left_frame, orient=tk.HORIZONTAL)
        
        self.tree2 = ttk.Treeview(left_frame, yscrollcommand=scrollbar_y2.set, xscrollcommand=scrollbar_x2.set)
        scrollbar_y2.config(command=self.tree2.yview)
        scrollbar_x2.config(command=self.tree2.xview)
        
        self.tree2.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_y2.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x2.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # 더블클릭 이벤트 바인딩
        self.tree2.bind("<Double-1>", self.on_query_result_double_click)
        
        # 오른쪽: 상세정보 패널
        right_frame = ttk.Frame(paned, padding="10")
        paned.add(right_frame, weight=2)
        
        # 상세정보 라벨
        ttk.Label(right_frame, text="상세 정보", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # 상세정보 텍스트 위젯 (읽기 전용)
        scrollbar_detail = ttk.Scrollbar(right_frame, orient=tk.VERTICAL)
        self.detail_text = tk.Text(right_frame, yscrollcommand=scrollbar_detail.set, wrap=tk.WORD, width=40, state=tk.DISABLED)
        scrollbar_detail.config(command=self.detail_text.yview)
        
        self.detail_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_detail.pack(side=tk.RIGHT, fill=tk.Y)
    
    def load_config(self):
        """config.txt 파일에서 API 키 로드"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith('#'):
                            continue
                        
                        if '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip()
                            
                            if key.lower() == 'clientid':
                                self.digikey_api.client_id = value
                            elif key.lower() == 'clientsecret':
                                self.digikey_api.client_secret = value
                            elif key.lower() == 'usesandbox' or key.lower() == 'sandbox':
                                # 샌드박스 환경 설정 (기본값: False, 프로덕션)
                                self.digikey_api.use_sandbox = value.lower() in ('true', '1', 'yes')
                                self.digikey_api.base_url = (
                                    self.digikey_api.SANDBOX_BASE_URL 
                                    if self.digikey_api.use_sandbox 
                                    else self.digikey_api.PRODUCTION_BASE_URL
                                )
                            elif key.lower() == 'redirecturi':
                                # RedirectURI는 저장만 하고 사용하지 않음 (필요시 사용 가능)
                                pass
            except Exception as e:
                print(f"config 파일 읽기 오류: {str(e)}")
    
    def save_config(self, client_id, client_secret, use_sandbox=True):
        """config.txt 파일에 API 키 저장 (중복 방지)"""
        try:
            # 기존 config 파일 읽기
            config_data = {}
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith('#'):
                            continue
                        if '=' in line:
                            key, value = line.split('=', 1)
                            key = key.strip()
                            value = value.strip()
                            # 중복 키 방지: 첫 번째 값만 사용
                            if key not in config_data:
                                config_data[key] = value
            
            # API 키 및 환경 설정 업데이트 (기존 값 덮어쓰기)
            config_data['ClientID'] = client_id
            config_data['ClientSecret'] = client_secret
            config_data['UseSandbox'] = 'true' if use_sandbox else 'false'
            
            # RedirectURI는 유지 (있는 경우)
            if 'RedirectURI' not in config_data:
                config_data['RedirectURI'] = 'https://localhost'
            
            # config 파일에 저장 (순서대로)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                # 주요 설정을 먼저 저장
                if 'ClientID' in config_data:
                    f.write(f"ClientID={config_data['ClientID']}\n")
                if 'ClientSecret' in config_data:
                    f.write(f"ClientSecret={config_data['ClientSecret']}\n")
                if 'RedirectURI' in config_data:
                    f.write(f"RedirectURI={config_data['RedirectURI']}\n")
                if 'UseSandbox' in config_data:
                    f.write(f"UseSandbox={config_data['UseSandbox']}\n")
        except Exception as e:
            print(f"config 파일 저장 오류: {str(e)}")
    
    def check_api_config(self):
        """디지키 API 설정 확인 및 자동 설정 다이얼로그 표시"""
        # API 설정 확인은 나중에 수행 (유저폼 이후)
        pass
    
    def show_initial_setup(self):
        """시작 시 엑셀 파일 및 시트 선택 유저폼 표시"""
        # 메인 윈도우를 뒤로 보내기
        self.root.lower()
        
        # 독립적인 윈도우로 생성
        setup_window = tk.Toplevel(self.root)
        setup_window.title("엑셀 파일 및 시트 선택")
        setup_window.geometry("500x300")
        setup_window.resizable(False, False)
        
        # 창 중앙 배치
        setup_window.update_idletasks()
        x = (setup_window.winfo_screenwidth() // 2) - (setup_window.winfo_width() // 2)
        y = (setup_window.winfo_screenheight() // 2) - (setup_window.winfo_height() // 2)
        setup_window.geometry(f"+{x}+{y}")
        
        # 모달 다이얼로그로 만들기
        setup_window.transient(self.root)
        setup_window.grab_set()
        setup_window.focus_set()
        setup_window.lift()
        setup_window.attributes('-topmost', True)  # 최상위로 설정
        setup_window.update()  # 강제 업데이트
        
        # 메인 프레임
        main_frame = ttk.Frame(setup_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(main_frame, text="엑셀 파일 및 시트를 선택하세요", font=("Arial", 12, "bold"))
        title_label.pack(pady=(0, 20))
        
        # 파일 선택 섹션
        file_frame = ttk.LabelFrame(main_frame, text="엑셀 파일", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50, state="readonly")
        file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        def browse_file():
            filename = filedialog.askopenfilename(
                title="엑셀 파일 선택",
                filetypes=[
                    ("엑셀 통합문서", "*.xlsx"),
                    ("엑셀 매크로 포함 문서", "*.xlsm"),
                    ("엑셀 파일 (통합문서 및 매크로 포함)", "*.xlsx *.xlsm"),
                    ("엑셀 97-2003 통합문서", "*.xls"),
                    ("모든 엑셀 파일", "*.xlsx *.xlsm *.xls"),
                    ("모든 파일", "*.*")
                ]
            )
            if filename:
                self.file_path_var.set(filename)
                try:
                    self.excel_handler.load_file(filename)
                    sheets = self.excel_handler.get_sheet_names()
                    sheet_combo['values'] = sheets
                    if sheets:
                        sheet_var.set(sheets[0])
                except Exception as e:
                    messagebox.showerror("오류", f"파일 로드 중 오류가 발생했습니다:\n{str(e)}")
                    self.file_path_var.set("")
        
        ttk.Button(file_frame, text="찾아보기...", command=browse_file).pack(side=tk.LEFT, padx=5)
        
        # 시트 선택 섹션
        sheet_frame = ttk.LabelFrame(main_frame, text="시트 선택", padding="10")
        sheet_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(sheet_frame, text="시트:").pack(side=tk.LEFT, padx=5)
        
        sheet_var = tk.StringVar()
        sheet_combo = ttk.Combobox(sheet_frame, textvariable=sheet_var, state="readonly", width=40)
        sheet_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        def confirm_setup():
            file_path = self.file_path_var.get().strip()
            selected_sheet = sheet_var.get().strip()
            
            if not file_path:
                messagebox.showwarning("경고", "엑셀 파일을 선택해주세요.")
                return
            
            if not selected_sheet:
                messagebox.showwarning("경고", "시트를 선택해주세요.")
                return
            
            try:
                # 시트 로드
                self.current_df = self.excel_handler.load_sheet(selected_sheet)
                setup_window.destroy()
                # 메인 윈도우를 앞으로 가져오기
                self.root.lift()
                self.root.focus_force()
                # 데이터 표시
                self.finish_setup(selected_sheet)
                # 유저폼 완료 후 API 설정 확인
                self.check_api_config_after_setup()
            except Exception as e:
                messagebox.showerror("오류", f"시트 로드 중 오류가 발생했습니다:\n{str(e)}")
        
        def cancel_setup():
            if messagebox.askyesno("확인", "프로그램을 종료하시겠습니까?"):
                setup_window.destroy()
                self.root.quit()
        
        ttk.Button(button_frame, text="확인", command=confirm_setup, width=15,).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="취소", command=cancel_setup, width=15).pack(side=tk.LEFT, padx=5)
        
        # 파일 변경 시 시트 목록 업데이트
        def on_file_change(*args):
            if self.file_path_var.get():
                try:
                    sheets = self.excel_handler.get_sheet_names()
                    sheet_combo['values'] = sheets
                    if sheets:
                        sheet_var.set(sheets[0])
                except:
                    pass
        
        # 취소 버튼으로 창 닫기 방지 (확인 또는 취소 버튼만 사용)
        setup_window.protocol("WM_DELETE_WINDOW", cancel_setup)
    
    def finish_setup(self, sheet_name):
        """초기 설정 완료 후 실행"""
        if hasattr(self, 'sheet_label'):
            self.sheet_label.config(text=f"선택된 시트: {sheet_name}")
        if self.current_df is not None:
            self.display_sheet_data()
    
    def load_excel_file(self):
        """엑셀 파일 로드"""
        filename = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[
                ("엑셀 통합문서", "*.xlsx"),
                ("엑셀 매크로 포함 문서", "*.xlsm"),
                ("엑셀 파일 (통합문서 및 매크로 포함)", "*.xlsx *.xlsm"),
                ("엑셀 97-2003 통합문서", "*.xls"),
                ("모든 엑셀 파일", "*.xlsx *.xlsm *.xls"),
                ("모든 파일", "*.*")
            ]
        )
        
        if filename:
            try:
                self.excel_handler.load_file(filename)
                self.current_df = None
                messagebox.showinfo("성공", f"파일이 로드되었습니다: {filename}")
                self.select_sheet()
            except Exception as e:
                messagebox.showerror("오류", f"파일 로드 중 오류가 발생했습니다:\n{str(e)}")
    
    def select_sheet(self):
        """시트 선택 다이얼로그"""
        if not self.excel_handler.file_loaded():
            messagebox.showwarning("경고", "먼저 엑셀 파일을 열어주세요.")
            return
        
        sheets = self.excel_handler.get_sheet_names()
        if not sheets:
            messagebox.showwarning("경고", "사용 가능한 시트가 없습니다.")
            return
        
        # 시트 선택 다이얼로그
        sheet_window = tk.Toplevel(self.root)
        sheet_window.title("시트 선택")
        sheet_window.geometry("300x200")
        
        ttk.Label(sheet_window, text="시트를 선택하세요:").pack(pady=10)
        
        sheet_var = tk.StringVar(value=sheets[0])
        sheet_combo = ttk.Combobox(sheet_window, textvariable=sheet_var, values=sheets, state="readonly")
        sheet_combo.pack(pady=10)
        
        def load_sheet():
            selected_sheet = sheet_var.get()
            try:
                self.current_df = self.excel_handler.load_sheet(selected_sheet)
                self.sheet_label.config(text=f"선택된 시트: {selected_sheet}")
                self.display_sheet_data()
                sheet_window.destroy()
            except Exception as e:
                messagebox.showerror("오류", f"시트 로드 중 오류가 발생했습니다:\n{str(e)}")
        
        ttk.Button(sheet_window, text="확인", command=load_sheet).pack(pady=10)
    
    def display_sheet_data(self):
        """시트 데이터를 트리뷰에 표시"""
        if self.current_df is None or self.current_df.empty:
            return
        
        # 기존 항목 삭제
        for item in self.tree1.get_children():
            self.tree1.delete(item)
        
        # 컬럼 설정
        columns = list(self.current_df.columns)
        self.tree1["columns"] = columns
        self.tree1["show"] = "headings"
        
        # 헤더 설정
        for col in columns:
            self.tree1.heading(col, text=col)
            self.tree1.column(col, width=150, anchor=tk.W)
        
        # 데이터 삽입
        for index, row in self.current_df.iterrows():
            values = [str(val) for val in row.values]
            self.tree1.insert("", tk.END, values=values, iid=index)
    
    def clean_part_number(self, part_number: str) -> str:
        """
        파트넘버 기본 정리 (안전한 정리만 수행)
        
        Args:
            part_number: 원본 파트넘버
            
        Returns:
            str: 정리된 파트넘버
        """
        if not part_number:
            return part_number
        
        # 앞뒤 공백 제거
        cleaned = part_number.strip()
        
        # 줄바꿈, 탭 문자 제거
        cleaned = cleaned.replace('\n', '').replace('\r', '').replace('\t', '')
        
        # 연속된 공백을 하나로 (단, 파트넘버 내부 공백은 유지)
        # 예: "ABC  123" -> "ABC 123" (너무 공격적이지 않게)
        
        return cleaned
    
    def is_query_failed(self, result: dict) -> bool:
        """
        조회 결과가 실패인지 판단
        
        Args:
            result: 조회 결과 딕셔너리
            
        Returns:
            bool: 실패 여부
        """
        if not result:
            return True
        
        # Manufacturer가 "검색 결과 없음"인 경우
        if result.get('Manufacturer') == '검색 결과 없음':
            return True
        
        # Error 필드가 있는 경우
        if 'Error' in result or 'error' in result:
            return True
        
        # Manufacturer가 "API 오류" 또는 "조회 실패"인 경우
        manufacturer = result.get('Manufacturer', '')
        if manufacturer in ['API 오류', '조회 실패']:
            return True
        
        return False
    
    def _normalize_mounting_to_smt_imt(self, value: str) -> str:
        """
        API 장착유형 문자열을 엑셀용 코드로 변환.
        표면실장 → "SMT", 스루홀 → "IMT", 그 외 → "".
        """
        if not value or not isinstance(value, str):
            return ""
        v = value.strip().lower()
        if not v or v in ('n/a', 'nan'):
            return ""
        if 'surface' in v or 'smt' in v or v == 'smt':
            return "SMT"
        if 'through' in v or 'thru' in v or 'through-hole' in v or '스루홀' in v or v == 'imt':
            return "IMT"
        return ""
    
    def _find_part_type_column(self) -> str | None:
        """시트에서 '자재 유형' 컬럼을 찾아 이름 반환. 없으면 None."""
        if self.current_df is None or self.current_df.empty:
            return None
        candidates = ['자재 유형', '자재유형', '유형', '부품유형', 'Part Type', '자재유형 ']
        for col in self.current_df.columns:
            c = str(col).strip()
            if not c:
                continue
            if c in candidates:
                return col
            if '자재' in c and '유형' in c:
                return col
        return None
    
    def _is_chip_resistor_row(self, row_index: int, part_type_col: str | None) -> bool:
        """해당 행의 자재 유형이 '칩저항'인지 여부. (조회 안 된 항목 중 칩저항만 SMT 처리용)"""
        if part_type_col is None or self.current_df is None:
            return False
        try:
            val = self.current_df.iloc[row_index].get(part_type_col, "")
            return "칩저항" in str(val).strip()
        except Exception:
            return False
    
    def _apply_mounting_column_to_sheet(self, part_number_col: str, query_results: list):
        """
        조회 결과를 반영해 current_df에 '장착유형' 열을 파트넘버 열 오른쪽에 삽입/갱신.
        값: 조회 성공 시 SMT/IMT, 조회 실패 시 자재 유형이 칩저항이면 SMT, 아니면 빈 문자열.
        """
        if self.current_df is None or not query_results:
            return
        part_type_col = self._find_part_type_column()
        mapping = {}
        for r in query_results:
            idx = r['Row']
            full = r.get('FullData') or {}
            if not self.is_query_failed(full):
                mapping[idx] = self._normalize_mounting_to_smt_imt(r.get('MountingType', '') or '')
            else:
                mapping[idx] = "SMT" if self._is_chip_resistor_row(idx, part_type_col) else ""
        series = pd.Series(mapping).reindex(self.current_df.index, fill_value="")
        col_name = "장착유형"
        if col_name in self.current_df.columns:
            self.current_df.drop(columns=[col_name], inplace=True)
        col_pos = self.current_df.columns.get_loc(part_number_col)
        self.current_df.insert(col_pos + 1, col_name, series)
        self.display_sheet_data()
    
    def _similarity_ratio(self, a: str, b: str) -> float:
        """두 문자열의 유사도 반환 (0~1). 대소문자 무시."""
        if not a or not b:
            return 0.0
        return difflib.SequenceMatcher(None, a.lower().strip(), b.lower().strip()).ratio()
    
    def show_similar_parts_selection_dialog(self, original_part_number: str, row_index: int, similar_list: list) -> tuple:
        """
        유사 자재 목록을 보여주고 사용자가 하나 선택하도록 함.
        
        Args:
            original_part_number: 원본 파트넘버
            row_index: 행 인덱스
            similar_list: 유사 자재 dict 목록 (각 dict에 PartNumber, Manufacturer, MountingType, Description, Similarity 등)
            
        Returns:
            tuple: (선택한 결과 dict, True) 또는 (None, False) 취소 시
        """
        if not similar_list:
            return (None, False)
        sel_window = tk.Toplevel(self.root)
        sel_window.title("유사 자재 선택")
        sel_window.geometry("750x400")
        sel_window.resizable(True, True)
        sel_window.transient(self.root)
        sel_window.grab_set()
        sel_window.focus_set()
        sel_window.lift()
        main_frame = ttk.Frame(sel_window, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text=f"'{original_part_number}' (Row {row_index})에 대한 유사 자재입니다. 하나를 선택하세요.", font=("Arial", 10, "bold")).pack(pady=(0, 10))
        columns = ("similarity", "part", "manufacturer", "mounting", "description")
        tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=12, selectmode="browse")
        tree.heading("similarity", text="유사도(%)")
        tree.heading("part", text="파트넘버")
        tree.heading("manufacturer", text="제조사")
        tree.heading("mounting", text="마운팅타입")
        tree.heading("description", text="설명")
        tree.column("similarity", width=80)
        tree.column("part", width=150)
        tree.column("manufacturer", width=120)
        tree.column("mounting", width=100)
        tree.column("description", width=250)
        scroll_y = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll_y.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        for i, item in enumerate(similar_list):
            sim = item.get("Similarity", 0)
            pct = f"{sim * 100:.0f}%" if isinstance(sim, (int, float)) else str(sim)
            desc = (item.get("Description") or "N/A")[:50]
            if len((item.get("Description") or "")) > 50:
                desc += "..."
            tree.insert("", tk.END, iid=i, values=(
                pct,
                item.get("PartNumber", ""),
                item.get("Manufacturer", "N/A"),
                item.get("MountingType", "N/A"),
                desc
            ))
        selected_result = [None]
        def on_ok():
            sel = tree.selection()
            if sel:
                idx = int(sel[0])
                if 0 <= idx < len(similar_list):
                    selected_result[0] = similar_list[idx]
            sel_window.destroy()
        def on_cancel():
            selected_result[0] = None
            sel_window.destroy()
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="선택", command=on_ok, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=on_cancel, width=12).pack(side=tk.LEFT, padx=5)
        sel_window.protocol("WM_DELETE_WINDOW", on_cancel)
        tree.bind("<Double-1>", lambda e: on_ok())
        sel_window.wait_window()
        return (selected_result[0], selected_result[0] is not None)
    
    def show_part_number_edit_dialog(self, original_part_number: str, row_index: int) -> tuple:
        """
        조회 실패 시 파트넘버 수정 및 웹 검색 다이얼로그 표시
        
        Args:
            original_part_number: 원본 파트넘버
            row_index: 행 인덱스
            
        Returns:
            tuple: (수정된 파트넘버, 웹 검색 여부)
                   사용자가 취소하면 (None, False) 반환
                   웹 검색을 했을 때는 (파트넘버, False) 반환 (API 재조회하지 않음)
        """
        edit_window = tk.Toplevel(self.root)
        edit_window.title("파트넘버 조회 실패")
        edit_window.geometry("550x280")
        edit_window.resizable(False, False)
        
        # 창 중앙 배치
        edit_window.update_idletasks()
        x = (edit_window.winfo_screenwidth() // 2) - (edit_window.winfo_width() // 2)
        y = (edit_window.winfo_screenheight() // 2) - (edit_window.winfo_height() // 2)
        edit_window.geometry(f"+{x}+{y}")
        
        # 모달 다이얼로그로 만들기
        edit_window.transient(self.root)
        edit_window.grab_set()
        edit_window.focus_set()
        edit_window.lift()
        
        # 메인 프레임
        main_frame = ttk.Frame(edit_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="파트넘버 조회 실패", 
            font=("Arial", 12, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 안내 메시지
        info_label = ttk.Label(
            main_frame,
            text=f"파트넘버 '{original_part_number}' (Row {row_index})를 찾을 수 없습니다.\n"
                 f"파트넘버에 오타나 불필요한 문자가 있을 수 있습니다.\n"
                 f"수정 후 Google에서 검색하거나 건너뛸 수 있습니다.",
            justify=tk.LEFT,
            font=("Arial", 9)
        )
        info_label.pack(pady=(0, 15))
        
        # 입력 프레임
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(input_frame, text="파트넘버:", width=12).pack(side=tk.LEFT, padx=5)
        part_var = tk.StringVar(value=original_part_number)
        part_entry = ttk.Entry(input_frame, textvariable=part_var, width=40)
        part_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        part_entry.select_range(0, tk.END)  # 전체 선택
        part_entry.focus()
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        result = [None, False]  # [수정된 파트넘버, 웹 검색 여부]
        
        def on_web_search():
            """웹에서 검색 버튼 클릭"""
            modified = part_var.get().strip()
            if not modified:
                modified = original_part_number
            
            # Google에서 디지키 사이트 검색
            search_query = f"{modified}"
            google_url = f"https://www.google.com/search?q={search_query}"
            webbrowser.open(google_url)
            
            # 웹 검색을 했으므로 결과 저장 (API 재조회하지 않음)
            result[0] = modified
            result[1] = False  # False = API 재조회하지 않음
            edit_window.destroy()
        
        def on_skip():
            """건너뛰기 버튼 클릭"""
            result[0] = original_part_number
            result[1] = False
            edit_window.destroy()
        
        def on_cancel():
            """취소 버튼 클릭"""
            result[0] = None
            result[1] = False
            edit_window.destroy()
        
        # Enter 키로 웹 검색
        part_entry.bind('<Return>', lambda e: on_web_search())
        
        ttk.Button(button_frame, text="웹에서 검색", command=on_web_search, width=14).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="건너뛰기", command=on_skip, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="취소", command=on_cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        # 창 닫기 프로토콜
        edit_window.protocol("WM_DELETE_WINDOW", on_cancel)
        
        # 다이얼로그가 닫힐 때까지 대기
        edit_window.wait_window()
        
        return tuple(result)
    
    def query_part_with_retry(self, part_number: str, row_index: int, progress_window=None) -> tuple:
        """
        파트넘버 조회 (실패 시 정리 및 재시도, 웹 검색 옵션 제공)
        
        Args:
            part_number: 조회할 파트넘버
            row_index: 행 인덱스 (다이얼로그 표시용)
            progress_window: 진행 상황 창 (선택적)
            
        Returns:
            tuple: (조회 결과 dict, API 호출 횟수)
        """
        original_part_number = part_number
        api_call_count = 0  # 실제 API 호출 횟수 추적
        
        # 1차: 원본 파트넘버로 조회
        result = None
        
        # 먼저 DB에서 조회
        db_result = self.part_db.get_part(part_number)
        if db_result:
            # DB에 결과가 있으면 (성공/실패 상관없이) 그 결과를 반환
            # 이미 조회한 적이 있으므로 다시 API를 호출하지 않음
            return (db_result, 0)  # DB에서 조회한 경우 API 호출 없음
        
        # DB에 없는 경우에만 API 조회
        try:
            api_result = self.digikey_api.search_part(part_number)
            if api_result:
                api_call_count += 1
                self.part_db.increment_api_call()
                if not self.is_query_failed(api_result):
                    # 성공한 경우 DB에 저장
                    self.part_db.save_part(api_result)
                    return (api_result, api_call_count)
                # 실패한 경우도 DB에 저장 (중복 조회 방지)
                result = api_result
                self.part_db.save_part(api_result)
        except RateLimitExceeded:
            raise  # 상위로 전달
        except Exception as e:
            result = {
                'PartNumber': part_number,
                'Manufacturer': '조회 실패',
                'MountingType': '조회 실패',
                'error': str(e)
            }
            # API 호출은 했지만 예외가 발생한 경우도 저장
            if api_call_count > 0:
                self.part_db.save_part(result)
        
        # 조회 실패 시
        if self.is_query_failed(result):
            # 2차: 기본 정리 후 재조회
            cleaned_part = self.clean_part_number(part_number)
            if cleaned_part != part_number:
                # 정리된 버전으로 재조회 시도
                try:
                    cleaned_result = self.digikey_api.search_part(cleaned_part)
                    if cleaned_result:
                        api_call_count += 1
                        self.part_db.increment_api_call()
                        if not self.is_query_failed(cleaned_result):
                            # 성공한 경우 DB에 저장
                            self.part_db.save_part(cleaned_result)
                            return (cleaned_result, api_call_count)
                        # 실패한 경우도 DB에 저장 (중복 조회 방지)
                        self.part_db.save_part(cleaned_result)
                except RateLimitExceeded:
                    raise
                except Exception:
                    pass  # 정리된 버전도 실패하면 계속 진행
            
            # 2.5차: 유사 자재 검색 (최대 10개, 유사도 60% 이상) → 사용자 선택
            try:
                similar_raw = self.digikey_api.search_part_multiple(original_part_number, 15)
                api_call_count += 1
                self.part_db.increment_api_call()
                similar_with_ratio = []
                for r in similar_raw:
                    ratio = self._similarity_ratio(original_part_number, r.get("PartNumber", ""))
                    if ratio >= SIMILAR_PARTS_MIN_RATIO:
                        r_copy = dict(r)
                        r_copy["Similarity"] = ratio
                        similar_with_ratio.append(r_copy)
                similar_with_ratio.sort(key=lambda x: x.get("Similarity", 0), reverse=True)
                similar_filtered = similar_with_ratio[:SIMILAR_PARTS_MAX_COUNT]
                if similar_filtered:
                    if progress_window:
                        progress_window.withdraw()
                    selected_result, did_select = self.show_similar_parts_selection_dialog(
                        original_part_number, row_index, similar_filtered
                    )
                    if progress_window:
                        progress_window.deiconify()
                    if did_select and selected_result:
                        self.part_db.save_part(selected_result)
                        return (selected_result, api_call_count)
            except RateLimitExceeded:
                raise
            except Exception:
                pass
            
            # 3차: 사용자에게 웹 검색 옵션 제공
            # 진행 창이 있으면 일시적으로 숨김
            if progress_window:
                progress_window.withdraw()
            
            modified_part, web_searched = self.show_part_number_edit_dialog(
                original_part_number, row_index
            )
            
            # 진행 창 다시 표시
            if progress_window:
                progress_window.deiconify()
            
            # 웹 검색을 했거나 건너뛰기를 선택한 경우
            if modified_part:
                # 건너뛰기 선택한 경우: 이미 API를 호출했으므로 실패 결과를 DB에 저장
                final_result = result if result else {
                    'PartNumber': original_part_number,
                    'Manufacturer': '검색 결과 없음',
                    'MountingType': 'N/A',
                    'Description': '파트넘버를 찾을 수 없습니다.'
                }
                # API를 호출한 경우에만 저장 (중복 조회 방지)
                if api_call_count > 0:
                    self.part_db.save_part(final_result)
                return (final_result, api_call_count)
            else:
                # 취소 선택한 경우 원본 결과 반환
                final_result = result if result else {
                    'PartNumber': original_part_number,
                    'Manufacturer': '검색 결과 없음',
                    'MountingType': 'N/A',
                    'Description': '파트넘버를 찾을 수 없습니다.'
                }
                return (final_result, api_call_count)
        
        # 결과가 있으면 반환 (실패한 경우도 포함)
        final_result = result if result else {
            'PartNumber': original_part_number,
            'Manufacturer': '검색 결과 없음',
            'MountingType': 'N/A',
            'Description': '파트넘버를 찾을 수 없습니다.'
        }
        # API를 호출한 경우 실패 결과도 DB에 저장 (중복 조회 방지)
        if api_call_count > 0 and self.is_query_failed(final_result):
            self.part_db.save_part(final_result)
        return (final_result, api_call_count)
    
    def on_part_double_click(self, event):
        """파트넘버 더블클릭 이벤트 처리"""
        selection = self.tree1.selection()
        if not selection:
            return
        
        # 선택된 행의 인덱스 가져오기
        selected_item = selection[0]
        try:
            row_index = int(selected_item)
        except ValueError:
            messagebox.showerror("오류", "행 인덱스를 가져올 수 없습니다.")
            return
        
        if self.current_df is None or row_index >= len(self.current_df):
            return
        
        # 파트넘버 컬럼 찾기 (대소문자 무시)
        part_number_col = self.find_part_number_column()
        
        if part_number_col is None:
            # 컬럼을 찾지 못한 경우 사용자에게 선택하게 함
            part_number_col = self.select_part_number_column()
            if part_number_col is None:
                return  # 사용자가 취소한 경우
        
        # 선택한 행부터 아래로 순환하며 조회
        self.query_parts_from_row(row_index, part_number_col)
    
    def find_part_number_column(self):
        """파트넘버 컬럼 자동 찾기"""
        if self.current_df is None or self.current_df.empty:
            return None
        
        # 다양한 패턴으로 파트넘버 컬럼 찾기
        possible_patterns = [
            # 정확한 매칭
            lambda col: 'part' in col.lower() and 'number' in col.lower(),
            # 파트넘버 (한글)
            lambda col: '파트' in col and '넘버' in col,
            lambda col: '파트' in col and '번호' in col,
            # Part Number (공백 포함)
            lambda col: col.lower().replace(' ', '') == 'partnumber',
            lambda col: col.lower().replace('_', '') == 'partnumber',
            # Part만 포함
            lambda col: col.lower() == 'part',
            lambda col: col.lower() == 'partno',
            lambda col: col.lower() == 'part_no',
            # Number만 포함 (일부 경우)
            lambda col: col.lower() == 'number' and 'part' not in col.lower(),
        ]
        
        for pattern in possible_patterns:
            for col in self.current_df.columns:
                if pattern(col):
                    return col
        
        return None
    
    def _fill_treeview_from_df(self, tree, df, max_rows=None):
        """
        DataFrame 내용으로 Treeview를 채움 (미리보기용).
        max_rows가 지정되면 상위 N행만 표시.
        """
        for item in tree.get_children():
            tree.delete(item)
        if df is None or df.empty:
            return
        cols = list(df.columns)
        tree["columns"] = cols
        tree["show"] = "headings"
        for c in cols:
            tree.heading(c, text=str(c))
            tree.column(c, width=120, anchor=tk.W)
        subset = df.head(max_rows) if max_rows else df
        for index, row in subset.iterrows():
            values = [str(val) for val in row.values]
            tree.insert("", tk.END, values=values)

    def select_part_number_column(self):
        """사용자에게 파트넘버 컬럼 선택하게 함 (표 형태 미리보기 포함)"""
        if self.current_df is None or self.current_df.empty:
            return None
        
        columns = list(self.current_df.columns)
        
        # 컬럼 선택 다이얼로그 (미리보기 영역 포함으로 크기 확대)
        col_window = tk.Toplevel(self.root)
        col_window.title("파트넘버 컬럼 선택")
        col_window.geometry("680x480")
        col_window.transient(self.root)
        col_window.grab_set()
        col_window.focus_set()
        
        main_frm = ttk.Frame(col_window, padding="10")
        main_frm.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frm, text="파트넘버 컬럼을 선택하세요 (아래 표에서 확인):", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        col_var = tk.StringVar()
        if columns:
            col_var.set(columns[0])
        row_sel = ttk.Frame(main_frm)
        row_sel.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(row_sel, text="컬럼:").pack(side=tk.LEFT, padx=(0, 5))
        col_combo = ttk.Combobox(row_sel, textvariable=col_var, values=columns, state="readonly", width=40)
        col_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 표 형태 미리보기 (A: 앱 내 테이블)
        preview_frm = ttk.LabelFrame(main_frm, text="시트 미리보기 (상위 {}행)".format(COLUMN_PREVIEW_MAX_ROWS), padding="5")
        preview_frm.pack(fill=tk.BOTH, expand=True, pady=5)
        prev_tree = ttk.Treeview(preview_frm, height=12, show="headings", selectmode="none")
        scroll_y = ttk.Scrollbar(preview_frm)
        scroll_x = ttk.Scrollbar(preview_frm, orient=tk.HORIZONTAL)
        prev_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.configure(command=prev_tree.yview)
        scroll_x.configure(command=prev_tree.xview)
        prev_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        preview_frm.columnconfigure(0, weight=1)
        preview_frm.rowconfigure(0, weight=1)
        self._fill_treeview_from_df(prev_tree, self.current_df, max_rows=COLUMN_PREVIEW_MAX_ROWS)
        
        selected_col = [None]
        
        def confirm():
            selected_col[0] = col_var.get()
            col_window.destroy()
        
        def cancel():
            col_window.destroy()
        
        btn_frm = ttk.Frame(main_frm)
        btn_frm.pack(pady=10)
        ttk.Button(btn_frm, text="확인", command=confirm, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frm, text="취소", command=cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        # 창 중앙 배치
        col_window.update_idletasks()
        w, h = 680, 480
        x = (col_window.winfo_screenwidth() // 2) - (w // 2)
        y = (col_window.winfo_screenheight() // 2) - (h // 2)
        col_window.geometry(f"{w}x{h}+{x}+{y}")
        
        col_window.wait_window()
        return selected_col[0]
    
    def query_parts_from_row(self, start_row, part_number_col):
        """선택한 행부터 아래로 순환하며 파트넘버 조회"""
        if self.current_df is None:
            return
        
        query_results = []
        db_hits = 0  # DB에서 조회한 횟수
        api_calls = 0  # API 호출 횟수
        
        # v1.2.4: 전체 조회할 파트넘버 개수 계산 (빈 값 제외)
        total_parts = 0
        for idx in range(start_row, len(self.current_df)):
            part_number = str(self.current_df.iloc[idx][part_number_col]).strip()
            if part_number and part_number != 'nan':
                total_parts += 1
        
        # 진행 상황 표시
        progress_window = tk.Toplevel(self.root)
        progress_window.title("조회 중...")
        progress_window.geometry("450x150")
        progress_window.resizable(False, False)
        
        # 창 중앙 배치
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
        y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
        progress_window.geometry(f"+{x}+{y}")
        
        progress_label = ttk.Label(progress_window, text="파트넘버를 조회하고 있습니다...", font=("Arial", 10, "bold"))
        progress_label.pack(pady=10)
        
        # v1.2.4: 진행 상황 표시 (전체/조회 완료)
        progress_count_label = ttk.Label(progress_window, text=f"조회 진행: 0 / {total_parts}개", font=("Arial", 9))
        progress_count_label.pack(pady=2)
        
        progress_detail = ttk.Label(progress_window, text="", font=("Arial", 8))
        progress_detail.pack(pady=5)
        
        # 진행률 바 추가
        progress_bar = ttk.Progressbar(progress_window, length=400, mode='determinate', maximum=total_parts)
        progress_bar.pack(pady=5)
        
        self.root.update()
        
        queried_count = 0  # 조회 완료한 개수
        
        try:
            # 선택한 행부터 끝까지 순환
            for idx in range(start_row, len(self.current_df)):
                part_number = str(self.current_df.iloc[idx][part_number_col]).strip()
                
                # 빈 값 건너뛰기
                if not part_number or part_number == 'nan':
                    continue
                
                queried_count += 1
                
                # v1.2.5: 진행 상황 업데이트 (전체/조회 완료 개수 표시)
                progress_count_label.config(text=f"조회 진행: {queried_count} / {total_parts}개")
                progress_detail.config(text=f"조회 중: {part_number} (Row {idx})\nDB: {db_hits}건 | API: {api_calls}건")
                progress_bar['value'] = queried_count
                self.root.update()
                
                result = None
                part_api_calls = 0  # 이 파트넘버 조회 시 실제 API 호출 횟수
                
                # v1.2.6: 새로운 재시도 로직 사용 (정리 → 웹 검색 옵션 제공)
                try:
                    # DB에서 먼저 조회 (성공한 경우만)
                    db_result = self.part_db.get_part(part_number)
                    if db_result and not self.is_query_failed(db_result):
                        result = db_result
                        db_hits += 1
                    else:
                        # DB에 없거나 실패한 경우 재시도 로직 사용
                        result, part_api_calls = self.query_part_with_retry(part_number, idx, progress_window)
                        api_calls += part_api_calls  # 실제 API 호출 횟수 추가
                            
                except RateLimitExceeded as e:
                    # API 호출 한도 초과 시 조회 중단
                    progress_window.destroy()
                    
                    # 재시도 시간 정보 포함 메시지
                    retry_info = ""
                    if e.retry_after:
                        hours = e.retry_after // 3600
                        minutes = (e.retry_after % 3600) // 60
                        if hours > 0:
                            retry_info = f"\n\n재시도 가능 시간: 약 {hours}시간 {minutes}분 후"
                        elif minutes > 0:
                            retry_info = f"\n\n재시도 가능 시간: 약 {minutes}분 후"
                        else:
                            retry_info = f"\n\n재시도 가능 시간: 약 {e.retry_after}초 후"
                    
                    messagebox.showwarning(
                        "API 호출 한도 초과",
                        f"API 일일 호출 제한에 도달했습니다.\n\n"
                        f"조회가 중단되었습니다.\n"
                        f"현재까지 조회된 결과: {len(query_results)}개\n"
                        f"DB 조회: {db_hits}건 | API 호출: {api_calls}건\n\n"
                        f"상세 정보:\n{str(e)}{retry_info}\n\n"
                        f"※ Product Information API 제한:\n"
                        f"  - 일일 최대: 1,000회\n"
                        f"  - 분당 최대: 120회"
                    )
                    
                    # 현재까지 조회된 결과 표시
                    if query_results:
                        self.query_results = query_results
                        self.display_query_results()
                        self.notebook.select(1)
                    
                    return  # 조회 중단
                except Exception as e:
                    # 기타 오류 시에도 기본 정보는 추가 (오류 메시지 포함)
                    error_msg = str(e)
                    print(f"파트넘버 조회 오류 ({part_number}): {error_msg}")  # 콘솔에 오류 출력
                    result = {
                        'PartNumber': part_number,
                        'Manufacturer': '조회 실패',
                        'MountingType': '조회 실패',
                        'error': error_msg
                    }
                
                # 결과가 있으면 query_results에 추가
                if result:
                    query_results.append({
                        'Row': idx,
                        'PartNumber': part_number,
                        'Manufacturer': result.get('Manufacturer', 'N/A'),
                        'MountingType': result.get('MountingType', 'N/A'),
                        'FullData': result
                    })
            
            self.query_results = query_results
            self._apply_mounting_column_to_sheet(part_number_col, query_results)
            self.display_query_results()
            
            # 조회 탭으로 전환
            self.notebook.select(1)
            
            progress_window.destroy()
            
            # API 통계 라벨 업데이트
            self.update_api_stats_label()
            
            # 조회 결과 통계 메시지
            today_total_calls = self.part_db.get_today_api_calls()
            remaining_calls = max(0, self.api_daily_limit - today_total_calls)
            
            messagebox.showinfo(
                "조회 완료", 
                f"{len(query_results)}개의 파트넘버 조회가 완료되었습니다.\n\n"
                f"조회 통계:\n"
                f"  - 데이터베이스에서 조회: {db_hits}건\n"
                f"  - API 호출: {api_calls}건\n\n"
                f"오늘 API 호출 현황:\n"
                f"  - 총 호출: {today_total_calls}/{self.api_daily_limit}회\n"
                f"  - 남은 호출: {remaining_calls}회\n\n"
                f"※ API 호출 횟수가 절약되었습니다!"
            )
            
        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("오류", f"조회 중 오류가 발생했습니다:\n{str(e)}")
    
    def display_query_results(self):
        """조회 결과를 트리뷰에 표시"""
        # 기존 항목 삭제
        for item in self.tree2.get_children():
            self.tree2.delete(item)
        
        if not self.query_results:
            return
        
        # 컬럼 설정
        columns = ['Row', 'PartNumber', 'Manufacturer', 'MountingType']
        self.tree2["columns"] = columns
        self.tree2["show"] = "headings"
        
        # 헤더 설정
        self.tree2.heading('Row', text='Row')
        self.tree2.heading('PartNumber', text='파트넘버')
        self.tree2.heading('Manufacturer', text='제조업체')
        self.tree2.heading('MountingType', text='마운팅타입')
        
        # 컬럼 너비 설정
        self.tree2.column('Row', width=60, anchor=tk.CENTER)
        self.tree2.column('PartNumber', width=200, anchor=tk.W)
        self.tree2.column('Manufacturer', width=200, anchor=tk.W)
        self.tree2.column('MountingType', width=150, anchor=tk.W)
        
        # 데이터 삽입 (장착유형: 시트와 동일하게 SMT/IMT/조회실패 표시)
        part_type_col = self._find_part_type_column()
        for i, result in enumerate(self.query_results):
            full = result.get('FullData') or {}
            if not self.is_query_failed(full):
                mt_display = self._normalize_mounting_to_smt_imt(result.get('MountingType', '') or '') or result.get('MountingType', 'N/A')
            else:
                mt_display = "SMT" if self._is_chip_resistor_row(result['Row'], part_type_col) else "조회 실패"
            values = [
                str(result['Row']),
                result['PartNumber'],
                result['Manufacturer'],
                mt_display
            ]
            self.tree2.insert("", tk.END, values=values, iid=i)
    
    def on_query_result_double_click(self, event):
        """조회 결과 더블클릭 시 상세정보 표시"""
        selection = self.tree2.selection()
        if not selection:
            return
        
        try:
            item_index = int(selection[0])
            if 0 <= item_index < len(self.query_results):
                result = self.query_results[item_index]
                self.display_detail_info(result)
        except (ValueError, IndexError) as e:
            messagebox.showerror("오류", f"상세정보를 가져올 수 없습니다: {str(e)}")
    
    def display_detail_info(self, result_data):
        """상세정보를 텍스트 위젯에 표시"""
        # 텍스트 위젯을 편집 가능하게 변경
        self.detail_text.config(state=tk.NORMAL)
        self.detail_text.delete(1.0, tk.END)
        
        # 기존 태그 삭제
        for tag in self.detail_text.tag_names():
            self.detail_text.tag_delete(tag)
        
        full_data = result_data.get('FullData', {})
        
        # 기본 정보
        info_text = "=== 파트넘버 상세정보 ===\n\n"
        
        # 데이터 출처 표시
        source = full_data.get('Source', 'API')
        info_text += f"데이터 출처: {source}\n"
        if source == 'Database':
            info_text += f"생성일시: {full_data.get('CreatedAt', 'N/A')}\n"
            info_text += f"수정일시: {full_data.get('UpdatedAt', 'N/A')}\n"
        info_text += "\n"
        
        info_text += f"Row: {result_data.get('Row', 'N/A')}\n"
        info_text += f"파트넘버: {result_data.get('PartNumber', 'N/A')}\n"
        info_text += f"제조업체: {result_data.get('Manufacturer', 'N/A')}\n"
        info_text += f"마운팅타입: {result_data.get('MountingType', 'N/A')}\n\n"
        
        self.detail_text.insert(tk.END, info_text)
        
        # URL 정보 처리 (링크로 표시)
        if 'error' not in full_data:
            self.detail_text.insert(tk.END, "--- 추가 정보 ---\n")
            
            # 제품 URL
            product_url = full_data.get('ProductUrl', '')
            if product_url:
                self.detail_text.insert(tk.END, "제품 상세정보: ")
                url_start = self.detail_text.index(tk.INSERT)
                self.detail_text.insert(tk.END, product_url)
                url_end = self.detail_text.index(tk.INSERT)
                self.detail_text.insert(tk.END, "\n")
                
                # 링크 태그 생성 및 적용
                tag_name = f"product_url_{id(product_url)}"
                self.detail_text.tag_add(tag_name, url_start, url_end)
                self.detail_text.tag_config(tag_name, foreground="blue", underline=True)
                self.detail_text.tag_bind(tag_name, "<Button-1>", lambda e, url=product_url: self.open_url(url))
                self.detail_text.tag_bind(tag_name, "<Enter>", lambda e: self.detail_text.config(cursor="hand2"))
                self.detail_text.tag_bind(tag_name, "<Leave>", lambda e: self.detail_text.config(cursor=""))
            
            # 데이터시트 URL
            datasheet_url = full_data.get('DatasheetUrl', '')
            if datasheet_url:
                self.detail_text.insert(tk.END, "데이터시트 URL: ")
                url_start = self.detail_text.index(tk.INSERT)
                self.detail_text.insert(tk.END, datasheet_url)
                url_end = self.detail_text.index(tk.INSERT)
                self.detail_text.insert(tk.END, "\n")
                
                # 링크 태그 생성 및 적용
                tag_name = f"datasheet_url_{id(datasheet_url)}"
                self.detail_text.tag_add(tag_name, url_start, url_end)
                self.detail_text.tag_config(tag_name, foreground="blue", underline=True)
                self.detail_text.tag_bind(tag_name, "<Button-1>", lambda e, url=datasheet_url: self.open_url(url))
                self.detail_text.tag_bind(tag_name, "<Enter>", lambda e: self.detail_text.config(cursor="hand2"))
                self.detail_text.tag_bind(tag_name, "<Leave>", lambda e: self.detail_text.config(cursor=""))
            
            # 기타 정보
            for key, value in full_data.items():
                # 이미 표시한 항목이나 URL 항목 제외
                if key not in ['Manufacturer', 'MountingType', 'ProductUrl', 'DatasheetUrl', 'Source', 'CreatedAt', 'UpdatedAt']:
                    # 딕셔너리나 리스트인 경우 문자열로 변환
                    if isinstance(value, dict):
                        value_str = ", ".join([f"{k}: {v}" for k, v in value.items()])
                        self.detail_text.insert(tk.END, f"{key}: {value_str}\n")
                    elif isinstance(value, list):
                        value_str = ", ".join([str(v) for v in value])
                        self.detail_text.insert(tk.END, f"{key}: {value_str}\n")
                    else:
                        self.detail_text.insert(tk.END, f"{key}: {value}\n")
        else:
            # v1.2.4: 조회 실패한 파트넘버에 웹 검색 링크 추가
            error_msg = full_data.get('error', '알 수 없는 오류')
            self.detail_text.insert(tk.END, f"\n오류: {error_msg}\n\n")
            
            # 디지키 웹 검색 링크 추가
            part_number = result_data.get('PartNumber', '')
            if part_number:
                self.detail_text.insert(tk.END, "--- 웹 검색 ---\n")
                self.detail_text.insert(tk.END, "디지키 웹사이트에서 검색: ")
                
                # 디지키 검색 URL 생성
                search_url = f"https://www.digikey.com/en/products?keywords={part_number}"
                url_start = self.detail_text.index(tk.INSERT)
                self.detail_text.insert(tk.END, search_url)
                url_end = self.detail_text.index(tk.INSERT)
                self.detail_text.insert(tk.END, "\n")
                
                # 링크 태그 생성 및 적용
                tag_name = f"search_url_{id(search_url)}"
                self.detail_text.tag_add(tag_name, url_start, url_end)
                self.detail_text.tag_config(tag_name, foreground="blue", underline=True)
                self.detail_text.tag_bind(tag_name, "<Button-1>", lambda e, url=search_url: self.open_url(url))
                self.detail_text.tag_bind(tag_name, "<Enter>", lambda e: self.detail_text.config(cursor="hand2"))
                self.detail_text.tag_bind(tag_name, "<Leave>", lambda e: self.detail_text.config(cursor=""))
        
        # 텍스트 위젯을 다시 읽기 전용으로 변경
        self.detail_text.config(state=tk.DISABLED)
    
    def open_url(self, url):
        """웹브라우저에서 URL 열기"""
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("오류", f"URL을 열 수 없습니다:\n{url}\n\n{str(e)}")
    
    def check_api_config_after_setup(self):
        """유저폼 완료 후 API 설정 확인"""
        # config 파일에서 읽은 후에도 API 설정이 없을 때만 물어봄
        if not self.digikey_api.is_configured():
            # 메인 윈도우가 표시된 후 API 설정 안내
            self.root.after(300, self.show_api_settings_with_message)
    
    def show_api_settings_with_message(self):
        """API 설정 다이얼로그를 메시지와 함께 표시"""
        # 메인 윈도우가 표시되어 있는지 확인
        if not self.root.winfo_viewable():
            self.root.deiconify()
        
        response = messagebox.askyesno(
            "API 설정 필요",
            "디지키 API를 사용하려면 API 키를 설정해야 합니다.\n\n"
            "지금 설정하시겠습니까?\n\n"
            "(나중에 메뉴 > 설정 > 디지키 API 설정에서도 설정할 수 있습니다.)"
        )
        if response:
            self.show_api_settings()
    
    def show_api_settings(self):
        """디지키 API 설정 다이얼로그"""
        # 메인 윈도우가 표시되어 있는지 확인하고 표시
        if not self.root.winfo_viewable():
            self.root.deiconify()
        
        settings_window = tk.Toplevel(self.root)
        settings_window.title("디지키 API 설정")
        settings_window.geometry("500x400")
        settings_window.transient(self.root)
        settings_window.grab_set()  # 모달 다이얼로그
        settings_window.focus_set()  # 포커스 설정
        settings_window.lift()  # 다른 창 위로 올리기
        
        # 창 중앙 배치
        settings_window.update_idletasks()
        x = (settings_window.winfo_screenwidth() // 2) - (settings_window.winfo_width() // 2)
        y = (settings_window.winfo_screenheight() // 2) - (settings_window.winfo_height() // 2)
        settings_window.geometry(f"+{x}+{y}")
        
        # 메인 프레임
        main_frame = ttk.Frame(settings_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 제목 및 안내
        ttk.Label(main_frame, text="디지키 API 설정", font=("Arial", 12, "bold")).pack(pady=(0, 10))
        ttk.Label(
            main_frame, 
            text="디지키 개발자 포털(developer.digikey.com)에서 받은\nClient ID와 Client Secret을 입력하세요.",
            justify=tk.CENTER,
            foreground="gray"
        ).pack(pady=(0, 10))
        
        # 환경 선택 (샌드박스/프로덕션)
        env_frame = ttk.Frame(main_frame)
        env_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(env_frame, text="환경:", width=15).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        env_var = tk.BooleanVar(value=self.digikey_api.use_sandbox)
        ttk.Radiobutton(env_frame, text="샌드박스", variable=env_var, value=True).grid(row=0, column=1, padx=5, sticky=tk.W)
        ttk.Radiobutton(env_frame, text="프로덕션", variable=env_var, value=False).grid(row=0, column=2, padx=5, sticky=tk.W)
        
        ttk.Label(
            main_frame, 
            text="※ 샌드박스 키는 샌드박스 환경, 프로덕션 키는 프로덕션 환경을 선택하세요",
            justify=tk.CENTER,
            foreground="red",
            font=("Arial", 8)
        ).pack(pady=(0, 10))
        
        # 입력 프레임
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(input_frame, text="Client ID:", width=15).grid(row=0, column=0, padx=5, pady=10, sticky=tk.W)
        client_id_entry = ttk.Entry(input_frame, width=35)
        client_id_entry.grid(row=0, column=1, padx=5, pady=10, sticky=(tk.W, tk.E))
        if self.digikey_api.client_id:
            client_id_entry.insert(0, self.digikey_api.client_id)
        
        ttk.Label(input_frame, text="Client Secret:", width=15).grid(row=1, column=0, padx=5, pady=10, sticky=tk.W)
        client_secret_entry = ttk.Entry(input_frame, width=35, show="*")
        client_secret_entry.grid(row=1, column=1, padx=5, pady=10, sticky=(tk.W, tk.E))
        if self.digikey_api.client_secret:
            client_secret_entry.insert(0, self.digikey_api.client_secret)
        
        input_frame.columnconfigure(1, weight=1)
        
        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=30, fill=tk.X)
        
        def save_settings():
            client_id = client_id_entry.get().strip()
            client_secret = client_secret_entry.get().strip()
            use_sandbox = env_var.get()
            
            if client_id and client_secret:
                self.digikey_api.set_credentials(client_id, client_secret)
                self.digikey_api.use_sandbox = use_sandbox
                self.digikey_api.base_url = (
                    self.digikey_api.SANDBOX_BASE_URL 
                    if use_sandbox 
                    else self.digikey_api.PRODUCTION_BASE_URL
                )
                # config 파일에 저장
                self.save_config(client_id, client_secret, use_sandbox)
                messagebox.showinfo("성공", "API 설정이 저장되었습니다.\nconfig.txt 파일에도 저장되었습니다.")
                settings_window.destroy()
            else:
                messagebox.showwarning("경고", "Client ID와 Client Secret을 모두 입력해주세요.")
        
        # 저장 버튼 (더 눈에 띄게)
        save_btn = ttk.Button(button_frame, text="저장", command=save_settings, width=18)
        save_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        # 취소 버튼
        cancel_btn = ttk.Button(button_frame, text="취소", command=settings_window.destroy, width=18)
        cancel_btn.pack(side=tk.LEFT, padx=10, expand=True)
    
    def update_api_stats_label(self):
        """API 통계 라벨 업데이트"""
        if hasattr(self, 'api_stats_label') and hasattr(self, 'part_db') and self.part_db:
            try:
                today_calls = self.part_db.get_today_api_calls()
                remaining = max(0, self.api_daily_limit - today_calls)
                self.api_stats_label.config(
                    text=f"오늘 API 호출: {today_calls}/{self.api_daily_limit} (남은: {remaining})",
                    foreground="red" if remaining < 100 else "blue"
                )
            except:
                pass
    
    def show_db_stats(self):
        """데이터베이스 통계 정보 표시"""
        if not hasattr(self, 'part_db') or not self.part_db:
            messagebox.showwarning("경고", "데이터베이스가 초기화되지 않았습니다.")
            return
        
        try:
            stats = self.part_db.get_stats()
            today_calls = stats.get('today_api_calls', 0)
            remaining = max(0, self.api_daily_limit - today_calls)
            
            messagebox.showinfo(
                "데이터베이스 통계",
                f"파트넘버 데이터베이스 통계\n\n"
                f"총 저장된 파트넘버: {stats.get('total_parts', 0):,}개\n"
                f"제조사 수: {stats.get('total_manufacturers', 0):,}개\n"
                f"마운팅 타입 수: {stats.get('total_mounting_types', 0):,}개\n\n"
                f"오늘 API 호출: {today_calls}/{self.api_daily_limit}회\n"
                f"남은 호출: {remaining}회\n\n"
                f"※ 데이터베이스 파일: parts_cache.db"
            )
        except Exception as e:
            messagebox.showerror("오류", f"통계 정보를 가져오는 중 오류가 발생했습니다:\n{str(e)}")
    
    def show_api_stats(self):
        """API 호출 통계 상세 정보 표시"""
        if not hasattr(self, 'part_db') or not self.part_db:
            messagebox.showwarning("경고", "데이터베이스가 초기화되지 않았습니다.")
            return
        
        try:
            # 최근 30일 통계 조회
            recent_stats = self.part_db.get_api_call_stats(limit=30)
            today_calls = self.part_db.get_today_api_calls()
            remaining = max(0, self.api_daily_limit - today_calls)
            
            # 통계 윈도우 생성
            stats_window = tk.Toplevel(self.root)
            stats_window.title("API 호출 통계")
            stats_window.geometry("500x600")
            stats_window.transient(self.root)
            
            # 창 중앙 배치
            stats_window.update_idletasks()
            x = (stats_window.winfo_screenwidth() // 2) - (stats_window.winfo_width() // 2)
            y = (stats_window.winfo_screenheight() // 2) - (stats_window.winfo_height() // 2)
            stats_window.geometry(f"+{x}+{y}")
            
            # 메인 프레임
            main_frame = ttk.Frame(stats_window, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 제목
            ttk.Label(main_frame, text="API 호출 통계", font=("Arial", 14, "bold")).pack(pady=(0, 10))
            
            # 오늘 통계 프레임
            today_frame = ttk.LabelFrame(main_frame, text="오늘 통계", padding="10")
            today_frame.pack(fill=tk.X, pady=10)
            
            ttk.Label(today_frame, text=f"총 호출: {today_calls}회", font=("Arial", 11)).pack(anchor=tk.W, pady=2)
            ttk.Label(today_frame, text=f"일일 한도: {self.api_daily_limit}회", font=("Arial", 11)).pack(anchor=tk.W, pady=2)
            ttk.Label(
                today_frame, 
                text=f"남은 호출: {remaining}회",
                font=("Arial", 11, "bold"),
                foreground="red" if remaining < 100 else "green"
            ).pack(anchor=tk.W, pady=2)
            
            # 진행률 바
            progress_frame = ttk.Frame(today_frame)
            progress_frame.pack(fill=tk.X, pady=5)
            
            progress_bar = ttk.Progressbar(
                progress_frame, 
                length=400, 
                mode='determinate',
                maximum=self.api_daily_limit,
                value=today_calls
            )
            progress_bar.pack(fill=tk.X)
            
            percentage = (today_calls / self.api_daily_limit * 100) if self.api_daily_limit > 0 else 0
            ttk.Label(progress_frame, text=f"{percentage:.1f}% 사용", font=("Arial", 9)).pack(anchor=tk.E, pady=2)
            
            # 최근 호출 통계 프레임
            history_frame = ttk.LabelFrame(main_frame, text="최근 30일 호출 이력", padding="10")
            history_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            # 트리뷰 및 스크롤바
            scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL)
            
            stats_tree = ttk.Treeview(
                history_frame,
                columns=('date', 'calls'),
                show='headings',
                yscrollcommand=scrollbar.set,
                height=15
            )
            scrollbar.config(command=stats_tree.yview)
            
            stats_tree.heading('date', text='날짜')
            stats_tree.heading('calls', text='호출 횟수')
            
            stats_tree.column('date', width=200, anchor=tk.CENTER)
            stats_tree.column('calls', width=150, anchor=tk.CENTER)
            
            # 데이터 삽입
            for stat in recent_stats:
                stats_tree.insert('', tk.END, values=(stat['date'], f"{stat['count']:,}회"))
            
            stats_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # 닫기 버튼
            ttk.Button(main_frame, text="닫기", command=stats_window.destroy, width=15).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("오류", f"API 통계를 가져오는 중 오류가 발생했습니다:\n{str(e)}")
    
    def on_closing(self):
        """프로그램 종료 처리"""
        # 열려있는 모든 Toplevel 윈도우 닫기
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Toplevel):
                try:
                    widget.destroy()
                except:
                    pass
        
        # 데이터베이스 연결 종료
        if hasattr(self, 'part_db') and self.part_db:
            self.part_db.close()
        
        # 메인 윈도우 종료
        self.root.quit()
        self.root.destroy()


def main():
    """메인 함수"""
    root = tk.Tk()
    app = DigikeyViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
