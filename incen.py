import os
import sys
import pandas as pd
import numpy as np
import json
import traceback
from datetime import datetime
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, StringVar
from tkinter import ttk
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path

# 사용자 문서 폴더 사용
docs_path = os.path.join(Path.home(), "Documents", "IncentiveCalculator")
os.makedirs(docs_path, exist_ok=True)
file_path = os.path.join(docs_path, "incentive_calculator_settings.json")

def format_excel_sheet(ws):
    """엑셀 시트에 서식 적용"""
    # 제목 행: 남색 배경 + 흰색 글씨
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")

    # 소계 행: 옅은 초록색 배경
    subtotal_fill = PatternFill(start_color="D8EAD3", end_color="D8EAD3", fill_type="solid")
    
    # 총계 행: 옅은 파란색 배경
    total_fill = PatternFill(start_color="D0E0E3", end_color="D0E0E3", fill_type="solid")

    # 제목 행 서식 적용
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    # 헤더 행에서 열 인덱스 찾기
    date_columns = []
    money_columns = []
    지부명_column = None
    매출월_column = None
    
    for col_idx, cell in enumerate(ws[1], start=1):
        column_name = str(cell.value).strip() if cell.value else ""
        if column_name in ['접수일자', '입금일']:
            date_columns.append(col_idx)
        if column_name in ['금액', '입금액', '차감액', '인센티브금액', '공급가액', '세액', '수수료']:
            money_columns.append(col_idx)
        if column_name == '지부명':
            지부명_column = col_idx
        if column_name == '매출월':
            매출월_column = col_idx
    
    print(f"날짜 컬럼: {date_columns}, 금액 컬럼: {money_columns}, 지부명 컬럼: {지부명_column}, 매출월 컬럼: {매출월_column}")

    # 각 행에 대해 서식 적용
    for row_idx in range(2, ws.max_row + 1):
        # 소계/총계 배경색 적용
        is_subtotal = False
        is_total = False
        
        if 지부명_column:
            지부명_value = ws.cell(row=row_idx, column=지부명_column).value
            if 지부명_value == '소계':
                is_subtotal = True
        
        if 매출월_column:
            매출월_value = ws.cell(row=row_idx, column=매출월_column).value
            if 매출월_value == '총계':
                is_total = True
        
        # 소계 행 서식
        if is_subtotal:
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = subtotal_fill
        
        # 총계 행 서식
        if is_total:
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = total_fill
        
        # 날짜 서식 적용
        for col_idx in date_columns:
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value:
                cell.number_format = 'yyyy-mm-dd'
        
        # 금액 서식 적용
        for col_idx in money_columns:
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                # 문자열이면 숫자로 변환 시도
                if isinstance(cell.value, str):
                    try:
                        cell.value = float(cell.value.replace(',', ''))
                    except ValueError:
                        pass
                
                # 차감액이 있는 경우 옅은 빨간색 배경 적용
                deduction_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")

                # 금액 서식 적용
                for col_idx in money_columns:
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value is not None:
                        # 문자열이면 숫자로 변환 시도
                        if isinstance(cell.value, str):
                            try:
                                cell.value = float(cell.value.replace(',', ''))
                            except ValueError:
                                pass

                        # 숫자 서식 적용
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0'

                        # 차감액 컬럼이며 값이 양수인 경우 배경색 적용
                        if ws.cell(row=1, column=col_idx).value == '차감액' and cell.value > 0:
                            cell.fill = deduction_fill


    # 모든 열 너비 자동 조정
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

class SettingsDialog(tk.Toplevel):
    """설정 관리 다이얼로그"""
    
    def __init__(self, calculator, settings):
        # calculator는 IncentiveCalculator 인스턴스, 실제 윈도우 부모는 calculator.root
        super().__init__(calculator.root)
        self.calculator = calculator  # IncentiveCalculator 참조 저장
        self.settings = settings.copy()
        
        # 영업담당 목록 가져오기
        self.available_sales_reps = self.calculator.get_sales_reps_list()
        
        # UI 구성
        self.title("설정")
        self.geometry("600x500")
        self.resizable(True, True)
        
        # 노트북(탭) 생성
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 제외 영업담당 탭
        self.excluded_reps_frame = tk.Frame(self.notebook)
        self.notebook.add(self.excluded_reps_frame, text="제외 영업담당")
        self.setup_excluded_reps_tab()
        
        # 지부별 인센티브율 탭
        self.branch_rates_frame = tk.Frame(self.notebook)
        self.notebook.add(self.branch_rates_frame, text="지부별 인센티브율")
        self.setup_branch_rates_tab()
        
        # 기본 영업담당 탭 추가
        self.default_reps_frame = tk.Frame(self.notebook)
        self.notebook.add(self.default_reps_frame, text="기본 영업담당")
        self.setup_default_reps_tab()
        
        # 지부 설정 탭 추가
        self.branch_setup_frame = tk.Frame(self.notebook)
        self.notebook.add(self.branch_setup_frame, text="지부 설정")
        self.setup_branch_setup_tab()

        # 컬럼 선택 탭 추가
        self.column_select_frame = tk.Frame(self.notebook)
        self.notebook.add(self.column_select_frame, text='표시할 컬럼 선택')
        self.setup_column_select_tab()
        
        # 버튼 프레임
        button_frame = tk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 저장 버튼
        save_btn = tk.Button(button_frame, text="저장", command=self.save_settings, width=10)
        save_btn.pack(side=tk.RIGHT, padx=5)
        
        # 취소 버튼
        cancel_btn = tk.Button(button_frame, text="취소", command=self.destroy, width=10)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def setup_column_select_tab(self):
        """표시할 컬럼 선택 탭"""
        frame = tk.Frame(self.column_select_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(frame, text="기본 컬럼 외에 표시할 추가 컬럼을 선택하세요.").pack(anchor=tk.W, pady=(0, 10))

        list_frame = tk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 왼쪽: 추가 가능한 컬럼
        available_frame = tk.LabelFrame(list_frame, text="추가 가능한 컬럼")
        available_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        self.available_column_listbox = tk.Listbox(available_frame, selectmode=tk.MULTIPLE)
        self.available_column_listbox.pack(fill=tk.BOTH, expand=True)

        # 오른쪽: 선택된 컬럼
        selected_frame = tk.LabelFrame(list_frame, text="표시할 컬럼")
        selected_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        self.selected_column_listbox = tk.Listbox(selected_frame, selectmode=tk.MULTIPLE)
        self.selected_column_listbox.pack(fill=tk.BOTH, expand=True)

        # 가운데 버튼
        middle_btn_frame = tk.Frame(list_frame)
        middle_btn_frame.pack(side=tk.LEFT, padx=5)

        tk.Button(middle_btn_frame, text="▶", command=self.add_selected_columns).pack(pady=5)
        tk.Button(middle_btn_frame, text="◀", command=self.remove_selected_columns).pack(pady=5)

        # 초기화
        additional_columns = self.calculator.additional_columns if hasattr(self.calculator, "additional_columns") else []
        selected_columns = self.settings.get('custom_columns', [])
        for col in additional_columns:
            if col not in selected_columns:
                self.available_column_listbox.insert(tk.END, col)
        for col in selected_columns:
            self.selected_column_listbox.insert(tk.END, col)

    def add_selected_columns(self):
        """왼쪽 → 오른쪽으로 이동"""
        selected = self.available_column_listbox.curselection()
        for i in reversed(selected):
            val = self.available_column_listbox.get(i)
            self.available_column_listbox.delete(i)
            self.selected_column_listbox.insert(tk.END, val)

    def remove_selected_columns(self):
        """오른쪽 → 왼쪽으로 이동"""
        selected = self.selected_column_listbox.curselection()
        for i in reversed(selected):
            val = self.selected_column_listbox.get(i)
            self.selected_column_listbox.delete(i)
            self.available_column_listbox.insert(tk.END, val)
    
    def setup_excluded_reps_tab(self):
        """제외 영업담당 설정 탭 구성"""
        # 제외 영업담당 프레임
        frame = tk.Frame(self.excluded_reps_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 설명 레이블
        tk.Label(frame, text="계산에서 제외할 영업담당 목록을 관리합니다.", anchor=tk.W).pack(fill=tk.X, pady=(0, 10))
        
        # 좌우 프레임 나누기
        list_frames = tk.Frame(frame)
        list_frames.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 왼쪽 프레임 (사용 가능한 영업담당)
        available_frame = tk.LabelFrame(list_frames, text="사용 가능한 영업담당")
        available_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # 오른쪽 프레임 (제외된 영업담당)
        excluded_frame = tk.LabelFrame(list_frames, text="제외된 영업담당")
        excluded_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        # 사용 가능한 영업담당 리스트박스
        self.available_listbox = tk.Listbox(available_frame, selectmode=tk.MULTIPLE)
        self.available_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar1 = ttk.Scrollbar(available_frame, orient="vertical", command=self.available_listbox.yview)
        scrollbar1.pack(side=tk.RIGHT, fill=tk.Y)
        self.available_listbox.configure(yscrollcommand=scrollbar1.set)
        
        # 제외된 영업담당 리스트박스
        self.excluded_listbox = tk.Listbox(excluded_frame, selectmode=tk.MULTIPLE)
        self.excluded_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar2 = ttk.Scrollbar(excluded_frame, orient="vertical", command=self.excluded_listbox.yview)
        scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)
        self.excluded_listbox.configure(yscrollcommand=scrollbar2.set)
        
        # 중간 버튼 프레임
        middle_btn_frame = tk.Frame(list_frames)
        middle_btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 이동 버튼들
        move_right_btn = tk.Button(middle_btn_frame, text="►", command=self.move_to_excluded, width=3)
        move_right_btn.pack(pady=5)
        
        move_left_btn = tk.Button(middle_btn_frame, text="◄", command=self.move_to_available, width=3)
        move_left_btn.pack(pady=5)
        
        # 영업담당 목록 초기화
        self.update_sales_reps_lists()
    
    def update_sales_reps_lists(self):
        """사용 가능한 영업담당와 제외된 영업담당 목록 업데이트"""
        # 기존 항목 제거
        self.available_listbox.delete(0, tk.END)
        self.excluded_listbox.delete(0, tk.END)
        
        # 현재 제외된 영업담당 목록
        excluded_reps = set(self.settings.get('excluded_sales_reps', []))
        
        # 영업담당 목록이 있는 경우
        if self.available_sales_reps:
            for rep in sorted(self.available_sales_reps):
                if rep not in excluded_reps:
                    self.available_listbox.insert(tk.END, rep)
        
        # 제외된 영업담당 표시
        for rep in sorted(excluded_reps):
            self.excluded_listbox.insert(tk.END, rep)
    
    def move_to_excluded(self):
        """선택한 영업담당을 제외 목록으로 이동"""
        selected_indices = self.available_listbox.curselection()
        if not selected_indices:
            return
        
        # 선택된 항목들을 가져와서 제외 목록으로 이동
        selected_items = [self.available_listbox.get(i) for i in selected_indices]
        
        # 역순으로 삭제 (인덱스 변경 방지)
        for i in sorted(selected_indices, reverse=True):
            self.available_listbox.delete(i)
        
        # 제외 목록에 추가
        for item in selected_items:
            self.excluded_listbox.insert(tk.END, item)
    
    def move_to_available(self):
        """선택한 영업담당을 사용 가능 목록으로 이동"""
        selected_indices = self.excluded_listbox.curselection()
        if not selected_indices:
            return
        
        # 선택된 항목들을 가져와서 사용 가능 목록으로 이동
        selected_items = [self.excluded_listbox.get(i) for i in selected_indices]
        
        # 역순으로 삭제 (인덱스 변경 방지)
        for i in sorted(selected_indices, reverse=True):
            self.excluded_listbox.delete(i)
        
        # 사용 가능 목록에 추가
        for item in selected_items:
            self.available_listbox.insert(tk.END, item)
    
    def setup_branch_rates_tab(self):
        """지부별 인센티브율 설정 탭 구성"""
        # 지부별 인센티브율 프레임
        frame = tk.Frame(self.branch_rates_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 설명 레이블
        tk.Label(frame, text="지부별 특별 인센티브율을 설정합니다.", anchor=tk.W).pack(fill=tk.X, pady=(0, 10))
        
        # 트리뷰 프레임
        treeview_frame = tk.Frame(frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 지부별 인센티브율 트리뷰
        columns = ('지부명', '1% 적용 시', '2% 적용 시', '4% 적용 시')
        self.branch_rates_tree = ttk.Treeview(treeview_frame, columns=columns, show='headings')
        
        # 열 설정
        for col in columns:
            self.branch_rates_tree.heading(col, text=col)
            self.branch_rates_tree.column(col, width=100)
        
        self.branch_rates_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.branch_rates_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.branch_rates_tree.configure(yscrollcommand=scrollbar.set)
        
        # 기존 지부별 인센티브율 표시
        branch_rates = self.settings.get('branch_rates', {})
        for branch, rates in branch_rates.items():
            self.branch_rates_tree.insert('', tk.END, values=(
                branch, 
                f"{rates.get('1', 1) * 100:.1f}%", 
                f"{rates.get('2', 2) * 100:.1f}%", 
                f"{rates.get('4', 4) * 100:.1f}%"
            ))
        
        # 버튼 프레임
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        # 추가 버튼
        add_btn = tk.Button(btn_frame, text="추가", command=self.add_branch_rate)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # 수정 버튼
        edit_btn = tk.Button(btn_frame, text="수정", command=self.edit_branch_rate)
        edit_btn.pack(side=tk.LEFT, padx=5)
        
        # 제거 버튼
        remove_btn = tk.Button(btn_frame, text="제거", command=self.remove_branch_rate)
        remove_btn.pack(side=tk.LEFT, padx=5)
    
    def setup_default_reps_tab(self):
        """기본 영업담당 설정 탭"""
        frame = tk.Frame(self.default_reps_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 설명 레이블
        tk.Label(frame, text="기본으로 표시할 영업담당을 설정합니다.", anchor=tk.W).pack(fill=tk.X, pady=(0, 10))
        
        # 기본 영업담당 리스트
        list_frame = tk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 영업담당 목록 리스트박스
        self.default_reps_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE)
        self.default_reps_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.default_reps_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.default_reps_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 버튼 프레임
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        # 추가 버튼
        add_btn = tk.Button(btn_frame, text="추가", command=self.add_default_rep)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # 제거 버튼
        remove_btn = tk.Button(btn_frame, text="제거", command=self.remove_default_rep)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # 기존 기본 영업담당 표시
        default_reps = self.settings.get('default_sales_reps', [])
        for rep in default_reps:
            self.default_reps_listbox.insert(tk.END, rep)
    
    def setup_branch_setup_tab(self):
        """지부 설정 탭"""
        frame = tk.Frame(self.branch_setup_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 설명 레이블
        tk.Label(frame, text="지부 목록을 관리합니다.", anchor=tk.W).pack(fill=tk.X, pady=(0, 10))
        
        # 지부 목록 트리뷰
        columns = ('지부명', '설명')
        self.branch_tree = ttk.Treeview(frame, columns=columns, show='headings')
        
        # 열 설정
        for col in columns:
            self.branch_tree.heading(col, text=col)
            self.branch_tree.column(col, width=100)
        
        self.branch_tree.pack(fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.branch_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.branch_tree.configure(yscrollcommand=scrollbar.set)
        
        # 버튼 프레임
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        # 추가 버튼
        add_btn = tk.Button(btn_frame, text="추가", command=self.add_branch)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # 수정 버튼
        edit_btn = tk.Button(btn_frame, text="수정", command=self.edit_branch)
        edit_btn.pack(side=tk.LEFT, padx=5)
        
        # 제거 버튼
        remove_btn = tk.Button(btn_frame, text="제거", command=self.remove_branch)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # 기존 지부 표시
        branches = self.settings.get('branches', {})
        for branch_name, desc in branches.items():
            self.branch_tree.insert('', tk.END, values=(branch_name, desc))
    
    def add_branch_rate(self):
        """지부별 인센티브율 추가"""
        branch_name = simpledialog.askstring("지부 추가", "지부명을 입력하세요:")
        if not branch_name:
            return
        
        # 중복 확인
        for item in self.branch_rates_tree.get_children():
            values = self.branch_rates_tree.item(item, 'values')
            if values[0] == branch_name:
                messagebox.showwarning("중복", f"'{branch_name}'은(는) 이미 목록에 있습니다.")
                return
        
        # 인센티브율 입력
        rate_1 = simpledialog.askfloat("인센티브율 설정", f"'{branch_name}'의 1% 적용 시 인센티브율(%):", minvalue=0, maxvalue=100, initialvalue=1)
        if rate_1 is None:
            return
        
        rate_2 = simpledialog.askfloat("인센티브율 설정", f"'{branch_name}'의 2% 적용 시 인센티브율(%):", minvalue=0, maxvalue=100, initialvalue=2)
        if rate_2 is None:
            return
        
        rate_4 = simpledialog.askfloat("인센티브율 설정", f"'{branch_name}'의 4% 적용 시 인센티브율(%):", minvalue=0, maxvalue=100, initialvalue=4)
        if rate_4 is None:
            return
        
        # 트리뷰에 추가
        self.branch_rates_tree.insert('', tk.END, values=(
            branch_name, 
            f"{rate_1:.1f}%", 
            f"{rate_2:.1f}%", 
            f"{rate_4:.1f}%"
        ))
    
    def edit_branch_rate(self):
        """지부별 인센티브율 수정"""
        selected = self.branch_rates_tree.selection()
        if not selected:
            messagebox.showinfo("알림", "수정할 항목을 선택하세요.")
            return
        
        item = selected[0]
        values = self.branch_rates_tree.item(item, 'values')
        branch_name = values[0]
        
        # 기존 값에서 %를 제거하고 숫자만 추출
        current_rate_1 = float(values[1].rstrip('%'))
        current_rate_2 = float(values[2].rstrip('%'))
        current_rate_4 = float(values[3].rstrip('%'))
        
        # 인센티브율 수정
        rate_1 = simpledialog.askfloat("인센티브율 설정", f"'{branch_name}'의 1% 적용 시 인센티브율(%):", 
                                       minvalue=0, maxvalue=100, initialvalue=current_rate_1)
        if rate_1 is None:
            return
        
        rate_2 = simpledialog.askfloat("인센티브율 설정", f"'{branch_name}'의 2% 적용 시 인센티브율(%):", 
                                       minvalue=0, maxvalue=100, initialvalue=current_rate_2)
        if rate_2 is None:
            return
        
        rate_4 = simpledialog.askfloat("인센티브율 설정", f"'{branch_name}'의 4% 적용 시 인센티브율(%):", 
                                       minvalue=0, maxvalue=100, initialvalue=current_rate_4)
        if rate_4 is None:
            return
        
        # 트리뷰 업데이트
        self.branch_rates_tree.item(item, values=(
            branch_name, 
            f"{rate_1:.1f}%", 
            f"{rate_2:.1f}%", 
            f"{rate_4:.1f}%"
        ))
    
    def remove_branch_rate(self):
        """지부별 인센티브율 제거"""
        selected = self.branch_rates_tree.selection()
        if not selected:
            messagebox.showinfo("알림", "제거할 항목을 선택하세요.")
            return
        
        # 선택한 항목 제거
        for item in selected:
            self.branch_rates_tree.delete(item)
    
    def add_default_rep(self):
        """기본 영업담당 추가"""
        rep_name = simpledialog.askstring("영업담당 추가", "추가할 영업담당 이름을 입력하세요:")
        if not rep_name:
            return
        
        # 중복 확인
        for i in range(self.default_reps_listbox.size()):
            if self.default_reps_listbox.get(i) == rep_name:
                messagebox.showwarning("중복", f"'{rep_name}'은(는) 이미 목록에 있습니다.")
                return
        
        # 리스트박스에 추가
        self.default_reps_listbox.insert(tk.END, rep_name)
    
    def remove_default_rep(self):
        """기본 영업담당 제거"""
        selected_indices = self.default_reps_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("알림", "제거할 영업담당을 선택하세요.")
            return
        
        # 선택된 항목들을 역순으로 제거 (인덱스 변경 방지)
        for i in sorted(selected_indices, reverse=True):
            self.default_reps_listbox.delete(i)
    
    def add_branch(self):
        """지부 추가"""
        branch_name = simpledialog.askstring("지부 추가", "지부명을 입력하세요:")
        if not branch_name:
            return
        
        # 중복 확인
        for item in self.branch_tree.get_children():
            values = self.branch_tree.item(item, 'values')
            if values[0] == branch_name:
                messagebox.showwarning("중복", f"'{branch_name}'은(는) 이미 목록에 있습니다.")
                return
        
        # 설명 입력
        description = simpledialog.askstring("지부 설명", f"'{branch_name}'의 설명을 입력하세요:")
        if description is None:
            description = ""
        
        # 트리뷰에 추가
        self.branch_tree.insert('', tk.END, values=(branch_name, description))
    
    def edit_branch(self):
        """지부 수정"""
        selected = self.branch_tree.selection()
        if not selected:
            messagebox.showinfo("알림", "수정할 지부를 선택하세요.")
            return
        
        item = selected[0]
        values = self.branch_tree.item(item, 'values')
        branch_name = values[0]
        current_desc = values[1]
        
        # 새 설명 입력
        new_desc = simpledialog.askstring("지부 설명 수정", f"'{branch_name}'의 새 설명을 입력하세요:", initialvalue=current_desc)
        if new_desc is None:
            return
        
        # 트리뷰 업데이트
        self.branch_tree.item(item, values=(branch_name, new_desc))
    
    def remove_branch(self):
        """지부 제거"""
        selected = self.branch_tree.selection()
        if not selected:
            messagebox.showinfo("알림", "제거할 지부를 선택하세요.")
            return
        
        # 선택한 항목 제거
        for item in selected:
            self.branch_tree.delete(item)
    
    def save_settings(self):
        """설정 저장"""
        # 제외 영업담당 설정 저장
        excluded_reps = list(self.excluded_listbox.get(0, tk.END))
        self.settings['excluded_sales_reps'] = excluded_reps
        
        # 지부별 인센티브율 설정 저장
        branch_rates = {}
        for item in self.branch_rates_tree.get_children():
            values = self.branch_rates_tree.item(item, 'values')
            branch_name = values[0]
            rate_1 = float(values[1].rstrip('%')) / 100
            rate_2 = float(values[2].rstrip('%')) / 100
            rate_4 = float(values[3].rstrip('%')) / 100
            
            branch_rates[branch_name] = {
                '1': rate_1,
                '2': rate_2,
                '4': rate_4
            }
        
        self.settings['branch_rates'] = branch_rates
        
        # 기본 영업담당 설정 저장
        default_reps = list(self.default_reps_listbox.get(0, tk.END))
        self.settings['default_sales_reps'] = default_reps
        
        # 지부 설정 저장
        branches = {}
        for item in self.branch_tree.get_children():
            values = self.branch_tree.item(item, 'values')
            branch_name = values[0]
            description = values[1]
            branches[branch_name] = description
        
        self.settings['branches'] = branches
        
        # custom_columns 저장
        selected_columns = list(self.selected_column_listbox.get(0, tk.END))
        self.settings['custom_columns'] = selected_columns

        # 부모 윈도우에 설정 전달
        self.calculator.update_settings(self.settings)
        
        messagebox.showinfo("알림", "설정이 저장되었습니다.")
        self.destroy()
    
class IncentiveCalculator:
    """인센티브 계산기 메인 클래스"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("매출/수금 인센티브 계산기")
        self.root.geometry("1000x700")
        
        # 열 이름 매핑 (변경된 이름 적용)
        self.column_mapping = {
            '접수일자': '매출일자',
            '입금일': '수금일자',
            '영업담당': '영업담당',
            '수수료': '금액',
            '입금액': '입금액',
            '차감액': '차감액'
        }
        self.reverse_mapping = {v: k for k, v in self.column_mapping.items()}  # 역방향 매핑
        
        # 팀 정의 - 특정 영업담당들을 하나의 팀으로 그룹화
        self.teams = {
            "서울센터": ["오세중", "조봉현", "오석현", "장동주"],
            "본사": ["본사접수"],
            "마케팅서": ["마케팅"],
            "경북센터": ["엄상흠"]
        }
        
        # 결과, 파일 경로 초기화
        self.result_df = None
        self.df = None
        self.file_paths = []  # 선택된 파일 경로들을 저장
        self.sales_rep_results = {}  # 영업담당별 결과 저장
        self.team_results = {}       # 팀별 결과 저장
        
        # 연도 및 월 리스트
        self.available_years = list(range(datetime.now().year - 5, datetime.now().year + 2))
        self.available_months = list(range(1, 13))
        
        # 설정값 초기화 (기본 설정)
        self.settings = {
            'excluded_sales_reps': [],
            'branch_rates': {},
            'default_sales_reps': ["오세중", "조봉현", "오석현", "장동주", "엄상흠"],
            'branches': {
                "서울센터": "서울 지역 담당",
                "본사": "본사 직접 영업",
                "경북센터": "경북 지역 담당"
            },
            'reference_year': datetime.now().year,
            'reference_month': datetime.now().month
        }
        
        # 기본 컬럼 정의
        self.default_columns = ['시험분야', '검사목적', '영업담당', '접수일자', '결재상태', '지부명', '거래처', '거래처 주소',
                                '제품/시료명', '식품종류', '공급가액', '세액', '수수료', '입금여부', '입금일',
                                '입금액', '차감율(%)', '차감액', '인센티브율', '인센티브금액', '잔액', '금액 조정 분', '시험기간']
        self.additional_columns = []
        
        # 설정 파일 로드
        self.load_settings()
        
        # 영업담당 초기화
        self.all_sales_reps = []
        default_reps = self.settings.get('default_sales_reps', [])
        if default_reps:
            self.all_sales_reps = list(default_reps)
        
        # 여기서 self.update_teams_from_settings() 호출 - 올바른 위치
        self.update_teams_from_settings()
        
        # 기준 연월 StringVar (UI 연동용)
        self.year_var = StringVar()
        self.month_var = StringVar()
        
        # UI 구성
        self.setup_ui()
    
    def load_settings(self):
        """설정 파일 로드"""
        try:
            if os.path.exists(file_path):  # 여기를 파일 경로 변수로 변경
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    # 기존 설정을 새로 로드한 설정으로 업데이트 (기본값 유지)
                    for key, value in loaded_settings.items():
                        self.settings[key] = value
                    
            # 팀 정의 업데이트
            self.update_teams_from_settings()
        except Exception as e:
            messagebox.showerror("설정 로드 오류", f"설정 파일을 로드하는 중 오류가 발생했습니다: {str(e)}")
    
    def save_settings(self):
        """설정 파일 저장"""
        try:
            # 디렉토리 확인
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"설정 저장 오류: {str(e)}")  # 콘솔에도 오류 출력
            messagebox.showerror("설정 저장 오류", f"설정 파일을 저장하는 중 오류가 발생했습니다: {str(e)}")
            
    def update_settings(self, new_settings):
        """설정 업데이트 및 저장"""
        self.settings.update(new_settings)
        try:
            self.save_settings()
        except:
            messagebox.showwarning("경고", "설정을 저장하는 데 실패했습니다. 프로그램은 계속 실행됩니다.")
    
    # 팀 정의 업데이트 (기본 영업담당 설정 변경 시)
    # self.update_teams_from_settings()  # 이 부분을 주석 처리 또는 삭제
    
    def update_teams_from_settings(self):
        """설정에서 팀 정의 업데이트"""
        # 기존 팀 정의 유지하면서 추가/업데이트
        branches = self.settings.get('branches', {})
        default_reps = self.settings.get('default_sales_reps', [])
        
        # 팀 정의 업데이트 (지부와 기본 영업담당 연결)
        for branch_name in branches.keys():
            if branch_name not in self.teams:
                # 새 지부는 빈 목록으로 초기화
                self.teams[branch_name] = []
        
        # 팀 멤버 자동 할당 시도 (기본 영업담당 중 팀에 할당되지 않은 영업담당을 찾아서 할당)
        assigned_reps = set()
        for team_members in self.teams.values():
            assigned_reps.update(team_members)
    
    def get_sales_reps_list(self):
        """영업담당 목록 반환 (설정 대화상자용)"""
        # 파일에서 로드된 영업담당과 기본 설정된 영업담당을 합쳐서 반환
        all_reps = set(self.all_sales_reps)
        default_reps = set(self.settings.get('default_sales_reps', []))
        return sorted(list(all_reps.union(default_reps)))
    
    def open_settings(self):
        """설정 다이얼로그 오픈"""
        try:
            if self.df is None or self.df.empty:
                messagebox.showwarning("경고", "먼저 엑셀 파일을 로드하세요.")
                return
            
            # 추가 컬럼 목록 갱신
            self.additional_columns = [col for col in self.df.columns if col not in self.default_columns]
            print("추가 컬럼 목록:", self.additional_columns)  # 디버깅용
            
            # 설정 다이얼로그 생성 및 표시
            settings_dialog = SettingsDialog(self, self.settings)
            settings_dialog.transient(self.root)  # 부모 창 지정
            settings_dialog.grab_set()            # 모달 창
            settings_dialog.lift()                # 최상위로 올림
        except Exception as e:
            print("설정 대화상자 오류:", str(e))
            traceback.print_exc()
            messagebox.showerror("오류", f"설정 창을 여는 중 오류가 발생했습니다: {str(e)}")
    
    def setup_ui(self):
        """UI 구성"""
        # 메뉴 바 생성
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)
        
        # 파일 메뉴
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="엑셀 파일 선택", command=self.select_files)
        file_menu.add_separator()
        file_menu.add_command(label="결과 저장", command=self.save_result)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        
        # 설정 메뉴
        settings_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="설정", menu=settings_menu)
        settings_menu.add_command(label="환경설정", command=self.open_settings)
        
        # 상단 프레임
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 파일 선택 버튼
        self.select_btn = tk.Button(top_frame, text="엑셀 파일 선택", command=self.select_files)
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        # 파일 목록 프레임
        self.files_frame = tk.Frame(self.root)
        self.files_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 파일 목록 레이블
        tk.Label(self.files_frame, text="선택된 파일:").pack(anchor=tk.W)
        
        # 파일 목록 표시 영역 (스크롤 가능)
        self.files_listbox_frame = tk.Frame(self.files_frame)
        self.files_listbox_frame.pack(fill=tk.X, pady=5)
        
        self.files_listbox = tk.Listbox(self.files_listbox_frame, height=3)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 파일 목록 스크롤바
        files_scrollbar = ttk.Scrollbar(self.files_listbox_frame, orient="vertical", command=self.files_listbox.yview)
        files_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.files_listbox.configure(yscrollcommand=files_scrollbar.set)
        
        # 파일 제거 버튼
        self.remove_file_btn = tk.Button(self.files_frame, text="선택 파일 제거", command=self.remove_selected_file)
        self.remove_file_btn.pack(anchor=tk.W, pady=5)
        
        # 중간 프레임 (기준월 선택)
        mid_frame = tk.Frame(self.root)
        mid_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 기준월 선택 레이블
        tk.Label(mid_frame, text="기준연월:").pack(side=tk.LEFT, padx=5)
        
        # 연도 선택 콤보박스
        self.year_var = StringVar()
        self.year_dropdown = ttk.Combobox(mid_frame, textvariable=self.year_var, state="readonly", width=6)
        self.year_dropdown['values'] = self.available_years
        self.year_dropdown.pack(side=tk.LEFT, padx=2)
        self.year_var.set(str(self.settings.get('reference_year', datetime.now().year)))
        
        tk.Label(mid_frame, text="년").pack(side=tk.LEFT)
        
        # 월 선택 콤보박스
        self.month_var = StringVar()
        self.month_dropdown = ttk.Combobox(mid_frame, textvariable=self.month_var, state="readonly", width=4)
        self.month_dropdown['values'] = self.available_months
        self.month_dropdown.pack(side=tk.LEFT, padx=2)
        self.month_var.set(str(self.settings.get('reference_month', datetime.now().month)))
        
        tk.Label(mid_frame, text="월").pack(side=tk.LEFT, padx=(0, 10))
        
        # 계산 버튼
        self.calc_btn = tk.Button(mid_frame, text="인센티브 계산", command=self.calculate_incentive_safe)
        self.calc_btn.pack(side=tk.LEFT, padx=10)
        
        # 탭 분리 버튼
        self.split_tabs_btn = tk.Button(mid_frame, text="탭 분리", command=self.split_tabs_by_sales_rep)
        self.split_tabs_btn.pack(side=tk.LEFT, padx=5)
        
        # 탭 저장 버튼
        self.save_tabs_btn = tk.Button(mid_frame, text="탭 저장", command=self.save_tabs)
        self.save_tabs_btn.pack(side=tk.LEFT, padx=5)
        
        # 모든 저장 버튼
        self.save_btn = tk.Button(mid_frame, text="모두 저장", command=self.save_result)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # 탭 컨트롤 생성
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
       
        # 전체 결과 탭
        self.all_results_frame = tk.Frame(self.notebook)
        self.notebook.add(self.all_results_frame, text="전체 결과")
        
        # 버튼 프레임 추가
        button_frame = tk.Frame(self.all_results_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 엑셀 변환 버튼 추가 - 여기를 수정하세요
        export_btn = tk.Button(button_frame, text="엑셀로 변환", command=self.export_current_view_to_excel)
        export_btn.pack(side=tk.RIGHT, padx=5, pady=5)
    
        # 트리뷰 + 스크롤바를 포함할 프레임 생성
        tree_frame = tk.Frame(self.all_results_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # 트리뷰 생성
        self.all_tree = ttk.Treeview(tree_frame)
        self.all_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 수직 스크롤바
        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.all_tree.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.all_tree.configure(yscrollcommand=scrollbar_y.set)

        # 수평 스크롤바 (아래쪽에 배치해야 하므로 self.all_results_frame에 직접 추가)
        scrollbar_x = ttk.Scrollbar(self.all_results_frame, orient="horizontal", command=self.all_tree.xview)
        scrollbar_x.pack(fill=tk.X)
        self.all_tree.configure(xscrollcommand=scrollbar_x.set)

    def select_files(self):
        """엑셀 파일 선택"""
        file_paths = filedialog.askopenfilenames(
            title="엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls")]
        )
        if file_paths:
            self.file_paths = list(file_paths)
            # 파일 목록 리스트박스에 표시
            self.files_listbox.delete(0, tk.END)
            for path in self.file_paths:
                self.files_listbox.insert(tk.END, os.path.basename(path))
            self.load_files_and_update_months()
    
    def remove_selected_file(self):
        """선택된 파일 제거"""
        selected_indices = self.files_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("알림", "제거할 파일을 선택하세요.")
            return
        
        # 선택된 항목들을 역순으로 제거 (인덱스 변경 방지)
        for i in sorted(selected_indices, reverse=True):
            del self.file_paths[i]
            self.files_listbox.delete(i)
        
        # 파일 목록이 변경되었으므로 월 목록 업데이트
        if self.file_paths:
            self.load_files_and_update_months()
        else:
            # 모든 파일이 제거되었을 경우 초기화
            self.df = None
            self.all_sales_reps = list(self.settings.get('default_sales_reps', []))  # 기본 영업담당으로 초기화
    
    def load_files_and_update_months(self):
        """여러 엑셀 파일을 로드하고 월 목록을 업데이트합니다."""
        if not self.file_paths:
            messagebox.showerror("오류", "선택된 파일이 없습니다.")
            return
        
        try:
            # 데이터프레임 목록
            dfs = []
            
            # 각 파일 로드
            for file_path in self.file_paths:
                try:
                    temp_df = pd.read_excel(file_path)
                    
                    # 필요한 열이 있는지 확인 (변경된 열 이름 사용)
                    required_columns = [self.reverse_mapping.get(col, col) for col in ['매출일자', '수금일자', '금액', '입금액', '지부명', '영업담당']]
                    missing_columns = [col for col in required_columns if col not in temp_df.columns]
                    
                    if missing_columns:
                        messagebox.showerror("오류", f"'{os.path.basename(file_path)}'에 다음 열이 없습니다: {', '.join(missing_columns)}")
                        return
                    
                    # 열 이름 변경 (내부적으로는 원래 이름 사용)
                    col_mapping = {k: v for k, v in self.column_mapping.items() if k in temp_df.columns}
                    temp_df = temp_df.rename(columns=col_mapping)
                    
                    # 날짜 열 형식 변환 - 오류 처리 강화
                    # 1. 날짜 컬럼이 이미 datetime 타입인지 확인
                    # 2. 문자열인 경우 다양한 형식 시도
                    # 3. 변환 실패 시 오류 대신 NaT(Not a Time) 사용
                    date_columns = ['매출일자', '수금일자']
                    for date_col in date_columns:
                        if date_col in temp_df.columns:
                            # 이미 datetime 타입인지 확인
                            if not pd.api.types.is_datetime64_any_dtype(temp_df[date_col]):
                                try:
                                    # errors='coerce'는 변환 실패 시 NaT로 설정
                                    temp_df[date_col] = pd.to_datetime(temp_df[date_col], errors='coerce')
                                    
                                    # NaT 값이 있는지 확인하고 사용자에게 알림
                                    nat_count = temp_df[date_col].isna().sum()
                                    if nat_count > 0:
                                        print(f"경고: '{date_col}' 열에서 {nat_count}개의 날짜 형식이 잘못되었습니다.")
                                except Exception as e:
                                    # 변환 중 예외 발생 시 처리
                                    print(f"날짜 변환 오류 ({date_col}): {str(e)}")
                                    messagebox.showwarning("날짜 형식 경고", 
                                                        f"'{date_col}' 열의 날짜 형식을 변환하는 중 문제가 발생했습니다.\n"
                                                        f"일부 데이터가 올바르게 처리되지 않을 수 있습니다.")
                    
                    # 데이터프레임 목록에 추가
                    dfs.append(temp_df)
                    
                except Exception as e:
                    messagebox.showerror("오류", f"'{os.path.basename(file_path)}' 파일 로드 중 오류: {str(e)}")
                    return
            
            # 모든 데이터프레임 병합
            if dfs:
                self.df = pd.concat(dfs, ignore_index=True)
                
                # 추가 컬럼 목록 갱신
                self.additional_columns = [col for col in self.df.columns if col not in self.default_columns]
                
                # 영업담당 목록 추출 (설정에서 사용)
                # 기본 영업담당과 파일에서 추출한 영업담당 합치기
                file_sales_reps = sorted(self.df['영업담당'].dropna().unique().tolist())
                default_sales_reps = self.settings.get('default_sales_reps', [])
                self.all_sales_reps = sorted(list(set(file_sales_reps + default_sales_reps)))
                
                messagebox.showinfo("알림", f"{len(self.file_paths)}개 파일이 로드되었습니다. 기준월을 선택하고 인센티브를 계산하세요.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 로드 중 오류가 발생했습니다: {str(e)}")
     
    def calculate_incentive_safe(self):
        """예외 처리를 강화한 인센티브 계산 함수"""
        try:
            print("계산 시작")  # 디버깅용 메시지
            if self.df is None or self.df.empty:
                messagebox.showerror("오류", "먼저 엑셀 파일을 선택해주세요.")
                return False

            # 열 이름 불일치 문제 확인
            self.check_column_mappings()

            # 기준 연월 가져오기
            try:
                reference_year = int(self.year_var.get())
                reference_month = int(self.month_var.get())
            except ValueError:
                messagebox.showerror("오류", "유효한 연도와 월을 선택해주세요.")
                return False

            # 설정에 기준 연월 저장
            self.settings['reference_year'] = reference_year
            self.settings['reference_month'] = reference_month
            self.save_settings()

            # 기준월 형식 변환 (YYYY-MM)
            reference_month_str = f"{reference_year}-{reference_month:02d}"
            reference_date = pd.to_datetime(reference_month_str + '-01')

            # 제외할 영업담당 필터링
            excluded_reps = self.settings.get('excluded_sales_reps', [])

            # 원본 데이터 복사
            df_temp = self.df.copy()

            # 0원 또는 누락 데이터 제거
            cols_to_check = ['공급가액', '세액', '금액']
            if all(col in df_temp.columns for col in cols_to_check):
                df_temp = df_temp[~(
                    (df_temp['공급가액'].fillna(0) == 0) &
                    (df_temp['세액'].fillna(0) == 0) &
                    (df_temp['금액'].fillna(0) == 0)
                )].copy()

            # 필수 열 확인
            required_columns = ['매출일자', '수금일자', '금액', '입금액', '영업담당']
            missing_columns = [col for col in required_columns if col not in df_temp.columns]
            if missing_columns:
                messagebox.showerror("오류", f"데이터에 다음 필수 열이 없습니다: {', '.join(missing_columns)}")
                return False

            # 제외 리스트 적용
            if excluded_reps:
                df_temp = df_temp[~df_temp['영업담당'].isin(excluded_reps)].copy()

            # 데이터 타입 정리
            df_filtered = self.fix_data_types(df_temp)

            # 기준월 이전 데이터 필터링
            mask_filtered = (
                (~pd.isna(df_filtered['매출일자']) & (df_filtered['매출일자'] < reference_date)) &
                (~pd.isna(df_filtered['수금일자']) & (df_filtered['수금일자'] < reference_date))
            )
            df_filtered = df_filtered[~mask_filtered].copy()

            # 월차이 계산
            df_filtered['월차이'] = np.nan
            mask = ~pd.isna(df_filtered['수금일자'])
            df_filtered.loc[mask, '월차이'] = (
                (df_filtered.loc[mask, '수금일자'].dt.year - df_filtered.loc[mask, '매출일자'].dt.year) * 12 +
                (df_filtered.loc[mask, '수금일자'].dt.month - df_filtered.loc[mask, '매출일자'].dt.month)
            )

            # 차감액 초기화
            if '차감액' not in df_filtered.columns:
                df_filtered['차감액'] = 0.0

            # 인센티브 계산 및 정렬
            df_filtered = self.calculate_incentive_with_deduction(df_filtered, reference_date)
            df_filtered = df_filtered.sort_values(
                by=['인센티브율', '매출일자'],
                ascending=[True, True]
            ).copy()
            # 월 컬럼 복원 (팀/사원별 그룹화 용)
            df_filtered['매출월'] = df_filtered['매출일자'].dt.strftime('%Y-%m')

            # 인센티브율별 그룹화 및 소계/총계 생성
            grouped = df_filtered.groupby('인센티브율', sort=False)
            result_dfs = []
            for rate, group in grouped:
                result_dfs.append(group.copy())
                subtotal = pd.DataFrame({
                    '인센티브율':    [rate],
                    '금액':         [group['금액'].sum()],
                    '입금액':       [group['입금액'].sum()],
                    '인센티브금액':  [group['인센티브금액'].sum()],
                    '차감액':       [group['차감액'].sum()],
                    '매출월':       [None],
                    '매출일자':     [None],
                    '수금일자':     [None],
                    '지부명':       ["소계"],
                    '영업담당':     [None],
                    '월차이':       [None],
                })
                result_dfs.append(subtotal)
            total = pd.DataFrame({
                '인센티브율':    [None],
                '금액':         [df_filtered['금액'].sum()],
                '입금액':       [df_filtered['입금액'].sum()],
                '인센티브금액':  [df_filtered['인센티브금액'].sum()],
                '차감액':       [df_filtered['차감액'].sum()],
                '지부명':       ["총계"],
                '매출월':       [None],
                '매출일자':     [None],
                '수금일자':     [None],
                '영업담당':     [None],
                '월차이':       [None],
            })
            non_empty_dfs = [df for df in result_dfs if not df.empty]
            self.result_df = pd.concat(non_empty_dfs + [total], ignore_index=True)

            # 팀/영업담당별 결과 생성 및 표시
            self.create_team_and_sales_rep_results(df_filtered)
            self.debug_dataframe(self.result_df, "최종 결과")
            self.display_results()
            return True

        except Exception as e:
            error_msg = f"계산 중 오류가 발생했습니다: {str(e)}"
            print(error_msg)  # 콘솔에 오류 출력
            print(traceback.format_exc())  # 상세 오류 스택 출력
            messagebox.showerror("오류", error_msg)
            return False
    
    def check_column_mappings(self):
        """열 이름 매핑을 확인하고 필요한 경우 수정"""
        print("현재 column_mapping:", self.column_mapping)
        print("현재 reverse_mapping:", self.reverse_mapping)
        
        # 데이터프레임의 실제 열 목록 확인
        if self.df is not None:
            print("데이터프레임 열 목록:", self.df.columns.tolist())
        
        # 매핑 업데이트 (필요한 경우)
        self.column_mapping = {
            '접수일자': '매출일자',  # UI에 표시할 이름: 데이터프레임 내부 이름
            '입금일': '수금일자',
            '영업담당': '영업담당',
            '수수료': '금액',  # '금액'을 '수수료'로 변경
            '입금액': '입금액',  # 인센티브 계산 기준으로 사용할 새 열
            '차감율': '차감율',
            '차감액': '차감액'   # 차감액 열 추가
        }
        self.reverse_mapping = {v: k for k, v in self.column_mapping.items()}
        
        print("업데이트된 column_mapping:", self.column_mapping)
        print("업데이트된 reverse_mapping:", self.reverse_mapping)
        
        return True
    
    # 정렬 조건 개선 - '연구용역'과 '지부명'을 포함한 정렬

    # calculate_advanced_incentive 메서드 수정
    # 정렬 조건 수정 - 인센티브율 및 매출일 기준 정렬

    def calculate_advanced_incentive(self, df_filtered, reference_date):
        """
        고급 인센티브 계산 함수
        1. 인센티브의 우선 적용
        2. 특정 지부 맞춤 인센티브 규칙 적용
        3. 상세 소계 및 그룹화 로직 구현
        """
        # 필수 열 초기화
        required_columns = [
            '인센티브율', '인센티브금액', '차감액', '접수일자', '수금일자', 
            '금액', '입금액', '지부명', '검사목적'
        ]
        for col in required_columns:
            if col not in df_filtered.columns:
                df_filtered[col] = np.nan

        # 특별 지부 인센티브 규칙 로드
        branch_rates = self.settings.get('branch_rates', {})
        
        # 등록된 지부 목록 가져오기
        registered_branches = self.settings.get('branches', {}).keys()
        
        def calculate_branch_specific_incentive(row):
            """특정 지부 맞춤 인센티브 계산"""
            branch_name = str(row.get('지부명', '')).strip()
            
            # 등록된 지부 확인
            is_registered_branch = branch_name in registered_branches
            
            # 특정 지부 맞춤 규칙 우선 적용
            if branch_name in branch_rates:
                branch_rate = branch_rates[branch_name]
                return max(branch_rate.values()) if branch_rate else default_incentive_calculation(row, is_registered_branch)
            
            # 기본 인센티브 계산 로직
            return default_incentive_calculation(row, is_registered_branch)

        def default_incentive_calculation(row, is_registered_branch=False):
            """기본 인센티브 계산 로직"""
            # 수금일자에 따른 인센티브율 계산
            if pd.isna(row['수금일자']):
                # 미수금 처리 로직
                return 0.0
            
            # 월차이 계산
            month_diff = (row['수금일자'].year - row['접수일자'].year) * 12 + \
                        (row['수금일자'].month - row['접수일자'].month)
            
            # 인센티브율 결정 로직
            if is_registered_branch:
                # 등록된 지부는 월차이와 상관없이 기본 인센티브 적용
                if month_diff <= 0:
                    return 0.04
                elif month_diff == 1:
                    return 0.02
                else:  # 2개월 이상
                    return 0.01
            else:
                # 일반 지부는 기존 규칙 적용
                if month_diff < 0:  # 선수금 처리
                    return 0.04
                elif month_diff == 0:
                    return 0.04
                elif month_diff == 1:
                    return 0.02
                elif month_diff == 2:
                    return 0.01
                else:
                    return 0.0

        # 인센티브율 및 금액 계산
        df_filtered['인센티브율'] = df_filtered.apply(calculate_branch_specific_incentive, axis=1)
        df_filtered['인센티브금액'] = df_filtered['입금액'] * df_filtered['인센티브율']

        # 디버깅 코드
        print(f"필터링 전 데이터 수: {len(df_filtered)}")
        print("필터링 전 인센티브율이 0인 항목:")
        print(df_filtered[df_filtered['인센티브율'] == 0][['지부명', '접수일자', '인센티브율']])

        # 인센티브가 0인 데이터 제외
        df_filtered = df_filtered[df_filtered['인센티브율'] > 0].copy()

        # 디버깅 코드
        print(f"필터링 후 데이터 수: {len(df_filtered)}")
        print("필터링 후 인센티브율 목록:")
        print(df_filtered['인센티브율'].unique())

        # 고급 정렬: 인센티브율 순 + 각 인센티브율 그룹 내 날짜 순 정렬
        df_filtered_sorted = df_filtered.sort_values(
            by=['인센티브율', '접수일자'], 
            ascending=[True, True]  # 1% → 2% → 4%, 날짜는 오름차순
        )

        # 인센티브율별 그룹화 및 소계 생성 (1%, 2%, 4% 순서)
        result_dfs = []
        
        for rate in [0.01, 0.02, 0.04]:
            rate_group = df_filtered_sorted[df_filtered_sorted['인센티브율'] == rate]
            
            if not rate_group.empty:
                # 원본 데이터 추가
                result_dfs.append(rate_group)
                
                # 소계 행 생성
                subtotal = pd.DataFrame({
                    '접수일자': [None],
                    '수금일자': [None],
                    '지부명': [f"{rate*100:.0f}% 소계"],
                    '인센티브율': [rate],
                    '금액': [rate_group['금액'].sum()],
                    '입금액': [rate_group['입금액'].sum()],
                    '인센티브금액': [rate_group['인센티브금액'].sum()],
                    '차감액': [rate_group['차감액'].sum()]
                }, index=[0])
                
                result_dfs.append(subtotal)

        # 최종 데이터프레임 생성
        if result_dfs:
            final_df = pd.concat(result_dfs, ignore_index=True)
            
            # 총계 행 추가
            total_row = pd.DataFrame({
                '접수일자': [None],
                '수금일자': [None],
                '지부명': ['총계'],
                '인센티브율': [None],
                '금액': [final_df['금액'].sum()],
                '입금액': [final_df['입금액'].sum()],
                '인센티브금액': [final_df['인센티브금액'].sum()],
                '차감액': [final_df['차감액'].sum()]
            }, index=[0])
            
            final_df = pd.concat([final_df, total_row], ignore_index=True)
            
            return final_df
        
        return pd.DataFrame()  # 데이터가 없을 경우 빈 데이터프레임 반환

    def calculate_incentive_with_deduction(self, df_filtered, reference_date):
        """
        인센티브 계산 (차감액 처리 포함)
        - 접수월 기준 4~5개월째에 입금액이 수수료와 다른 경우 차감 적용
        """
        branch_rates = self.settings.get('branch_rates', {})

        # 필요한 컬럼 보장
        for col in ['인센티브율', '인센티브금액', '차감액']:
            if col not in df_filtered.columns:
                df_filtered[col] = 0.0

        # '검사목적' 컬럼이 없으면 빈 문자열로
        if '검사목적' not in df_filtered.columns:
            df_filtered['검사목적'] = ''

        # 각 행 반복
        for idx, row in df_filtered.iterrows():
            # 1. 입금 여부 및 입금액/수수료 비교
            # 입금액이 있고 0보다 큰 경우
            has_payment = not pd.isna(row['입금액']) and float(row['입금액'] or 0) > 0
            
            # 입금액과 수수료(금액)가 동일한지 확인
            fee = float(row['금액'] or 0)
            payment = float(row['입금액'] or 0) if has_payment else 0
            is_fully_paid = abs(payment - fee) < 0.01  # 부동소수점 오차 고려
            
            # 2. 월차이 계산 - 기준월을 포함하여 계산
            month_diff = 0
            if not pd.isna(row['매출일자']):
                # 경우 1: 매출년도와 기준년도가 같은 경우
                if reference_date.year == row['매출일자'].year:
                    # 기준월 - 매출월 + 1 (기준월 포함)
                    month_diff = reference_date.month - row['매출일자'].month + 1
                # 경우 2: 연도가 넘어가는 경우
                else:
                    # (12 - 매출월) + 기준월 + 1 (기준월 포함)
                    month_diff = (12 - row['매출일자'].month) + reference_date.month + 1
                    
                    # 2개년도 이상 차이나는 경우 추가 계산
                    year_diff = reference_date.year - row['매출일자'].year - 1
                    if year_diff > 0:
                        month_diff += year_diff * 12
            
            # 3. 차감 대상 여부 확인 
            # - 4~5개월 차이이고 
            # - 입금액과 수수료가 다른 경우(미입금 포함)
            is_deduction_range = (4 <= month_diff <= 5)
            is_deduction_target = is_deduction_range and not is_fully_paid
            
            if is_deduction_target:
                # 차감 대상인 경우
                if not has_payment:  
                    # 미입금인 경우 - 수수료의 3%를 차감
                    deduction_amount = fee * 0.03
                else:  
                    # 일부만 입금된 경우 - (수수료 - 입금액)을 차감
                    deduction_amount = fee - payment
                    # 음수가 나오면 0으로 처리 (입금액이 수수료보다 많은 경우)
                    deduction_amount = max(0, deduction_amount)
                    
                df_filtered.at[idx, '차감액'] = deduction_amount
                
                # 미입금인 경우 인센티브 없음
                if not has_payment:
                    df_filtered.at[idx, '인센티브율'] = 0.0
                    df_filtered.at[idx, '인센티브금액'] = 0.0
                    continue
            else:
                # 차감 대상이 아닌 경우
                df_filtered.at[idx, '차감액'] = 0.0
            
            # 4. 입금된 경우 인센티브 계산 (차감 여부와 무관)
            if has_payment:
                # 수금일자와 매출일자 간의 월차이 (인센티브율 계산용)
                sales_payment_diff = 0
                if not pd.isna(row['수금일자']) and not pd.isna(row['매출일자']):
                    # 같은 년도인 경우
                    if row['수금일자'].year == row['매출일자'].year:
                        sales_payment_diff = row['수금일자'].month - row['매출일자'].month
                    else:
                        # 연도가 넘어가는 경우
                        sales_payment_diff = (12 - row['매출일자'].month) + row['수금일자'].month
                        
                        # 2개년도 이상 차이나는 경우
                        year_diff = row['수금일자'].year - row['매출일자'].year - 1
                        if year_diff > 0:
                            sales_payment_diff += year_diff * 12
                
                branch_name = str(row.get('지부명', '')).strip()
                purpose = str(row.get('검사목적', '')).strip()
                
                # 인센티브율 결정
                if purpose == '연구용역':
                    rate = 0.02
                else:
                    # 기본율(월차이)
                    if sales_payment_diff <= 0:
                        base = 0.04
                    elif sales_payment_diff == 1:
                        base = 0.02
                    elif sales_payment_diff == 2:
                        base = 0.01
                    else:
                        base = 0.0
                    
                    # 지부별 커스터마이징
                    if branch_name in branch_rates:
                        br = branch_rates[branch_name]
                        if base == 0.04:
                            rate = br.get('4', base)
                        else:
                            # 기본율이 0,1,2% 구간은 모두 '2' 키로 통합
                            rate = br.get('2', 0.02)
                    else:
                        rate = base
                
                df_filtered.at[idx, '인센티브율'] = rate
                df_filtered.at[idx, '인센티브금액'] = payment * rate
            else:
                # 미입금 경우
                df_filtered.at[idx, '인센티브율'] = 0.0
                df_filtered.at[idx, '인센티브금액'] = 0.0

        return df_filtered

    def compute_deduction_rate(
        self, 
        receipt_date: datetime, 
        payment_date: datetime
    ) -> float:
        """
        차감 규정 함수
        - 접수월 기준 4개월, 5개월째 입금액 없을 경우 수수료의 3% 차감
        """
        # 월차이 계산
        month_diff = (
            (payment_date.year - receipt_date.year) * 12 +
            (payment_date.month - receipt_date.month)
        )
        
        # 4개월 또는 5개월째 입금 없으면 차감
        if month_diff in [4, 5]:
            return 0.03  # 3% 차감
        
        return 0.0  # 차감 없음

    def fix_data_types(self, df):
        """필요한 열의 데이터 타입을 숫자로 변환 (쉼표 제거 포함)"""
        numeric_columns = ['공급가액', '세액', '금액', '입금액', '차감액', '인센티브금액']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce')
        return df

    def debug_dataframe(self, df, title="디버그"):
        """디버깅을 위한 데이터프레임 출력 함수"""
        print(f"===== {title} =====")
        if df is not None:
            print(df.head())
            print(df.info())
        else:
            print("데이터 없음")

    def apply_special_incentive_rules(self, row, standard_rate, rate_key):
        """특수 인센티브율 규칙: 연구용역/지부별 대상 1% → 2% 상향"""
        if standard_rate == 0.01:
            is_research = str(row.get('검사목적', '')).strip() == '연구용역'
            branch_name = str(row.get('지부명', '')).strip()
            branch_rates = self.settings.get('branch_rates', {})
            group_names = {"서울센터", "본사", "경북센터"}

            is_branch_target = branch_name in branch_rates and branch_name not in group_names

            if is_research or is_branch_target:
                return 0.02

        if standard_rate == 0.04:
            return 0.04

        return standard_rate

    def create_team_and_sales_rep_results(self, df_filtered):
        """팀 및 영업담당별 결과 데이터프레임 생성 (인센티브율별 소계/총계)"""
        # 영업담당 목록 추출
        all_sales_reps = df_filtered['영업담당'].dropna().unique().tolist()
        
        # 초기화
        self.team_results = {}
        self.sales_rep_results = {}
        
        # 1) 팀별 결과
        for team_name, members in self.teams.items():
            team_data = df_filtered[df_filtered['영업담당'].isin(members)].copy()
            if team_data.empty:
                continue
            
            result_dfs = []
            # 인센티브율별 그룹화 (0.0, 0.01, 0.02, 0.04 순서)
            for rate in sorted(team_data['인센티브율'].unique()):
                group = team_data[team_data['인센티브율'] == rate]
                # 원본 행
                result_dfs.append(group.copy())
                # 소계 행
                subtotal = pd.DataFrame({
                    '인센티브율':    [rate],
                    '금액':         [group['금액'].sum()],
                    '입금액':       [group['입금액'].sum()],
                    '인센티브금액':  [group['인센티브금액'].sum()],
                    '차감액':       [group['차감액'].sum()],
                    '지부명':       ['소계'],
                    '영업담당':     [team_name],
                    '매출일자':     [None],
                    '수금일자':     [None],
                    '월차이':       [None],
                })
                result_dfs.append(subtotal)
            
            # 팀 전체 총계 행
            total = pd.DataFrame({
                '인센티브율':    [None],
                '금액':         [team_data['금액'].sum()],
                '입금액':       [team_data['입금액'].sum()],
                '인센티브금액':  [team_data['인센티브금액'].sum()],
                '차감액':       [team_data['차감액'].sum()],
                '지부명':       ['총계'],
                '영업담당':     [team_name],
                '매출일자':     [None],
                '수금일자':     [None],
                '월차이':       [None],
            })
            
            # 빈 DataFrame 제외 후 합치기
            non_empty = [df for df in result_dfs if not df.empty]
            team_df = pd.concat(non_empty + [total], ignore_index=True)
            self.team_results[team_name] = team_df
        
        # 2) 영업담당별 결과
        for rep in all_sales_reps:
            rep_data = df_filtered[df_filtered['영업담당'] == rep].copy()
            if rep_data.empty:
                continue
            
            result_dfs = []
            for rate in sorted(rep_data['인센티브율'].unique()):
                group = rep_data[rep_data['인센티브율'] == rate]
                result_dfs.append(group.copy())
                subtotal = pd.DataFrame({
                    '인센티브율':    [rate],
                    '금액':         [group['금액'].sum()],
                    '입금액':       [group['입금액'].sum()],
                    '인센티브금액':  [group['인센티브금액'].sum()],
                    '차감액':       [group['차감액'].sum()],
                    '지부명':       ['소계'],
                    '영업담당':     [rep],
                    '매출일자':     [None],
                    '수금일자':     [None],
                    '월차이':       [None],
                })
                result_dfs.append(subtotal)
            
            total = pd.DataFrame({
                '인센티브율':    [None],
                '금액':         [rep_data['금액'].sum()],
                '입금액':       [rep_data['입금액'].sum()],
                '인센티브금액':  [rep_data['인센티브금액'].sum()],
                '차감액':       [rep_data['차감액'].sum()],
                '지부명':       ['총계'],
                '영업담당':     [rep],
                '매출일자':     [None],
                '수금일자':     [None],
                '월차이':       [None],
            })
            
            non_empty = [df for df in result_dfs if not df.empty]
            rep_df = pd.concat(non_empty + [total], ignore_index=True)
            self.sales_rep_results[rep] = rep_df

    def display_results(self):
        """전체 결과 및 팀/영업담당별 결과 표시"""
        # 모든 탭 초기화
        self.clear_all_tabs()
        
        # 전체 결과 표시
        self.display_result_in_tree(self.all_tree, self.result_df)
        
        # 팀별 탭 및 결과 생성
        for team_name, team_df in self.team_results.items():
            # 팀 탭 생성
            team_frame = tk.Frame(self.notebook)
            self.notebook.add(team_frame, text=f"{team_name}")
            
            # 팀 트리뷰 생성
            team_tree = ttk.Treeview(team_frame)
            team_tree.pack(fill=tk.BOTH, expand=True)
            
            # 스크롤바 추가
            scrollbar_y = ttk.Scrollbar(team_tree, orient="vertical", command=team_tree.yview)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            team_tree.configure(yscrollcommand=scrollbar_y.set)
            
            scrollbar_x = ttk.Scrollbar(team_tree, orient="horizontal", command=team_tree.xview)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            team_tree.configure(xscrollcommand=scrollbar_x.set)
            
            # 결과 표시
            self.display_result_in_tree(team_tree, team_df)
    
    def display_result_in_tree(self, tree, df, override_columns=None):
        """트리뷰에 결과 표시 (컬럼 제목 + 삭제버튼 포함)"""

        # 기존 컬럼 헤더 프레임이 있으면 제거
        parent_frame = tree.master
        for widget in parent_frame.winfo_children():
            if hasattr(widget, 'header_frame_tag') and widget.header_frame_tag:
                widget.destroy()

        # 열 설정 (override_columns가 있으면 그것을 사용)
        columns_to_show = override_columns if override_columns is not None else (
            self.default_columns + self.settings.get('custom_columns', [])
        )

        # 컬럼 이름 매핑 및 유효성 검사
        selected_columns_mapped = []
        for col in columns_to_show:
            mapped_col = self.column_mapping.get(col, col)
            if mapped_col in df.columns:
                selected_columns_mapped.append(mapped_col)

        # 트리뷰 컬럼 설정
        tree["columns"] = selected_columns_mapped
        tree.heading("#0", text="No.")
        tree.column("#0", width=50)

        for col in selected_columns_mapped:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        # 기존 데이터 제거
        tree.delete(*tree.get_children())

        # 트리뷰에 데이터 삽입
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            values = [row.get(col, "") for col in selected_columns_mapped]
            tree.insert("", "end", text=str(i), values=values)

        # 컬럼 삭제 버튼 UI 하단에 추가
        button_frame = tk.Frame(parent_frame)
        button_frame.header_frame_tag = True
        button_frame.pack(fill="x")

        for col in selected_columns_mapped:
            frame = tk.Frame(button_frame)
            frame.pack(side="left", padx=1)

            label = tk.Label(frame, text=col)
            label.pack(side="left")

            btn = tk.Button(frame, text="✕", fg="red", command=lambda c=col: self.remove_column_from_view(c, tree, df))
            btn.pack(side="left")

        # 데이터 추가
        for i, row in df.iterrows():
            values = []
            # 내부 표시 컬럼 정의
            internal_columns = selected_columns_mapped
            for col in internal_columns:
                if col in ['매출일자', '수금일자'] and not pd.isna(row.get(col)):
                    values.append(row[col].strftime('%Y-%m-%d'))
                elif col == '인센티브율' and not pd.isna(row.get(col)):
                    values.append(f"{row[col]*100:.1f}%")
                elif col in ['인센티브금액', '금액', '입금액', '차감액'] and not pd.isna(row.get(col)):
                    values.append(f"{int(row[col]):,}")
                else:
                    values.append(row.get(col, ''))

            if (isinstance(row.get('지부명'), str) and row['지부명'] == '소계') or \
            (isinstance(row.get('매출월'), str) and row['매출월'] == '총계'):
                tree.insert("", tk.END, text=str(i+1), values=values, tags=('subtotal',))
            else:
                tree.insert("", tk.END, text=str(i+1), values=values)

        tree.tag_configure('subtotal', background='lightgray', font=('Arial', 10, 'bold'))

    def confirm_remove_column(self, column, tree, df):
        """컬럼 제거 확인 대화상자"""
        result = messagebox.askyesno("컬럼 삭제", f"'{column}' 컬럼을 삭제하시겠습니까?")
        if result:
            # 사용자가 '예'를 선택한 경우
            self.remove_column_from_view(column, tree, df)

    def remove_column_from_view(self, column, tree, df):
        """트리뷰에서 컬럼 제거"""
        current_columns = list(tree["columns"])
        if column in current_columns:
            current_columns.remove(column)
            self.display_result_in_tree(tree, df, override_columns=current_columns)
            messagebox.showinfo("알림", f"‘{column}’ 컬럼이 제거되었습니다.")
        
    def clear_all_tabs(self):
        """모든 탭 초기화"""
        # 전체 결과 트리뷰 초기화
        for item in self.all_tree.get_children():
            self.all_tree.delete(item)
        
        # 기존 탭 제거 (첫 번째 탭 제외)
        while len(self.notebook.tabs()) > 1:
            self.notebook.forget(1)  # 첫 번째 탭(전체 결과)은 유지하고 나머지 제거    

    def export_current_view_to_excel(self):
        """현재 보이는 화면을 엑셀로 내보내기"""
        if self.result_df is None:
            messagebox.showerror("오류", "저장할 결과가 없습니다.")
            return
        
        # 파일 저장 대화상자 열기
        file_path = filedialog.asksaveasfilename(
            title="엑셀로 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            # 엑셀 작성자 생성
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 현재 표시된 컬럼 (전체 결과 화면)
                displayed_df = self.result_df.copy()
                
                # 열 이름 변경 (UI에 표시할 이름으로)
                for original_col, ui_col in self.reverse_mapping.items():
                    if original_col in displayed_df.columns:
                        displayed_df = displayed_df.rename(columns={original_col: ui_col})
                
                # 엑셀로 저장
                displayed_df.to_excel(writer, sheet_name='전체 결과', index=False)
            
            messagebox.showinfo("알림", f"결과가 {file_path}에 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 변환 중 오류가 발생했습니다: {str(e)}")
            traceback.print_exc()
    
    def split_tabs_by_sales_rep(self):
        """영업담당별로 탭을 분리하는 함수"""
        if not hasattr(self, 'sales_rep_results') or not self.sales_rep_results:
            messagebox.showinfo("알림", "먼저 인센티브를 계산해주세요.")
            return
        
        # 모든 탭 초기화
        self.clear_all_tabs()
        
        # 전체 결과 표시
        self.display_result_in_tree(self.all_tree, self.result_df)
        
        # 팀별 탭 먼저 생성 - 팀 구조를 보여주기 위해
        for team_name, team_df in self.team_results.items():
            # 팀 탭 생성
            team_frame = tk.Frame(self.notebook)
            self.notebook.add(team_frame, text=f"팀: {team_name}")
            
            # 팀 트리뷰 생성
            team_tree = ttk.Treeview(team_frame)
            team_tree.pack(fill=tk.BOTH, expand=True)
            
            # 스크롤바 추가
            scrollbar_y = ttk.Scrollbar(team_tree, orient="vertical", command=team_tree.yview)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            team_tree.configure(yscrollcommand=scrollbar_y.set)
            
            scrollbar_x = ttk.Scrollbar(team_tree, orient="horizontal", command=team_tree.xview)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            team_tree.configure(xscrollcommand=scrollbar_x.set)
            
            # 결과 표시
            self.display_result_in_tree(team_tree, team_df)
        
        # 영업담당별 탭 생성
        processed_reps = set()  # 중복 처리 방지
        for rep_key, rep_df in self.sales_rep_results.items():
            # 영업담당 이름 및 팀 이름 추출
            if '_' in rep_key:
                team_name, rep_name = rep_key.split('_', 1)
                tab_name = f"{rep_name} ({team_name})"  # 팀 정보 포함한 탭 이름
            else:
                rep_name = rep_key
                tab_name = rep_name
            
            # 중복 처리 방지
            if rep_name in processed_reps:
                continue
            
            processed_reps.add(rep_name)
            
            # 영업담당 탭 생성
            rep_frame = tk.Frame(self.notebook)
            self.notebook.add(rep_frame, text=tab_name)
            
            # 영업담당 트리뷰 생성
            rep_tree = ttk.Treeview(rep_frame)
            rep_tree.pack(fill=tk.BOTH, expand=True)
            
            # 스크롤바 추가
            scrollbar_y = ttk.Scrollbar(rep_tree, orient="vertical", command=rep_tree.yview)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            rep_tree.configure(yscrollcommand=scrollbar_y.set)
            
            scrollbar_x = ttk.Scrollbar(rep_tree, orient="horizontal", command=rep_tree.xview)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            rep_tree.configure(xscrollcommand=scrollbar_x.set)
            
            # 결과 표시
            self.display_result_in_tree(rep_tree, rep_df)
        
        messagebox.showinfo("알림", f"{len(self.team_results)}개의 팀 탭과 {len(processed_reps)}개의 영업담당별 탭이 생성되었습니다.")
    
    def save_tabs(self):
        """각 탭의 내용을 별도 파일로 저장"""
        if not self.sales_rep_results:
            messagebox.showinfo("알림", "먼저 인센티브를 계산해주세요.")
            return

        save_dir = filedialog.askdirectory(title="저장할 폴더 선택")
        if not save_dir:
            return

        reference_year = int(self.year_var.get())
        reference_month = int(self.month_var.get())

        try:
            saved_count = 0
            team_files = {
                team_name: {
                    'path': os.path.join(save_dir, f"{reference_year}{reference_month:02d}_{team_name}.xlsx"),
                    'members': []
                }
                for team_name in self.teams.keys()
            }

            selected_columns_raw = self.default_columns + self.settings.get('custom_columns', [])
            selected_columns = [self.column_mapping.get(col, col) for col in selected_columns_raw]

            # 영업담당자별 저장
            for rep_key, rep_df in self.sales_rep_results.items():
                rep_name = rep_key.split('_', 1)[1] if '_' in rep_key else rep_key
                rep_file_path = os.path.join(save_dir, f"{reference_year}{reference_month:02d}_{rep_name}.xlsx")

                available_columns = [col for col in selected_columns if col in rep_df.columns]
                rep_df_for_save = rep_df[available_columns].copy()

                for original_col, ui_col in self.reverse_mapping.items():
                    if original_col in rep_df_for_save.columns:
                        rep_df_for_save = rep_df_for_save.rename(columns={original_col: ui_col})

                rep_df_for_save.to_excel(rep_file_path, index=False)

                # 서식 적용
                wb = openpyxl.load_workbook(rep_file_path)
                ws = wb.active
                format_excel_sheet(ws)
                wb.save(rep_file_path)

                saved_count += 1

                # 팀 멤버 기록
                if '_' in rep_key:
                    team_name = rep_key.split('_', 1)[0]
                    if team_name in team_files:
                        team_files[team_name]['members'].append(rep_name)

            # 팀 파일 생성
            for team_name, team_info in team_files.items():
                if team_name in self.team_results and team_info['members']:
                    team_df = self.team_results[team_name]
                    team_path = team_info['path']

                    with pd.ExcelWriter(team_path, engine='openpyxl') as writer:
                        available_columns = [col for col in selected_columns if col in team_df.columns]
                        team_df_for_save = team_df[available_columns].copy()

                        for original_col, ui_col in self.reverse_mapping.items():
                            if original_col in team_df_for_save.columns:
                                team_df_for_save = team_df_for_save.rename(columns={original_col: ui_col})

                        team_df_for_save.to_excel(writer, sheet_name=team_name, index=False)

                        for member in team_info['members']:
                            member_key = f"{team_name}_{member}"
                            if member_key in self.sales_rep_results:
                                member_df = self.sales_rep_results[member_key]
                                available_columns = [col for col in selected_columns if col in member_df.columns]
                                member_df_for_save = member_df[available_columns].copy()

                                for original_col, ui_col in self.reverse_mapping.items():
                                    if original_col in member_df_for_save.columns:
                                        member_df_for_save = member_df_for_save.rename(columns={original_col: ui_col})

                                sheet_name = member[:31]  # 시트 이름 제한
                                for char in ['[', ']', '*', '?', ':', '/']:
                                    sheet_name = sheet_name.replace(char, '_')

                                member_df_for_save.to_excel(writer, sheet_name=sheet_name, index=False)

                    # 팀 파일 서식 적용
                    wb = openpyxl.load_workbook(team_path)
                    ws = wb.active
                    format_excel_sheet(ws)
                    wb.save(team_path)

                    saved_count += 1

            messagebox.showinfo("알림", f"{saved_count}개의 파일이 저장되었습니다.")

        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다: {str(e)}")
            traceback.print_exc()


    
    def setup_column_select_tab(self):
        """표시할 컬럼 선택 탭"""
        frame = tk.Frame(self.column_select_frame)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(frame, text="기본 컬럼 외에 표시할 추가 컬럼을 선택하세요.").pack(anchor=tk.W, pady=(0, 10))

        list_frame = tk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 왼쪽: 추가 가능한 컬럼
        available_frame = tk.LabelFrame(list_frame, text="추가 가능한 컬럼")
        available_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        self.available_column_listbox = tk.Listbox(available_frame, selectmode=tk.MULTIPLE)
        self.available_column_listbox.pack(fill=tk.BOTH, expand=True)

        # 오른쪽: 선택된 컬럼
        selected_frame = tk.LabelFrame(list_frame, text="표시할 컬럼")
        selected_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        self.selected_column_listbox = tk.Listbox(selected_frame, selectmode=tk.MULTIPLE)
        self.selected_column_listbox.pack(fill=tk.BOTH, expand=True)

        # 가운데 버튼
        middle_btn_frame = tk.Frame(list_frame)
        middle_btn_frame.pack(side=tk.LEFT, padx=5)

        tk.Button(middle_btn_frame, text="▶", command=self.add_selected_columns).pack(pady=5)
        tk.Button(middle_btn_frame, text="◀", command=self.remove_selected_columns).pack(pady=5)

        # 초기화
        additional_columns = self.calculator.additional_columns if hasattr(self.calculator, "additional_columns") else []
        selected_columns = self.settings.get('custom_columns', [])
        for col in additional_columns:
            if col not in selected_columns:
                self.available_column_listbox.insert(tk.END, col)
        for col in selected_columns:
            self.selected_column_listbox.insert(tk.END, col)

    def add_selected_columns(self):
        """왼쪽 → 오른쪽으로 이동"""
        selected = self.available_column_listbox.curselection()
        for i in reversed(selected):
            val = self.available_column_listbox.get(i)
            self.available_column_listbox.delete(i)
            self.selected_column_listbox.insert(tk.END, val)

    def remove_selected_columns(self):
        """오른쪽 → 왼쪽으로 이동"""
        selected = self.selected_column_listbox.curselection()
        for i in reversed(selected):
            val = self.selected_column_listbox.get(i)
            self.selected_column_listbox.delete(i)
            self.available_column_listbox.insert(tk.END, val)


    def save_result(self):
        """모든 결과를 하나의 엑셀 파일로 저장"""
        if self.result_df is None:
            messagebox.showerror("오류", "저장할 결과가 없습니다.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="결과 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")]
        )
        
        if file_path:
            try:
                # 엑셀 작성자 생성
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # 표시할 컬럼 선택 (기본 컬럼 + 사용자가 추가 선택한 컬럼)
                    selected_columns_raw = self.default_columns + self.settings.get('custom_columns', [])
                    selected_columns = []
                    
                    # 매핑된 열 이름으로 변환
                    for col in selected_columns_raw:
                        # 컬럼 매핑이 있으면 변환된 이름 사용, 없으면 원래 이름 사용
                        mapped_col = self.column_mapping.get(col, col)
                        selected_columns.append(mapped_col)
                    
                    # 데이터프레임의 실제 열만 필터링
                    available_columns = [col for col in selected_columns if col in self.result_df.columns]
                    result_df_for_save = self.result_df[available_columns].copy()
                    
                    # 열 이름 변경 (UI에 표시할 이름으로)
                    for original_col, ui_col in self.reverse_mapping.items():
                        if original_col in result_df_for_save.columns:
                            result_df_for_save = result_df_for_save.rename(columns={original_col: ui_col})
                    
                    # 전체 결과 시트에 저장
                    result_df_for_save.to_excel(writer, sheet_name='전체 결과', index=False)
                    
                    # 팀별 결과 저장
                    for team_name, team_df in self.team_results.items():
                        # 시트 이름 처리
                        sheet_name = str(team_name)
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:31]
                        
                        # 특수문자 제거
                        for char in ['[', ']', '*', '?', ':', '/']:
                            sheet_name = sheet_name.replace(char, '_')
                        
                        # 데이터프레임의 실제 열만 필터링
                        available_columns = [col for col in selected_columns if col in team_df.columns]
                        team_df_for_save = team_df[available_columns].copy()
                        
                        # 열 이름 변경
                        for original_col, ui_col in self.reverse_mapping.items():
                            if original_col in team_df_for_save.columns:
                                team_df_for_save = team_df_for_save.rename(columns={original_col: ui_col})
                        
                        team_df_for_save.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 개별 영업담당별 결과 저장
                    for rep_key, rep_df in self.sales_rep_results.items():
                        # 영업담당 이름 추출
                        if '_' in rep_key:
                            sheet_name = rep_key.split('_', 1)[1]  # 팀 이름 이후의 담당자 이름만 추출
                        else:
                            sheet_name = rep_key
                        
                        # 시트 이름 길이 제한
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:31]
                        
                        # 특수문자 제거
                        for char in ['[', ']', '*', '?', ':', '/']:
                            sheet_name = sheet_name.replace(char, '_')
                        
                        # 데이터프레임의 실제 열만 필터링
                        available_columns = [col for col in selected_columns if col in rep_df.columns]
                        rep_df_for_save = rep_df[available_columns].copy()
                        
                        # 열 이름 변경
                        for original_col, ui_col in self.reverse_mapping.items():
                            if original_col in rep_df_for_save.columns:
                                rep_df_for_save = rep_df_for_save.rename(columns={original_col: ui_col})
                        
                        rep_df_for_save.to_excel(writer, sheet_name=sheet_name, index=False)

                # 모든 시트에 서식 적용
                wb = openpyxl.load_workbook(file_path)
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    format_excel_sheet(ws)
                wb.save(file_path)
                
                messagebox.showinfo("알림", f"결과가 {file_path}에 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류가 발생했습니다: {str(e)}")
                traceback.print_exc()

def main():
    root = tk.Tk()
    app = IncentiveCalculator(root)
    root.mainloop()


# incen.py 파일 마지막 부분을 다음과 같이 수정
if __name__ == "__main__":
    import tkinter as tk
    root = tk.Tk()
    app = IncentiveCalculator(root)
    root.mainloop()