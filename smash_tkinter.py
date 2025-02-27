import tkinter as tk
from tkinter import ttk, messagebox
import os
import webbrowser  # 웹브라우저 열기 위한 모듈 추가
import re
from dotenv import load_dotenv
import pandas as pd
import yaml
from excel_to_notion import ExcelToNotionImporter
import time
import sys


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class SmashTeamGenerator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("스매시 조 편성기 v2.0")
        self.geometry("1200x800")
        self.config_file = "smash_config.json"
        # .env 파일 읽기
        load_dotenv()
        # 스매시 페이지 토큰값
        self.NOTION_TOKEN = os.getenv("NOTION_TOKEN")
        if not self.NOTION_TOKEN:
            raise ValueError("NOTION_TOKEN이 .env에 설정되지 않았습니다.")
        # Notion 관련 정보 설정
        self.parent_page_id = "1a39f0ea074e80e085a5dbe8bfa5404f"
        self.template_page_id = "1a79f0ea074e807b9b18c586b0890893"
        # 사용자 문서 폴더에 저장

        # Windows: `C:\Users\사용자이름`
        # macOS: `/ Users/사용자이름`
        documents_path = os.path.join(os.path.expanduser("~"), "Documents")
        self.excel_file_path = os.path.join(documents_path, "팀_구성_결과.xlsx")
        self.notion_page_url = None  # 생성된 Notion 페이지 URL 저장

        # GUI 스타일 초기화
        self.init_style()
        # 메인 위젯 생성
        self.create_widgets()

    def init_style(self):
        """스타일 설정"""
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f5f5f5')
        self.style.configure('TButton', font=('맑은 고딕', 10), padding=5)
        self.style.configure('Header.TLabel', font=('맑은 고딕', 11, 'bold'))

    def create_widgets(self):
        """메인 위젯 구성"""
        main_frame = ttk.Frame(self)
        main_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)

        # 1. 입력 영역
        self.create_input_section(main_frame)

        # 2. 결과 영역
        self.create_result_section(main_frame)

        # 3. 상태 표시바
        self.status_bar = ttk.Label(self, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_input_section(self, parent):
        """텍스트 입력 영역"""
        input_frame = ttk.LabelFrame(parent, text="■ 입력 데이터 붙여넣기")
        input_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 接龙 내용 입력
        self.jielong_text = self.create_text_area(
            input_frame, "接龙 내용 (接龙 전체 붙여넣기):", 0, "jielong")

        # 제외 명단 입력
        self.lesson_text = self.create_text_area(
            input_frame, "제외 명단 (코치/레슨생):", 1, "lesson")

        input_frame.grid_columnconfigure(0, weight=1)  # 접룡 영역 가중치 1
        input_frame.grid_columnconfigure(1, weight=1)  # 제외명단 영역 가중치 1
        input_frame.grid_rowconfigure(0, weight=1)

    def create_text_area(self, parent, label_text, column, example):
        """재사용 가능한 텍스트 영역 생성"""
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=column, sticky="nsew")

        lbl = ttk.Label(frame, text=label_text, style='Header.TLabel')
        lbl.pack(anchor=tk.W)

        text_widget = tk.Text(frame, wrap=tk.WORD, height=10,
                              font=('맑은 고딕', 10))
        # tmp 예시추가
        if example == "jielong":
            try:
                file_path = resource_path('examples/jielong.txt')
                with open(file_path, 'r', encoding='utf-8') as f:
                    example_text = f.read()
                    text_widget.insert(tk.END, example_text)
            except FileNotFoundError:
                print("jielong 예시 파일이 존재하지 않습니다")
        elif example == "lesson":
            try:
                file_path = resource_path('examples/lesson.txt')
                with open(file_path, 'r', encoding='utf-8') as f:
                    example_text = f.read()
                    text_widget.insert(tk.END, example_text)
            except FileNotFoundError:
                print("lesson 예시 파일이 존재하지 않습니다")

        scroll = ttk.Scrollbar(frame, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scroll.set)

        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        return text_widget

    def create_result_section(self, parent):
        """결과 표시 영역"""
        result_frame = ttk.LabelFrame(parent, text="■ 팀 구성 결과")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 상단: 트리뷰 영역 (좌우)
        tree_frame = ttk.Frame(result_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        #  왼쪽: 페어링 프레임
        pairing_frame = ttk.LabelFrame(tree_frame, text=" 페어링 결과")
        pairing_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 페어링 테이블
        self.pairing_tree = ttk.Treeview(
            pairing_frame,
            columns=('index', 't1', 't2'),
            show='headings',
        )
        self.pairing_tree.heading('index', text='No.', anchor='center')
        self.pairing_tree.column(
            'index', width=50, anchor='center', stretch=tk.NO)
        self.pairing_tree.heading('t1', text='팀 1', anchor='center')
        self.pairing_tree.column('t1', width=120, anchor='center')
        self.pairing_tree.heading('t2', text='팀 2', anchor='center')
        self.pairing_tree.column('t2', width=120, anchor='center')

        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(pairing_frame, orient=tk.VERTICAL,
                                  command=self.pairing_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.pairing_tree.configure(yscroll=scrollbar.set)
        self.pairing_tree.pack(fill='both', expand=True, padx=0, pady=0)

        # 오른쪽: 조 편성 프레임
        group_frame = ttk.LabelFrame(tree_frame, text=" 조 편성 결과 (A,B,C 조)")
        group_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        # 조 편성 테이블
        self.group_tree = ttk.Treeview(
            group_frame,
            columns=('index', 'A', 'B', 'C'),
            show='headings',
        )
        self.group_tree.heading('index', text='No.', anchor='center')
        self.group_tree.column(
            'index', width=50, anchor='center', stretch=tk.NO)
        self.group_tree.heading('A', text='A조', anchor='center')
        self.group_tree.column('A', width=100, anchor='center')
        self.group_tree.heading('B', text='B조', anchor='center')
        self.group_tree.column('B', width=100, anchor='center')
        self.group_tree.heading('C', text='C조', anchor='center')
        self.group_tree.column('C', width=100, anchor='center')

        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(group_frame, orient=tk.VERTICAL,
                                  command=self.group_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.group_tree.configure(yscroll=scrollbar.set)
        self.group_tree.pack(fill='both', expand=True, padx=0, pady=0)

        # 하단: 제어 버튼 영역
        btn_frame = ttk.Frame(result_frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="팀 생성",
                   command=self.generate_teams).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="추출하기",
                   command=self.export_to_excel).pack(side=tk.LEFT, padx=5)

        # Notion 업로드 버튼 추가
        ttk.Button(btn_frame, text="Notion 업로드",
                   command=self.upload_to_notion).pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="초기화",
                   command=self.clear_inputs).pack(side=tk.LEFT, padx=5)

        # Notion 링크 표시 레이블 추가
        self.notion_link_frame = ttk.Frame(result_frame)
        self.notion_link_frame.pack(pady=5)

        self.notion_link_label = ttk.Label(
            self.notion_link_frame,
            text="",
            foreground="blue",
            cursor="hand2"
        )
        self.notion_link_label.pack(side=tk.LEFT)

        # Notion 상태 표시 레이블
        self.notion_status_label = ttk.Label(self.notion_link_frame, text="")
        self.notion_status_label.pack(side=tk.LEFT, padx=10)

    def generate_teams(self):
        """팀 생성 핵심 로직"""
        try:
            # 1. 입력 데이터 처리
            # 1.1 입력 데이터 가져오기
            jielong_content = self.jielong_text.get("1.0", tk.END).strip()
            lesson_content = self.lesson_text.get("1.0", tk.END).strip()

            # 接龙 입력값 검증
            if not jielong_content:
                raise ValueError("接龙 내용을 입력해주세요")

            # 1.2 接龙 내용 추출
            jielong_start = jielong_content.find('1.')
            if jielong_start == -1:
                raise ValueError("접룡 내용에 '1.' 표시가 없습니다")

            # 接龙 명단 추출
            jielong_names_list = re.findall(
                r'(?:게스트\s\([가-힣]+\)|[가-힣]+[A-Z]?)', jielong_content[jielong_start:])
            print(f"접룡 인원: {len(jielong_names_list)} 명")
            # print(jielong_names_list)

            # 1.3 레슨 내용 추출
            lesson_names_list = []
            if lesson_content:
                lesson_start = lesson_content.find('1.')
                if lesson_start == -1:
                    raise ValueError("레슨 인원에 '1.' 표시가 없습니다")
            # 레슨 명단 추출
                lesson_names_list = re.findall(
                    r'[가-힣]+[A-Z]?', lesson_content[lesson_start:])
                lesson_names_list = [
                    name for name in lesson_names_list if "코치" not in name]
                print(f"레슨 인원: {len(lesson_names_list)} 명")
                # print(lesson_names_list)

            # 2. 그룹 데이터 로드
            yaml_path = resource_path('groups.yaml')
            with open(yaml_path, 'r', encoding='utf-8') as f:
                groups_data = yaml.safe_load(f)

            # 3. 1st 페어링 편성
            # 3.1 접룡 인원 정렬
            ordered_list = []
            for group in groups_data['groups'].values():
                ordered_list.extend(group)
            jielong_ordered_list = sorted(
                jielong_names_list, key=lambda x: ordered_list.index(x))

            # 3.2 접룡 인원에서 레슨 인원 제외
            pairing_list = [
                name for name in jielong_ordered_list if name not in lesson_names_list]

            # 홀수 짝수 확인, 홀수면 单打 인원 선택/입력 창 호출
            solo_player = None
            if len(pairing_list) % 2 != 0:
                # 单打 인원 선택/입력 창
                selected_player = self.show_solo_selection_dialog(pairing_list)
                if not selected_player:
                    raise ValueError("팀 생성에 필요한 单打 인원을 선택해야 합니다")

                solo_player = selected_player
                pairing_list.remove(solo_player)  # 单打 인원 제외

            # 3.3 정렬된 인원 페어링 편성
            half = len(pairing_list) // 2
            first_column = pairing_list[:half]
            second_column = pairing_list[half:][::-1]

            # 3.4 트리뷰 초기화 및 데이터 삽입
            self.pairing_tree.delete(*self.pairing_tree.get_children())

            # 페어링 인원 추가
            for i, (t1, t2) in enumerate(zip(first_column, second_column), start=1):
                self.pairing_tree.insert('', tk.END, values=(i, t1, t2))

            # 单打 인원 추가 (빨간색)
            if solo_player:
                last_index = len(self.pairing_tree.get_children()) + 1
                self.pairing_tree.insert(
                    '', tk.END,
                    values=(last_index, solo_player, ""),
                    tags=('solo',)
                )
                self.pairing_tree.tag_configure('solo', background='#FFE4E1')

            # 레슨 인원 추가 (노란색으로 변경)
            if lesson_names_list:
                lesson_pairs = []
                for i in range(0, len(lesson_names_list), 2):
                    pair = (
                        lesson_names_list[i],
                        lesson_names_list[i+1] if i+1 < len(lesson_names_list) else '')
                    lesson_pairs.append(pair)
                last_index = len(self.pairing_tree.get_children())
                for i, (t1, t2) in enumerate(lesson_pairs, start=1):
                    self.pairing_tree.insert(
                        '', tk.END,
                        values=(i+last_index, t1, t2),
                        tags=('lesson',)
                    )
                self.pairing_tree.tag_configure(
                    'lesson', background='#FFEB9C')

            # 4. 2nd 조 편성
            # 4.1 접룡 인원 그룹별로 분리
            filtered_groups = {
                'A': [name for name in groups_data['groups']['A'] if name in jielong_names_list],
                'B': [name for name in groups_data['groups']['B'] if name in jielong_names_list],
                'C': [name for name in groups_data['groups']['C'] if name in jielong_names_list]
            }

            # 4.2 트리뷰에 표시
            self.group_tree.delete(*self.group_tree.get_children())

            max_length = max(len(filtered_groups['A']),
                             len(filtered_groups['B']),
                             len(filtered_groups['C']))

            for i in range(max_length):
                #fmt: off
                a = filtered_groups['A'][i] if i < len(filtered_groups['A']) else ''
                b = filtered_groups['B'][i] if i < len(filtered_groups['B']) else ''
                c = filtered_groups['C'][i] if i < len(filtered_groups['C']) else ''
                #fmt: on
                self.group_tree.insert('', tk.END, values=(i+1, a, b, c))

            self.status_bar.config(text="팀 생성 완료")

        except Exception as e:
            messagebox.showerror("오류", str(e))

    def export_to_excel(self):
        """Excel 추출 기능 (시트 분리)"""
        try:
            group_data = [
                self.group_tree.item(item)['values'][1:]  # 인덱스 제거
                for item in self.group_tree.get_children()
            ]
            # 결과 데이터 추출
            pair_data = []
            self.lesson_indices = []  # 레슨 행 인덱스
            self.solo_indices = []    # 단식 행 인덱스

            for i, item in enumerate(self.pairing_tree.get_children()):
                values = self.pairing_tree.item(item)['values'][1:]  # 인덱스 제거
                pair_data.append(values)

                # 태그 확인하여 인덱스 저장
                tags = self.pairing_tree.item(item, 'tags')
                if 'lesson' in tags:
                    self.lesson_indices.append(i)
                if 'solo' in tags:
                    self.solo_indices.append(i)

            # Excel 파일 생성
            with pd.ExcelWriter(self.excel_file_path, engine='xlsxwriter') as writer:
                # 페어링 데이터프레임 생성
                pairing_df = pd.DataFrame(pair_data, columns=['t1', 't2'])

                # 페어링 시트에 데이터 쓰기
                pairing_df.to_excel(
                    writer,
                    sheet_name='페어링',
                    index=False,
                )

                # 워크북과 워크시트 객체 가져오기
                workbook = writer.book
                pairing_sheet = writer.sheets['페어링']

                # 포맷 생성
                yellow_format = workbook.add_format(
                    {'bg_color': '#FFEB9C'})  # 노란색 (레슨)
                red_format = workbook.add_format(
                    {'bg_color': '#FFE4E1'})     # 빨간색 (단식)

                # 레슨 행에 노란색 배경 적용
                for row_idx in self.lesson_indices:
                    pairing_sheet.write(
                        row_idx + 1, 0, pairing_df.iloc[row_idx, 0], yellow_format)
                    pairing_sheet.write(
                        row_idx + 1, 1, pairing_df.iloc[row_idx, 1], yellow_format)

                # 단식 행에 빨간색 배경 적용
                for row_idx in self.solo_indices:
                    pairing_sheet.write(
                        row_idx + 1, 0, pairing_df.iloc[row_idx, 0], red_format)
                    pairing_sheet.write(
                        row_idx + 1, 1, pairing_df.iloc[row_idx, 1], red_format)

                group_df = pd.DataFrame(group_data, columns=['A', 'B', 'C'])
                group_df.to_excel(
                    writer,
                    sheet_name='조편성',
                    index=False,
                )

            self.status_bar.config(text="엑셀 추출 완료: 팀_구성_결과.xlsx")
            messagebox.showinfo("성공", "엑셀 파일이 성공적으로 저장되었습니다")

        except Exception as e:
            messagebox.showerror("오류", str(e))

    def clear_inputs(self):
        """입력 초기화"""
        self.jielong_text.delete("1.0", tk.END)
        self.lesson_text.delete("1.0", tk.END)
        self.pairing_tree.delete(*self.pairing_tree.get_children())
        self.group_tree.delete(*self.group_tree.get_children())
        self.notion_link_label.config(text="")
        self.notion_status_label.config(text="")
        self.notion_link_label.unbind("<Button-1>")
        self.notion_page_url = None
        self.status_bar.config(text="입력 초기화 완료")

    def show_solo_selection_dialog(self, players):
        """单打 인원 선택/입력 창"""
        dialog = tk.Toplevel(self)
        dialog.title("单打 인원 선택창")
        dialog.geometry("300x120")

        dialog.transient(self)  # 메인 윈도우의 자식 창으로 설정
        dialog.grab_set()  # 모든 이벤트를 이 창으로 포커스

        dialog.resizable(False, False)  # 크기 변경 비활성화

        # 창 위치를 메인 윈도우 중앙에
        self.update_idletasks()
        main_x = self.winfo_x()
        main_y = self.winfo_y()
        main_width = self.winfo_width()
        main_height = self.winfo_height()
        dialog.geometry(
            f"+{main_x + (main_width//2 - 150)}+{main_y + (main_height//2 - 60)}")

        selected_player = tk.StringVar()

        ttk.Label(dialog, text="인원이 홀수입니다.\n单打 인원을 선택/입력 하세요 :",
                  anchor='center').pack(pady=5)
        cb = ttk.Combobox(dialog, textvariable=selected_player, values=players)
        cb.pack(pady=5)

        def on_confirm():
            dialog.selected_player = selected_player.get().strip()
            dialog.destroy()

        ttk.Button(dialog, text="확인", command=on_confirm).pack(pady=5)

        dialog.wait_window()  # 다이얼로그 닫힐 때까지 대기
        return getattr(dialog, 'selected_player', None)

    def upload_to_notion(self):
        """Notion에 업로드하는 기능"""
        # 엑셀 파일이 존재하는지 확인
        if not os.path.exists(self.excel_file_path):
            messagebox.showerror("오류", "엑셀 파일이 존재하지 않습니다. 먼저 '추출하기'를 실행해주세요.")
            return

        try:
            # 기존 바인딩 제거 및 레이블 초기화
            self.notion_link_label.config(text="")
            self.notion_link_label.unbind("<Button-1>")

            # 업로드 시작 표시
            self.notion_status_label.config(
                text="Notion 업로드 중...", foreground="orange")
            self.update_idletasks()  # UI 업데이트

            # 실행 시간 측정 시작
            t1 = time.time()

            # lesson_indices 확인 - 없으면 빈 리스트로 초기화
            lesson_indices = getattr(self, 'lesson_indices', [])

            # ExcelToNotionImporter 인스턴스 생성 및 실행
            cloner = ExcelToNotionImporter(
                self.NOTION_TOKEN,
                self.parent_page_id,
                self.template_page_id,
                self.excel_file_path,
                self.lesson_indices
            )

            # 템플릿 페이지 복제
            self.notion_page_url = cloner.duplicate_template_page()

            # 블록 업데이트
            cloner.update_block()

            # 실행 시간 측정 종료
            t2 = time.time()
            elapsed_time = t2 - t1

            # 링크 레이블 새로 설정
            self.notion_link_label.config(
                text=f"Notion 페이지 바로가기",
                font=('맑은 고딕', 10, 'underline'),
                foreground="blue",
                cursor="hand2"
            )
            # 바인딩 다시 설정
            self.notion_link_label.bind("<Button-1>", self.open_notion_page)

            # 상태 및 소요 시간 표시
            self.notion_status_label.config(
                text=f"업로드 완료 (소요 시간: {elapsed_time:.1f}초)",
                foreground="green"
            )

            # 상태바 업데이트
            self.status_bar.config(
                text=f"Notion 업로드 완료: {self.notion_page_url}")

            # 성공 메시지 표시
            messagebox.showinfo("성공", "Notion에 팀 구성 데이터가 성공적으로 업로드되었습니다.")

        except Exception as e:
            # 오류 발생 시 처리
            self.notion_link_label.config(text="")
            self.notion_status_label.config(text="업로드 실패", foreground="red")
            self.status_bar.config(text=f"Notion 업로드 오류: {str(e)}")
            messagebox.showerror("Notion 업로드 오류", str(e))

    def open_notion_page(self, event):
        """Notion 페이지 링크 클릭 시 웹 브라우저에서 열기"""
        if self.notion_page_url:
            webbrowser.open(self.notion_page_url)


if __name__ == "__main__":
    app = SmashTeamGenerator()
    app.mainloop()
