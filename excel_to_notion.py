# import os
# import time
from dotenv import load_dotenv
from notion_client import Client
from datetime import datetime
from pprint import pprint
import pandas as pd
import concurrent.futures  # 병렬 처리를 위한 라이브러리 추가


class ExcelToNotionImporter:
    def __init__(self, notion_token, parent_page_id, template_page_id, excel_file_path=None, lesson_indices=None,
                 max_workers=3, use_concurrent=True):
        # 토큰 및 페이지 ID 설정
        self.NOTION_TOKEN = notion_token
        self.parent_page_id = parent_page_id
        self.template_page_id = template_page_id
        self.table_id_list = []
        self.lesson_indices = lesson_indices or []
        self.total_people = 0
        # excel 관련
        self.excel_file_path = excel_file_path
        self.excel_data = None

        # 성능 최적화 관련 파라미터
        self.max_workers = max_workers  # 병렬 작업자 수
        self.use_concurrent = use_concurrent  # 병렬 처리 사용 여부

        # Excel 데이터 로드 (파일 경로가 제공된 경우)
        if excel_file_path:
            self.load_excel_data()

        # Notion 클라이언트 초기화
        self.notion = Client(auth=self.NOTION_TOKEN)

    def load_excel_data(self):
        """Excel 파일에서 데이터를 읽어오는 함수"""
        try:
            # 페어링과 조편성 시트 모두 로드
            self.excel_data = {
                'pairing': pd.read_excel(self.excel_file_path, sheet_name='페어링'),
                'teams': pd.read_excel(self.excel_file_path, sheet_name='조편성')
            }
            print(
                f"Excel 데이터 로드 완료: 페어링 {self.excel_data['pairing'].shape[0]}행, 조편성 {self.excel_data['teams'].shape[0]}행")
            # 총인원 계산
            for _, row in self.excel_data['pairing'].iterrows():
                # t1과 t2 컬럼에서 비어있지 않은 값의 수를 더함
                if not pd.isna(row['t1']) and row['t1'] != "":
                    self.total_people += 1
                if not pd.isna(row['t2']) and row['t2'] != "":
                    self.total_people += 1
        except Exception as e:
            print(f"Excel 데이터 로드 중 오류 발생: {e}")
            self.excel_data = None

    def process_block(self, block):
        """여러 child 블록 처리 (재귀 함수)"""
        block_type = block['type']
        block_id = block['id']

        # 자식 블록이 있는 경우
        if block['has_children']:
            children_blocks = self.notion.blocks.children.list(block_id)[
                'results']
            children_data = []
            # 각 자식 블록 처리
            for child in children_blocks:
                child_data = self.process_block(child)
                if child_data:
                    children_data.append(child_data)

            # 블록 속성 복사 및 자식 추가
            block_properties = block[block_type].copy()
            block_properties['children'] = children_data
            return {
                "object": "block",
                "type": block_type,
                block_type: block_properties,
            }
        else:
            return {
                "object": "block",
                "type": block_type,
                block_type: block[block_type]
            }

    def update_block(self):
        """테이블 블록 업데이트 함수 (최적화 버전)"""

        def get_all_block(block):
            """새로운 페이지에 있는 모든 블록 조회 & table_id_list 추가"""
            block_id = block['id']
            block_type = block['type']
            if block_type == 'table':
                self.table_id_list.append(block_id)

            if block['has_children']:
                children_blocks = self.notion.blocks.children.list(block_id)[
                    'results']
                for child in children_blocks:
                    get_all_block(child)

        # 모든 테이블 ID 가져오기
        block_list = self.notion.blocks.children.list(self.new_page_id)
        for block in block_list['results']:
            get_all_block(block)

        # 병렬 처리를 위한 함수 정의
        def update_table(table_index):
            table_id = self.table_id_list[table_index]
            if table_index == 0:  # 기본 정보 테이블
                return self._update_basic_info_table(table_id)
            elif table_index == 1 and 'pairing' in self.excel_data:  # 페어링 테이블
                return self._update_pairing_table(table_id)
            elif table_index == 2 and 'teams' in self.excel_data:  # 조편성 테이블
                return self._update_teams_table(table_id)
            return None

        # 병렬 처리 수행 (선택적)
        if self.use_concurrent and len(self.table_id_list) > 1:
            with concurrent.futures.ThreadPoolExecutor(max_workers=min(len(self.table_id_list), self.max_workers)) as executor:
                futures = [executor.submit(update_table, i)
                           for i in range(len(self.table_id_list))]
                concurrent.futures.wait(futures)
        else:
            # 순차 처리
            for i in range(len(self.table_id_list)):
                update_table(i)

    def _update_basic_info_table(self, table_id):
        """기본 정보 테이블 업데이트"""
        basic_info = ["", "", self.total_people, "21:00-23:00"]
        cells = []
        for cell_value in basic_info:
            cell = [{
                "type": "text",
                "text": {"content": str(cell_value)},
                "annotations": {
                    "bold": False, "code": False, "color": "default",
                    "italic": False, "strikethrough": False, "underline": False
                }
            }]
            cells.append(cell)

        self.notion.blocks.children.append(
            block_id=table_id,
            children=[{
                "object": "block",
                "type": "table_row",
                "table_row": {"cells": cells}
            }]
        )
        return "basic_info_updated"

    def _update_pairing_table(self, table_id):
        """페어링 테이블 업데이트"""
        df = self.excel_data['pairing']
        columns = df.columns.tolist()
        print(f"페어링 시트 데이터 추가 중... (총 {len(df)}행)")

        all_rows = []
        for idx, row in df.iterrows():
            cells = []
            # 인덱스 (1부터 시작)
            cells.append([{"type": "text", "text": {"content": str(idx + 1)}}])

            # 레슨생 여부 확인
            is_lesson = idx in self.lesson_indices

            # 컬럼 데이터 추가
            for col in columns:
                cell_value = str(row[col]) if not pd.isna(row[col]) else ""

                # 레슨생인 경우 노랑색 배경 적용
                text_obj = {
                    "type": "text",
                    "text": {"content": cell_value}
                }

                if is_lesson:
                    text_obj["annotations"] = {
                        "bold": False, "code": False, "color": "yellow_background",
                        "italic": False, "strikethrough": False, "underline": False
                    }

                cells.append([text_obj])

            all_rows.append({
                "object": "block",
                "type": "table_row",
                "table_row": {"cells": cells}
            })

        # 한 번에 모든 행 추가
        self.notion.blocks.children.append(
            block_id=table_id,
            children=all_rows
        )
        print(f"페어링 테이블 업데이트 완료 ({len(df)} 행 추가)")
        return "pairing_updated"

    def _update_teams_table(self, table_id):
        """조편성 테이블 업데이트"""
        df = self.excel_data['teams']
        columns = df.columns.tolist()
        print(f"조편성 시트 데이터 추가 중... (총 {len(df)}행)")

        all_rows = []
        for idx, row in df.iterrows():
            cells = []
            # 인덱스 (1부터 시작)
            cells.append([{"type": "text", "text": {"content": str(idx + 1)}}])

            # 컬럼 데이터 추가 (배경색 지정)
            for col_idx, col in enumerate(columns):
                cell_value = str(row[col]) if not pd.isna(row[col]) else ""

                # 컬럼 순서에 따라 다른 배경색 적용
                if col_idx == 0:
                    bg_color = "green_background"
                elif col_idx == 1:
                    bg_color = "blue_background"
                else:
                    bg_color = "purple_background"

                cells.append([{
                    "type": "text",
                    "text": {"content": cell_value},
                    "annotations": {
                        "bold": False, "code": False, "color": bg_color,
                        "italic": False, "strikethrough": False, "underline": False
                    }
                }])

            all_rows.append({
                "object": "block",
                "type": "table_row",
                "table_row": {"cells": cells}
            })

        # 한 번에 모든 행 추가
        self.notion.blocks.children.append(
            block_id=table_id,
            children=all_rows
        )
        print(f"조편성 테이블 업데이트 완료 ({len(df)} 행 추가)")
        return "teams_updated"

    def duplicate_template_page(self):
        # 페이지 생성 로직
        today = datetime.now()
        weekday_ko = ["월", "화", "수", "목", "금", "토", "일"]
        formatted_date = f"{today:%y}.{today.month}.{today.day} ({weekday_ko[today.weekday()]})"

        new_page = self.notion.pages.create(
            parent={
                "type": "page_id",
                "page_id": self.parent_page_id
            },
            properties={
                "title": {
                    "title": [
                        {
                            "text": {"content": formatted_date}
                        }
                    ]
                }
            }
        )
        self.new_page_id = new_page['id']

        # 템플릿 페이지 child 블록 가져오기
        block_list = self.notion.blocks.children.list(self.template_page_id)

        # 모든 블록을 한 번에 처리하기 위한 리스트
        all_blocks = []
        for block in block_list['results']:
            processed_block = self.process_block(block)
            if processed_block:
                all_blocks.append(processed_block)

        # 모든 블록을 한 번에 추가 (API 호출 최소화)
        if all_blocks:
            self.notion.blocks.children.append(
                block_id=self.new_page_id,
                children=all_blocks
            )

        new_page_url = f"https://notion.so/{self.new_page_id.replace('-', '')}"
        return new_page_url


if __name__ == "__main__":
    """ # .env 파일 읽기
    load_dotenv()
    # 스매시 페이지 토큰값
    NOTION_TOKEN = os.getenv("NOTION_TOKEN")
    if not NOTION_TOKEN:
        raise ValueError("NOTION_TOKEN이 .env에 설정되지 않았습니다.")
    # 조 편성 부모 페이지
    parent_page_id = "1a39f0ea074e80e085a5dbe8bfa5404f"
    # 조 편성 양식 템플릿
    template_page_id = "1a79f0ea074e807b9b18c586b0890893"
    excel_file_path = "team_composition_result.xlsx"

    t1 = time.time()
    cloner = ExcelToNotionImporter(
        NOTION_TOKEN, parent_page_id, template_page_id, excel_file_path)
    new_page_url = cloner.duplicate_template_page()
    print(f"새 페이지가 생성되었습니다: {new_page_url}")
    cloner.update_block()
    t2 = time.time()
    print(f"notion_api 총 소요 시간: {t2 - t1:.1f}초") """
