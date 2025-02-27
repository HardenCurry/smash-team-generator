import os
from dotenv import load_dotenv
from notion_client import Client
import time
from datetime import datetime
from pprint import pprint
import pandas as pd


class ExcelToNotionImporter:
    def __init__(self, notion_token, parent_page_id, template_page_id, excel_file_path=None, lesson_indices=None):
        # 토큰 및 페이지 ID 설정
        self.NOTION_TOKEN = notion_token
        self.parent_page_id = parent_page_id
        self.template_page_id = template_page_id
        self.table_id_list = []
        self.lesson_indices = lesson_indices
        self.total_people = 0
        # excel 관련
        self.excel_file_path = excel_file_path
        self.excel_data = None

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
        """테이블 블록 업데이트 함수"""

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

        block_list = self.notion.blocks.children.list(self.new_page_id)
        for block in block_list['results']:
            get_all_block(block)

        for i, table_id in enumerate(self.table_id_list):
            cells = []
            all_rows = []
            # 맨위 테이블
            if i == 0:
                # 코트,레슨 코트,총 인원,시간
                basic_info = ["", "", self.total_people, "21:00-23:00"]
                for cell_value in basic_info:
                    # 셀 텍스트 배열 생성
                    cell = [{
                        "type": "text",
                        "text": {
                            "content": str(cell_value)
                        },
                        "annotations": {
                            "bold": False,
                            "code": False,
                            "color": "default",
                            "italic": False,
                            "strikethrough": False,
                            "underline": False
                        }
                    }]
                    cells.append(cell)
                # 테이블에 새 행 추가
                self.notion.blocks.children.append(
                    block_id=table_id,
                    children=[{
                        "object": "block",
                        "type": "table_row",
                        "table_row": {
                            "cells": cells
                        }
                    }]
                )

            # 페어링 sheet 넣기
            elif i == 1 and 'pairing' in self.excel_data:
                df = self.excel_data['pairing']
                print(f"페어링 시트 데이터 추가 중... (총 {len(df)}행)")

                columns = df.columns.tolist()
                # 각 행 추가
                for idx, row in df.iterrows():
                    cells = []

                    # 첫 번째 셀: 인덱스 (1부터 시작)
                    cells.append([{
                        "type": "text",
                        "text": {
                            "content": str(idx + 1)
                        }
                    }])

                    # 레슨생 여부 확인 - 인덱스가 lesson_indices에 있는지 확인
                    is_lesson = idx in self.lesson_indices

                    # 두 번째와 세 번째 셀: Excel 데이터의 두 컬럼
                    for col in columns:
                        cell_value = str(row[col]) if not pd.isna(
                            row[col]) else ""

                        # 레슨생인 경우 노랑색 배경 적용
                        if is_lesson:
                            cells.append([{
                                "type": "text",
                                "text": {
                                    "content": cell_value
                                },
                                "annotations": {
                                    "bold": False,
                                    "code": False,
                                    "color": "yellow_background",
                                    "italic": False,
                                    "strikethrough": False,
                                    "underline": False
                                }
                            }])
                        else:
                            cells.append([{
                                "type": "text",
                                "text": {
                                    "content": cell_value
                                }
                            }])

                    all_rows.append({
                        "object": "block",
                        "type": "table_row",
                        "table_row": {
                            "cells": cells
                        }
                    })
                # 테이블에 새 행 추가
                self.notion.blocks.children.append(
                    block_id=table_id,
                    children=all_rows
                )

                print(f"페어링 테이블 업데이트 완료 ({len(df)} 행 추가)")

            # 조편성 sheet 넣기
            elif i == 2 and 'teams' in self.excel_data:
                df = self.excel_data['teams']
                columns = df.columns.tolist()
                print(f"조편성 시트 데이터 추가 중... (총 {len(df)}행)")

                # 빈 all_rows 배열 초기화
                all_rows = []

                # 각 행 추가
                for idx, row in df.iterrows():
                    cells = []
                    # 첫 번째 셀: 인덱스 (1부터 시작) - 배경색 없음
                    cells.append([{
                        "type": "text",
                        "text": {
                            "content": str(idx + 1)
                        }
                    }])

                    # Excel 데이터의 컬럼들 - 순서에 따라 다른 배경색 적용
                    for col_idx, col in enumerate(columns):
                        cell_value = str(row[col]) if not pd.isna(
                            row[col]) else ""

                        # 컬럼 순서에 따라 다른 배경색 적용
                        if col_idx == 0:  # 첫 번째 데이터 열
                            bg_color = "green_background"
                        elif col_idx == 1:  # 두 번째 데이터 열
                            bg_color = "blue_background"
                        else:  # 세 번째 데이터 열 (및 그 이상)
                            bg_color = "purple_background"

                        cells.append([{
                            "type": "text",
                            "text": {
                                "content": cell_value
                            },
                            "annotations": {
                                "bold": False,
                                "code": False,
                                "color": bg_color,
                                "italic": False,
                                "strikethrough": False,
                                "underline": False
                            }
                        }])

                    all_rows.append({
                        "object": "block",
                        "type": "table_row",
                        "table_row": {
                            "cells": cells
                        }
                    })

                # 테이블에 새 행 추가
                self.notion.blocks.children.append(
                    block_id=table_id,
                    children=all_rows
                )
                print(f"조편성 테이블 업데이트 완료 ({len(df)} 행 추가)")

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
        for block in block_list['results']:
            processed_block = self.process_block(block)
            if processed_block:
                # 부모 table 블록 생성
                self.notion.blocks.children.append(
                    block_id=self.new_page_id,
                    children=[processed_block]
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
    excel_file_path = "팀_구성_결과.xlsx"

    t1 = time.time()
    cloner = ExcelToNotionImporter(
        NOTION_TOKEN, parent_page_id, template_page_id, excel_file_path)
    new_page_url = cloner.duplicate_template_page()
    print(f"새 페이지가 생성되었습니다: {new_page_url}")
    cloner.update_block()
    t2 = time.time()
    print(f"notion_api 총 소요 시간: {t2 - t1:.1f}초") """
