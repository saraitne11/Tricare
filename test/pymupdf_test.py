import fitz  # PyMuPDF
import pandas as pd


def parse_with_pymupdf(pdf_path):
    # PDF 문서 열기
    doc = fitz.open(pdf_path)
    page = doc[0]  # 첫 번째 페이지

    # 1. 테이블 찾기 (기본 설정으로 감지)
    # vertical_strategy, horizontal_strategy 등을 자동 처리하지만,
    # 필요하다면 세부 설정도 가능합니다.
    tables = page.find_tables()

    print(f"총 {len(tables.tables)}개의 테이블을 감지했습니다.\n")

    for i, table in enumerate(tables):
        print(f"--- Table {i + 1} ---")

        # PyMuPDF의 테이블 객체를 바로 Pandas DataFrame으로 변환 가능
        df = table.to_pandas()

        # 데이터프레임 출력
        print(df)
        print("\n")

    # 2. 텍스트 블록 추출 (위치 정보 포함)
    # 표가 아닌 일반 텍스트를 위치 좌표와 함께 가져올 때 유용합니다.
    text_blocks = page.get_text("blocks")
    print("--- Text Blocks Sample (Top 3) ---")
    for block in text_blocks[:3]:
        print(block)  # (x0, y0, x1, y1, text, block_no, block_type)


if __name__ == '__main__':
    # file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Justin/PT chart_Misheff_Kenai_July_3_AT-0001361137.pdf"  # Validity Date (s) 부분 2줄
    file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Soyeon/PT chart_Kinkade_Jeremy_July_1_3.pdf"  # 테이블 2개
    # file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Kyo/PT chart_Masson_Jonathan_July_2_AT-0001489716.pdf"    # Diagnosis 2줄
    parse_with_pymupdf(file_path)