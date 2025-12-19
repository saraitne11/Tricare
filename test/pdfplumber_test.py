import pdfplumber
import pandas as pd


def parse_pdf_tables(pdf_path):
    # PDF 파일 열기
    with pdfplumber.open(pdf_path) as pdf:
        # 첫 번째 페이지 선택 (파일에 따라 페이지 순회 가능)
        page = pdf.pages[0]

        # 1. 테이블 추출 (표 형태 데이터)
        # table_settings를 통해 감지 정확도를 높일 수 있습니다.
        tables = page.extract_tables(table_settings={
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "intersection_y_tolerance": 10,
        })

        print(f"총 {len(tables)}개의 테이블을 찾았습니다.\n")

        for i, table in enumerate(tables):
            # None 값 제거 및 데이터프레임 변환
            clean_table = [[cell.replace('\n', ' ') if cell else '' for cell in row] for row in table]
            df = pd.DataFrame(clean_table)

            print(f"--- Table {i + 1} ---")
            print(df)
            print("\n")

            # 필요시 엑셀로 저장
            # df.to_excel(f"table_{i+1}.xlsx")

        # 2. 텍스트 추출 (표가 아닌 부분 포함)
        # layout=True로 하면 원본 위치를 최대한 보존하여 텍스트를 추출합니다.
        text = page.extract_text(layout=True)
        print("--- Layout Preserved Text ---")
        print(text)


# # 파일 경로 입력
# file_path = 'PT chart_Cho_Sungwoo_June_27_AT-0001500153.pdf'
# parse_pdf_tables(file_path)


if __name__ == '__main__':
    parse_pdf_tables("../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Justin/PT chart_Misheff_Kenai_July_3_AT-0001361137.pdf")