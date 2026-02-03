import pdfplumber
import pandas as pd


TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "intersection_y_tolerance": 10,
}


def parse_with_pdfplumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]

        tables = page.extract_tables(table_settings=TABLE_SETTINGS) or []
        print(f"총 {len(tables)}개의 테이블을 감지했습니다.\n")

        for i, table in enumerate(tables):
            df = pd.DataFrame(table)
            if df.empty:
                print(f"--- Table {i + 1} (empty) ---\n")
                continue

            df = df.fillna("")
            df = df.applymap(lambda v: v.replace("\n", " ").strip() if isinstance(v, str) else v)
            header_row = df.iloc[0].astype(str)
            df = df.iloc[1:].reset_index(drop=True)
            df.columns = header_row

            print(f"--- Table {i + 1} ---")
            print(df)
            print("\n")

        text = page.extract_text(layout=True)
        print("--- Layout Preserved Text ---")
        print(text)


if __name__ == '__main__':
    # file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Justin/PT chart_Misheff_Kenai_July_3_AT-0001361137.pdf"  # Validity Date (s) 부분 2줄
    file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Soyeon/PT chart_Kinkade_Jeremy_July_1_3.pdf"  # 테이블 2개
    # file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Kyo/PT chart_Masson_Jonathan_July_2_AT-0001489716.pdf"    # Diagnosis 2줄
    parse_with_pdfplumber(file_path)