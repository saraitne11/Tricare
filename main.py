import re
import fitz
import pandas as pd
import datetime as dt

from pathlib import Path

from typing import List, Dict, Any


# --- 사용자 설정 -------------------------------------------------
PDF_DRI = r"C:\Users\10391\Documents\Workspaces\Tricare\data\pt-charts\PT chart (Jun. 30 - Jul. 5)(200개)(249)".replace("\\", "/")  # PDF가 들어 있는 최상위 폴더
INPUT_XLSX = r"C:\Users\10391\Documents\Workspaces\Tricare\data\pt-list\pt list.xlsx".replace("\\", "/")
SHEET_NAME = "Jun 30 _ Jul 5"
COLUMNS = "A:G"

PT_LIST_MERGE_XLSX = r"pt_list_merge.xlsx"
PDF_SUMMARY_XLSX = r"pdf_summary.xlsx"
# ----------------------------------------------------------------


HEADER = re.compile(r"Dr\.?\s*Joung[`'’]?s\s*Clinic\s*&\s*Physical\s*Therapy\s*Center")
VISIT_NO = re.compile(r"#\s*([0-9\s]+/\s*[0-9]+)")
AUTH_NO = re.compile(r"\(\s*(AT-[^)]+)\s*\)")

target_fields = [
    {"name": "Patient Name", "pattern": r"\s*Patient\s*Name\s*"},
    {"name": "DOB", "pattern": r"\s*DOB\s*"},
    {"name": "Diagnosis/CC", "pattern": r"\s*Diagnosis/CC\s*"},
    {"name": "Therapist", "pattern": r"\s*Therapist\s*"},
    {"name": "DOS", "pattern": r"\s*DOS\s*"},
    {"name": "Visit No.", "pattern": r"\s*Visit\s*No\s*"}
]


def convert_dob(date_txt) -> str:
    """'April 5, 1973' ➜ '1973. 4. 5'"""
    for f in ("%B %d, %Y", "%b %d, %Y"):
        try:
            d = dt.datetime.strptime(date_txt, f)
            return d.strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def convert_dos(date_str) -> str:
    """'6/24/2025' ➜ '2025. 06. 24'"""
    try:
        # 월/일/년 형식 파싱
        d = dt.datetime.strptime(date_str, "%m/%d/%Y")
        return d.strftime("%Y-%m-%d")
    except ValueError:
        return ""


def normalize_spaces(text: pd.Series) -> pd.Series:
    """
    Collapse whitespace (tabs/newlines/multiple spaces) to a single space, strip ends,
    and keep original NaN as missing (avoid turning them into the literal 'nan').
    """
    is_na = text.isna()
    cleaned = text.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    return cleaned.where(~is_na, "")


def extract_data(df: pd.DataFrame, pdf_path: str) -> dict[Any, Any] | None:
    # 한 명(또는 한 차트)의 정보를 담을 딕셔너리
    row_data = {}

    labels = df.iloc[:, 0]
    values = df.iloc[:, 1]
    labels = labels.astype(str).str.strip().str.replace('\n', '', regex=False)
    values = values.astype(str).str.strip().str.replace('\n', '', regex=False)

    # 타겟 필드 하나씩 순회하며 값 찾기
    for field in target_fields:
        # 라벨 컬럼에서 해당 키워드를 포함하는 행 찾기 (대소문자 무시)
        # na=False: NaN 값 있어도 에러 안 나게 처리
        mask = labels.str.contains(field["pattern"], case=False, na=False, regex=True)

        # 매칭되는 행이 하나라도 있으면 값 가져오기
        if mask.any():
            # mask가 True인 첫 번째 행의 값을 가져옴
            extracted_value = values[mask].iloc[0]
            if field["name"] == "DOB":
                row_data[field["name"]] = convert_dob(extracted_value)
            elif field["name"] == "DOS":
                row_data[field["name"]] = convert_dos(extracted_value)
            elif field["name"] == "Visit No.":
                visit_match = VISIT_NO.search(extracted_value)
                auth_match = AUTH_NO.search(extracted_value)
                row_data["Visit No"] = visit_match.group(1).strip() if visit_match else None
                row_data["Authorization No"] = auth_match.group(1).strip() if auth_match else None
            else:
                row_data[field["name"]] = extracted_value
        else:
            row_data[field["name"]] = None  # 없으면 비워둠

    # 유효한 데이터가 있으면(예: 이름이 있으면) 리스트에 추가
    if row_data.get("Patient Name"):
        row_data["File"] = pdf_path
        return row_data
    else:
        return None


def parse_pdf(pdf_path: str) -> List[Dict[str, Any]]:
    doc = fitz.open(pdf_path)
    data_list = []
    for i in range(doc.page_count):
        page = doc[i]

        tables = page.find_tables()
        for j, table in enumerate(tables):
            df = table.to_pandas()

            # 한 페이지에 테이블이 좌우로 두개 추출되는지 검사
            col_indices = []
            for k, col in enumerate(df.columns):
                if HEADER.search(col):
                    col_indices.append(k)

            if len(col_indices) >= 2:
                # 테이블이 2개 이상 추출됐을 경우
                for c in range(len(col_indices)-1):
                    # 좌측 테이블부터 순서대로 처리
                    split_df = df.iloc[:, col_indices[c]:col_indices[c+1]]
                    d = extract_data(split_df, pdf_path)
                    if d:
                        data_list.append(d)
                # 마지막 우측 테이블 처리
                split_df = df.iloc[:, col_indices[-1]:]
                d = extract_data(split_df, pdf_path)
                if d:
                    data_list.append(d)
            else:
                # 테이블이 1개만 있을 경우
                d = extract_data(df, pdf_path)
                if d:
                    data_list.append(d)

    return data_list



if __name__ == '__main__':

    root_dir = Path(PDF_DRI)
    pdf_list = list(root_dir.rglob("*.pdf"))

    data = []
    for i, file in enumerate(pdf_list):
        rows = parse_pdf(file)
        data.extend(rows)
        print(f"%d/%d, File: %s, rows: %d" % (i + 1, len(pdf_list), file, len(rows)))

    df_pdf = pd.DataFrame(data)
    # 공백/개행/탭을 통일해서 매칭 실패를 줄인다
    for col in ["Patient Name", "Diagnosis/CC", "Authorization No"]:
        if col in df_pdf.columns:
            df_pdf[col] = normalize_spaces(df_pdf[col])

    df_excel = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME, usecols=COLUMNS)
    df_excel["Weekly pt. tx list"] = normalize_spaces(df_excel["Weekly pt. tx list"])
    df_excel["Diagnosis"] = normalize_spaces(df_excel["Diagnosis"])
    df_excel["Authorization number"] = normalize_spaces(df_excel["Authorization number"])
    df_excel["Date of birth"] = pd.to_datetime(df_excel["Date of birth"]).dt.strftime("%Y-%m-%d")
    df_excel["Date of Therapy"] = pd.to_datetime(df_excel["Date of Therapy"]).dt.strftime("%Y-%m-%d")
    df_excel["Visit No"] = ""
    df_excel["File"] = ""
    cnt = 0

    for idx, row in df_excel.iterrows():
        retrieves = df_pdf[
            (df_pdf["Patient Name"] == row["Weekly pt. tx list"]) &
            (df_pdf["DOB"] == row["Date of birth"]) &
            (df_pdf["Diagnosis/CC"] == row["Diagnosis"]) &
            (df_pdf["Authorization No"] == row["Authorization number"]) &
            (df_pdf["DOS"] == row["Date of Therapy"])
            ]

        if len(retrieves) == 1:
            retrieve = retrieves.iloc[0]
            df_excel.loc[idx, "Visit No"] = retrieve["Visit No"]
            df_excel.loc[idx, "File"] = retrieve["File"]
            cnt += 1


    df_pdf.to_excel(PDF_SUMMARY_XLSX, index=False)
    print(f"PDF 데이터 {len(df_pdf)}건 → {PDF_SUMMARY_XLSX}")

    df_excel.to_excel(PT_LIST_MERGE_XLSX, index=False)
    print(f"원본 엑셀 {len(df_excel)}건, 매칭된 데이터 {cnt}건 → {PT_LIST_MERGE_XLSX}")

    # file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Justin/PT chart_Misheff_Kenai_July_3_AT-0001361137.pdf"  # Validity Date (s) 부분 2줄
    # file_path = "PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Soyeon/PT chart_Kinkade_Jeremy_July_1_3.pdf"  # 테이블 2개
    # file_path = "../PT chart (Jun. 30 - Jul. 5)(200개)(249)/PT chart - Kyo/PT chart_Masson_Jonathan_July_2_AT-0001489716.pdf"    # Diagnosis 2줄
    # print(parse_pdf(file_path))