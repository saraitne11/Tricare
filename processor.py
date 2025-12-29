import datetime as dt
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple
import json
import argparse

import fitz
import pandas as pd

HEADER = re.compile(r"Dr\.?\s*Joung[`'’]?s\s*Clinic\s*&\s*Physical\s*Therapy\s*Center")
VISIT_NO = re.compile(r"#\s*([0-9]+\s*(?:/\s*[0-9]+)?)")
AUTH_NO = re.compile(r"\(\s*(AT-[^)]+)\s*\)")

REQUIRED_EXCEL_COLS = [
    "Weekly pt. tx list",
    "Date of birth",
    "Diagnosis",
    "Authorization number",
    "Date of Therapy",
]

target_fields = [
    {"name": "Patient Name", "pattern": r"\s*Patient\s*Name\s*"},
    {"name": "DOB", "pattern": r"\s*DOB\s*"},
    {"name": "Diagnosis/CC", "pattern": r"\s*Diagnosis/CC\s*"},
    {"name": "Therapist", "pattern": r"\s*Therapist\s*"},
    {"name": "DOS", "pattern": r"\s*DOS\s*"},
    {"name": "Visit No.", "pattern": r"\s*Visit\s*No\s*"}
]


def _ensure_abs_path(path_str: str, kind: str) -> Path:
    path = Path(path_str)
    if not path.is_absolute():
        raise ValueError(f"{kind} 경로는 절대경로로 입력해야 합니다: {path_str}")
    if kind == "file" and not path.is_file():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path_str}")
    if kind == "dir" and not path.is_dir():
        raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {path_str}")
    return path


def convert_dob(date_txt) -> str:
    for f in ("%B %d, %Y", "%b %d, %Y"):
        try:
            d = dt.datetime.strptime(date_txt, f)
            return d.strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def convert_dos(date_str) -> str:
    try:
        d = dt.datetime.strptime(date_str, "%m/%d/%Y")
        return d.strftime("%Y-%m-%d")
    except ValueError:
        return ""


def normalize_spaces(text: pd.Series) -> pd.Series:
    is_na = text.isna()
    cleaned = text.astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    return cleaned.where(~is_na, "")


def _read_excel_with_header_detection(
    input_path: Path, sheet_name: str, usecols: str
) -> pd.DataFrame:
    """
    의미 없는 상단 행이 있어도 실제 헤더가 있는 행을 찾아서 DataFrame을 생성한다.
    요구 헤더: REQUIRED_EXCEL_COLS
    """
    df_raw = pd.read_excel(input_path, sheet_name=sheet_name, usecols=usecols, header=None)

    def _is_header_row(row: pd.Series) -> bool:
        values = [str(v).strip().lower() for v in row.tolist() if pd.notna(v)]
        return all(any(val == req.lower() for val in values) for req in REQUIRED_EXCEL_COLS)

    header_idx = None
    for idx, row in df_raw.iterrows():
        if _is_header_row(row):
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError("엑셀에서 필요한 헤더 행을 찾지 못했습니다. (Weekly pt. tx list 등)")

    header_values = [str(v).strip() if pd.notna(v) else "" for v in df_raw.iloc[header_idx].tolist()]
    df_excel = df_raw.iloc[header_idx + 1 :].copy()
    df_excel.columns = header_values
    df_excel = df_excel.dropna(how="all")  # 전체 빈 행 제거
    return df_excel


def extract_data(df: pd.DataFrame, pdf_path: str) -> dict[Any, Any] | None:
    row_data = {}

    labels = df.iloc[:, 0]
    values = df.iloc[:, 1]
    labels = labels.astype(str).str.strip().str.replace('\n', '', regex=False)
    values = values.astype(str).str.strip().str.replace('\n', '', regex=False)

    for field in target_fields:
        mask = labels.str.contains(field["pattern"], case=False, na=False, regex=True)
        if mask.any():
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
            row_data[field["name"]] = None

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
        for table in tables:
            df = table.to_pandas()
            col_indices = []
            for k, col in enumerate(df.columns):
                if HEADER.search(col):
                    col_indices.append(k)

            if len(col_indices) >= 2:
                for c in range(len(col_indices) - 1):
                    split_df = df.iloc[:, col_indices[c]:col_indices[c + 1]]
                    d = extract_data(split_df, pdf_path)
                    if d:
                        data_list.append(d)
                split_df = df.iloc[:, col_indices[-1]:]
                d = extract_data(split_df, pdf_path)
                if d:
                    data_list.append(d)
            else:
                d = extract_data(df, pdf_path)
                if d:
                    data_list.append(d)

    return data_list


def run_matching(
    pdf_dir: str | list[str] | None,
    input_xlsx: str | None,
    sheet_name: str,
    columns: str,
    progress_cb: Callable[[int, int, Path, int], None] | None = None,
    stop_flag: Callable[[], bool] | None = None,
) -> Tuple[pd.DataFrame, pd.DataFrame, int]:
    if not pdf_dir:
        raise ValueError("PDF 폴더 경로가 필요합니다.")
    if not input_xlsx:
        raise ValueError("입력 엑셀 경로가 필요합니다.")

    pdf_dirs: List[Path] = []
    if isinstance(pdf_dir, (list, tuple)):
        pdf_dirs = [_ensure_abs_path(p, "dir") for p in pdf_dir if p]
    else:
        pdf_dirs = [_ensure_abs_path(pdf_dir, "dir")]
    if not pdf_dirs:
        raise ValueError("PDF 폴더 경로가 필요합니다.")

    input_path = _ensure_abs_path(input_xlsx, "file")

    pdf_list: List[Path] = []
    for root_dir in pdf_dirs:
        pdf_list.extend(root_dir.rglob("*.pdf"))
    total = len(pdf_list)
    data: List[Dict[str, Any]] = []

    for i, file in enumerate(pdf_list):
        if stop_flag and stop_flag():
            raise RuntimeError("사용자 중지")
        rows = parse_pdf(str(file))
        data.extend(rows)
        if progress_cb:
            progress_cb(i + 1, total, file, len(rows))

    df_pdf = pd.DataFrame(data)

    # Drop unused columns after matching
    for col in ("Times", "Therapist"):
        if col in df_pdf.columns:
            df_pdf = df_pdf.drop(columns=col)

    for col in ["Patient Name", "Diagnosis/CC", "Authorization No"]:
        if col in df_pdf.columns:
            df_pdf[col] = normalize_spaces(df_pdf[col])

    df_excel = _read_excel_with_header_detection(input_path, sheet_name=sheet_name, usecols=columns)
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
            # (df_pdf["Authorization No"] == row["Authorization number"]) &
            (df_pdf["DOS"] == row["Date of Therapy"])
        ]
        if len(retrieves) == 1:
            retrieve = retrieves.iloc[0]
            df_excel.loc[idx, "Visit No"] = retrieve["Visit No"]
            df_excel.loc[idx, "File"] = retrieve["File"]
            cnt += 1

    # Drop unused columns and reorder for final output
    drop_cols = [c for c in ("Times", "Therapist") if c in df_excel.columns]
    if drop_cols:
        df_excel = df_excel.drop(columns=drop_cols)

    # Format date columns for final output
    if "Date of birth" in df_excel.columns:
        dob_dt = pd.to_datetime(df_excel["Date of birth"], errors="coerce")
        df_excel["Date of birth"] = (
            dob_dt.dt.strftime("%b %d, %Y")
            .str.replace(r"^([A-Za-z]+) 0", r"\1 ", regex=True)
        )
    if "Date of Therapy" in df_excel.columns:
        dos_dt = pd.to_datetime(df_excel["Date of Therapy"], errors="coerce")
        df_excel["Date of Therapy"] = dos_dt.dt.strftime("%Y.%m.%d")

    desired_order = [
        "Weekly pt. tx list",
        "Authorization number",
        "Date of Therapy",
        "Visit No",
        "Diagnosis",
        "Date of birth",
        "File",
    ]
    existing_cols = [c for c in desired_order if c in df_excel.columns]
    if existing_cols:
        df_excel = df_excel[existing_cols]

    return df_pdf, df_excel, cnt


def _cli() -> None:
    parser = argparse.ArgumentParser(
        description="단일 PDF 파싱 테스트",
        epilog="예시:\n  python processor.py C:\\\\path\\\\to\\\\chart.pdf",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("pdf_path", help="테스트할 PDF 절대경로")
    args = parser.parse_args()

    pdf_path = Path(args.pdf_path)
    if not pdf_path.is_absolute():
        raise SystemExit("pdf_path는 절대경로로 입력하세요.")
    if not pdf_path.is_file():
        raise SystemExit(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")

    rows = parse_pdf(str(pdf_path))
    print(f"파일: {pdf_path}")
    print(f"추출 건수: {len(rows)}")
    for i, row in enumerate(rows):
        print(f"\n--- Row {i+1} ---")
        print(json.dumps(row, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _cli()

