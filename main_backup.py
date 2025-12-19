# pip install pdfplumber pandas openpyxl
import datetime as dt
import os
import re

from pathlib import Path
import pandas as pd
import pdfplumber

# --- 사용자 설정 -------------------------------------------------
PDF_DRI = r"C:\Users\sarai\Downloads\PT chart (Jun. 30 - Jul. 5)(200개)(249)"   # PDF가 들어 있는 최상위 폴더
INPUT_XLSX = r"C:\Users\sarai\Downloads\pt list.xlsx"
SHEET_NAME = "Jun 30 _ Jul 5"
COLUMNS = "A:G"

PT_LIST_MERGE_XLSX = r"pt_list_merge.xlsx"
PDF_SUMMARY_XLSX = r"pdf_summary.xlsx"
# ----------------------------------------------------------------

# 라벨별 정규식
PAT_PATIENT = re.compile(r"Patient Name\s+([^\n]+)")
PAT_DOB = re.compile(r"DOB\.\s+([^\n]+)")
PAT_DIAG = re.compile(r"Diagnosis/CC\s+([^\n]+)")
PAT_DOS = re.compile(r"DOS\.\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})")
PAT_VISIT = re.compile(r"Visit No\.\s*#\s*([0-9\s]+/\s*[0-9]+)\s*\(\s*(AT-[^)]+)\s*\)")


def convert_dob(date_txt):
    """'April 5, 1973' ➜ '1973. 4. 5'"""
    for f in ("%B %d, %Y", "%b %d, %Y"):
        try:
            d = dt.datetime.strptime(date_txt, f)
            return d.strftime("%Y-%m-%d")
        except ValueError:
            pass
    return ""


def convert_dos(date_str):
    """'6/24/2025' ➜ '2025. 06. 24'"""
    try:
        # 월/일/년 형식 파싱
        d = dt.datetime.strptime(date_str, "%m/%d/%Y")
        return d.strftime("%Y-%m-%d")
    except ValueError:
        return ""


def strip_dup(raw, label):
    """'…Patient Name Baek…' → 첫 번째 값만 보존"""
    return raw.split(label)[0].strip()


def parse_pdf(path):
    r = []
    with pdfplumber.open(path) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    # 공통 필드
    m_name = PAT_PATIENT.search(text)
    m_dob = PAT_DOB.search(text)
    m_diag = PAT_DIAG.search(text)
    if not (m_name and m_dob and m_diag):
        # 필수 필드 없으면 파일 경로, 파일 명만
        r.append([path, os.path.basename(path)])
        return r

    name = strip_dup(m_name.group(1), "Patient Name")
    dob = strip_dup(m_dob.group(1), "DOB.")
    diag = strip_dup(m_diag.group(1), "Diagnosis/CC")

    # ▶ DOS·Visit를 각각 전부 수집
    dos_list = PAT_DOS.findall(text)
    visit_list = PAT_VISIT.findall(text)

    # ▶ 두 리스트를 순서대로 짝지워 한 행 생성
    for dos, (vis_raw, code) in zip(dos_list, visit_list):
        r.append([
            path,
            os.path.basename(path),
            name,
            convert_dob(dob),
            diag,
            code,
            convert_dos(dos),

            vis_raw,   # '3 / 40' ➜ '3/40'

        ])
    return r


# === PDF 데이터 추출 ===
root_dir = Path(PDF_DRI)
pdf_list = list(root_dir.rglob("*.pdf"))

data = []
for i, file in enumerate(pdf_list):
    rows = parse_pdf(file)
    data.extend(rows)
    print(f"%d/%d, File: %s, rows: %d" % (i+1, len(pdf_list), file, len(rows)))

pdfs = pd.DataFrame(data, columns=[
    "Full Path",
    "File",
    "Weekly pt. tx list (PDF)",
    "Date of birth (PDF)",
    "Diagnosis (PDF)",
    "Authorization number (PDF)",
    "Date of Therapy (PDF)",
    "Visit No (#/total) (PDF)"
])
pdfs.to_excel(PDF_SUMMARY_XLSX, index=False)
print(f"PDF 데이터 추출 완료, PDF 파일 {len(pdf_list)}건, 추출된 데이터 {len(pdfs)}건 → {PT_LIST_MERGE_XLSX}")


# === 원본 엑셀과 PDF 추출 데이터 매칭 ===
df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME, usecols=COLUMNS)
df["Weekly pt. tx list"] = df["Weekly pt. tx list"].str.strip()
df["Diagnosis"] = df["Diagnosis"].str.strip()
df["Authorization number"] = df["Authorization number"].str.strip()
df["Date of birth"] = pd.to_datetime(df["Date of birth"]).dt.strftime("%Y-%m-%d")
df["Date of Therapy"] = pd.to_datetime(df["Date of Therapy"]).dt.strftime("%Y-%m-%d")
df["Visit No"] = ""
df["File"] = ""
cnt = 0
for idx, row in df.iterrows():
    retrieves = pdfs[
        (pdfs["Weekly pt. tx list (PDF)"] == row["Weekly pt. tx list"]) &
        (pdfs["Date of birth (PDF)"] == row["Date of birth"]) &
        (pdfs["Diagnosis (PDF)"] == row["Diagnosis"]) &
        (pdfs["Authorization number (PDF)"] == row["Authorization number"]) &
        (pdfs["Date of Therapy (PDF)"] == row["Date of Therapy"])
    ]

    if len(retrieves) == 1:
        retrieve = retrieves.iloc[0]
        df.loc[idx, "Visit No"] = retrieve["Visit No (#/total) (PDF)"]
        df.loc[idx, "File"] = retrieve["Full Path"]
        cnt += 1

df.to_excel(PT_LIST_MERGE_XLSX, index=False)
print(f"원본 엑셀 {len(df)}건, 매칭된 데이터 {cnt}건 → {PT_LIST_MERGE_XLSX}")
