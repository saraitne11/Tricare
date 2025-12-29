import io
import datetime as dt
import tempfile
import zipfile
import shutil
from pathlib import Path

import pandas as pd
import streamlit as st

from processor import run_matching


def _to_excel_bytes(df: pd.DataFrame) -> bytes:
    """DataFrame을 엑셀 바이너리로 변환."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.read()


def _render_results(df_pdf: pd.DataFrame, df_excel: pd.DataFrame, matched: int, ts: str | None = None) -> None:
    st.subheader("결과 요약")
    n_pdf_files = df_pdf["File"].nunique() if "File" in df_pdf.columns else len(df_pdf)
    col1, col2, col3, col4 = st.columns(4)
    col1.metric(label="PDF 파일 수", value=n_pdf_files)
    col2.metric(label="PDF 추출 건수", value=len(df_pdf))
    col3.metric(label="엑셀 행 수", value=len(df_excel))
    col4.metric(label="매칭 성공 건수", value=matched)

    st.divider()
    st.subheader("PDF 추출 결과 미리보기")
    st.dataframe(df_pdf.head(200), use_container_width=True, height=400)

    st.subheader("병합된 엑셀 미리보기")
    st.dataframe(df_excel.head(200), use_container_width=True, height=400)

    st.divider()
    st.subheader("다운로드")
    pdf_bytes = _to_excel_bytes(df_pdf)
    excel_bytes = _to_excel_bytes(df_excel)
    ts = ts or dt.datetime.now().strftime("%Y%m%d%H%M%S")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button(
            "PDF 요약 엑셀 다운로드",
            data=pdf_bytes,
            file_name=f"pdf_summary_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_d2:
        st.download_button(
            "병합 엑셀 다운로드",
            data=excel_bytes,
            file_name=f"pt_list_merge_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def _cleanup_temp_dirs():
    temp_dirs = st.session_state.get("temp_dirs", [])
    for d in temp_dirs:
        try:
            shutil.rmtree(d, ignore_errors=True)
        except Exception:
            pass
    st.session_state["temp_dirs"] = []


def _extract_zip_recursive(zip_file: Path, dest_dir: Path, max_depth: int = 3) -> None:
    """ZIP 내부에 중첩된 ZIP이 있을 때까지 안전하게 풀어준다."""
    dest_dir = dest_dir.resolve()
    queue: list[tuple[Path, int]] = [(zip_file, 0)]

    def _safe_extract(zf: zipfile.ZipFile) -> None:
        for info in zf.infolist():
            target = (dest_dir / info.filename).resolve()
            if not str(target).startswith(str(dest_dir)):
                raise ValueError("ZIP 내 경로가 대상 폴더 밖으로 벗어납니다.")
        zf.extractall(dest_dir)

    while queue:
        src, depth = queue.pop()
        if depth > max_depth:
            raise ValueError("ZIP 중첩 깊이 한도를 초과했습니다.")
        with zipfile.ZipFile(src) as zf:
            _safe_extract(zf)
            for info in zf.infolist():
                if info.filename.lower().endswith(".zip"):
                    nested = dest_dir / info.filename
                    if nested.is_file():
                        queue.append((nested, depth + 1))


def main():
    st.set_page_config(page_title="PT 차트 매칭 도구", layout="wide")
    st.title("PT 차트 매칭 도구")
    st.write("왼쪽 사이드바에서 PDF/엑셀 입력 방식을 선택한 뒤 실행하세요.")

    default_sheet = "Jun 30 _ Jul 5"
    default_cols = "A:G"

    if "temp_dirs" not in st.session_state:
        st.session_state.temp_dirs = []
    if "pdf_mode" not in st.session_state:
        st.session_state.pdf_mode = "ZIP 업로드"

    with st.sidebar:
        st.header("입력 설정")
        st.caption("PDF와 엑셀 입력 방식을 선택하세요.")

        pdf_mode = st.radio(
            "PDF 입력 방식",
            ["ZIP 업로드", "경로 스캔 (재귀)", "절대경로 입력"],
            index=0,
            key="pdf_mode",
        )
        pdf_zip = None
        pdf_dir_input = None
        if pdf_mode == "경로 스캔 (재귀)":
            st.text_area(
                "PDF 루트 경로 (여러 줄 입력 가능)",
                value=st.session_state.get("pdf_paths_raw", ""),
                placeholder=r"C:\data\pt\pdf_root",
                key="pdf_paths_raw",
                height=90,
                help="여러 경로를 입력하면 모두 재귀적으로 스캔합니다.",
            )
        elif pdf_mode == "ZIP 업로드":
            pdf_zip = st.file_uploader(
                "PDF ZIP 업로드",
                type=["zip"],
                accept_multiple_files=False,
                key="pdf_zip_upload",
                help="ZIP 내 모든 PDF를 자동으로 풀어 처리합니다.",
            )
        else:
            pdf_dir_input = st.text_input(
                "PDF 폴더 절대경로 (기존 방식)",
                value=st.session_state.get("pdf_dir_input", ""),
                key="pdf_dir_input",
                placeholder=r"C:\data\pt\pdf_root",
            )

        excel_mode = st.radio(
            "엑셀 입력 방식",
            ["파일 업로드", "절대경로 입력"],
            horizontal=True,
            key="excel_mode",
        )
        excel_file = None
        excel_path_input = None
        if excel_mode == "파일 업로드":
            excel_file = st.file_uploader(
                "입력 엑셀 업로드",
                type=["xlsx", "xls"],
                accept_multiple_files=False,
                key="excel_file_upload",
            )
        else:
            excel_path_input = st.text_input(
                "입력 엑셀 절대경로 (기존 방식)",
                value=st.session_state.get("excel_path_input", ""),
                key="excel_path_input",
                placeholder=r"C:\data\pt\input.xlsx",
            )

        sheet_name = st.text_input("시트명", value=default_sheet, key="sheet_name")
        columns = st.text_input("열 범위 (예: A:G)", value=default_cols, key="columns")

        col_run, col_stop = st.columns(2)
        run_clicked = col_run.button("실행", type="primary", use_container_width=True)
        stop_clicked = col_stop.button("중지", type="secondary", use_container_width=True)

    if "results" not in st.session_state:
        st.session_state.results = None
    if "log_lines" not in st.session_state:
        st.session_state.log_lines = []
    if "stop_requested" not in st.session_state:
        st.session_state.stop_requested = False
    if "run_ts" not in st.session_state:
        st.session_state.run_ts = ""
    if "pdf_paths_raw" not in st.session_state:
        st.session_state.pdf_paths_raw = ""

    if stop_clicked:
        st.session_state.stop_requested = True
        st.warning("중지 요청됨: 다음 작업 지점에서 중단합니다.")

    if run_clicked:
        # 이전 임시 폴더 정리
        _cleanup_temp_dirs()

        resolved_pdf_dir: str | list[str] | None = None
        resolved_xlsx: str | None = None

        if pdf_mode == "경로 스캔 (재귀)":
            pdf_lines = [p.strip() for p in st.session_state.get("pdf_paths_raw", "").splitlines() if p.strip()]
            if not pdf_lines:
                st.error("PDF 루트 경로를 한 줄 이상 입력해 주세요.")
                return
            resolved_pdf_dir = pdf_lines if len(pdf_lines) > 1 else pdf_lines[0]
        elif pdf_mode == "ZIP 업로드":
            if not pdf_zip:
                st.error("PDF ZIP 파일을 업로드해 주세요.")
                return
            tmp_dir = Path(tempfile.mkdtemp(prefix="tricare_pdf_zip_"))
            try:
                upload_zip_path = tmp_dir / (pdf_zip.name or "upload.zip")
                upload_zip_path.write_bytes(pdf_zip.getvalue())
                _extract_zip_recursive(upload_zip_path, tmp_dir, max_depth=5)
            except (zipfile.BadZipFile, ValueError) as e:
                st.error(f"ZIP 해제 실패: {e}")
                shutil.rmtree(tmp_dir, ignore_errors=True)
                return
            resolved_pdf_dir = str(tmp_dir)
            st.session_state.temp_dirs.append(tmp_dir)
        else:
            pdf_dir_value = st.session_state.get("pdf_dir_input", "")
            if not pdf_dir_value:
                st.error("PDF 폴더 절대경로를 입력해 주세요.")
                return
            resolved_pdf_dir = pdf_dir_value

        if excel_mode == "파일 업로드":
            if not excel_file:
                st.error("입력 엑셀 파일을 업로드해 주세요.")
                return
            tmp_dir = Path(tempfile.mkdtemp(prefix="tricare_excel_"))
            excel_path = tmp_dir / excel_file.name
            excel_path.write_bytes(excel_file.getvalue())
            resolved_xlsx = str(excel_path)
            st.session_state.temp_dirs.append(tmp_dir)
        else:
            xlsx_value = st.session_state.get("excel_path_input", "")
            if not xlsx_value:
                st.error("입력 엑셀 절대경로를 입력해 주세요.")
                return
            resolved_xlsx = xlsx_value

        st.session_state.log_lines = []
        st.session_state.results = None
        st.session_state.run_ts = dt.datetime.now().strftime("%Y%m%d%H%M%S")
        st.session_state.stop_requested = False

        progress = st.progress(0.0, text="대기 중")
        status_box = st.expander("로그 메시지", expanded=False)
        status_log_placeholder = status_box.empty()

        def render_status_logs():
            html = """
            <div style="height:200px; overflow-y:auto; background:#f0f0f0; color:#111;
                        padding:8px; border-radius:4px; font-family:monospace;
                        white-space:pre-wrap; border:1px solid #d0d0d0;">{logs}</div>
            """.format(logs="\n".join(st.session_state.log_lines))
            status_log_placeholder.markdown(html, unsafe_allow_html=True)

        def append_log(msg: str):
            st.session_state.log_lines.append(msg)
            st.session_state.log_lines = st.session_state.log_lines[-500:]
            render_status_logs()

        def on_progress(done: int, total: int, file: Path, rows: int):
            pct = done / total if total else 1
            progress.progress(pct, text=f"{done}/{total} 처리 중: {file.name} (rows={rows})")
            parent_name = file.parent.name or file.parent
            msg = f"{done}/{total} | {parent_name}\\{file.name} | rows={rows}"
            append_log(msg)

        try:
            with st.spinner("처리 중..."):
                df_pdf, df_excel, matched = run_matching(
                    pdf_dir=resolved_pdf_dir,
                    input_xlsx=resolved_xlsx,
                    sheet_name=sheet_name,
                    columns=columns,
                    progress_cb=on_progress,
                    stop_flag=lambda: st.session_state.get("stop_requested", False),
                )
            st.session_state.results = (df_pdf, df_excel, matched)
            append_log(f"완료 - 매칭 성공: {matched}건")
        except Exception as e:
            err_msg = f"오류: {e}"
            st.error(err_msg)
            append_log(err_msg)
            return

    if st.session_state.results:
        df_pdf, df_excel, matched = st.session_state.results
        _render_results(df_pdf, df_excel, matched, ts=st.session_state.get("run_ts"))


if __name__ == "__main__":
    main()

