import io
import datetime as dt
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


def main():
    st.set_page_config(page_title="PT 차트 매칭 도구", layout="wide")
    st.title("PT 차트 매칭 도구")
    st.write("왼쪽 사이드바에서 경로 입력 또는 업로드를 설정한 뒤 실행하세요.")

    default_sheet = "Jun 30 _ Jul 5"
    default_cols = "A:G"

    with st.sidebar:
        st.header("입력 설정")
        st.caption("경로를 절대경로로 입력해 주세요.")

        pdf_dir = st.text_input("PDF 폴더 절대경로", value=st.session_state.get("pdf_dir", ""), key="pdf_dir")
        input_xlsx = st.text_input("입력 엑셀 절대경로", value=st.session_state.get("input_xlsx", ""), key="input_xlsx")

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

    if stop_clicked:
        st.session_state.stop_requested = True
        st.warning("중지 요청됨: 다음 작업 지점에서 중단합니다.")

    if run_clicked:
        if not pdf_dir:
            st.error("PDF 폴더 절대경로를 입력해 주세요.")
            return
        if not input_xlsx:
            st.error("입력 엑셀 절대경로를 입력해 주세요.")
            return

        st.session_state.log_lines = []
        st.session_state.results = None
        st.session_state.run_ts = dt.datetime.now().strftime("%Y%m%d%H%M%S")
        st.session_state.stop_requested = False

        # status_badge = st.empty()  # 상태 배지 사용을 원하면 주석을 해제하세요.

        progress = st.progress(0.0, text="대기 중")
        status_box = st.expander("로그 메시지", expanded=False)
        status_log_placeholder = status_box.empty()

        def render_status_logs():
            # 200px 높이, 내부 스크롤 적용
            html = """
            <div style="height:200px; overflow-y:auto; background:#f0f0f0; color:#111;
                        padding:8px; border-radius:4px; font-family:monospace;
                        white-space:pre-wrap; border:1px solid #d0d0d0;">{logs}</div>
            """.format(logs="\n".join(st.session_state.log_lines))
            status_log_placeholder.markdown(html, unsafe_allow_html=True)

        def append_log(msg: str):
            st.session_state.log_lines.append(msg)
            # 최근 500줄만 유지
            st.session_state.log_lines = st.session_state.log_lines[-500:]
            render_status_logs()

        def on_progress(done: int, total: int, file: Path, rows: int):
            pct = done / total if total else 1
            progress.progress(pct, text=f"{done}/{total} 처리 중: {file.name} (rows={rows})")
            parent_name = file.parent.name or file.parent
            msg = f"{done}/{total} | {parent_name}\{file.name} | rows={rows}"
            append_log(msg)

        try:
            with st.spinner("처리 중..."):
                df_pdf, df_excel, matched = run_matching(
                    pdf_dir=pdf_dir or None,
                    input_xlsx=input_xlsx or None,
                    sheet_name=sheet_name,
                    columns=columns,
                    progress_cb=on_progress,
                    stop_flag=lambda: st.session_state.get("stop_requested", False),
                )
            st.session_state.results = (df_pdf, df_excel, matched)
            # 실행 시점 타임스탬프를 결과와 함께 유지
            # if status_badge:
            #     status_badge.success("완료")
            append_log(f"완료 - 매칭 성공: {matched}건")
        except Exception as e:
            err_msg = f"오류: {e}"
            # if status_badge:
            #     status_badge.error("실패")
            st.error(err_msg)
            append_log(err_msg)
            return

    if st.session_state.results:
        df_pdf, df_excel, matched = st.session_state.results
        _render_results(df_pdf, df_excel, matched, ts=st.session_state.get("run_ts"))


if __name__ == "__main__":
    main()

