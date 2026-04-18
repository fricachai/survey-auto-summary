#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地上傳介面：
- 上傳資料檔、設定檔、模型檔、標籤檔
- 提供樣本檔下載
- 直接呼叫 survey_auto_summary.py 執行分析
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import streamlit as st


ROOT = Path(__file__).resolve().parent
WORK_DIR = ROOT / "ui_runs"
INPUT_DIR = WORK_DIR / "latest_input"
OUTPUT_DIR = WORK_DIR / "latest_output"

SAMPLE_FILES = {
    "資料檔樣本": ROOT / "sample_data_template.csv",
    "設定檔樣本": ROOT / "questionnaire_config_template.csv",
    "模型檔樣本": ROOT / "analysis_model_template.csv",
    "標籤檔樣本": ROOT / "sample_variable_labels.csv",
}


def reset_dirs() -> None:
    if INPUT_DIR.exists():
        shutil.rmtree(INPUT_DIR)
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR)
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def save_uploaded_file(uploaded_file, target_path: Path) -> Path:
    target_path.write_bytes(uploaded_file.getbuffer())
    return target_path


def render_upload_row(
    title: str,
    description: str,
    upload_key: str,
    sample_label: str,
    required: bool = False,
):
    col1, col2 = st.columns([3, 2])
    with col1:
        st.markdown(f"**{title}**")
        st.caption(description + ("（必填）" if required else "（選填）"))
        uploaded = st.file_uploader(
            label=title,
            type=["csv", "xlsx", "xls"],
            key=upload_key,
            label_visibility="collapsed",
        )
    with col2:
        sample_path = SAMPLE_FILES[sample_label]
        st.markdown("**樣本檔下載**")
        st.download_button(
            label=f"下載 {sample_path.name}",
            data=sample_path.read_bytes(),
            file_name=sample_path.name,
            mime="application/octet-stream",
            key=f"download_{upload_key}",
            use_container_width=True,
        )
    return uploaded


def run_analysis(data_path: Path, config_path: Path, model_path: Path | None, label_path: Path | None) -> tuple[bool, str]:
    command = [
        sys.executable,
        str(ROOT / "survey_auto_summary.py"),
        "--data",
        str(data_path),
        "--config",
        str(config_path),
        "--outdir",
        str(OUTPUT_DIR),
        "--straightline-check",
    ]
    if model_path:
        command.extend(["--models", str(model_path)])
    if label_path:
        command.extend(["--labels", str(label_path)])

    result = subprocess.run(
        command,
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    output_text = "\n".join(part for part in [result.stdout.strip(), result.stderr.strip()] if part)
    return result.returncode == 0, output_text


def render_output_downloads() -> None:
    if not OUTPUT_DIR.exists():
        return

    files = sorted([path for path in OUTPUT_DIR.iterdir() if path.is_file()])
    chart_dir = OUTPUT_DIR / "charts"
    if not files and not chart_dir.exists():
        return

    st.subheader("分析結果下載")
    for path in files:
        st.download_button(
            label=f"下載 {path.name}",
            data=path.read_bytes(),
            file_name=path.name,
            mime="application/octet-stream",
            key=f"result_{path.name}",
            use_container_width=True,
        )

    if chart_dir.exists():
        chart_files = sorted(chart_dir.glob("*"))
        if chart_files:
            st.markdown("**圖表檔案**")
            for path in chart_files:
                st.download_button(
                    label=f"下載 {path.name}",
                    data=path.read_bytes(),
                    file_name=path.name,
                    mime="application/octet-stream",
                    key=f"chart_{path.name}",
                    use_container_width=True,
                )


def main() -> None:
    st.set_page_config(page_title="Survey Auto Summary 上傳工具", layout="wide")
    st.title("Survey Auto Summary 上傳工具")
    st.write("上傳你的問卷資料與設定檔後，系統會在本機執行分析，並提供結果下載。")

    st.markdown("### 上傳檔案")
    data_file = render_upload_row(
        "1. 問卷資料檔",
        "支援 CSV / Excel。欄位名稱需對應設定檔中的 variable。",
        "data_file",
        "資料檔樣本",
        required=True,
    )
    config_file = render_upload_row(
        "2. 題項設定檔",
        "需包含 variable、type；量表題建議含 construct、reverse、label。",
        "config_file",
        "設定檔樣本",
        required=True,
    )
    model_file = render_upload_row(
        "3. 模型設定檔",
        "若要跑多元迴歸、階層迴歸、中介、調節，請上傳此檔。",
        "model_file",
        "模型檔樣本",
        required=False,
    )
    label_file = render_upload_row(
        "4. 中文標籤檔",
        "若要讓輸出表格顯示較友善的中文名稱，請上傳此檔。",
        "label_file",
        "標籤檔樣本",
        required=False,
    )

    st.markdown("### 執行")
    st.caption("按下後會在本機產生 output 檔案，並在下方提供下載按鈕。")
    run_clicked = st.button("開始分析", type="primary", use_container_width=True)

    if run_clicked:
        if data_file is None or config_file is None:
            st.error("請至少上傳資料檔與設定檔。")
        else:
            reset_dirs()
            data_path = save_uploaded_file(data_file, INPUT_DIR / data_file.name)
            config_path = save_uploaded_file(config_file, INPUT_DIR / config_file.name)
            model_path = save_uploaded_file(model_file, INPUT_DIR / model_file.name) if model_file else None
            label_path = save_uploaded_file(label_file, INPUT_DIR / label_file.name) if label_file else None

            with st.spinner("分析執行中，請稍候..."):
                success, logs = run_analysis(data_path, config_path, model_path, label_path)

            st.markdown("### 執行紀錄")
            st.code(logs or "無額外訊息。")

            if success:
                st.success("分析完成，可以下載結果。")
            else:
                st.error("分析失敗，請檢查上方紀錄。")

    render_output_downloads()

    st.markdown("### 使用提醒")
    st.markdown(
        "- 資料檔欄名必須和設定檔 `variable` 一致。\n"
        "- 模型檔與標籤檔可不提供；若未提供，系統仍可執行基本分析。\n"
        "- 若某分析無法執行，請優先檢查執行紀錄與輸出中的備註工作表。"
    )


if __name__ == "__main__":
    main()
