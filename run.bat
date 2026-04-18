@echo off
setlocal

cd /d "%~dp0"

if "%~1"=="" (
    set DATA_FILE=sample_data_template.csv
) else (
    set DATA_FILE=%~1
)

if "%~2"=="" (
    set CONFIG_FILE=questionnaire_config_template.csv
) else (
    set CONFIG_FILE=%~2
)

if "%~3"=="" (
    set OUTPUT_DIR=output
) else (
    set OUTPUT_DIR=%~3
)

if "%~4"=="" (
    set MODEL_FILE=analysis_model_template.csv
) else (
    set MODEL_FILE=%~4
)

if "%~5"=="" (
    set LABEL_FILE=sample_variable_labels.csv
) else (
    set LABEL_FILE=%~5
)

python survey_auto_summary.py --data "%DATA_FILE%" --config "%CONFIG_FILE%" --models "%MODEL_FILE%" --labels "%LABEL_FILE%" --outdir "%OUTPUT_DIR%" --straightline-check

if errorlevel 1 (
    echo.
    echo 執行失敗，請檢查錯誤訊息。
    pause
    exit /b 1
)

echo.
echo 執行完成，請查看 "%OUTPUT_DIR%" 資料夾。
pause
