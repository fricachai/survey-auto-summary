# Survey Auto Summary

本工具用於學術研究問卷資料之前處理、信效度檢查、差異分析、模型分析與摘要輸出。適用情境包含：

- 碩士論文問卷分析
- 教學實踐研究
- 課程回饋問卷
- 量表前測與正式施測資料整理

## 主要功能

- 題項數值清理與反向計分
- 缺失值比例篩選
- 直線作答檢查
- 題項描述統計
- 天花板 / 地板效果
- 構面平均分數表
- Cronbach's alpha
- 校正後題項總分相關（CITC）
- 刪題後 alpha
- KMO / Bartlett / EFA
- 人口統計變項次數、百分比、有效百分比、累積有效百分比
- 各構面 Pearson 相關矩陣與明細
- independent samples t-test
- one-way ANOVA
- 多元迴歸
- 階層迴歸
- 中介效果分析
- 調節效果分析
- APA 風格表格與表註
- PLS-SEM 前置整理
- 中文變項標籤對照
- 自動圖表輸出
- 論文段落模板輸出
- 本地上傳介面

## 專案檔案

- `survey_auto_summary.py`
  - 主程式
- `streamlit_app.py`
  - 本地上傳介面
- `run.bat`
  - 命令列快速執行
- `run_ui.bat`
  - 啟動上傳介面
- `requirements.txt`
  - 套件需求

### 範本檔

建議優先使用 `.xlsx` 範本，避免 Windows 直接開啟 CSV 時發生中文亂碼。

- `questionnaire_config_template.xlsx`
- `sample_data_template.xlsx`
- `analysis_model_template.xlsx`
- `sample_variable_labels.xlsx`
- `sample_config_demographic.xlsx`

若你要用 CSV，也保留了對應的 `.csv` 版本。

## 問卷設定檔格式

至少需要：

- `variable`
- `type`

量表題建議再提供：

- `construct`
- `subconstruct`
- `reverse`
- `label`

### `type` 說明

- `likert`
  - 量表題項
- `demographic`
  - 人口統計或分組變項
- `continuous`
  - 連續數值變項

### `reverse` 說明

- `0`：非反向題
- `1`：反向題

## 模型設定檔格式

欄位如下：

- `analysis_type`
- `model_name`
- `outcome`
- `predictor`
- `mediator`
- `moderator`
- `covariates`
- `step_predictors`
- `notes`

### `analysis_type` 可用值

- `regression`
- `hierarchical_regression`
- `mediation`
- `moderation`

### 欄位規則

- `predictor`
  - 多個變項以 `|` 分隔
- `covariates`
  - 多個控制變項以 `|` 分隔
- `step_predictors`
  - 階層迴歸使用
  - 步驟內變項以 `|` 分隔
  - 步驟與步驟之間以 `||` 分隔

## 中文標籤檔

欄位格式：

- `variable`
- `label`

若執行時提供 `--labels`，輸出表格會優先使用中文標籤。

## 安裝方式

建議 Python 3.10 以上版本。

```bash
pip install -r requirements.txt
```

## 執行方式

### 1. 命令列執行

```bash
python survey_auto_summary.py --data sample_data_template.xlsx --config questionnaire_config_template.xlsx --models analysis_model_template.xlsx --labels sample_variable_labels.xlsx --outdir output --straightline-check
```

### 2. Windows 快速執行

```bat
run.bat
```

### 3. 本地上傳介面

```bat
run_ui.bat
```

或：

```bash
streamlit run streamlit_app.py
```

上傳介面提供：

- 問卷資料檔上傳按鈕
- 題項設定檔上傳按鈕
- 模型設定檔上傳按鈕
- 中文標籤檔上傳按鈕
- 每一種檔案旁邊都有樣本檔下載按鈕
- 分析完成後直接下載結果

## 輸出內容

主要輸出包括：

- `survey_auto_summary_output.xlsx`
- `survey_auto_summary_report.docx`
- `survey_auto_summary_brief.txt`
- `論文段落模板.txt`
- `survey_auto_summary_notes.xlsx`
- `output/charts/`

另外，上傳介面模式會在：

- `ui_runs/latest_input`
- `ui_runs/latest_output`

保留最近一次執行的輸入與輸出。

## 若分析無法執行

請優先檢查：

1. Excel 中 `執行環境備註`
2. `survey_auto_summary_notes.xlsx`
3. 對應分析工作表是否為空表
4. 是否有以下常見原因：
   - 樣本數不足
   - 欄位不存在
   - 類別變項未先轉成模型可用的數值格式
   - 套件未安裝
   - 模型設定檔欄位填寫錯誤

## 重要說明：避免亂碼

若你在 Windows 直接雙擊 CSV，Excel 可能用系統預設編碼開啟，導致中文亂碼。建議：

1. 優先使用本專案提供的 `.xlsx` 範本
2. 若一定要用 CSV，請在 Excel 中用：
   - `資料` -> `自文字/CSV`
   - 編碼選 `UTF-8`

## 建議使用流程

1. 先下載 `.xlsx` 範本
2. 依範本填入自己的資料、設定、模型與中文標籤
3. 用 `run_ui.bat` 開啟上傳介面
4. 上傳檔案後執行分析
5. 下載 Excel、Word、圖表與段落模板結果

## 後續仍可擴充

- 事後比較（Scheffé、Tukey、Games-Howell）
- 類別變項自動 dummy coding
- ANCOVA
- 多重中介 / 多重調節
- PROCESS 類型模型
- SmartPLS / WarpPLS 匯入格式自動轉檔
