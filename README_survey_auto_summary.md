# Survey Auto Summary

本工具用於學術研究問卷資料之前處理、信效度檢查、差異分析、模型分析與摘要輸出。主要適用於：

- 碩士論文問卷分析
- 教學實踐研究
- 課程回饋問卷
- 量表前測與正式施測資料整理

## 一、目前可用功能

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
- Word / Excel / TXT 摘要輸出

## 二、專案檔案

- `survey_auto_summary.py`
  - 主程式
- `questionnaire_config_template.csv`
  - 基本設定範本
- `sample_data_template.csv`
  - 範例資料
- `analysis_model_template.csv`
  - 多元迴歸、階層迴歸、中介、調節之模型設定範本
- `sample_variable_labels.csv`
  - 中文變項標籤對照範本
- `sample_config_demographic.csv`
  - 含更多人口統計欄位的設定範例
- `requirements.txt`
  - 套件需求
- `run.bat`
  - Windows 快速執行檔
- `sample_output_說明.txt`
  - 輸出檔案說明

## 三、設定檔格式

### 1. 問卷設定檔

必要欄位：

- `variable`
- `type`

建議欄位：

- `construct`
- `subconstruct`
- `reverse`
- `label`

### 2. `type` 說明

- `likert`
  - 量表題項，會進行反向計分、信效度與構面平均分數整理
- `demographic`
  - 人口統計或類別分組變項，會進行人口統計表、t 檢定與 ANOVA
- `continuous`
  - 連續數值變項，會納入迴歸、中介、調節等模型分析資料集

### 3. `reverse` 說明

- `0`：非反向題
- `1`：反向題

### 4. 中文變項標籤檔

欄位格式：

- `variable`
- `label`

若提供 `--labels`，輸出表格中會優先使用對照後的中文標籤。

## 四、模型設定檔格式

模型設定檔欄位：

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
  - 每一個步驟之內變項以 `|` 分隔
  - 步驟與步驟之間以 `||` 分隔
  - 範例：`age||人格特質|使用動機||知覺價值`

## 五、設計假設與限制

1. 構面分數以同一 `construct` 之題項平均數計算。
2. Pearson 相關分析以構面平均分數為主。
3. t 檢定採 Welch t-test。
4. ANOVA 為單因子變異數分析，尚未內建事後比較。
5. 多元迴歸、階層迴歸、中介、調節目前以數值型變項或構面平均分數為主。
6. 若模型中使用類別變項，建議先自行轉為數值編碼後再納入模型。
7. 中介分析目前輸出 `a`、`b`、`c`、`c'`、Sobel 檢定與 bootstrap 信賴區間。
8. 調節分析目前使用平均中心化後的交互作用項。
9. KMO / Bartlett / EFA 需安裝 `factor_analyzer`。
10. Word 輸出需安裝 `python-docx`。
11. 圖表輸出需安裝 `matplotlib`。
12. 若某分析因樣本數不足、欄位缺漏或套件未安裝而無法執行，系統會保留其他輸出，並在結果中標註原因。

## 六、安裝方式

建議 Python 版本：`3.10` 以上。

```bash
pip install -r requirements.txt
```

## 七、執行方式

### 基本執行

```bash
python survey_auto_summary.py --data sample_data_template.csv --config questionnaire_config_template.csv --outdir output --straightline-check
```

### 含模型設定與中文標籤

```bash
python survey_auto_summary.py --data sample_data_template.csv --config questionnaire_config_template.csv --models analysis_model_template.csv --labels sample_variable_labels.csv --outdir output --straightline-check
```

### Windows 快速執行

```bat
run.bat
```

若要指定檔案：

```bat
run.bat your_data.csv your_config.csv output your_models.csv your_labels.csv
```

## 八、參數說明

- `--data`
  - 問卷資料檔路徑
- `--config`
  - 問卷設定檔路徑
- `--models`
  - 模型設定檔路徑
- `--labels`
  - 中文變項標籤檔路徑
- `--outdir`
  - 輸出資料夾
- `--scale-min`
  - Likert 最小值，預設 `1`
- `--scale-max`
  - Likert 最大值，預設 `5`
- `--missing-threshold`
  - 個案缺失比例刪除門檻，預設 `0.2`
- `--straightline-check`
  - 是否啟用直線作答檢查

## 九、輸出內容

### 1. `survey_auto_summary_output.xlsx`

主要工作表包括：

- `專案結構檢查`
- `執行環境備註`
- `樣本篩選摘要`
- `人口統計摘要`
- `題項描述統計`
- `構面分數資料`
- `構面描述統計`
- `信度摘要`
- `KMO_Bartlett`
- `相關分析矩陣`
- `相關分析明細`
- `t檢定摘要`
- `ANOVA摘要`
- `多元迴歸摘要`
- `多元迴歸係數`
- `階層迴歸摘要`
- `階層迴歸係數`
- `中介分析摘要`
- `中介分析明細`
- `調節分析摘要`
- `調節分析係數`
- `警示訊息`
- `APA表1_人口統計`
- `APA表2_構面描述`
- `APA表3_t檢定`
- `APA表4_ANOVA`
- `APA表5_相關分析`
- `APA表6_多元迴歸`
- `APA表7_階層迴歸`
- `APA表8_中介分析`
- `APA表9_調節分析`
- `APA表註`
- `圖表輸出清單`
- `PLS_題項資料`
- `PLS_Z標準化題項`
- `PLS_測量模型對照`
- `PLS_缺失與變異檢核`
- `PLS_構面Z分數`
- `PLS_VIF檢核`
- `PLS_SEM前置說明`

另會依構面輸出：

- `CITC_構面名`
- `EFA負荷_構面名`
- `EFA特徵值_構面名`

### 2. `survey_auto_summary_report.docx`

內容包括：

- 正式研究摘要
- APA 風格表格
- 表註
- 圖表輸出清單

### 3. `survey_auto_summary_brief.txt`

- 短版摘要

### 4. `論文段落模板.txt`

- 可直接整理進第四章的段落骨架

### 5. `output/charts`

預設可能包含：

- 人口統計長條圖
- 構面平均數長條圖
- 相關熱圖

### 6. `survey_auto_summary_notes.xlsx`

- 套件缺漏、Word 未輸出或圖表未輸出等備註

## 十、若分析無法執行

請優先查看：

1. Excel 中 `執行環境備註`
2. `survey_auto_summary_notes.xlsx`
3. 對應分析工作表是否為空表
4. 是否有以下常見原因：
   - 樣本數不足
   - 欄位不存在
   - 類別變項未先數值化卻被納入模型
   - 套件未安裝
   - 某模型設定欄位填寫錯誤

## 十一、建議研究使用流程

1. 依 `questionnaire_config_template.csv` 建立正式設定檔。
2. 若需要中文標籤，建立 `sample_variable_labels.csv` 類似格式之對照檔。
3. 若需要模型分析，依 `analysis_model_template.csv` 建立模型設定檔。
4. 先以小樣本測試輸出是否正常。
5. 再套用正式資料。
6. 檢查 `警示訊息`、`信度摘要`、`APA表註` 與 `論文段落模板.txt`。

## 十二、後續可再擴充

- 事後比較（Scheffé、Tukey、Games-Howell）
- 類別變項自動 dummy coding
- 共變數分析（ANCOVA）
- 多元中介 / 多元調節
- PROCESS 類型模型
- APA 指定版面與斜體格式細節
- SmartPLS / WarpPLS 匯入檔自動轉檔
