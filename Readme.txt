# 氣味檢測器市調分析系統

## 使用流程

```
氣味檢測器市調.xlsx
        │
        ▼
  analyze_market.py         ← 第一步：LLM 結構化提取
  (ollama + gemma4:e4b)
        │
        ▼
氣味檢測器市調_分析結果.xlsx
        │
        ▼
  cluster_analysis.py       ← 第二步：競爭聚類 + 四象限（選配）
        │
        ▼
氣味檢測器_競爭聚類分析.xlsx + 競爭熱圖.png
```

---

## 安裝依賴

```bash
# 必要
pip install pandas openpyxl ollama

# 選配（向量聚類 & 熱圖）
pip install sentence-transformers scikit-learn matplotlib
```

---

## 使用方式

### 第一步：LLM 結構化分析

1. 將 `氣味檢測器市調.xlsx` 放在同目錄
2. 確認 Ollama 服務正在執行，且已拉取模型：
   ```bash
   ollama serve
   ollama pull gemma3:12b
   ```
3. 執行分析：
   ```bash
   python analyze_market.py
   ```

### 第二步：競爭聚類分析（選配）

```bash
python cluster_analysis.py
```

---

## 輸出欄位說明

| 欄位 | 說明 | 可能值 |
|------|------|--------|
| `sensor_type` | 感測器技術 | MOS / MEMS / GC-MS / NDIR / 電化學 ... |
| `form_factor` | 產品型態 | 手持可攜式 / 固定式 / 嵌入式模組 ... |
| `precision_tier` | 精度等級 | 實驗室級 / 工業級 / 消費電子級 |
| `trl` | 技術成熟度 | 原型 / 研發中 / 小量商用 / 成熟商用 |
| `target_gases` | 目標氣體 | VOCs, H2S, NH3 ... （逗號分隔） |
| `output_type` | 數據輸出型態 | 分級顯示 / 精確數值 / 氣味指紋圖譜 |
| `ecosystem` | 生態系整合 | 無 / 藍牙App / IoT雲端 / AI驅動 / SaaS平台 |
| `application_segments` | 應用場景 | 食品品質, 環境監測 ... （逗號分隔） |
| `competitive_moat` | 競爭護城河 | 純硬體 / 軟硬整合 / 資料平台 ... |
| `key_features` | 獨特功能 | 最多3個關鍵詞 |
| `confidence` | 分析信心度 | 高 / 中 / 低 |
| `analyzed_at` | 分析時間戳 | ISO 8601 格式 |

---

## 常用設定調整

在 `analyze_market.py` 頂部修改：

```python
MODEL_NAME   = "gemma3:12b"   # 改為 gemma3:4b 加快速度
BATCH_DELAY  = 0.5            # 每筆間隔秒數（記憶體不足時加大）
FORCE_RERUN  = False          # True = 強制重跑所有列
```

---

## 持續更新工作流程

每次在 Google Sheets 新增資料後：
1. 匯出為 `氣味檢測器市調.xlsx`（覆蓋舊檔）
2. 執行 `python analyze_market.py`
   - 程式會自動跳過已有 `analyzed_at` 的列
   - 只分析新增資料，節省時間