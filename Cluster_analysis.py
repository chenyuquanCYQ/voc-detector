"""
競爭聚類分析（第二階段）
需先執行 analyze_market.py 產出分析結果，再跑此腳本
依賴：sklearn / matplotlib / openpyxl / ollama
Embedding 改用 Ollama nomic-embed-text，不需下載 HuggingFace 模型
"""

import pandas as pd
import numpy as np
import json
import time
from pathlib import Path

INPUT_FILE  = "氣味檢測器市調_分析結果.xlsx"
OUTPUT_FILE = "氣味檢測器_競爭聚類分析.xlsx"

# ─── 嘗試匯入選配套件 ───────────────────────────────────────────
try:
    import ollama
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False
    print("⚠️  ollama 套件未安裝：pip install ollama")

try:
    from sklearn.cluster import KMeans
    from sklearn.preprocessing import normalize
    SK_AVAILABLE = True
except ImportError:
    SK_AVAILABLE = False
    print("⚠️  scikit-learn 未安裝：pip install scikit-learn")

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    MPL_AVAILABLE = True
except ImportError:
    MPL_AVAILABLE = False


# ─── Ollama Embedding ────────────────────────────────────────────
EMBED_MODEL = "nomic-embed-text"   # ollama pull nomic-embed-text

def get_embeddings_ollama(texts: list[str]) -> np.ndarray:
    """用 Ollama nomic-embed-text 產生向量，完全本地端"""
    embeddings = []
    total = len(texts)
    for i, text in enumerate(texts):
        try:
            r = ollama.embeddings(model=EMBED_MODEL, prompt=text[:1000])
            embeddings.append(r["embedding"])
        except Exception as e:
            print(f"   ⚠️  第 {i+1} 筆 embedding 失敗：{e}，以零向量替代")
            # 用上一筆維度或暫時填 0（後面 normalize 會處理）
            dim = len(embeddings[-1]) if embeddings else 768
            embeddings.append([0.0] * dim)
        if i % 50 == 0:
            print(f"   Embedding 進度：{i+1}/{total}")
        time.sleep(0.05)   # 避免 Ollama OOM
    return np.array(embeddings, dtype=np.float32)


# ─── 向量聚類 ────────────────────────────────────────────────────
def run_clustering(df: pd.DataFrame, n_clusters: int = 6) -> pd.DataFrame:
    if not OLLAMA_AVAILABLE:
        print("⚠️  ollama 不可用，跳過向量聚類")
        return df
    if not SK_AVAILABLE:
        print("⚠️  scikit-learn 不可用，跳過向量聚類")
        return df

    desc_col = next((c for c in df.columns if "描述" in c or "特色" in c), None)
    if desc_col is None:
        print("⚠️  找不到描述欄位，跳過聚類")
        return df

    texts = df[desc_col].fillna("").tolist()
    print(f"🔄 使用 Ollama ({EMBED_MODEL}) 產生 {len(texts)} 筆向量...")
    print("   （首次執行請確認已執行過：ollama pull nomic-embed-text）")

    embeddings = get_embeddings_ollama(texts)
    embeddings = normalize(embeddings)

    print(f"🔄 KMeans 聚類（k={n_clusters}）...")
    km = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
    df["cluster_id"] = km.fit_predict(embeddings)

    # 自動標記聚類名稱（以各群最常見 sensor_type 命名）
    if "sensor_type" in df.columns:
        cluster_labels = (
            df.groupby("cluster_id")["sensor_type"]
            .agg(lambda x: x.value_counts().index[0] if len(x) > 0 else "未知")
            .to_dict()
        )
        df["cluster_label"] = df["cluster_id"].map(cluster_labels)

    return df


# ─── 競爭四象限分析 ──────────────────────────────────────────────
def quadrant_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """
    X 軸：生態系整合程度（無=0, 藍牙App=1, IoT雲端=2, AI驅動=3, SaaS=4）
    Y 軸：精度等級（消費電子=1, 工業=2, 實驗室=3）
    """
    ecosystem_score = {
        "無": 0, "藍牙App": 1, "IoT雲端": 2, "AI驅動": 3, "SaaS平台": 4, "多種整合": 4
    }
    precision_score = {"消費電子級": 1, "工業級": 2, "實驗室級": 3}

    df["eco_score"] = df.get("ecosystem", pd.Series()).map(
        lambda x: max((v for k, v in ecosystem_score.items() if k in str(x)), default=0)
    )
    df["precision_score"] = df.get("precision_tier", pd.Series()).map(
        lambda x: max((v for k, v in precision_score.items() if k in str(x)), default=1)
    )

    def assign_quadrant(row):
        x, y = row.get("eco_score", 0), row.get("precision_score", 1)
        if x >= 2 and y >= 2:  return "Q1：高精度+強生態（平台型）"
        if x >= 2 and y < 2:   return "Q2：低精度+強生態（IoT消費型）"
        if x < 2  and y >= 2:  return "Q3：高精度+弱生態（儀器型）"
        return                         "Q4：低精度+弱生態（基礎感測型）"

    df["quadrant"] = df.apply(assign_quadrant, axis=1)
    return df


# ─── 競爭強度熱圖 ────────────────────────────────────────────────
def plot_heatmap(df: pd.DataFrame):
    if not MPL_AVAILABLE:
        print("⚠️  matplotlib 未安裝，跳過熱圖：pip install matplotlib")
        return
    if "application_segments" not in df.columns or "sensor_type" not in df.columns:
        return

    rows = []
    for _, r in df.iterrows():
        segs = str(r.get("application_segments", "")).split(",")
        st   = str(r.get("sensor_type", "不明")).split("/")[0].strip()
        for s in segs:
            rows.append({"segment": s.strip(), "sensor_type": st})

    heat_df = pd.DataFrame(rows)
    pivot = heat_df.groupby(["segment", "sensor_type"]).size().unstack(fill_value=0)
    pivot = pivot.loc[pivot.sum(axis=1).nlargest(10).index]

    fig, ax = plt.subplots(figsize=(12, 6))
    im = ax.imshow(pivot.values, cmap="YlOrRd", aspect="auto")
    ax.set_xticks(range(len(pivot.columns)))
    ax.set_xticklabels(pivot.columns, rotation=45, ha="right", fontsize=9)
    ax.set_yticks(range(len(pivot.index)))
    ax.set_yticklabels(pivot.index, fontsize=9)
    ax.set_title("應用場景 × 感測器類型 競爭熱圖", fontsize=13, pad=12)
    plt.colorbar(im, ax=ax, label="產品數量")
    plt.tight_layout()
    plt.savefig("競爭熱圖.png", dpi=150)
    print("📊 競爭熱圖已儲存至：競爭熱圖.png")
    plt.close()


# ─── 主流程 ──────────────────────────────────────────────────────
def main():
    if not Path(INPUT_FILE).exists():
        raise FileNotFoundError(f"找不到 {INPUT_FILE}，請先執行 analyze_market.py")

    df = pd.read_excel(INPUT_FILE, sheet_name="原始+分析")
    print(f"✅ 讀入 {len(df)} 筆分析結果")

    df = quadrant_analysis(df)
    df = run_clustering(df)
    plot_heatmap(df)

    # 輸出
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="完整資料", index=False)

        # 四象限摘要
        if "quadrant" in df.columns:
            brand_col = next((c for c in df.columns if "品牌" in c or "brand" in c.lower()), None)
            if brand_col:
                q_summary = df.groupby("quadrant").agg(
                    產品數=("quadrant", "count"),
                    代表品牌=(brand_col, lambda x: "、".join(x.dropna().astype(str).unique()[:5]))
                ).reset_index()
                q_summary.to_excel(writer, sheet_name="競爭四象限", index=False)

        # 聚類摘要
        if "cluster_id" in df.columns:
            group_cols = ["cluster_id", "cluster_label"] if "cluster_label" in df.columns else ["cluster_id"]
            c_summary = df.groupby(group_cols).size().reset_index(name="產品數")
            c_summary.to_excel(writer, sheet_name="聚類結果", index=False)

    print(f"💾 聚類分析已儲存至：{OUTPUT_FILE}")


if __name__ == "__main__":
    main()