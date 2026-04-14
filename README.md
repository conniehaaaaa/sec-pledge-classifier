# SEC DEF 14A 質押／避險政策分類器

從 SEC DEF 14A 代理聲明書中，自動辨識公司是否具有**內部人質押（Pledging）**或**避險（Hedging）**政策，並分類其嚴格程度。

---

## 專案流程

```
Excel 資料（含網址）
        ↓
   隨機取樣（最多 12 筆）
        ↓
   爬取 SEC 網頁 HTML
        ↓
   Regex 粗篩相關段落
        ↓
   GPT-4o 分類（JSON 輸出）
        ↓
   輸出 Excel 結果檔
```

---

## 輸出欄位說明

| 欄位 | 說明 |
|------|------|
| `AI_Pledge2` | 反質押政策：0=無、1=嚴格禁止、2=禁止但可申請例外、-1=無法判斷 |
| `AI_Pledge3` | 反避險政策：0=無、1=嚴格禁止、2=禁止但可申請例外、-1=無法判斷 |
| `Pledge2_breadth` | 政策涵蓋範圍：1=全體員工、2=董事+高管、3=特定人員 |
| `Pledge2_strength` | 禁止力道：1=強（禁止）、2=中（須申請）、3=弱（建議不要） |
| `Pledge3_breadth` | 同上，針對避險政策 |
| `Pledge3_strength` | 同上，針對避險政策 |
| `AI_依據文字` | LLM 判斷所依據的關鍵句子 |
| `Regex命中區塊數` | Regex 命中的 HTML 區塊數量 |
| `Regex段落` | 實際傳給 LLM 的相關段落原文 |

---

## 環境設定

### 1. Clone 專案

```bash
git clone https://github.com/你的帳號/sec-pledge-classifier.git
cd sec-pledge-classifier
```

### 2. 安裝套件

```bash
pip install -r requirements.txt
```

### 3. 設定 API Key

```bash
cp .env.example .env
```

用文字編輯器打開 `.env`，填入你的 OpenAI API Key：

```
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxx
```

### 4. 放入資料檔

將 `DEF14A2023merged_Traning.xlsx` 放在專案根目錄下（與 `classify.py` 同層）。

---

## 執行

```bash
python classify.py
```

執行完成後，結果會輸出至 `LLM分類結果.xlsx`。

---

## 取樣邏輯

| 取樣條件 | 筆數 |
|----------|------|
| is_Pledge2 = 1（有質押禁令）| 3 筆 |
| is_Pledge2 = 0（無質押禁令）| 3 筆 |
| is_Pledge3 = 1（有避險禁令）| 3 筆 |
| is_Pledge3 = 0（無避險禁令）| 3 筆 |
| **合計（去重後）** | **最多 12 筆** |

隨機種子固定為 `RANDOM_SEED = 42`，每次執行結果相同。
若想每次抽到不同樣本，可將 `classify.py` 中的 `RANDOM_SEED = 42` 改為 `RANDOM_SEED = None`。

---

## 注意事項

- `.env` 含有 API Key，**絕對不要上傳 GitHub**（已加入 `.gitignore`）
- 資料檔 `.xlsx` 也已加入 `.gitignore`，需自行放入
- 每筆資料呼叫一次 GPT-4o，費用約 $0.01～0.05 USD/筆
- 程式每筆間隔 1 秒（`time.sleep(1)`）避免超過 API rate limit

---

## 專案結構

```
sec-pledge-classifier/
├── classify.py          # 主程式
├── requirements.txt     # 套件清單
├── .env.example         # API Key 範本（複製為 .env 後填入）
├── .gitignore           # 排除敏感檔案
└── README.md            # 說明文件（本檔）
```
