"""
SEC DEF 14A 質押／避險政策分類器
===================================
流程：從 Excel 讀取網址 → 爬取 SEC 網頁 → Regex 粗篩段落 → GPT-4o 分類 → 輸出結果

輸出欄位：
  CIK | FILEDATE | AI_Pledge2 | AI_Pledge3 |
  Pledge2_breadth | Pledge2_strength | Pledge3_breadth | Pledge3_strength |
  AI_依據文字 | Regex命中區塊數 | Regex段落
"""

import re
import json
import time
import random
import html
import openpyxl
import pandas as pd
import requests
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()  # 讀取 .env 中的 OPENAI_API_KEY

# ══════════════════════════════════════════════════════════════════════════════
# 設定區（依需求修改這裡）
# ══════════════════════════════════════════════════════════════════════════════

FILE_PATH   = 'DEF14A2023merged_Traning.xlsx'  # 放在同目錄下

# 取樣設定：
#   SAMPLE_N = None  → 跑全部資料
#   SAMPLE_N = 10    → 隨機抽 10 筆測試
SAMPLE_N    = 10

RANDOM_SEED = 42     # 固定取樣結果；改成 None 則每次不同

# ══════════════════════════════════════════════════════════════════════════════

random.seed(RANDOM_SEED)
client = OpenAI()  # 自動讀取環境變數 OPENAI_API_KEY

# ══════════════════════════════════════════════════════════════════════════════
# V5 粗篩 Regex
# ══════════════════════════════════════════════════════════════════════════════

# V5：廣泛關鍵字掃描（保留你原有的，修正 insider trading policy）
V5_PATTERN = re.compile(
    r'pledg\w+|short\s+sale\w*|derivative'
    r'|margin\s+account'
    r'|anti.?hedg|anti.?pledg'
    r'|zero.?cost\s+collar|prepaid\s+variable\s+forward'
    r'|monetization\s+transaction'
    r'|\bcollateral\b|\bmargin\s+loans?\b',
    re.IGNORECASE | re.DOTALL
)
# 注意：hedg\w+ 和 insider trading policy 移出來，
#    改用下面有上下文的 pattern 處理，避免誤判

# 需要上下文的 pattern
CONTEXT_PATTERNS = {
    'hedg_with_context': re.compile(
        r'\bprohibit\w*\s*[\S\s]*?\bhedg'   # prohibit + hedg
        r'|\banti.?hedg'                     # anti-hedging
        r'|\bhedg\w+.*\bpolic',              # hedging + policy
        re.IGNORECASE | re.DOTALL
    ),
    'insider_trading_policy': re.compile(
        r'\binsider\s*trading\s*policy\b[\S\s]*?\b(pledg|hedg|prohibit)\b',
        re.IGNORECASE | re.DOTALL
    ),
    'policy_against': re.compile(
        r'\b(policy|policies)\b[\S\s]*?\b(against|ban|restrict)\w*\b[\S\s]*?\b(pledg|hedg)\b',
        re.IGNORECASE | re.DOTALL
    ),
    'ownership_guideline': re.compile(
        r'\b(stock|equity)\s+ownership\s+guidelines?\b[\S\s]*?\b(pledg|hedg)\b',
        re.IGNORECASE | re.DOTALL
    ),
}

def classify_text(text: str) -> str | None:
    # 先跑廣泛關鍵字
    if V5_PATTERN.search(text):
        return 'keyword_hit'
    # 再跑需要上下文的條件
    for hit_type, pattern in CONTEXT_PATTERNS.items():
        if pattern.search(text):
            return hit_type
    return None  # 不輸出


# ══════════════════════════════════════════════════════════════════════════════
# 工具函式
# ══════════════════════════════════════════════════════════════════════════════

def get_cell(row, idx):
    v = row[idx].value
    return str(v).strip() if v is not None else ""

def fetch_html(url: str, timeout: int = 15) -> str:
    """抓取 SEC HTM 文件，回傳原始 HTML"""
    try:
        resp = requests.get(
            url, timeout=timeout,
            headers={'User-Agent': 'yizhen1426@gmail.com'}
        )
        resp.raise_for_status()
        return resp.text
    except Exception as e:
        print(f"  [fetch 失敗] {url}: {e}")
        return ""

def extract_relevant_passages(raw_html: str) -> str:
    """
    以 HTML 區塊標籤為單位切割文件，
    對每個區塊的純文字跑 V5 Regex，命中的區塊完整保留。
    """
    BLOCK_TAG_RE = re.compile(
        r'<(p|div|li|td|th|span|section|article)[^>]*>(.*?)</\1>',
        re.IGNORECASE | re.DOTALL
    )

    def strip_tags(text: str) -> str:
        text = re.sub(r'<[^>]+>', ' ', text)
        text = html.unescape(text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    seen_texts = set() 
    passages   = []

    for m in BLOCK_TAG_RE.finditer(raw_html):
        block_text = strip_tags(m.group(2)) 

        if len(block_text) < 20: 
            continue
        if not V5_PATTERN.search(block_text):
            continue

        key = block_text[:120] #去重
        if key in seen_texts:
            continue
        seen_texts.add(key)

        passages.append(block_text) 

    return "\n\n---\n\n".join(passages) if passages else ""

def classify_with_llm(passages: str, cik: str, filedate: str) -> dict: 
    """
    呼叫 GPT-4o，同時判斷 is_Pledge2 / is_Pledge3 以及
    各自的 breadth（涵蓋範圍）和 strength（禁止力道），回傳結構化 JSON。
    """
    NULL_RESULT = {
        "is_Pledge2": -1, "is_Pledge3": -1,
        "Pledge2_breadth": None, "Pledge2_strength": None,
        "Pledge3_breadth": None, "Pledge3_strength": None,
        "evidence": "[No relevant passages found by Regex — unable to classify]"
    }

    if not passages.strip():
        return NULL_RESULT

    system_prompt = """You are an SEC DEF 14A analyst. Classify pledging and hedging policies from the provided passages.

=== LABELS ===

is_Pledge2 (Anti-Pledging): Does a formal policy prohibit insiders from pledging company shares as collateral?
is_Pledge3 (Anti-Hedging): Does a formal policy prohibit insiders from hedging company securities (e.g., short sales, collars, swaps, forwards)?
  0 = No policy
  1 = Strict prohibition, no exceptions
  2 = Prohibited but exceptions/pre-approval permitted

=== CLASSIFICATION RULES ===

→ Label 1: passage explicitly says insiders are prohibited/may not/shall not pledge or hedge.
  - Banning margin accounts = anti-pledging policy (label 1), even without the word "pledge."

→ Label 2: prohibited BUT pre-approval or exceptions exist ("unless pre-approved," "subject to approval").
  - Any mention of a pre-approval process = label 2, not 1.

→ Label 0 (these are NOT policies — common false positives to avoid):
  - Ownership disclosure: "X shares pledged as collateral" — this is a disclosure, not a prohibition.
  - Equity award plan restriction: "Participant may not sell, transfer, or pledge unvested awards" — plan-level transferability rule, not a company policy.
  - SPAC/shareholder lock-up: "Sponsors agreed not to sell or pledge shares during lock-up" — not an insider trading policy.

=== POLICY QUALITY (only when is_Pledge2 or is_Pledge3 >= 1; else null) ===

breadth — who is covered?
  1 = All employees / firm-wide
  2 = Directors and executive officers / senior management only
  3 = Narrow group only (NEOs, specific individuals)

strength — how binding is the language?
  1 = Strong: "prohibited," "may not," "shall not," "are not permitted"
  2 = Moderate: "generally prohibits," "not permitted without approval," "subject to pre-clearance"
  3 = Weak: "discouraged," "should not," "not recommended"

=== OUTPUT ===

Respond ONLY with valid JSON, no markdown, no extra text:
{
  "is_Pledge2": <0|1|2>,
  "is_Pledge3": <0|1|2>,
  "Pledge2_breadth": <1|2|3|null>,
  "Pledge2_strength": <1|2|3|null>,
  "Pledge3_breadth": <1|2|3|null>,
  "Pledge3_strength": <1|2|3|null>,
  "evidence": "<up to 3 key sentences separated by | >"
}"""

    user_prompt = f"""The following passages were extracted from the DEF 14A proxy statement filed by CIK={cik} on {filedate}.
They were selected by a keyword filter (Regex) for relevance to pledging and hedging policies.

--- EXTRACTED PASSAGES ---
{passages[:8000]}  
--- END OF PASSAGES ---

Based solely on the passages above, classify all fields."""

    try:  
        response = client.chat.completions.create(
            model="gpt-4o",
            max_tokens=1000, # 足夠讓模型回傳完整 JSON
            messages=[
                {"role": "system", "content": system_prompt},  
                {"role": "user",   "content": user_prompt}
            ],
            temperature=0, # 固定輸出，避免同一文本每次不同
        )
        raw = response.choices[0].message.content.strip() 
        raw = re.sub(r'^```json\s*', '', raw)  
        raw = re.sub(r'\s*```$', '', raw)

        try:
            return json.loads(raw)  # 嘗試解析 JSON
        except json.JSONDecodeError:
            def grab(pattern, text):
                m = re.search(pattern, text)
                return int(m.group(1)) if m else None

            p2 = grab(r'"is_Pledge2"\s*:\s*([012])', raw)  # 嘗試用 Regex 抓關鍵欄位，看看能不能補救 JSON 解析失敗的情況
            p3 = grab(r'"is_Pledge3"\s*:\s*([012])', raw)
            pb = grab(r'"Pledge2_breadth"\s*:\s*([123])', raw)
            ps = grab(r'"Pledge2_strength"\s*:\s*([123])', raw)
            hb = grab(r'"Pledge3_breadth"\s*:\s*([123])', raw)
            hs = grab(r'"Pledge3_strength"\s*:\s*([123])', raw)
            ev = re.search(r'"evidence"\s*:\s*"([^"]{0,300})', raw)

            if p2 is not None and p3 is not None:
                print("  [JSON 不完整，Regex fallback 成功]")
                return {
                    "is_Pledge2": p2, "is_Pledge3": p3,
                    "Pledge2_breadth": pb, "Pledge2_strength": ps,
                    "Pledge3_breadth": hb, "Pledge3_strength": hs,
                    "evidence": ev.group(1) + "...[truncated]" if ev else "evidence truncated"
                }
            else:
                print(f"  [JSON 解析失敗]\n  原始回應：{raw[:200]}")
                return {**NULL_RESULT, "evidence": f"parse failed: {raw[:100]}"}

    except Exception as e:
        print(f"  [API 呼叫失敗] {e}")
        return {**NULL_RESULT, "evidence": f"API error: {str(e)}"}

# ══════════════════════════════════════════════════════════════════════════════
# 讀取資料
# ══════════════════════════════════════════════════════════════════════════════

print("讀取資料中...")
wb = openpyxl.load_workbook(FILE_PATH, read_only=True, data_only=True)
sheet = wb.active
headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))] 

def find_col(name):
    try:
        return headers.index(name)
    except ValueError:
        raise ValueError(f"找不到欄位「{name}」\n現有欄位：{headers}")

idx_cik      = find_col('CIK')
idx_filedate = find_col('FILEDATE')
idx_http     = find_col('http')

all_data = [] # 先讀取所有資料到 list，再轉 DataFrame，避免一次讀取大量資料導致記憶體問題
for row in sheet.iter_rows(min_row=2): 
    all_data.append({
        'CIK':      get_cell(row, idx_cik),
        'FILEDATE': get_cell(row, idx_filedate),
        'http':     get_cell(row, idx_http),
    })
wb.close()

df = pd.DataFrame(all_data)
df = df[df['http'].str.startswith('http', na=False)].copy()
df = df.reset_index(drop=True)
print(f"有效筆數（含網址）：{len(df)}")

# ══════════════════════════════════════════════════════════════════════════════
# 取樣
# ══════════════════════════════════════════════════════════════════════════════

if SAMPLE_N is not None and SAMPLE_N < len(df):
    samples = df.sample(n=SAMPLE_N, random_state=RANDOM_SEED).reset_index(drop=True)
    print(f"隨機抽樣 {SAMPLE_N} 筆（共 {len(df)} 筆中）")
else:
    samples = df.copy()
    print(f"執行全部 {len(df)} 筆")

print(samples[['CIK', 'FILEDATE']].to_string())

# ══════════════════════════════════════════════════════════════════════════════
# 主流程：fetch → Regex → LLM → 輸出
# ══════════════════════════════════════════════════════════════════════════════

results = [] # 最終結果列表，每筆是一個 dict，最後會轉成 DataFrame

for i, row in samples.iterrows(): 
    cik      = row['CIK']
    filedate = row['FILEDATE']
    url      = row['http']

    print(f"\n[{i+1}/{len(samples)}] CIK={cik} DATE={filedate}")
    print(f"  URL: {url}")

    # Step 1: 抓 SEC 文件
    raw_html = fetch_html(url)
    if not raw_html: 
        print("  → 抓取失敗，跳過")
        results.append({
            'CIK': cik, 'FILEDATE': filedate,
            'AI_Pledge2': -1, 'AI_Pledge3': -1,  # 無法分類的預設值
            'Pledge2_breadth': None, 'Pledge2_strength': None,
            'Pledge3_breadth': None, 'Pledge3_strength': None,
            'AI_依據文字': '文件抓取失敗',
            'Regex命中區塊數': 0,
            'Regex段落': '',
        })
        continue

    # Step 2: Regex 粗篩
    passages  = extract_relevant_passages(raw_html)
    hit_count = passages.count('\n\n---\n\n') + 1 if passages else 0 # 計算命中區塊數（以分隔符為基準）
    print(f"  Regex 命中 {hit_count} 個區塊，擷取 {len(passages)} 字元")

    # Step 3: LLM 分類
    llm_result     = classify_with_llm(passages, cik, filedate)
    ai_p2          = llm_result.get('is_Pledge2', -1) 
    ai_p3          = llm_result.get('is_Pledge3', -1)
    ai_p2_breadth  = llm_result.get('Pledge2_breadth', None)
    ai_p2_strength = llm_result.get('Pledge2_strength', None)
    ai_p3_breadth  = llm_result.get('Pledge3_breadth', None)
    ai_p3_strength = llm_result.get('Pledge3_strength', None)
    evidence       = llm_result.get('evidence', '')

    print(f"  AI判斷: Pledge2={ai_p2} (breadth={ai_p2_breadth}, strength={ai_p2_strength})")
    print(f"          Pledge3={ai_p3} (breadth={ai_p3_breadth}, strength={ai_p3_strength})")
    print(f"  依據: {evidence[:120]}")

    results.append({
        'CIK':              cik,
        'FILEDATE':         filedate,
        'AI_Pledge2':       ai_p2,
        'AI_Pledge3':       ai_p3,
        'Pledge2_breadth':  ai_p2_breadth,
        'Pledge2_strength': ai_p2_strength,
        'Pledge3_breadth':  ai_p3_breadth,
        'Pledge3_strength': ai_p3_strength,
        'AI_依據文字':      evidence,
        'Regex命中區塊數':  hit_count,
        'Regex段落':        passages[:6000],
    })

    time.sleep(1)  # 避免 rate limit

# ══════════════════════════════════════════════════════════════════════════════
# 輸出結果
# ══════════════════════════════════════════════════════════════════════════════

result_df = pd.DataFrame(results)

print("\n" + "="*70)
print("【分類結果摘要】")
print("="*70)
print(result_df[['CIK', 'FILEDATE', 'AI_Pledge2', 'AI_Pledge3']].to_string(index=False))

out_path = 'LLM分類結果.xlsx'
result_df.to_excel(out_path, index=False)
print(f"\n結果已輸出至 {out_path}")