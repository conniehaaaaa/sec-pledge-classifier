"""
SEC DEF 14A 質押／避險政策分類器
===================================
流程：從 Excel 讀取網址 → 爬取 SEC 網頁 → Regex 粗篩段落 → GPT-4o 分類 → 輸出結果

輸出欄位：
  CIK | FILEDATE | AI_Pledge2 | AI_Pledge3 |
  Pledge2_breadth | Pledge2_strength | Pledge3_breadth | Pledge3_strength |
  AI_依據文字 | Regex命中區塊數 | Regex段落

修改紀錄：
  v2 - 新增段落噪音過濾器（DISCARD_PATTERN），修正 LLM system prompt false positive 規則
  v3 - 補充 6 條 DISCARD_PATTERN
  v4 - 移除 DISCARD_PATTERN，改由 LLM system prompt 統一處理噪音判斷
       好處：不需要再手動維護 Regex pattern，新年份遇到新噪音只需補充 prompt 描述
  v5 - 修正 prompt 中 G、H、I、J 四條噪音規則
  v6 - 將 A-J 十條細則收斂為兩條核心原則（Filter 1: 對象是誰、Filter 2: 是規定還是事實）
       大幅簡化 prompt，降低錯殺忠臣的風險
"""

import re
import json
import time
import random
import html
import os
import pandas as pd
import requests
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════════════════════════════════════════
# 設定區（依需求修改這裡）
# ══════════════════════════════════════════════════════════════════════════════

FILE_PATH   = 'CYZ_DEF14A.xlsx'
SHEET_NAME  = 'DEF14A_06_25'
YEAR        = 2007           # ← 修改這裡換年份

SAMPLE_N    = None             # None = 跑全部；數字 = 隨機抽樣
RANDOM_SEED = 42

# ══════════════════════════════════════════════════════════════════════════════

random.seed(RANDOM_SEED)
client = OpenAI()

CHECKPOINT_FILE = f'checkpoint_{YEAR}.xlsx'
PROCESSED_FILE  = f'processed_{YEAR}.txt'
OUTPUT_FILE     = f'LLM分類結果_{YEAR}.xlsx'
SAVE_EVERY      = 20

# ══════════════════════════════════════════════════════════════════════════════
# Regex 粗篩（唯一用途：把完全沒有相關字眼的段落濾掉，降低送 LLM 的量）
# 不做語意判斷，語意判斷全部交給 LLM
# ══════════════════════════════════════════════════════════════════════════════

V5_PATTERN = re.compile(
    r'pledg\w+|short\s+sale\w*|derivative'
    r'|margin\s+account'
    r'|anti.?hedg|anti.?pledg'
    r'|zero.?cost\s+collar|prepaid\s+variable\s+forward'
    r'|monetization\s+transaction'
    r'|\bcollateral\b|\bmargin\s+loans?\b',
    re.IGNORECASE | re.DOTALL
)

# ══════════════════════════════════════════════════════════════════════════════
# LLM System Prompt
#
# 核心設計：
#   Step 1 — 讓 LLM 先判斷每個段落的「文件性質」，過濾噪音
#   Step 2 — 只對真正是 insider trading policy 的段落進行分類
#
# 好處：遇到新的噪音類型，只需在 [NOISE TYPES] 下方加一行描述，不需要改 code
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are an SEC DEF 14A analyst specializing in pledging and hedging policies.

════════════════════════════════════════════════════════════
STEP 1 — FILTER: IS THIS A RELEVANT PASSAGE?
════════════════════════════════════════════════════════════

The passages you receive were pulled by a keyword search and may contain noise.
Before classifying, apply these two filters. If EITHER filter applies, ignore the
passage entirely — treat it as if it does not exist.

FILTER 1 — Wrong subject: the restriction does not apply to individual insiders.
  A relevant passage restricts what individual insiders (employees, officers,
  directors) may do with their own company securities.
  IGNORE a passage if the subject is any of the following instead:
    • The company itself  (e.g., "The Company shall not invest in short sales...")
    • A fund entity       (e.g., "The Fund may not pledge its assets...")
    • Unissued equity awards  (e.g., "Unvested RSUs may not be pledged...")
    • LLC/partnership units   (e.g., "Members may not transfer their Units...")
    • A specific large shareholder defined by ownership percentage
                          (e.g., "Any 5% Stockholder may not pledge shares...")
    • SPAC sponsors or IPO-bound individuals under a one-time lock-up agreement
                          (e.g., "Sponsors may not pledge shares until 180 days after IPO...")

FILTER 2 — Wrong type: the passage is a factual statement, not a policy rule.
  A relevant passage states a rule about what insiders must or must not do.
  IGNORE a passage if it is instead:
    • A disclosure of current facts  (e.g., "No shares are currently pledged.")
    • A loan contract trigger clause (e.g., "If participant pledges shares,
      interest rate will be adjusted...")
    • An explicit statement that no restriction exists
                          (e.g., "We do not restrict pledging.")

If a passage passes both filters → it is relevant. Proceed to Step 2.

════════════════════════════════════════════════════════════
STEP 2 — CLASSIFY: LABEL THE POLICY
════════════════════════════════════════════════════════════

Classify based only on passages that passed Step 1.

--- LABELS ---

is_Pledge2 (Anti-Pledging):
  Does a formal policy prohibit insiders from pledging company shares as collateral?

is_Pledge3 (Anti-Hedging):
  Does a formal policy prohibit insiders from hedging company securities
  (e.g., short sales, collars, swaps, forwards, puts, calls)?

  0 = No relevant policy found
  1 = Prohibited, no exceptions allowed
  2 = Prohibited, but exceptions or pre-approval are permitted

--- RULES ---

→ Label 1: passage says insiders "are prohibited," "may not," "shall not," or
  "are not permitted" to pledge or hedge, with no exception process mentioned.
  - Banning margin accounts = anti-pledging (label 1), even without the word "pledge."

→ Label 2: there IS a prohibition, but an exception or pre-approval process exists.
  ("unless pre-approved," "subject to prior approval," "with General Counsel consent")
  - IMPORTANT: pre-clearance or approval alone, with no prohibition language = label 0,
    not label 2. The prohibition must be present.

→ Label 0: no relevant passage found after filtering, OR no prohibition language found.

--- POLICY QUALITY (fill only when is_Pledge2 or is_Pledge3 ≥ 1; else set null) ---

breadth — who is covered?
  1 = All employees / firm-wide
  2 = Directors and executive officers / senior management only
  3 = Narrow group only (NEOs, specific named individuals)

strength — how binding is the language?
  1 = Strong:   "prohibited," "may not," "shall not," "are not permitted"
  2 = Moderate: "generally prohibits," "not permitted without approval," "subject to pre-clearance"
  3 = Weak:     "discouraged," "should not," "not recommended"

════════════════════════════════════════════════════════════
OUTPUT
════════════════════════════════════════════════════════════

Respond ONLY with valid JSON, no markdown, no extra text:
{
  "is_Pledge2": <0|1|2>,
  "is_Pledge3": <0|1|2>,
  "Pledge2_breadth": <1|2|3|null>,
  "Pledge2_strength": <1|2|3|null>,
  "Pledge3_breadth": <1|2|3|null>,
  "Pledge3_strength": <1|2|3|null>,
  "evidence": "<key sentence(s) from the actual policy, up to 3, separated by | . Write NONE if all passages were noise.>"
}"""


# ══════════════════════════════════════════════════════════════════════════════
# 工具函式
# ══════════════════════════════════════════════════════════════════════════════

def fetch_html(url: str, timeout: int = 15) -> str:
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


def extract_relevant_passages(raw_html: str) -> tuple[str, int]:
    """
    Regex 粗篩：只負責找有關鍵字的段落，不做語意過濾
    回傳 (passages_str, hit_count)
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
    kept       = []

    for m in BLOCK_TAG_RE.finditer(raw_html):
        block_text = strip_tags(m.group(2))
        if len(block_text) < 20:
            continue
        if not V5_PATTERN.search(block_text):
            continue
        key = block_text[:120]
        if key in seen_texts:
            continue
        seen_texts.add(key)
        kept.append(block_text)

    passages = "\n\n---\n\n".join(kept) if kept else ""
    return passages, len(kept)


def classify_with_llm(passages: str, cik: str, filedate: str) -> dict:
    NULL_RESULT = {
        "is_Pledge2": -1, "is_Pledge3": -1,
        "Pledge2_breadth": None, "Pledge2_strength": None,
        "Pledge3_breadth": None, "Pledge3_strength": None,
        "evidence": "[No relevant passages found by Regex — unable to classify]"
    }

    if not passages.strip():
        return NULL_RESULT

    user_prompt = f"""The following passages were extracted from the DEF 14A proxy statement
filed by CIK={cik} on {filedate}.

They were selected by a keyword filter (Regex) for containing words like
"pledge," "hedge," "margin," "derivative," "short sale," or "collateral."
The passages may include both genuine insider trading policies AND noise
(disclosures, plan documents, fund agreements, etc.).

Please apply Step 1 (filter noise) and Step 2 (classify) as instructed.

--- EXTRACTED PASSAGES ---
{passages[:8000]}
--- END OF PASSAGES ---"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            max_tokens=1000,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": user_prompt}
            ],
            temperature=0,
        )
        raw = response.choices[0].message.content.strip()
        raw = re.sub(r'^```json\s*', '', raw)
        raw = re.sub(r'\s*```$',     '', raw)

        try:
            result = json.loads(raw)
            if not result.get('evidence'):
                result['evidence'] = '[LLM returned no evidence]'
            return result

        except json.JSONDecodeError:
            def grab(pattern, text):
                m = re.search(pattern, text)
                return int(m.group(1)) if m else None

            p2 = grab(r'"is_Pledge2"\s*:\s*([012])', raw)
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
                    "evidence": (ev.group(1) + "...[truncated]") if ev else "[evidence truncated]"
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

print(f"讀取資料中（篩選年份：{YEAR}）...")
df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, index_col=None)
df.columns = df.columns.str.strip()
print(df.columns.tolist())

df['FILEDATE'] = pd.to_datetime(df['FILEDATE'], format='%Y/%m/%d', errors='coerce')
df = df[df['FILEDATE'].dt.year == YEAR].copy()
df = df.reset_index(drop=True)
print(f"{YEAR} 年有效筆數：{len(df)}")

# ══════════════════════════════════════════════════════════════════════════════
# 取樣
# ══════════════════════════════════════════════════════════════════════════════

if SAMPLE_N is not None and SAMPLE_N < len(df):
    samples = df.sample(n=SAMPLE_N, random_state=RANDOM_SEED).reset_index(drop=True)
    print(f"隨機抽樣 {SAMPLE_N} 筆（共 {len(df)} 筆中）")
else:
    samples = df.copy()
    print(f"執行全部 {len(df)} 筆")

# ══════════════════════════════════════════════════════════════════════════════
# 斷點續跑
# ══════════════════════════════════════════════════════════════════════════════

if os.path.exists(CHECKPOINT_FILE):
    df_done = pd.read_excel(CHECKPOINT_FILE)
    results = df_done.to_dict('records')
    print(f"讀取斷點，已有結果 {len(results)} 筆")
else:
    results = []

if os.path.exists(PROCESSED_FILE):
    with open(PROCESSED_FILE, 'r') as f:
        done_urls = set(f.read().splitlines())
    print(f"已處理 URL：{len(done_urls)} 筆")
else:
    done_urls = set()

# ══════════════════════════════════════════════════════════════════════════════
# 主流程：fetch → Regex 粗篩 → LLM（含噪音過濾＋分類）→ 輸出
# ══════════════════════════════════════════════════════════════════════════════

actual_count = 0

for i, row in samples.iterrows():
    cik      = row['CIK']
    filedate = row['FILEDATE']
    url      = row['filename']

    if url in done_urls:
        continue

    print(f"\n[{i+1}/{len(samples)}] CIK={cik} DATE={filedate}")
    print(f"  URL: {url}")

    # Step 1: 抓 SEC 文件
    raw_html = fetch_html(url)
    if not raw_html:
        print("  → 抓取失敗，跳過（下次重試）")
        time.sleep(0.5)
        continue

    # Step 2: Regex 粗篩（只找關鍵字段落，不做語意過濾）
    passages, hit_count = extract_relevant_passages(raw_html)
    print(f"  Regex 命中 {hit_count} 個區塊，擷取 {len(passages)} 字元")

    # Step 3: LLM（自行過濾噪音 + 分類）
    llm_result     = classify_with_llm(passages, cik, filedate)
    ai_p2          = llm_result.get('is_Pledge2', -1)
    ai_p3          = llm_result.get('is_Pledge3', -1)
    ai_p2_breadth  = llm_result.get('Pledge2_breadth', None)
    ai_p2_strength = llm_result.get('Pledge2_strength', None)
    ai_p3_breadth  = llm_result.get('Pledge3_breadth', None)
    ai_p3_strength = llm_result.get('Pledge3_strength', None)
    evidence       = llm_result.get('evidence', '[LLM returned no evidence]')

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

    done_urls.add(url)
    with open(PROCESSED_FILE, 'a') as f:
        f.write(url + '\n')

    actual_count += 1
    time.sleep(1)

    if actual_count % SAVE_EVERY == 0:
        pd.DataFrame(results).to_excel(CHECKPOINT_FILE, index=False)
        print(f"  已存斷點：實際處理 {actual_count} 筆")

# ══════════════════════════════════════════════════════════════════════════════
# 最終輸出
# ══════════════════════════════════════════════════════════════════════════════

result_df = pd.DataFrame(results)

print("\n" + "="*70)
print(f"【{YEAR} 年分類結果摘要】")
print("="*70)
print(result_df[['CIK', 'FILEDATE', 'AI_Pledge2', 'AI_Pledge3']].to_string(index=False))

result_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n結果已輸出至 {OUTPUT_FILE}")