"""
SEC DEF 14A 質押／避險政策分類器
===================================
流程：從 Excel 讀取網址 → 爬取 SEC 網頁 → Regex 粗篩段落 → GPT-4o 分類 → 輸出結果

輸出欄位：
  CIK | FILEDATE | AI_Pledge2 | AI_Pledge3 |
  Pledge2_breadth | Pledge2_strength | Pledge3_breadth | Pledge3_strength |
  AI_依據文字 | Regex命中區塊數 | Regex過濾後區塊數 | Regex段落

修改紀錄：
  v2 - 新增段落噪音過濾器（DISCARD_PATTERN），修正 LLM system prompt false positive 規則
  v3 - 補充 6 條 DISCARD_PATTERN（Fund Advisor 合約、REIT 投資限制、Shareholder/Member Agreement、
       IPO lock-up Transfer Restriction、Loan Program 補版、X% Stockholder 股份限制）
       補強 KEEP_OVERRIDE（加入 'it is our policy'）
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
# V5 粗篩 Regex
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
# 段落噪音過濾器（Regex 命中後、送 LLM 前執行）
#
# 凡符合以下任一 pattern 的段落視為噪音，直接丟棄，不送 LLM：
#   1. 純揭露：「沒有人質押」→ 描述現狀，不是政策
#   2. Equity award/grant plan 移轉限制 → 計畫層級，非 insider trading policy
#   3. 基金資產質押限制（Fund's assets）→ 基金操作規定，非公司 insider policy
#   4. Fund margin account 限制 → 同上
#   5. LLC / Partnership Unit 移轉協議 → 非 insider trading policy
#   6. Loan Program 質押條款 → 合約觸發條款，非禁止政策
#   7. SPAC / lock-up 協議 → 不適用
#   8. 明確說「不限制質押」→ 無政策
# ══════════════════════════════════════════════════════════════════════════════

DISCARD_PATTERN = re.compile(
    # 1. 純揭露：「沒有 executive/director 質押」
    r'no\s+(executive|director|officer|named\s+executive).{0,80}(pledged|margin)'

    # 2. Award/Grant 計畫層級移轉限制（unvested shares 禁止轉讓）
    r'|(award|grant|RSU|option|restricted\s+stock).{0,120}'
    r'(may\s+not|shall\s+not|not\s+be|cannot).{0,80}(pledg|transfer|assign|encumber)'
    r'|(may\s+not|shall\s+not|cannot).{0,80}(sell|assign|transfer|pledge)'
    r'.{0,80}(award|grant|RSU|restricted\s+stock|unvested)'

    # 3. 基金資產質押
    r'|fund.{0,60}(pledging|hypothecat|mortgag)'
    r'|pledging.{0,60}fund.{0,30}assets'

    # 4. Fund margin account（帳戶名義為基金）
    r'|no\s+margin\s+accounts?\s+(are\s+to\s+be|to\s+be)\s+maintained\s+in\s+the\s+name\s+of\s+the\s+fund'

    # 5. LLC / Partnership Unit 移轉協議
    r'|(section\s+\d+|unit\s+assignment\s+policy).{0,200}(pledge|transfer)\s+of.{0,30}units'

    # 6. Loan Program 質押觸發條款（原版：sells or pledges ... loan program）
    r'|(sells?\s+or\s+pledges?|pledges?\s+or\s+sells?).{0,60}loan\s+program'
    r'|loan\s+program.{0,200}(sells?\s+or\s+pledges?|pledges?\s+or\s+sells?)'

    # 7. SPAC / lock-up
    r'|lock.?up.{0,80}(pledg|transfer)'
    r'|sponsor.{0,60}(pledg|not\s+to\s+sell)'

    # 8. 明確說不限制
    r'|do\s+not\s+restrict\s+pledges'

    # ── v3 新增 ──────────────────────────────────────────────────────────────

    # 9. Fund Advisor 合約：Advisor agrees ... pledge/borrow/short position ... Fund
    r'|advisor\s+agrees.{0,150}(pledge|borrow|short\s+position).{0,100}fund'

    # 10. REIT / 公司整體投資組合限制（Articles of Incorporation / Investment Restrictions）
    r'|company\s+shall\s+not\s+make\s+investments\s+in.{0,150}short\s+sales'

    # 11. Shareholder / Member Agreement 移轉或質押禁止（合約條款，非 insider policy）
    r'|(shareholder|member).{0,50}(shall\s+not|will\s+not|prohibited).{0,100}'
    r'(encumber|pledg|transfer).{0,80}(stock|shares|units)'

    # 12. IPO lock-up / Transfer Restrictions：covered managing director + hedging
    r'|transfer\s+restrictions?.{0,200}(covered\s+managing|managing\s+director)'
    r'.{0,300}(hedg|short\s+sale|derivative)'

    # 13. Loan Program 補版：in the event a participant sells or pledges ... loan program
    r'|in\s+the\s+event\s+a\s+participant\s+(sells|pledges).{0,100}loan\s+program'

    # 14. X% Stockholder 持股移轉限制（非 insider trading policy）
    r'|\d+%\s+stockholder\s+(cannot|may\s+not).{0,100}(pledge|sell|transfer)'
    r'.{0,100}without.{0,50}(board|approval)',

    re.IGNORECASE | re.DOTALL
)

# 白名單：就算觸發 DISCARD，只要同時出現正式政策語言，仍然保留
KEEP_OVERRIDE = re.compile(
    r'insider\s+trading\s+policy'
    r'|securities\s+trading\s+policy'
    r'|company\s+policy\s+prohibit'
    r'|it\s+is\s+our\s+policy'                              # v3 新增
    r'|board.{0,30}adopted.{0,60}(policy|guideline)'
    r'|compensation\s+committee.{0,60}(policy|guideline|prohibit)',
    re.IGNORECASE | re.DOTALL
)


def should_discard_passage(text: str) -> bool:
    """回傳 True 表示這段應該被丟棄（噪音）"""
    if DISCARD_PATTERN.search(text):
        if KEEP_OVERRIDE.search(text):
            return False  # 白名單蓋過，保留
        return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# LLM System Prompt（已補強 false positive 規則）
# ══════════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are an SEC DEF 14A analyst. Classify pledging and hedging policies from the provided passages.

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
  - IMPORTANT: pre-clearance ALONE (without an explicit prohibition) is NOT label 2.
    A passage saying only "transactions require pre-clearance" with no prohibition → label 0.

→ Label 0 — the following are NOT policies (common false positives — read carefully):

  A. Ownership disclosure / current status:
     - "X shares pledged as collateral" → disclosure of current holdings, not a prohibition.
     - "No executive officer has pledged any shares" → describes current status, NOT a policy.
       Even though phrased negatively, it is still a factual disclosure, not a rule.

  B. Equity award / stock plan transferability restrictions:
     - "Participant may not sell, transfer, or pledge unvested awards/RSUs/options/grants" →
       This is a plan-level transferability rule for unissued instruments, NOT a company insider
       trading policy. Key signal: the subject is an "award," "grant," "RSU," "option," or
       "restricted stock" and the restriction applies specifically to those instruments → label 0.

  C. Fund-level operating restrictions:
     - "No margin accounts to be maintained in the name of the Fund" → fund operating rule, not
       an insider trading policy → label 0.
     - "Fund may not pledge its assets" → same → label 0.

  D. LLC / Partnership Unit transfer agreements:
     - "Section 9 requires Director consent for any pledge of Units" → this is an LLC operating
       agreement on member unit transfers, NOT a personal insider trading policy → label 0.

  E. Loan Program pledge clauses:
     - "Interest rate is adjusted if participant pledges shares purchased under the Loan Program" →
       this is a loan contract term, not a prohibition on pledging → label 0.

  F. SPAC / lock-up / IPO Transfer Restriction agreements:
     - "Sponsors agreed not to sell or pledge shares during lock-up" → contractual lock-up,
       NOT an insider trading policy → label 0.
     - "Transfer Restrictions. Each covered managing director agrees not to transfer... including
       short sales, hedging arrangements..." → IPO personal lock-up agreement, NOT a company
       insider trading policy → label 0.

  G. Fund Advisor contracts:
     - "The Advisor agrees that neither it nor its employees shall pledge the Fund's assets" →
       this is an investment advisory agreement restricting the fund manager's use of fund assets,
       NOT a company insider trading policy → label 0.

  H. REIT / company-level investment restrictions (Articles of Incorporation):
     - "The Company shall not make investments in: (i) foreign currency; (ii) short sales..." →
       this restricts the company's investment portfolio, NOT individual insider trading → label 0.

  I. Shareholder / Member agreements:
     - "Shareholder shall not encumber or pledge any Stock except upon written consent" →
       this is a shareholder agreement clause, NOT an insider trading policy → label 0.
     - "5% Stockholder cannot purchase, sell, pledge... without Board approval" →
       ownership transfer restriction for large holders, NOT an insider trading policy → label 0.

  J. Explicit statement of no restriction:
     - "We do not restrict pledges but require pre-clearance" → explicitly says no restriction
       exists → label 0.

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


def extract_relevant_passages(raw_html: str) -> tuple[str, int, int]:
    """
    回傳 (filtered_passages_str, regex_hit_count, after_filter_count)
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

    seen_texts  = set()
    regex_hits  = []
    kept        = []

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
        regex_hits.append(block_text)

    regex_hit_count = len(regex_hits)

    # 噪音過濾
    for passage in regex_hits:
        if should_discard_passage(passage):
            continue
        kept.append(passage)

    filtered_passages = "\n\n---\n\n".join(kept) if kept else ""
    return filtered_passages, regex_hit_count, len(kept)


def classify_with_llm(passages: str, cik: str, filedate: str) -> dict:
    NULL_RESULT = {
        "is_Pledge2": -1, "is_Pledge3": -1,
        "Pledge2_breadth": None, "Pledge2_strength": None,
        "Pledge3_breadth": None, "Pledge3_strength": None,
        "evidence": "[No relevant passages found by Regex — unable to classify]"
    }

    if not passages.strip():
        return NULL_RESULT

    user_prompt = f"""The following passages were extracted from the DEF 14A proxy statement filed by CIK={cik} on {filedate}.
They were selected by a keyword filter (Regex) for relevance to pledging and hedging policies.

--- EXTRACTED PASSAGES ---
{passages[:8000]}
--- END OF PASSAGES ---

Based solely on the passages above, classify all fields."""

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
        raw = re.sub(r'\s*```$', '', raw)

        try:
            result = json.loads(raw)
            # evidence 保底：確保不會是 None/空
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
# 主流程：fetch → Regex → 噪音過濾 → LLM → 輸出
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

    # Step 2: Regex 粗篩 + 噪音過濾
    passages, regex_hit_count, kept_count = extract_relevant_passages(raw_html)
    print(f"  Regex 命中 {regex_hit_count} 個區塊 → 噪音過濾後保留 {kept_count} 個，擷取 {len(passages)} 字元")

    # Step 3: LLM 分類
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
        'CIK':               cik,
        'FILEDATE':          filedate,
        'AI_Pledge2':        ai_p2,
        'AI_Pledge3':        ai_p3,
        'Pledge2_breadth':   ai_p2_breadth,
        'Pledge2_strength':  ai_p2_strength,
        'Pledge3_breadth':   ai_p3_breadth,
        'Pledge3_strength':  ai_p3_strength,
        'AI_依據文字':       evidence,
        'Regex命中區塊數':   regex_hit_count,
        'Regex過濾後區塊數': kept_count,     # 新增欄位：方便事後診斷過濾率
        'Regex段落':         passages[:6000],
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