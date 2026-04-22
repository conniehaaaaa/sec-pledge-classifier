"""
用 206_withURL.xlsx 裡的 URL 重新爬取並分類
輸出：206_rerun_result.xlsx
"""

import re
import json
import time
import html
import os
import pandas as pd
import requests
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════════════════════════════
# 設定區
# ══════════════════════════════════════════════════════════════════

INPUT_FILE      = '2006_withURL.xlsx'
OUTPUT_FILE     = '2006_rerun_result.xlsx'
CHECKPOINT_FILE = '2006_checkpoint.xlsx'
PROCESSED_FILE  = '2006_processed.txt'
SAVE_EVERY      = 20

# ══════════════════════════════════════════════════════════════════
# Regex 粗篩
# ══════════════════════════════════════════════════════════════════

V5_PATTERN = re.compile(
    r'pledg\w+|short\s+sale\w*|derivative'
    r'|margin\s+account'
    r'|anti.?hedg|anti.?pledg'
    r'|zero.?cost\s+collar|prepaid\s+variable\s+forward'
    r'|monetization\s+transaction'
    r'|\bcollateral\b|\bmargin\s+loans?\b',
    re.IGNORECASE | re.DOTALL
)

# ══════════════════════════════════════════════════════════════════
# LLM System Prompt
# ══════════════════════════════════════════════════════════════════

SYSTEM_PROMPT = """You are a finance professor and analyst specializing in anti-hedging and anti-pledging policies disclosed in SEC DEF 14A filings available through the EDGAR database.

## STEP 1 — FILTER (ignore passage if either applies)
FILTER 1 — Classify the passage as the wrong subject if the restriction does not apply to corporate insiders.
IGNORE if it applies to: the company itself, a fund entity, unissued/unvested equity, LLC/partnership units, or SPAC sponsors under a one-time IPO lock-up.

FILTER 2 — Classify the passage as a wrong type if the statement is not a policy rule. 
IGNORE if it is: a factual disclosure, a loan contract trigger clause, or an explicit statement that no restrictions on pledging or hedging exists.

## STEP 2 — CLASSIFY
**is_Pledge2** (Anti-Pledging): formal policy prohibits insiders from pledging shares as collateral.
**is_Pledge3** (Anti-Hedging): formal policy prohibits hedging (e.g., puts, calls, swaps, collars, forwards).

Values: 0 = no relevant policy | 1 = prohibited, no exceptions | 2 = prohibited with exceptions/pre-approval

Rules:
- Label 1: passage says insiders "are prohibited," "may not," "shall not," or
  "are not permitted" to pledge or hedge, with no exception process mentioned.
  Note: banning margin accounts = anti-pledging, label 1.
- Label 2: prohibition exists AND exception or pre-approval process is mentioned.
  Note: pre-clearance alone with NO prohibition language = label 0, not label 2.
- Label 0: no relevant passage after filtering, or no prohibition language.

## STEP 3 — POLICY QUALITY (only when is_Pledge2 or is_Pledge3 ≥ 1; else null)

breadth: 1 = narrow (NEOs, specific individual) | 2 = directors & executive officers  | 3 = all employees/firm-wide 
strength: 1 = weak ("discouraged") | 2 = moderate ("generally prohibits," "subject to pre-clearance") | 3 = strong ("prohibited," "may not," "shall not")

## OUTPUT — valid JSON only, no markdown:
{
  "is_Pledge2": <0|1|2>,
  "is_Pledge3": <0|1|2>,
  "Pledge2_breadth": <1|2|3|null>,
  "Pledge2_strength": <1|2|3|null>,
  "Pledge3_breadth": <1|2|3|null>,
  "Pledge3_strength": <1|2|3|null>,
  "evidence": "<up to 3 key sentences separated by | . NONE if all passages were noise.>"
}
"""

# ══════════════════════════════════════════════════════════════════

client = OpenAI()

NULL_RESULT = {
    "is_Pledge2": -1, "is_Pledge3": -1,
    "Pledge2_breadth": None, "Pledge2_strength": None,
    "Pledge3_breadth": None, "Pledge3_strength": None,
    "evidence": "[No relevant passages found by Regex — unable to classify]"
}


def fetch_html(url: str, timeout: int = 15) -> str:
    try:
        resp = requests.get(url, timeout=timeout,
                            headers={'User-Agent': 'yizhen1426@gmail.com'})
        resp.raise_for_status()
        return resp.text
    except Exception as e:
        print(f"  [fetch 失敗] {url}: {e}")
        return ""


def extract_relevant_passages(raw_html: str) -> tuple:
    BLOCK_TAG_RE = re.compile(
        r'<(p|div|li|td|th|span|section|article)[^>]*>(.*?)</\1>',
        re.IGNORECASE | re.DOTALL
    )

    def strip_tags(text):
        text = re.sub(r'<[^>]+>', ' ', text)
        text = html.unescape(text)
        return re.sub(r'\s+', ' ', text).strip()

    seen, kept = set(), []
    for m in BLOCK_TAG_RE.finditer(raw_html):
        block = strip_tags(m.group(2))
        if len(block) < 20 or not V5_PATTERN.search(block):
            continue
        key = block[:120]
        if key in seen:
            continue
        seen.add(key)
        kept.append(block)

    return "\n\n---\n\n".join(kept), len(kept)


def classify_with_llm(passages: str, cik, filedate) -> dict:
    if not passages or not str(passages).strip():
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
        raw = re.sub(r'\s*```$', '', raw)

        try:
            result = json.loads(raw)
            if not result.get('evidence'):
                result['evidence'] = '[LLM returned no evidence]'
            return result
        except json.JSONDecodeError:
            def grab(pat, t):
                m = re.search(pat, t)
                return int(m.group(1)) if m else None
            p2 = grab(r'"is_Pledge2"\s*:\s*([012])', raw)
            p3 = grab(r'"is_Pledge3"\s*:\s*([012])', raw)
            if p2 is not None and p3 is not None:
                print("  [JSON 不完整，Regex fallback 成功]")
                return {
                    "is_Pledge2": p2, "is_Pledge3": p3,
                    "Pledge2_breadth":  grab(r'"Pledge2_breadth"\s*:\s*([123])', raw),
                    "Pledge2_strength": grab(r'"Pledge2_strength"\s*:\s*([123])', raw),
                    "Pledge3_breadth":  grab(r'"Pledge3_breadth"\s*:\s*([123])', raw),
                    "Pledge3_strength": grab(r'"Pledge3_strength"\s*:\s*([123])', raw),
                    "evidence": "[fallback parsed]"
                }
            print(f"  [JSON 解析失敗] {raw[:200]}")
            return {**NULL_RESULT, "evidence": f"parse failed: {raw[:100]}"}

    except Exception as e:
        print(f"  [API 呼叫失敗] {e}")
        return {**NULL_RESULT, "evidence": f"API error: {str(e)}"}


# ══════════════════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════════════════

print(f"讀取 {INPUT_FILE} ...")
df = pd.read_excel(INPUT_FILE)
print(f"共 {len(df)} 筆")

# 斷點續跑
results   = pd.read_excel(CHECKPOINT_FILE).to_dict('records') if os.path.exists(CHECKPOINT_FILE) else []
done_urls = set(open(PROCESSED_FILE).read().splitlines()) if os.path.exists(PROCESSED_FILE) else set()
print(f"已處理：{len(done_urls)} 筆")

actual_count = 0

for i, row in df.iterrows():
    cik      = row['CIK']
    filedate = row['FILEDATE']
    url      = row.get('filename', '')

    if pd.isna(url) or str(url).strip() == '':
        print(f"\n[{i+1}/{len(df)}] CIK={cik}   無 URL，跳過")
        continue

    if url in done_urls:
        continue

    print(f"\n[{i+1}/{len(df)}] CIK={cik}  DATE={filedate}")
    print(f"  URL: {url}")

    raw_html = fetch_html(url)
    if not raw_html:
        time.sleep(0.5)
        continue

    passages, hit_count = extract_relevant_passages(raw_html)
    print(f"  Regex 命中 {hit_count} 個區塊，擷取 {len(passages)} 字元")

    res            = classify_with_llm(passages, cik, filedate)
    ai_p2          = res.get('is_Pledge2', -1)
    ai_p3          = res.get('is_Pledge3', -1)
    ai_p2_breadth  = res.get('Pledge2_breadth')
    ai_p2_strength = res.get('Pledge2_strength')
    ai_p3_breadth  = res.get('Pledge3_breadth')
    ai_p3_strength = res.get('Pledge3_strength')
    evidence       = res.get('evidence', '[LLM returned no evidence]')

    print(f"  AI判斷: Pledge2={ai_p2} (breadth={ai_p2_breadth}, strength={ai_p2_strength})")
    print(f"          Pledge3={ai_p3} (breadth={ai_p3_breadth}, strength={ai_p3_strength})")
    print(f"  依據: {str(evidence)[:120]}")

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
        print(f"  已存斷點：{actual_count} 筆")

result_df = pd.DataFrame(results)
result_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n完成！共 {len(result_df)} 筆 → {OUTPUT_FILE}")