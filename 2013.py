import re
import os
import pandas as pd
import requests
import time
from tqdm import tqdm

year = 2006
# ── Patterns ──────────────────────────────────────────────
V5_PATTERN = re.compile(
    r'pledg\w+|short\s+sale\w*|derivative'
    r'|margin\s+account'
    r'|anti.?hedg|anti.?pledg'
    r'|zero.?cost\s+collar|prepaid\s+variable\s+forward'
    r'|monetization\s+transaction'
    r'|\bcollateral\b|\bmargin\s+loans?\b',
    re.IGNORECASE | re.DOTALL
)

CONTEXT_PATTERNS = {
    'hedg_with_context': re.compile(
        r'\bprohibit\w*\s*[\S\s]*?\bhedg'
        r'|\banti.?hedg'
        r'|\bhedg\w+.*\bpolic',
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

# ── 分類 + 回傳命中片段 ────────────────────────────────────
def classify_text_with_match(text: str) -> tuple[str | None, str | None]:
    m = V5_PATTERN.search(text)
    if m:
        start = max(0, m.start() - 50)
        end   = min(len(text), m.end() + 50)
        snippet = text[start:end].replace('\n', ' ').strip()
        return 'keyword_hit', snippet

    for hit_type, pattern in CONTEXT_PATTERNS.items():
        m = pattern.search(text)
        if m:
            start = max(0, m.start() - 50)
            end   = min(len(text), m.end() + 50)
            snippet = text[start:end].replace('\n', ' ').strip()
            return hit_type, snippet

    return None, None

# ── 下載文字 ───────────────────────────────────────────────
HEADERS = {'User-Agent': 'yizhen1426@gmail.com'}

def fetch_text(url: str) -> str | None:
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        r.raise_for_status()
        return r.text
    except Exception:
        return None

# ── 讀取資料 ───────────────────────────────────────────────
df = pd.read_excel('CYZ_DEF14A.xlsx', sheet_name='DEF14A_06_25', index_col=None)
df.columns = ['cik','cname', 'filedate', 'url']

df['filedate'] = pd.to_datetime(df['filedate'], errors='coerce')
df_year = df[df['filedate'].dt.year == year].copy()
print(f"{year}年筆數：{len(df_year)}")

# ── 先試跑前5筆 ────────────────────────────────────────────
df_test = df_year.head(5).copy()
print(f"\n測試筆數：{len(df_test)}")
print("-" * 80)

for i, (_, row) in enumerate(df_test.iterrows()):
    print(f"\n[{i+1}/5] CIK:{row['cik']} | {row['filedate'].date()}")
    print(f"URL: {row['url']}")

    text = fetch_text(row['url'])

    if text is None:
        print("下載失敗")
        continue

    print(f"下載成功，文字長度：{len(text):,} 字元")

    hit_type, snippet = classify_text_with_match(text)

    if hit_type:
        print(f"命中：{hit_type}")
        print(f"片段：...{snippet}...")
    else:
        print("未命中")

    time.sleep(0.1)

# ── 確認5筆沒問題後，跑全部 ───────────────────────────────
input("\n5筆測試完成，按 Enter 繼續跑全部 2013 年資料...")

CHECKPOINT_FILE = f'checkpoint_{year}.xlsx'

if os.path.exists(CHECKPOINT_FILE):
    df_done = pd.read_excel(CHECKPOINT_FILE)
    done_urls = set(df_done['url'].tolist())
    results = df_done.to_dict('records')
    print(f"讀取斷點，已完成 {len(done_urls)} 筆")
else:
    done_urls = set()
    results = []

SAVE_EVERY = 100

for i, (_, row) in enumerate(tqdm(df_year.iterrows(), total=len(df_year))):

    if row['url'] in done_urls:
        continue

    text = fetch_text(row['url'])

    if text is None:
        hit_type, snippet = 'fetch_error', None
    else:
        hit_type, snippet = classify_text_with_match(text)

    if hit_type is not None:
        results.append({
            'cik':          row['cik'],
            'filedate':     row['filedate'],
            'url':          row['url'],
            'hit_type':     hit_type,
            'matched_text': snippet,
        })

    done_urls.add(row['url'])
    time.sleep(0.1)

    if (i + 1) % SAVE_EVERY == 0:
        pd.DataFrame(results).to_excel(CHECKPOINT_FILE, index=False)
        print(f"  已存斷點：{i+1} 筆處理完畢")

# ── 最終輸出 ───────────────────────────────────────────────
df_result = pd.DataFrame(results)
df_result.to_excel(f'result_{year}.xlsx', index=False)
print(f"\n{year}完成！命中筆數：{len(df_result)}")
print(df_result['hit_type'].value_counts())