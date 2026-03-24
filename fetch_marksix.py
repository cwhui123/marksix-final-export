import requests
import pandas as pd
import re

# 官方 HKJC Mark Six 結果頁（英文）
url = "https://bet.hkjc.com/en/marksix/results"

headers = {
    "User-Agent": "Mozilla/5.0"
}

html = requests.get(url, headers=headers, timeout=20).text
print("HTML length:", len(html))

# ✅ 用正則直接抓開彩結果（穩定）
pattern = re.compile(
    r'(\d{2}/\d{3}).*?(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})',
    re.S
)

matches = pattern.findall(html)
print("Matches found:", len(matches))

records = []
for m in matches:
    records.append({
        "期數": m[0],
        "N1": int(m[1]),
        "N2": int(m[2]),
        "N3": int(m[3]),
        "N4": int(m[4]),
        "N5": int(m[5]),
        "N6": int(m[6]),
        "特別號": int(m[7]),
    })

print("✅ Records fetched:", len(records))

df = pd.DataFrame(records[::-1])
df.to_excel("data.xlsx", index=False)

print("✅ data.xlsx written")
