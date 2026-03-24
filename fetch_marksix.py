import requests
import pandas as pd
from bs4 import BeautifulSoup

url = "https://lottery.hk/en/mark-six/results/"
html = requests.get(url, timeout=20).text

print("HTML length:", len(html))  # ✅ debug

soup = BeautifulSoup(html, "html.parser")

tables = soup.find_all("table")
print("Tables found:", len(tables))  # ✅ debug

records = []

for table in tables:
    for row in table.find_all("tr"):
        cells = [c.get_text(strip=True) for c in row.find_all("td")]
        if len(cells) >= 9:
            try:
                records.append({
                    "期數": cells[0],
                    "日期": cells[1],
                    "N1": int(cells[2]),
                    "N2": int(cells[3]),
                    "N3": int(cells[4]),
                    "N4": int(cells[5]),
                    "N5": int(cells[6]),
                    "N6": int(cells[7]),
                    "特別號": int(cells[8]),
                })
            except:
                pass

print("✅ Records fetched:", len(records))

df = pd.DataFrame(records[::-1])
df.to_excel("data.xlsx", index=False)

print("✅ data.xlsx written")
