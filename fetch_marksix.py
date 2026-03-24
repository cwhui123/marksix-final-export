import pandas as pd

# 公開、已處理的 Mark Six 資料集（自動更新）
url = "https://raw.githubusercontent.com/icelam/mark-six-data-visualization/master/data/marksix.csv"

print("Downloading CSV...")

df = pd.read_csv(url)

print("Rows fetched:", len(df))

df = df.rename(columns={
    "draw": "期數",
    "no1": "N1",
    "no2": "N2",
    "no3": "N3",
    "no4": "N4",
    "no5": "N5",
    "no6": "N6",
    "extra": "特別號"
})

df = df[["期數", "N1", "N2", "N3", "N4", "N5", "N6", "特別號"]]

df.to_excel("data.xlsx", index=False)

print("✅ data.xlsx written")
