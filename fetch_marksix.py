import pandas as pd

# 多個備用資料來源（任何一個成功就用）
DATA_URLS = [
    # 來源 A（如將來恢復）
    "https://raw.githubusercontent.com/icelam/mark-six-data-visualization/master/data/marksix.csv",

    # 來源 B（你之後可以自行再加）
    # "https://raw.githubusercontent.com/xxx/xxx/master/marksix.csv",
]

df = None

for url in DATA_URLS:
    try:
        print("Trying:", url)
        df = pd.read_csv(url)
        if len(df) > 0:
            print("✅ Data loaded from:", url)
            break
    except Exception as e:
        print("❌ Failed:", url, e)

if df is None:
    raise RuntimeError("No valid data source available")

# 標準化欄位（如來源略有不同，你可以再微調）
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

print("✅ Records fetched:", len(df))

df.to_excel("data.xlsx", index=False)
print("✅ data.xlsx written")
