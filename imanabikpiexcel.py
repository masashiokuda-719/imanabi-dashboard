import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

# CSVデータの読み込み
df = pd.read_csv("imanabitestgraph.csv", encoding="utf-8")

# AB列データ（年, 新規）の抽出とクリーニング
ab = df.iloc[:, [0, 1]].dropna()
ab.columns = ["年", "新規"]

# DE列データ（X, 売上）の抽出とクリーニング
de = df.iloc[:, [3, 4]].dropna()
de.columns = ["売上X", "売上"]

# 棒グラフ（AB列）
plt.figure(figsize=(4, 3))
plt.bar(ab["年"].astype(str), ab["新規"].astype(int), color="skyblue")
plt.xlabel("年")
plt.ylabel("新規")
plt.title("年ごとの新規")
plt.tight_layout()
plt.savefig("bar_ab.png")
plt.close()

# 折れ線グラフ（DE列）
plt.figure(figsize=(4, 3))
plt.plot(de["売上X"], de["売上"], marker="o", color="orange")
plt.xlabel("X")
plt.ylabel("売上")
plt.title("売上推移")
plt.tight_layout()
plt.savefig("line_de.png")
plt.close()

# Excelファイル作成
wb = Workbook()
ws = wb.active
ws.title = "グラフ付きシート"

# 画像の貼り付け
img1 = XLImage("bar_ab.png")
img2 = XLImage("line_de.png")

# 折れ線グラフを左側に（A1セル）
ws.add_image(img2, "A1")
# 棒グラフを右側に（J1セルあたりに）
ws.add_image(img1, "J1")

wb.save("output_graphs.xlsx")
print("Excelファイル（output_graphs.xlsx）を作成しました。")