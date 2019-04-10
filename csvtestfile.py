from pathlib import Path
import csv

# CSVのパスを取得
downloads = Path(str(Path.home()) + r"\Downloads")
csvList = list(downloads.glob("*.csv"))

# CSVファイルを開く
csvFile = open(csvList.pop(),'r')

# 列を取得
aList = []
for row in csv.reader(csvFile):
    aList.append(row[1]) #CSVの行を指定

# 先頭行を削除しておく
del aList[0]

# テキストに書き込むテキストを作成
aText = ""
for a in reversed(aList):
	if a == "":
		continue
	aText += a
	aText += "\n\n"

#書き込む
desktop = str(Path.home()) + "\Desktop"
aFile = open(desktop + "\Aを出力.txt","w")
aFile.writelines(aText)
aFile.close()
