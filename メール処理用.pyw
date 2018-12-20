from pathlib import Path
from datetime import datetime
import sys,csv
from enum import Enum
 
class ErrorStatus(Enum):
	fileFailed = 1 #ファイルが無い時
	fileEncrypted = 2 #ファイルが暗号化されて読めない時
 
def failed(a):
	date = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
	#desktop = str(Path.home()) + "\Desktop"
	failureString = ""
	errorStatusDictionary = {ErrorStatus.fileFailed:"CSVファイルの取得に失敗しました。",ErrorStatus.fileEncrypted:"暗号化されて読めない可能性があります。"}
	failureString = date + "\n" + errorStatusDictionary[a]

	inquiryFile = open(desktop + r"\問い合わせを出力.txt","w")
	inquiryFile.writelines(failureString)
	inquiryFile.close()
	answerFile = open(desktop + r"\回答を出力.txt","w")
	answerFile.writelines(failureString)
	answerFile.close()
	sys.exit()
 
downloads = Path(str(Path.home()) + r"\Downloads")
desktop = str(Path.home()) + r"\Desktop"
csvList = list(downloads.glob("*.csv"))
 
# csvファイル名を取得できなかったら終了
if csvList == []:
	failed(ErrorStatus.fileFailed)
 
csvFile = open(csvList.pop(),'r')
inquiryList = []
answerList = []
 
 
try:
 	for row in csv.reader(csvFile):
		inquiryList.append(row[10]) #CSVの行を指定
		answerList.append(row[11]) #CSVの行を指定
except UnicodeDecodeError:
 	failed(ErrorStatus.fileEncrypted)
 
# 先頭行を削除しておく
del inquiryList[0]
del answerList[0]
 
inquiryText = ""
for inquiry in reversed(inquiryList):
 	if inquiry == "":
		continue
	inquiryText += inquiry
	inquiryText += "\n\n"
 
#書き込むときは暗号化されてても問題ないっぽい
inquiryFile = open("問い合わせを出力.txt","w")
inquiryFile.writelines(inquiryText)
inquiryFile.close()
 
answerText = ""
for answer in reversed(answerList):
	if answer == "":
		continue
	answerText += answer
	answerText += "\n\n"
answerFile = open("回答を出力.txt","w")
answerFile.writelines(answerText)
answerFile.close()