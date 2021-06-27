'''
*****************************************************************************
Process   PDFを分割する
Date      2021/04/15  K.Endo
Memo      PDFファイルを指定したページごとに分割する

*****************************************************************************
'''
import datetime
import PyPDF2 as pd
#import pathlib
import os
import tkinter, tkinter.filedialog, tkinter.messagebox

# ファイル選択ダイアログの表示
root = tkinter.Tk()
root.withdraw()
fTyp = [("", "*.pdf")]
iDir = os.path.abspath(os.path.dirname(__file__))
tkinter.messagebox.showinfo('PDF分割','分割したいファイルを選択してください。')
file = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)

# 処理ファイル名の出力
#tkinter.messagebox.showinfo('PDF分割',file)

reader = pd.PdfFileReader(file)
outDir = os.path.dirname(file)

num_pages = reader.getNumPages()  			# ページ数の取得
digits = len(str(num_pages))  				# ページ数の桁数の取得
fpad = '{0:0' + str(digits) + 'd}'  		# format用文字列作成

# 出力ファイル名 ex) yyyymmdd_1.pdf
today = datetime.date.today()
# 指定したPDFがある場所にoutフォルダを作っておく
os.makedirs(outDir + '/out', exist_ok = True)
											# out/yyyymmdd_
strFilePref = outDir + '/out/' + today.strftime('%Y%m%d') + '_'

writer = None
bPrint = False
iName = 1
for i in range(num_pages):
	page = reader.getPage(i)  				# ページを取得

	# 奇数ページ
	if (i + 1) % 2 == 1:
		writer = pd.PdfFileWriter()  		# 空のwriterオブジェクト作成
		bPrint = False

	writer.addPage(page)  					# writerオブジェクトにページを追加

	# 偶数ページ
	if (i + 1) % 2 == 0:
		fname = strFilePref + fpad.format(iName) + '.pdf'
		with open(fname, mode='wb') as f:
			writer.write(f)  				# 出力
		bPrint = True
		iName += 1

# 奇数ページで終わった場合、出力されていない１ページを出力
if bPrint == False:
	fname = strFilePref + fpad.format(iName) + '.pdf'
	with open(fname, mode='wb') as f:
		writer.write(f)  					# 出力
	