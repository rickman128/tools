'''
*****************************************************************************
Process   PDFを分割する
Date      2021/04/15  K.Endo
Memo      PDFファイルを指定したページごとに分割する

*****************************************************************************
'''
import datetime
import PyPDF2 as pd
import os
import PySimpleGUI as sg
import glob

def SplitPDF(file):
	reader = pd.PdfFileReader(file)
	outDir = os.path.dirname(file)

	num_pages = reader.getNumPages()  			# ページ数の取得
	digits = len(str(num_pages))  				# ページ数の桁数の取得
	fpad = '{0:0' + str(digits) + 'd}'  		# format用文字列作成

	# 出力ファイル名 ex) yyyymmdd_1.pdf
	today = datetime.date.today()
	# 指定したPDFがある場所にoutフォルダを作っておく
	try:
		os.makedirs(outDir + '/out', exist_ok = True)
	except FileExistsError as e:
		print('ディレクトリoutは既に存在します。')

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

	# 作成したファイルのリスト
	files = glob.glob(outDir + './out/*.pdf')
	# ファイル名の最後に定義した名前を付加する
	lst_name = ['吉川技士長','笠川CE','松岡CE','片岡CE','山﨑CE','堀田CE','中村CE','関戸CE','橋脇CE','川道CE','髙村CE','吉田CE','梅田CE','松中CE']

	if len(files) == len(lst_name):
		for (file, name) in zip(files, lst_name):
			after = file.replace('.pdf', '_' + name + '.pdf')
			before = file
			os.rename(before, after)
	else:
		sg.popup('分割したファイルが14人分ではないため、リネームできません。')

	sg.popup('処理結果', 'PDFファイルを分割しました。\n 出力先：' + outDir + '/out/')


# テーマ
sg.theme('DarkPurple1')

# GUIレイアウト
layout = [[sg.Text('指定したPDFファイルを2ページずつに分割します。\n分割するファイルを選択してください。')],
			[sg.Text('ファイル'), sg.Input(), sg.FileBrowse('ファイルを選択', key = 'inputFilePath')],
			[sg.Button('実行', key = 'exec'), sg.Button('閉じる', key = 'close')]]

# ウィンドウの生成
window = sg.Window('PDF分割', layout)


# イベント
while True:
	event, values = window.read()
	if event == sg.WIN_CLOSED or event == 'close':
		break
	elif event == 'inputFilePath':
		print('ファイル選択ボタンクリック')
	elif event == 'exec':
		SplitPDF(values[0])
		window.close()
window.close()