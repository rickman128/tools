'''
*****************************************************************************
Process   指定フォルダ内の画像を全てリサイズする
Date      2021/06/11  K.Endo
Memo      横幅サイズを指定すると、アスペクト比を維持してリサイズする

*****************************************************************************
'''
import os
from PIL import Image
import PySimpleGUI as sg

# アスペクト比が変わるリサイズ
def resize(img, size):
	return img.resize((size, size))

def scale_to_width(img, width):  # アスペクト比を固定して、幅が指定した値になるようリサイズする。
    height = round(img.height * width / img.width)
    return img.resize((width, height))

def main(path, size):
	dir = path
	dist = path + '/out'
	# outフォルダを作っておく
	os.makedirs(os.path.abspath(dist), exist_ok = True)
	files = os.listdir(dir)
	for file in files:
		try:
			img = Image.open(os.path.join(dir, file))
			# アスペクト比が変わるリサイズ
		#	new_img = resize(img, size)
			# widthだけ指定してアスペクト比が変わらないリサイズ
			new_img = scale_to_width(img, int(size))
			new_img.save(os.path.join(dist, file))
		except OSError as e:
			pass

	sg.popup('処理結果', 'リサイズが完了しました。\n 出力先：' + dist)

# テーマ
sg.theme('DarkPurple1')
cmbValues = ['大', '中', '小', 'カスタム']

# GUIレイアウト
layout = [[sg.Text('指定したフォルダ内の画像ファイルを指定サイズに変換します。\nフォルダを選択してください。')],
			[sg.Text('フォルダ', size=(10, 1)), sg.Input(), sg.FolderBrowse('フォルダを選択', key = 'inputFolderPath')],
			[sg.Text('サイズ　', size=(10, 1)), sg.Combo(cmbValues, default_value = '選択してください', enable_events=True, size=(16, 1), key = 'cmbSize')],
			[sg.Text('横幅サイズ', size=(10, 1)), sg.Input(key = 'size', size=(10, 1)), sg.Text('ピクセル')],
			[sg.Button('実行', key = 'exec'), sg.Button('閉じる', key = 'close')]]

# ウィンドウの生成
window = sg.Window('画像リサイズ', layout)

# イベント
while True:
	event, values = window.read()
	if event == sg.WIN_CLOSED or event == 'close':
		break
	elif event == 'inputFolderPath':
		print('フォルダ選択ボタンクリック')
	elif event == 'cmbSize':
		if values['cmbSize'] == '大':
			window['size'].update(value='800', disabled=True)
		elif values['cmbSize'] == '中':
			window['size'].update(value='500', disabled=True)
		elif values['cmbSize'] == '小':
			window['size'].update(value='300', disabled=True)
		else:
			window['size'].update(disabled=False)
	elif event == 'exec':
		if values['inputFolderPath'] != '':
			main(values['inputFolderPath'], values['size'])
			window.close()
window.close()