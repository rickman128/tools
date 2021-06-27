'''
*****************************************************************************
Process   動画ファイルの先頭・末尾をカットする
Date      2021/06/24  K.Endo
Memo      フォルダ指定
		  指定秒数は切り取りたい秒数。
		  ex) 10、20を指定→動画の最初の10秒、最後の20秒をカット
*****************************************************************************
'''
import os
import PySimpleGUI as sg
import glob
from moviepy.editor import *

# コーデック対応辞書
codecs = {
	'mpg'	:	'mpeg1video',
	'mpeg'	:	'mpeg1video',
	'ogv'	:	'libtheora',
	'webm'	:	'libvpx',
	'ogg'	:	'libvorbis',
	'mp3'	:	'pcm_s16le',
	'wav'	:	'libvorbis',
	'm4a'	:	'libfdk_aac'
}

# 動画のカット
def cut_movie(path, sec1, sec2):
	print('------------cut_movie----------------')
	cd = ''
	try:
		# 元のファイルがある場所にout/ファイル名として保存する
		dir = os.path.dirname(path)
		save_path = dir + '/out/' + os.path.basename(path)
		# カット位置（sec2は終了位置ではなく、切り取り秒数）
		# endをマイナスで指定すると、指定した秒数分切り取る
		start = sec1
		end = int(sec2) * -1
		#ビデオのカット開始
		video = VideoFileClip(path).subclip(start, end)

		root, ext = os.path.splitext(path)
		ext = ext.replace('.', '')
		try:
			cd = codecs[ext]
		except KeyError as e:
			print('KeyError: ' + ext)

		# 書き込み
		if cd != '':
			print('codec '+ cd + ': ' + path)
			print('codec '+ cd + ': ' + save_path)
			video.write_videofile(save_path, fps = 29, codec = cd)
		else:
			# デフォルトコーデックを使用
			print('default codec: ' + path)
			print('default codec: ' + save_path)
			video.write_videofile(save_path, fps = 29)
	except OSError as e:
		pass
#	except ValueError as e:
#		print(e)


# メイン
def main(path, sec1, sec2):
	dir = path
	dist = path + '/out'
	# outフォルダを作っておく
	os.makedirs(os.path.abspath(dist), exist_ok = True)
	# 対象ファイル取得
	files = os.listdir(dir)
	for file in files:
		# 動画をカット
		filepath = os.path.join(dir, file)
		filepath = filepath.replace(os.sep,'/')
		if os.path.isfile(filepath):
			cut_movie(os.path.join(dir, file), sec1, sec2)

	sg.popup('処理結果', '動画編集が完了しました。\n 出力先：' + dist)

# テーマ
sg.theme('DarkPurple1')

# GUIレイアウト
layout = [[sg.Text('指定したフォルダ内の動画から先頭・末尾をカットします。\nフォルダを選択してください。')],
			[sg.Text('フォルダ', size=(10, 1)), sg.Input(), sg.FolderBrowse('フォルダを選択', key = 'inputFolderPath')],
			[sg.Text('先頭', size=(10, 1)), sg.Input(key = 'sec1', size=(10, 1)), sg.Text('秒')],
			[sg.Text('末尾', size=(10, 1)), sg.Input(key = 'sec2', size=(10, 1)), sg.Text('秒')],
			[sg.Button('実行', key = 'exec'), sg.Button('閉じる', key = 'close')]]

# ウィンドウの生成
window = sg.Window('動画カット', layout)


# イベント
while True:
	event, values = window.read()
	if event == sg.WIN_CLOSED or event == 'close':
		break
	elif event == 'inputFolderPath':
		print('フォルダ選択ボタンクリック')
	elif event == 'exec':
		# 本当はここでsec1 < sec2かどうかチェックした方が良い
		main(values['inputFolderPath'], values['sec1'], values['sec2'])
		window.close()
window.close()