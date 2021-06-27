'''
*****************************************************************************
Process   バリテスから受講者CSVをエクスポートする
Date      2021/03/02  K.Endo
Memo      定義した配列の研修について、受講者CSVをエクスポートする

*****************************************************************************
'''
import sys
import time
from time import strftime
import pyautogui

'''
*****************************************************************************
	構造体用クラス
*****************************************************************************
'''
class picture:
	def __init__(self, strFileName, strMsg, conf = 0.8):
		self.strFileName = strFileName
		self.strMsg = strMsg
		self.conf = conf


'''
*****************************************************************************
 Name      clickPic
 Param     strPicFileName	: 対象画像ファイル名
           strMsg       	: 対象の名称
           confidence       : 閾値(default 0.8)
		   xOff				: x軸オフセット
		   yOff				: y軸オフセット
 Result    Nothing
 Memo      指定した画像ファイルの位置をクリックする
*****************************************************************************
'''
def clickPic(strPicFileName, strMsg, conf = 0.8, xOff = 0, yOff = 0):
	strPicFileName = 'C:\\Python\\export_baritess\\img\\' + strPicFileName
	ret = pyautogui.locateCenterOnScreen(strPicFileName, confidence = conf)
	if ret==None:
		print(strMsg + 'が見つからない')
		sys.exit()
	else:
		x, y = ret
		x = x + xOff
		y = y + yOff
		print('【' + strMsg + '】' + ' x: ' + str(x) + ' y: ' + str(y))
	pyautogui.click(x, y)
	pyautogui.sleep(1)


'''
*****************************************************************************
 Name      exportCSV
 Param     strPicFileName	: 対象画像ファイル名
           strMsg       	: 対象の名称
           confidence       : 閾値(default 0.8)
 Result    True or False
 Memo      指定した研修ファイルの位置をクリックする
*****************************************************************************
'''
def exportCSV(strPicFileName, strMsg, conf = 0.8):
	# 研修の受講者の位置をクリック
	clickPic(strPicFileName, strMsg, conf = conf, xOff = OFFSET_RIGHT)

	# アンケート
	clickPic('5_1_survey.png', 'アンケート', conf = 0.8)

	pyautogui.moveRel(0, OFFSET_DOWN_1REC)	# 回答済み
	pyautogui.click()
	pyautogui.sleep(1)

	# テスト
	clickPic('5_2_test.png', 'テスト', conf = 0.8)

	pyautogui.moveRel(0, OFFSET_DOWN_1REC)	# 合格
	pyautogui.click()
	pyautogui.sleep(1)

	# 表示件数
	clickPic('5_3_count.png', '表示件数', conf = 0.7)

	pyautogui.moveRel(0, OFFSET_DOWN_ALLREC)	# すべて
	pyautogui.click()
	pyautogui.sleep(1)

	# 検索
	clickPic('5_4_search.png', '検索ボタン', conf = 0.8)

	# CSVエクスポート
	clickPic('5_5_csvexport.png', 'CSVエクスポートボタン', conf = 0.8)

	# ログインID
	#clickPic('6_1_loginid.png', 'ログインID', conf = 0.8)

	# 確定
	#clickPic('6_2_submit.png', '確定ボタン', conf = 0.9)

	# 名前をつけて保存
	clickPic('7_1_save.png', '保存ボタン', conf = 0.9, xOff = 43)	# 保存の右端の三角をクリック

	pyautogui.moveRel(0, 46)	# 名前をつけて保存の位置に移動
	pyautogui.click()
	pyautogui.sleep(2)

	# pyautogui.press('Home')	# homeキーが押せない・・・
	pyautogui.hotkey('ctrl', 'c')
	pyautogui.write(strMsg + '_')
	pyautogui.hotkey('ctrl', 'v')
	pyautogui.sleep(1)

	# 保存
	clickPic('8_1_save.png', '保存ボタン', conf = 0.8)

	# 閉じる
	clickPic('9_1_close.png', '閉じるボタン', conf = 0.8)

	pyautogui.sleep(2)

	return True

'''
*****************************************************************************
 Name      clickKensyuPics
 Param     strPicFileName	: 対象画像ファイル名
           strMsg       	: 対象の名称
           confidence       : 閾値(default 0.8)
 Result    True or False
 Memo      指定した画像ファイルの位置をクリックする
*****************************************************************************
'''
def clickKensyuPics(tree, kensyu):
	clickPic(tree.strFileName, tree.strMsg, tree.conf)	# ツリーの項目クリック
	for item in kensyu:
		exportCSV(item.strFileName, item.strMsg, item.conf)

'''
*****************************************************************************
	MAIN
*****************************************************************************
'''
try:
	time_start = time.perf_counter()
	print(strftime('%Y/%m/%d %H:%M:%S', time.localtime()))

	OFFSET_RIGHT = 350
	OFFSET_DOWN_1REC = 15
	OFFSET_DOWN_ALLREC = 115

	PIC_TREE = [
		picture('2_sinpai.png', '人工心肺リンク', 0.7), 
		picture('3_kokyuki.png', '人工呼吸器リンク', 0.7),
		picture('4_josaido.png', '除細動装置リンク', 0.7)]

	PIC_KENSYU = [
		[picture('2_1_ecmo.png', 'sinpai_ecmo', 0.8), picture('2_2_iabp.png', 'sinpai_IABP', 0.7)],
		[picture('3_1_savina.png', 'kokyu_Savina', 0.8), picture('3_2_chinsei.png', 'kokyu_chinsei', 0.8), picture('3_3_ridatu.png', 'kokyu_ridatu', 0.8), picture('3_4_nagare.png', 'kokyu_nagare', 0.8)],
		[picture('4_1_ns.png', 'josaido_ns', 0.7), picture('4_2_conv.png', 'josaido_conv', 0.7), picture('4_3_dr.png', 'josaido_dr', 0.7)]]

	x, y = 100, 0
	# 研修・講義管理ボタン
	clickPic('1_kensyukanri.png', '研修・講義管理ボタン', 0.8)
	# ツリーをクリックして研修を順番にexportする
	for (tree, kensyu) in zip(PIC_TREE, PIC_KENSYU):
		clickKensyuPics(tree, kensyu)

	now = time.ctime()
	print(strftime('%Y/%m/%d %H:%M:%S', time.localtime()))

	time_end = time.perf_counter()
	time_dif = time_end - time_start
	print('所要時間: ' + str(time_dif) + '秒')

except KeyboardInterrupt:
	print('\n終了')