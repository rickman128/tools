'''
*****************************************************************************
Process   バリテスの研修参加者エクスポート
Date      2021/05/24  K.Endo
Memo      指定の年度の医療機器定期研修受講者CSVをエクスポートする


*****************************************************************************
'''
import datetime
import time
import os
import glob
import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager 

'''
*****************************************************************************
	Const
*****************************************************************************
'''
# ツリーのアイテム定義
TREE_LIST = ['【1 人工心肺装置及び補助循環装置】', '【2 人工呼吸器】', '【4 除細動装置（AEDを除く）】']

# 研修一覧のカラム定義
CLM_TITLE = 1
CLM_YEAR = 2
CLM_JUKOUSYA = 4

'''
*****************************************************************************
 Name      export_kensyu
 Param     title	: 研修タイトル
 Result    Nothing
 Memo      指定した研修の参加者CSVをエクスポートする
		　 この関数が呼ばれたときには別窓で１つの研修が表示されている
*****************************************************************************
'''
def export_kensyu(driver, title):
	try:
		old_handle = driver.current_window_handle
		handles = driver.window_handles
		# 最後に開かれたウインドウに切り替える
		driver.switch_to.window(handles[-1])

		# テスト=合格
		test = driver.find_element_by_id('test_status')
		Select(test).select_by_visible_text('合格')

		# アンケート=回答済み
		quest = driver.find_element_by_id('questionnaire_status')
		Select(quest).select_by_visible_text('回答済')

		# 表示件数=全件
		count = driver.find_element_by_name('per_page_size')
		Select(count).select_by_visible_text('全件')

		# 検索
		search = driver.find_element_by_id('search_student')
		search.click()
		time.sleep(1)

		# 受講者CSVエクスポート
		export = driver.find_element_by_xpath('/html/body/form[2]/div[1]/nobr/button[3]')
		export.click()
		time.sleep(2)

		# ダウンロードフォルダにlist_yyyymmdd.csvで保存されたので、名前を変える
		# DLフォルダ内のcsv取得
		dlfolder = 'C:\\Users\\ce4_user\\Downloads\\*.csv'
		files = glob.glob(dlfolder)
		# 最新ファイル
		latest_file = max(files, key = os.path.getctime)
		beforename = os.path.basename(latest_file)			# xxxx.csv
		now = datetime.datetime.now()
		after = os.path.splitext(latest_file)[0] + now.strftime('%H%M%S') + '_' + title + '.csv'
		before = latest_file
		os.rename(before, after)

	finally:
		driver.close()
		time.sleep(1)
		# 1つ前のウインドウに切り替える
		driver.switch_to.window(old_handle)

'''
*****************************************************************************
 Name      main
 Param     Nothing
 Result    Nothing
 Memo      メイン処理
*****************************************************************************
'''
def main(input_url, input_year):
	try:
		driver=webdriver.Chrome('C:\\webdriver\\chromedriver.exe')
		driver.get(input_url)
		time.sleep(1)

		# バリテスのアイコン（オレンジ色の花）
		baritess = driver.find_element_by_xpath('/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr/td[4]/a')
		baritess.click()

		# フレーム切り替え
		frame = driver.find_element_by_id('floatingpage')
		driver.switch_to.frame(frame)

		# 研修・講義管理ボタンクリック
		tab = driver.find_element_by_xpath('//*[@id="tblMainLink"]/tbody/tr/td[3]/a/nobr')
		tab.click()

		# ツリーの人工心肺～保育器ループ
		for item in TREE_LIST:
			node = driver.find_element_by_link_text(item)
			node.click()
			time.sleep(1)
			# table取得
			tbl = driver.find_element_by_class_name('padding_vertical0')

			# 研修タイトルじゃなくて「受講者」列のリンクをクリックするので、リンクテキストで特定できない
			row = 2
			while True:
				td_title = tbl.find_element_by_xpath('//*[@id="contents_wrapper"]/table/tbody/tr[' + str(row) + ']/td[' + str(CLM_TITLE) + ']')
				td_year = tbl.find_element_by_xpath('//*[@id="contents_wrapper"]/table/tbody/tr[' + str(row) + ']/td[' + str(CLM_YEAR) + ']')
				td_jukousya = tbl.find_element_by_xpath('//*[@id="contents_wrapper"]/table/tbody/tr[' + str(row) + ']/td[' + str(CLM_JUKOUSYA) + ']')
				# 当年度の研修だけ処理する（登録順に並んでいる）
				if int(td_year.text) == input_year:
					td_jukousya.click()
					time.sleep(2)
					# 研修1件export
					export_kensyu(driver, td_title.text)
				else:
					break

				row += 1
				# フレーム切り替え
				frame = driver.find_element_by_id('floatingpage')
				driver.switch_to.frame(frame)

	except Exception as e:
		sg.popup(e)
		print(e)
	finally:
		driver.quit()

'''
*****************************************************************************
	pysimplegui実行
*****************************************************************************
'''
sg.theme('HotDogStand')

i = 2018
today = datetime.date.today()
yearlist = []
while True:
	yearlist.append(i)
	if i == today.year:
		break
	i += 1

layout = [[sg.Text('URL'), sg.InputText(key = 'url')],
			[sg.Text('出力年度'), sg.Combo((yearlist), key = 'year')],
			[sg.Button('実行'), sg.Button('閉じる')]]

window = sg.Window('バリテスエクスポート', layout)

while True: 
	event, values = window.read() 
	if event == sg.WIN_CLOSED or event == '閉じる': 
		break 
	elif event == '実行': 
		url = values['url']
		year = values['year']
		main(url, year)
		sg.popup('処理が完了しました。')
window.close()

