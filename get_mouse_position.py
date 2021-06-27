import pyautogui

print('to exit press Ctrl + C')
try:
	while True:
		x,y = pyautogui.position()
		position = 'X:' + str(x).rjust(4) + ' Y:' + str(y).rjust(4)
		print(position, end='')
		print('\b' * len(position),end='',flush=True)
		pyautogui.sleep(1)
except KeyboardInterrupt:
	print('\n終了')