from time import sleep
from pywinauto import Desktop, Application

app = Application(backend="uia")
app.start("calc.exe")

dlg = Desktop(backend="uia")["電卓"]

dlg['1'].click()
sleep(1)
dlg['プラス'].click()
sleep(1)
dlg['2'].click()
sleep(1)
dlg['等号'].click()
sleep(1)

dlg.close()