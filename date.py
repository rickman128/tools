import matplotlib.pyplot as plt
import numpy as np
from matplotlib import rcParams
import matplotlib.dates as mdates
import datetime
from dateutil.relativedelta import relativedelta

rcParams['font.family'] = 'sans-serif'
rcParams['font.sans-serif'] = ['Hiragino Maru Gothic Pro', 'Yu Gothic', 'Meirio', 'Takao', 'IPAexGothic', 'IPAPGothic', 'Noto Sans CJK JP']

# 12ヶ月
time_range = 12

dates = [datetime.datetime(2021, 4, 1) + relativedelta(months = i) for i in range(time_range)]
vals = [0, 0, 0, 10, 20, 30, 35, 40, 50, 80, 110, 120]

ax = plt.subplot()
ax.plot(dates, vals)

# x軸の日付ラベル
xfmt = mdates.DateFormatter('%m月')

xloc = mdates.MonthLocator()

ax.xaxis.set_major_locator(xloc)
ax.xaxis.set_major_formatter(xfmt)

ax.grid(True)

plt.show()