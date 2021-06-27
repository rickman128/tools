import matplotlib.pyplot as plt
import numpy as np
from matplotlib import rcParams

rcParams['font.family'] = 'sans-serif'
rcParams['font.sans-serif'] = ['Hiragino Maru Gothic Pro', 'Yu Gothic', 'Meirio', 'Takao', 'IPAexGothic', 'IPAPGothic', 'Noto Sans CJK JP']

labels = ['人工心肺装置及び補助循環装置','人工呼吸器','血液浄化装置','除細動装置（AEDを除く）','閉鎖式保育器','診療用高エネルギー放射線発生装置','診療用放射線照射装置'
]
dummy1 = [20, 34, 30, 35, 27, 10, 10]
dummy2 = [25, 32, 34, 20, 25, 11, 11]

x = np.arange(len(labels))
width = 0.35

fig, ax = plt.subplots()
rects1 = ax.bar(x - width/2, dummy1, width, label = 'dummy1')
rects2 = ax.bar(x + width/2, dummy2, width, label = 'dummy2')

ax.set_ylabel('Scores')
ax.set_title('Scores by group and segment')
ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.legend()

ax.bar_label(rects1, padding=3)
ax.bar_label(rects2, padding=3)

fig.tight_layout()

plt.show()