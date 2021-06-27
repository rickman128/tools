import os
from PIL import Image

# アスペクト比が変わるリサイズ
def resize(img, size):
	return img.resize((size, size))

def scale_to_width(img, width):  # アスペクト比を固定して、幅が指定した値になるようリサイズする。
    height = round(img.height * width / img.width)
    return img.resize((width, height))

dir = '.\\img\\'
dist = '\\out'

# outフォルダを作っておく
os.makedirs(os.path.abspath(os.path.dirname(__file__)) + dist, exist_ok = True)

files = os.listdir(dir)
for file in files:
	try:
		img = Image.open(os.path.join(dir, file))
		# アスペクト比が変わるリサイズ
	#	new_img = resize(img, 300)
		# widthだけ指定してアスペクト比が変わらないリサイズ
		new_img = scale_to_width(img, 300)
		new_img.save(os.path.join('.\\out', file))
	except OSError as e:
		pass
