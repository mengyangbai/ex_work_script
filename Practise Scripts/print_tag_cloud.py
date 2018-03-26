# -*- coding: utf-8 -*-
import codecs
import random
from pytagcloud import create_tag_image, create_html_data, make_tags, LAYOUT_HORIZONTAL, LAYOUTS
from pytagcloud.colors import COLOR_SCHEMES
from pytagcloud.lang.counter import get_tag_counts
# from pylab import mpl
# mpl.rcParams['font.sans-serif'] = ['SimHei']#['FangSong'] # 指定默认字体
# mpl.rcParams['axes.unicode_minus'] = False # 解决保存图像是负号'-'显示为方块的问题
wd = []

fp=codecs.open("weight.txt", "r",'utf-8');

alllines=fp.readlines();

fp.close();

for eachline in alllines:
    line = eachline.split('\t')
    #print eachline,
    tuple= (line[0],float(line[1]))
    wd.append(tuple)

print(wd)


from operator import itemgetter
#swd = sorted(wd.items(), key=itemgetter(1), reverse=True)
tags = make_tags(wd,minsize = 50, maxsize = 240)
create_tag_image(tags, 'tag_cloud.png', background=(0, 0, 0, 255),
size=(2400, 1000),layout=LAYOUT_HORIZONTAL,
fontname="SimHei")