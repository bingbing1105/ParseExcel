
from cmath import isnan, nan, tan
import math
import pandas as pd
import numpy as np
import numpy as np
import matplotlib.pyplot as plt

df=pd.read_excel("ccc.xlsx","sheet1")
plt.plot(df['u-2σ'])
plt.plot(df['u+2σ'])
# 设置画布大小
plt.figure(figsize=(30, 9))
# 画第1条折线，参数看名字就懂，还可以自定义数据点样式等等。
# plt.plot(x, y1, color='#FF0000', label='u-2σ', linewidth=3.0)
# 画第2条折线
# plt.plot(x, y2, color='#00FF00', label='u+2σ', linewidth=3.0)

# 给第1条折线数据点加上数值，前两个参数是坐标，第三个是数值，ha和va分别是水平和垂直位置（数据点相对数值）。
""" for a, b in zip(x, y1):
    plt.text(a, b, '%0.2f'%b, ha='center', va= 'bottom', fontsize=0)
# 给第2条折线数据点加上数值
for a, b in zip(x, y2):
    plt.text(a, b, '%0.2f'%b, ha='center', va= 'bottom', fontsize=0)
#绘制网格线
plt.grid(color='b',linestyle='dotted',linewidth=0.5)
# 画水平横线，参数分别表示在y=3，x=0~len(x)-1处画直线。
# plt.hlines(3, 0, len(x)-1, colors = "#000000", linestyles = "dashed")
 
# 添加x轴和y轴刻度标签
plt.xticks([r for r in x], x_ticks, fontsize=10, rotation=20)
plt.yticks(fontsize=18) """
 
# 添加x轴和y轴标签
plt.xlabel(u'Distance', fontsize=25)
plt.ylabel(u'RMS', fontsize=25)
 
# 标题
plt.title(u'TL460-S1-E1', fontsize=25)
 
# 图例
plt.legend(fontsize=18)
 
# 保存图片
plt.savefig('./figure.pdf', bbox_inches='tight')
# 显示图片
plt.show()
