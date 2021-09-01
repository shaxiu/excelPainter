import win32com.client
from win32com.client import DispatchEx
from ctypes.wintypes import RGB
import cv2
import numpy as np

# color_RGB用于存储RGB颜色
color_total = []
# img_file为你要画的图片的路径，路径中不应含有中文
img_file = "D:/painter.jpg"
# 读取图片文件
img_a = cv2.imread(img_file)
# cv2默认为BGR顺序，将顺序转为RGB
img_color = cv2.cvtColor(img_a, cv2.COLOR_BGR2RGB)
# 返回height，width，以及通道数，不用所以省略掉
h, l, _ = img_a.shape
# 打印图片总行数和列数，即竖向有多少像素，横向有多少像素
print('行数%d，列数%d' % (h, l))
# 将颜色数据添加到color_total中，颜色数据方面采集完成
for i in img_color:
    color_total.append(i)

# Win32#打开EXCEL
excel = win32com.client.DispatchEx('Excel.Application')
# 要处理的excel文件路径
WinBook = excel.Workbooks.Open('D:/painter.xlsx')
# 要处理的excel页
WinSheet = WinBook.Worksheets('Sheet1')
# 设置单元格颜色
# excel中[1,1]代表的是第一行第一列的单元格，而数组中[0][0]代表的是第一行一列
# 其中color_total[x-1][y-1][0]对应的是第x行第y列图像R的值 color_total[x-1][y-1][1]代表G color_total[x-1][y-1][2]代表B
for x in range(1, h):
    for y in range(1, l):
        WinSheet.Cells(x, y).Interior.Color = RGB(
            color_total[x-1][y-1][0], color_total[x-1][y-1][1], color_total[x-1][y-1][2])
        # 打印正在进行描绘的像素的位置
        print(x, y)
# 保存
WinBook.save
# 关闭
WinBook.close
