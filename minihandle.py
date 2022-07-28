from tkinter import *

from PIL import Image
import re

import xlrd
from PIL import ImageGrab, ImageTk
from aip import AipOcr

config = {
    'appId': '25963617',
    'apiKey': 'FkImi13Uj6Ry0irvgcAhoxq8',
    'secretKey': 'nrowIKYG7RgKHw2N0V06fURyFtFWNvXo'
}

client = AipOcr(**config)


def get_file_content(file):
    with open(file, 'rb') as fp:
        return fp.read()


def img_to_str(image_path):
    image = get_file_content(image_path)
    result = client.basicGeneral(image)
    if 'words_result' in result:
        return '\n'.join([w['words'] for w in result['words_result']])


def btn_click():
    with open("log.txt", 'r+', encoding='utf-8') as f_log:
        log_ar = f_log.read().splitlines()
    allin = re.split(r'[,，]', log_ar[0])
    im = ImageGrab.grab(bbox=(int(allin[0]), int(allin[1]), int(allin[2]), int(allin[3])))  # 左上角和右下角坐标
    im.save("pic.png")
    imagepath = 'pic.png'
    res = img_to_str(imagepath)
    res = re.findall(r'(\w+)', res)
    res = ''.join(res)
    res = res[0:4]
    cleartext()
    FileContaceList = 'tiku.xls'
    FileName = FileContaceList
    # open file
    FileObj = xlrd.open_workbook(FileName)
    # 获取第一个工作表
    sheet = FileObj.sheets()[0]
    # 行数
    row_count = sheet.nrows
    # 列数
    col_count = sheet.ncols
    ct = 0
    # 搜索关键字符串
    for element in range(row_count):
        if str(res) in str(sheet.row_values(element)[0]):
            text.insert(INSERT, sheet.row_values(element)[0] + "\n")
            print(sheet.row_values(element)[0])
            print("答案：%s（请仔细核对选项，考试可能会乱序）" % sheet.row_values(element)[1])
            text.insert(INSERT, "答案：" + str(
                sheet.row_values(element)[1]) + "\n" + "\n")
            print("题库中第%s题" % sheet.row_values(element)[2])
            #  text.insert(INSERT, str(sheet.row_values(element)[2]) + "\n")
            ct = ct + 0
        else:
            ct = ct + 1
        if ct == row_count:
            print("搜索失败，请重新输入关键字")
            text.insert(INSERT, "搜索失败，请重新输入关键字！\n")
    print("搜索结束")


def btn_find():
    cleartext()
    FileContaceList = 'tiku.xls'
    FileName = FileContaceList
    # open file
    FileObj = xlrd.open_workbook(FileName)
    # 获取第一个工作表
    sheet = FileObj.sheets()[0]
    # 行数
    row_count = sheet.nrows
    # 列数
    col_count = sheet.ncols
    ct = 0
    global ipText
    print(ipText.get())
    # 搜索关键字符串
    for element in range(row_count):
        if str(ipText.get()) in str(sheet.row_values(element)[0]):
            text.insert(INSERT, sheet.row_values(element)[0] + "\n")
            print(sheet.row_values(element)[0])
            print("答案：%s（请仔细核对选项，考试可能会乱序）" % sheet.row_values(element)[1])
            text.insert(INSERT, "答案：" + str(
                sheet.row_values(element)[1]) + "\n" + "\n")
            print("题库中第%s题" % sheet.row_values(element)[2])
            #  text.insert(INSERT, str(sheet.row_values(element)[2]) + "\n")
            ct = ct + 0
        else:
            ct = ct + 1
        if ct == row_count:
            print("搜索失败，请重新输入关键字")
            text.insert(INSERT, "搜索失败，请重新输入关键字！\n")
    print("搜索结束")


# enter调用
def btn_click_enter(self):
    btn_find()


# 清空消息
def cleartext():
    text.delete('0.0', END)


# 创建窗口对象的背景色
root = Tk()
root.title('题库demo')
root.geometry('390x440')

# Frame为布局函数
main_frame = Frame(root)
text_frame = Frame(main_frame)
station_frame = Frame(main_frame)
botton_frame = Frame(station_frame)
# 建立列表
l1 = Label(station_frame, text='关键字：', width=8, height=5, font=('黑体', 14))
# l2 = Label(station_frame,text='')
ipText = Entry(station_frame, font=('黑体', 14))
# 字体显示
# ft = tkFont.Font(family='Fixdsys', size=10, weight=tkFont.BOLD)
# pack是加载到窗口
l1.pack(side='left')
ipText.pack(side='left', ipady='10')
ipText['width'] = 25

# l2.pack(side='left')

'''
两个函数的意义是既能enter运行，又可以点击运行，方便操作，扩大使用
bind绑定enter键
注意里面是return 而不是enter
'''
b = Button(text='自动搜索,如搜索不到则手动输入', command=btn_click)
b['width'] = 25
b['height'] = 2
b.pack(side='top')
# ipText.bind("<Return>", btn_click_enter)
d = Button(station_frame, text='手动搜索', command=btn_find)
d['width'] = 16
d['height'] = 2
d.pack(side='left')
ipText.bind("<Return>", btn_click_enter)
# 消息输入界面

text = Text(text_frame, width=34, height=12, font=('宋体', 16))
text.pack()
main_frame.pack()
c = Button(text='清空', command=cleartext)
c['width'] = 4
c['height'] = 1
c.pack(side='top')

# 输入框的位置
station_frame.pack(side='top', pady='1')
text_frame.pack(side='top')

# 进入消息循环

root.attributes("-topmost", 1)
root.mainloop()
