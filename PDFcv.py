from time import sleep
from sys import path as syspath
from os import path as ospath
from re import sub
from win32gui import GetForegroundWindow, GetWindowText
from pyperclip import paste, copy# 引入模块

syspath.append(ospath.abspath("SO_site-packages"))

recent_value = ""
tmp_value = ""  # 初始化（应该也可以没有这一行，感觉意义不大。但是对recent_value的初始化是必须的）
E_pun = u',.!?:;()"\''
C_pun = u'，。！？：；（）“‘'
tableE2C = {f: t for f, t in zip(E_pun, C_pun)}
tableC2E = {t: f for t, f in zip(C_pun, E_pun)}


def strQ2B(ustring):
    """全角转半角# https://www.cnblogs.com/kaituorensheng/p/3554571.html
    """
    rstring = ""
    for uchar in ustring:
        inside_code = ord(uchar)
        if inside_code == 12288:  # 全角空格直接转换
            inside_code = 32
        elif 65281 <= inside_code <= 65374:  # 全角字符（除空格）根据关系转化
            inside_code -= 65248

        rstring += chr(inside_code)
    return rstring


def strB2Q(ustring):
    """半角转全角"""
    rstring = ""
    for uchar in ustring:
        inside_code = ord(uchar)
        if inside_code == 32:  # 半角空格直接转化
            inside_code = 12288
        elif 32 <= inside_code <= 126:  # 半角字符（除空格）根据关系转化
            inside_code += 65248

        rstring += chr(inside_code)
    return rstring


def E_C_trans(string):
    """英文标点转中文标点 https://blog.csdn.net/nanbei2463776506/article/details/82967140
    """
    global E_pun, C_pun, tableE2C, tableC2E
    for ep in E_pun:
        pat = r"(?<=[^\x00-\xff])" + "\\" + ep
        string = sub(pat, tableE2C[ep], string)
    for cp in C_pun:
        string = sub(r"(?<=[\x00-\xff])" + cp, tableC2E[cp], string)
    return string.translate(tableE2C)


def details(string):
    new1 = string.replace("", "，")
    new2 = new1.replace("∙", ".")
    new3 = new2.replace("- ", "")
    return new3


while True:
    # 获取windows活动窗口、最前窗口的标题 https://blog.csdn.net/shjsfx/article/details/106089331
    winName = GetWindowText(GetForegroundWindow())
    if ("福昕阅读器" in winName) or ("WPS" in winName) or ("知云" in winName) or ("Xodo" in winName):
        tmp_value = paste()  # 读取剪切板复制的内容

        try:
            if tmp_value != recent_value:  # 如果检测到剪切板内容有改动，那么就进入文本的修改
                # recent_value = tmp_value

                out = sub(r"\s{2,}", " ", tmp_value)  # 将文本的换行符去掉，变成一个空格
                our1 = sub(r'(?<=[^\x00-\xff])\s(?=[^\x00-\xff])', '', out)  # 中文去空格
                out2 = details(our1)  # 细节
                out3 = strQ2B(out2)  # 全角转半角
                result = E_C_trans(out3)  # 中英文标点转换

                recent_value = result
                copy(result)  # 将修改后的文本写入系统剪切板中

                print("\n Value changed: %s" % recent_value)  # 输出已经去除换行符的文本
            sleep(0.1)
        except KeyboardInterrupt:  # 如果有ctrl+c，那么就退出这个程序。  （不过好像并没有用。无伤大雅）
            break

# https://blog.csdn.net/weixin_39278265/article/details/84194996
