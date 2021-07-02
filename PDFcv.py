from time import sleep
from sys import path as syspath
from os import path as ospath
import re
from win32gui import GetWindowText, GetForegroundWindow

syspath.append(ospath.abspath("SO_site-packages"))
import pyperclip  # 引入模块
import itertools, glob
from sysTrayIcon import SysTrayIcon

recent_value = ""
# tmp_value = ""  # 初始化（应该也可以没有这一行，感觉意义不大。但是对recent_value的初始化是必须的）
E_pun = u';,.!?:()"\''
C_pun = u'；，。！？：（）“‘'
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
        elif (inside_code >= 65281 and inside_code <= 65374):  # 全角字符（除空格）根据关系转化
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
        elif inside_code >= 32 and inside_code <= 126:  # 半角字符（除空格）根据关系转化
            inside_code += 65248

        rstring += chr(inside_code)
    return rstring


def E_C_trans(string):
    """英文标点转中文标点 https://blog.csdn.net/nanbei2463776506/article/details/82967140
    """
    global E_pun, C_pun, tableE2C, tableC2E
    for ep in E_pun:
        pat = "(?<=[^\\x00-\\xff])\\" + ep
        string = re.sub(pat, tableE2C[ep], string)
    for cp in C_pun:
        string = re.sub('(?<=[\\x00-\\xff])' + cp, tableC2E[cp], string)
    return string.translate(tableE2C)


def PDFcv():
    global recent_value
    # 获取windows活动窗口、最前窗口的标题 https://blog.csdn.net/shjsfx/article/details/106089331
    windowName = GetWindowText(GetForegroundWindow())
    while 1:
        if ("福昕阅读器" in windowName) or ("WPS" in windowName) or ("秒" in windowName):
            tmp_value = pyperclip.paste()  # 读取剪切板复制的内容

            try:
                if tmp_value != recent_value:  # 如果检测到剪切板内容有改动，那么就进入文本的修改
                    recent_value = tmp_value
                    out = re.sub(r"\s{2,}", " ", recent_value)  # 将文本的换行符去掉，变成一个空格
                    changed = re.sub(r'(?<=[^\\x00-\\xff]) (?=[^\\x00-\\xff])', '', out)  # 中文去空格
                    changed1 = strQ2B(changed)  # 全角转半
                    changed2 = E_C_trans(changed1)
                    pyperclip.copy(changed2)  # 将修改后的文本写入系统剪切板中
                    print("\n Value changed: %s" % str(changed2))  # 输出已经去除换行符的文本
                sleep(0.1)
            except KeyboardInterrupt:  # 如果有ctrl+c，那么就退出这个程序。  （不过好像并没有用。无伤大雅）
                break
            return tmp_value


# https://blog.csdn.net/weixin_39278265/article/details/84194996

if __name__ == '__main__':
    # recent_value = PDFcv()

    hover_text = "PDFcv"

    def bye(sysTrayIcon): print('Bye, then.')


    SysTrayIcon("001.ico", hover_text, (), on_quit=bye, default_menu_index=1, window_class_name=PDFcv)
