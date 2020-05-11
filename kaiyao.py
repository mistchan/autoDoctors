# -*- coding: utf-8 -*-
"""
Created on Wed Oct  9 20:30:08 2019

@author: Administrator
kaiyao.1.0
"""
import time
import win32clipboard as w
from tkinter import *

import pyautogui
import pyperclip
import win32con

pyautogui.PAUSE = 0.3


class App(object):
    ch_e = ''
    cho = ('血常规', '胸片', '腹部彩超', '脑电图', '处方', '结束', '1', '2')

    def __init__(self, root):
        self.root = root

        self.root.wm_attributes('-topmost', 1)
        self.root.overrideredirect(True)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()

        ww = 850
        wh = 20
        self.root.geometry(
            "%dx%d+%d+%d" %
            (ww, wh, (sw - ww) / 2, (sh - wh) / 2 * 1.92))

        self.frame = Frame(self.root)
        self.frame.grid()

        Button(
            self.frame,
            text=self.cho[0],
            command=self.c0,
            width=18,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=0)
        Button(
            self.frame,
            text=self.cho[1],
            command=self.c1,
            width=18,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=3)
        Button(
            self.frame,
            text=self.cho[2],
            command=self.c2,
            width=18,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=6)
        Button(
            self.frame,
            text=self.cho[3],
            command=self.c3,
            width=18,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=9)
        Button(
            self.frame,
            text=self.cho[4],
            command=self.c4,
            width=18,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=12)
        Button(
            self.frame,
            text=self.cho[5],
            command=self.c5,
            width=18,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=15)
        Button(
            self.frame,
            text=self.cho[6],
            command=self.c6,
            width=9,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=18)
        Button(
            self.frame,
            text=self.cho[7],
            command=self.c7,
            width=9,
            height=1,
            activebackground='grey',
            relief='groove').grid(
            row=0,
            column=21)

    def c0(self):
        self.ch_e = self.cho[0]
        self.root.destroy()

    def c1(self):
        self.ch_e = self.cho[1]
        self.root.destroy()

    def c2(self):
        self.ch_e = self.cho[2]
        self.root.destroy()

    def c3(self):
        self.ch_e = self.cho[3]
        self.root.destroy()

    def c4(self):
        self.ch_e = self.cho[4]
        self.root.destroy()

    def c5(self):
        self.ch_e = self.cho[5]
        self.root.destroy()

    def c6(self):
        self.ch_e = self.cho[6]
        self.root.destroy()

    def c7(self):
        self.ch_e = self.cho[7]
        self.root.destroy()


def loop_loc(im):
    while True:
        box = pyautogui.locateOnScreen(im, grayscale=True)
        if not box:
            pass
        else:
            cord = pyautogui.center(box)
            break
    return cord


def ttos(s):
    ss = time.strftime('%Mmin%Ssec', time.gmtime(s))
    return ss


def j_p(lists, area):
    # lists为列表类型数据，area为元组

    for each in lists:
        while True:
            box = pyautogui.locateOnScreen(each, grayscale=True, region=area)
            if not box:
                continue
            else:
                cord = pyautogui.center(box)
                pyautogui.click(cord)
                break


class p:
    def __init__(self, w, a):
        self.w = w
        self.a = a

    def klsfs(self):
        if 8 <= self.w <= 10:
            return 25
        elif 11 <= self.w <= 16:
            return 30
        elif 17 <= self.w <= 20:
            return 50

    def afh(self):
        n1 = [x for x in range(7, 9)]
        n2 = [x for x in range(9, 13)]
        n3 = [x for x in range(13, 20)]
        n4 = [x for x in range(20, 30)]
        if self.w == [5, 6]:
            return 2
        elif self.w in n1:
            return 3
        elif self.w in n2:
            return 4
        elif self.w in n3:
            return 6
        elif self.w in n4:
            return 9

    def zdsk(self):
        if self.w == 5:
            return 50
        elif self.w in [6, 7, 8]:
            return 60
        elif self.w in [9, 10, 11]:
            return 80
        elif self.w in [12, 13, 14, 15]:
            return 125
        elif self.w in [x1 for x1 in range(16, 21)]:
            return 160
        elif self.w in [x1 for x1 in range(21, 25)]:
            return 200
        elif self.w in [x1 for x1 in range(25, 30)]:
            return 250

    def bxp(self):
        if self.w in [12, 13, 14, 15]:
            return 0.125
        elif self.w in [x2 for x2 in range(16, 21)]:
            return 0.16
        elif self.w in [x2 for x2 in range(21, 25)]:
            return 0.2
        elif self.w in [x2 for x2 in range(25, 37)]:
            return 0.25
        elif self.w in [x2 for x2 in range(37, 50)]:
            return 0.375
        elif self.w in [x2 for x2 in range(50, 70)]:
            return 0.5

    def kwk(self):
        if self.w == 5:
            return 10
        elif self.w in [6, 7, 8, 9]:
            return 15
        elif self.w in [10, 11, 12]:
            return 22.5
        elif self.w in [13, 14, 15, 16, 17]:
            return 30
        elif self.w in [x3 for x3 in range(17, 21)]:
            return 37.5
        elif self.w in [x3 for x3 in range(21, 25)]:
            return 45
        elif self.w in [x3 for x3 in range(25, 37)]:
            return 60
        elif self.w in [x3 for x3 in range(37, 100)]:
            return 75

    def akn(self):

        if 1 < self.a <= 2:
            return 0.14
        elif self.a > 2:
            return 0.28

    def lhjn(self):
        if self.w in [x4 for x4 in range(10, 20)]:
            return 50
        elif self.w in [x3 for x3 in range(20, 40)]:
            return 100
        elif self.w in [x3 for x3 in range(40, 70)]:
            return 150

    def fkkl(self):
        if self.a <= 1:
            return 2
        elif 1 < self.a <= 4:
            return 3
        elif 4 < self.a <= 8:
            return 6

    def jbkfy(self):
        if self.a < 1:
            return 3
        elif 1 <= self.a <= 2:
            return 5
        elif 3 <= self.a <= 5:
            return 7.5
        else:
            return 10

    def dckfy(self):
        if self.a < 3:
            return 5
        elif 3 <= self.a < 7:
            return 10
        elif 7 <= self.a < 10:
            return 15
        else:
            return 20

    def qly(self):
        if self.a < 2:
            return 5
        elif 2 <= self.a < 7:
            return 7.5
        elif 7 <= self.a < 10:
            return 10
        else:
            return 15

    def jqkl(self):
        if 3 < self.a < 5:
            return 5
        elif self.a >= 5:
            return 7.5
        else:
            return 3.7

    def jby(self):
        if self.a <= 1:
            return 3
        elif 1 < self.a <= 3:
            return 5
        elif 4 < self.a <= 6:
            return 7.5
        else:
            return 10

    def rdz(self):
        if self.a <= 2:
            return 0.1
        elif self.w <= 18:
            return int(self.w * 0.6)
        elif 19 <= self.w <= 35:
            return 10
        else:
            return 15

    def mnz(self):
        if self.w in [8, 9, 10, 11, 12, 13, 14]:
            return 0.25
        else:
            return int(self.w * 0.2) / 10

    def jsz(self):
        if self.w in [8, 9, 10, 11, 12, 13, 14]:
            return 0.25
        else:
            return int(self.w * 0.2) / 10

    def xmn(self):
        if self.w in [8, 9, 10, 11, 12, 13, 14]:
            return 0.25
        else:
            return int(self.w * 0.2) / 10

    def dxz(self):
        if self.w in [5, 6, 7]:
            return 2.5
        elif self.w in [8, 9, 10, 11]:
            return 3
        elif self.w in [12, 13, 14, 15, 16, 17, 18, 19]:
            return 4
        elif self.w >= 20:
            return 5

    def tdz(self):
        if self.w <= 20:
            return int(self.w * 0.5) / 10
        elif 20 < self.w <= 40:

            return 1

        else:
            return 1.5

    def xdz(self):
        if self.w <= 33:
            return int(self.w * 0.3) / 10
        elif self.w > 33:
            return 1.0

    def scz(self):
        if self.w <= 20:
            return int(self.w * 0.5) / 10
        elif 20 < self.w <= 40:

            return 1

        else:
            return 1.5

    def tqz(self):
        if int(self.w * 0.5) / 10 >= 1.5:
            return 1.5
        elif 1 <= int(self.w * 0.5) / 10 < 1.5:
            return 1
        else:
            return int(self.w * 0.5) / 10

    def ndz(self):
        if int(self.w * 0.5) / 10 >= 1.5:
            return 1.5
        elif 1 <= int(self.w * 0.5) / 10 < 1.5:
            return 1
        else:
            return int(self.w * 0.5) / 10

    def aqz(self):
        return self.w * 0.01

    def qxz(self):
        return self.w * 0.01

    def lyz(self):
        if self.w in [4, 5]:
            return 0.2
        elif self.w in [6, 7]:
            return 0.3
        elif self.w in [8, 9]:
            return 0.4
        else:
            return 0.5

    def pkkm(self):
        if self.w < 20:
            return 1.5
        elif 20 <= self.w < 50:
            return 2.5
        else:
            return 5

    def psgd(self):
        if self.w < 20:
            return 0.7
        elif 20 <= self.w < 50:
            return 1
        else:
            return 2

    def pacj(self):
        if self.w < 30:
            return 0.03
        else:
            return 0.05

    def ylha(self):
        if self.w in [5, 6]:
            return 2.5

        elif self.w in [7, 8]:
            return 3
        elif self.w in [9, 10]:
            return 4
        elif self.w in [xl1 for xl1 in range(11, 20)]:
            return 5
        elif self.w in [xl for xl in range(20, 40)]:
            return 7.5
        else:
            return 10

    def kldllt(self):
        if self.a < 1:
            return 0.13
        elif 1 <= self.a <= 5:
            return 0.25
        elif 5 < self.a <= 11:
            return 0.5
        elif self.a >= 12:
            return 1

    def klmlst(self):
        if self.a < 2:
            return 2.5
        elif 2 <= self.a <= 5:
            return 4
        elif 5 < self.a <= 14:
            return 5

    def akz(self):
        dott = int(self.w * 0.3) / 10
        if dott <= 1:
            return dott
        elif 1 < dott < 1.5:
            return 1
        elif dott >= 1.5:
            return 1.5

    def gxz(self):
        return int(self.w * 0.5) * 0.01

    def wcz(self):
        if self.a < 3:
            return 0.5
        else:
            return 1

    def yhz(self):
        if self.w in [8, 9, 10, 11, 12, 13]:
            return 60
        elif self.w in [14, 15, 16, 17]:
            return 80
        elif self.w in [18, 19, 20]:
            return 100
        elif self.w in [21, 22, 23, 24, 25]:
            return 120


# 自动化
class Tv(object):
    p_count = 0
    dur_m = 0
    dur_s = 0

    end_t = 0
    start_t = 0
    speed = 0


class Pcode(object):
    choice = ''

    def __init__(self, root):
        self.root = root
        self.root.title('Patient ID')
        self.root.wm_attributes('-topmost', 1)

        self.root.wm_attributes('-alpha', 0.75)

        self.root.overrideredirect(False)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()

        ww = 400
        wh = 70
        self.root.geometry("%dx%d+%d+%d" %
                           (ww, wh, (sw - ww) / 2, int((sh - wh) / 2) * 0.8))
        self.root.after(1, lambda: self.root.focus_force())
        self.inf = StringVar()
        self.l = Label(self.root, textvariable=self.inf, font=('微软雅黑', 10))
        self.inf.set('Last consultion takes: ' + str(ttos(Tv.dur_s)) + '\nNo:' +
                     str(Tv.p_count + 1) + '(' + str(Tv.speed) + '/per patient)')

        self.l.pack()
        self.frame = Frame(self.root)
        self.frame.pack()
        self.e = Entry(
            self.frame,
            relief='groove',
            width=40,
            font=(
                '微软雅黑',
                10),
            bg='silver')
        self.e.focus_set()
        self.e.bind('<Key-Return>', self.c0)
        self.e.bind('<FocusIn>', self.c2)
        self.e.bind('<FocusOut>', self.c3)
        self.e.bind('<Control-Key-i>', self.min_)
        self.e.pack(side=LEFT)

        Button(
            self.frame,
            text='取消',
            command=self.c1,
            width=10,
            height=1,
            activebackground='grey',
            relief='groove',
            font=(
                '黑体',
                10)).pack(
            side=RIGHT)

    def c0(self, event):
        self.choice = self.e.get()
        self.root.destroy()

    def c1(self):
        self.choice = None
        self.root.destroy()

    def c2(self, event):
        self.e['bg'] = 'white'

    def c3(self, event):
        self.e['bg'] = 'silver'

    def min_(self, event):
        self.root.state('iconic')


class Drug(object):
    drug = {
        'klsfs': '头孢克肟颗粒',
        'zdsk': '克洛己新干混悬剂',
        'afh': '小儿氨酚黄那敏颗粒',
        'bxp': '头孢丙烯片',
        'kwk': '奥司他韦颗粒',
        'akn': '阿克诺片',
        'lhjn': '罗红霉素胶囊',
        'jbkfy': '解表口服液',
        'dckfy': '定喘口服液',
        'qly': '鱼腥草芩蓝合剂',
        'jqkl': '小儿金翘颗粒',
        'jby': '桔贝合剂',
        'pkkm': '克咳敏片',
        'psgd': '赛庚啶片',
        'pacj': '氨茶碱片',
        'ylha': '氯化铵',
        'klmlst': '孟鲁司特颗粒',
        'wcz': '维生素C',
        'rdz': '热毒宁',
        'mnz': '头孢米诺(福安)',
        'jsz': '头孢米诺(金石)',
        'xdz': '头孢西丁',
        'gxz': '更昔洛韦',
        'dxz': '地塞米松',
        'tdz': '头孢他啶',
        'ndz': '头孢曲松(那单)',
        'yhz': '炎琥宁',
        'tqz': '头孢曲松(立健松)',
        'aqz': '小阿奇',
        'qxz': '大阿奇',
        'lyz': '拉氧头孢0.5',
        'akz': '阿莫西林克拉维酸钾针',
        'xmn': '头孢米诺(欣峰)'}

    drugs = {
        'mnrd': (
            'mnz', 'rdz'), 'tqrd': (
            'tqz', 'rdz'), 'tdrd': (
                'tdz', 'rdz'), 'lygx': (
                    'lyz', 'gxz'), 'tqyh': (
                        'tqz', 'yhz'), 'xdrd': (
                            'xdz', 'rdz'), 'mnyh': (
                                'mnz', 'yhz'), 'akrd': (
                                    'akz', 'rdz'), 'fr': (
                                        'jqkl', 'afh', 'kwk', 'bxp'), 'ks': (
                                            'pkkm', 'psgd', 'ylha', 'bxp', 'jqkl'), 'zdkw': (
                                                'zdsk', 'kwk', 'jqkl', 'afh'), 'tqgx': (
                                                    'tqz', 'gxz')}

    age_s = ''
    weight_s = ''


while True:
    para = Drug()

    if Tv.dur_m != 0 and Tv.p_count != 0:
        Tv.speed = ttos(Tv.dur_m / Tv.p_count)
    else:
        Tv.speed = 0
    root3 = Tk()
    code_p = Pcode(root3)
    root3.mainloop()
    bingren = code_p.choice
    Tv.start_t = time.time()
    Tv.p_count += 1

    if not bingren:
        break
    else:

        pyautogui.click(loop_loc(r'.\\shot\就诊卡.png'))
        pyautogui.typewrite(str(bingren))
        pyautogui.press('enter')

        pyautogui.press('space')
        pyautogui.press('enter')
        pyautogui.press('space')

    while True:

        root1 = Tk()

        root1.overrideredirect(True)
        root1.wm_attributes('-topmost', 1)
        sw1 = root1.winfo_screenwidth()
        sh1 = root1.winfo_screenheight()

        ww1 = 200
        wh1 = 670
        root1.geometry("%dx%d+%d+%d" %
                       (ww1, wh1, (sw1 - ww1) / 2, int((sh1 - wh1) / 2) * 1.4))

        root1.after(1, lambda: root1.focus_force())
        zd = StringVar()

        def bb(event):
            global zdg

            zdg = []
            for each in lb1.curselection():
                zdg.append(lb1.get(each))

            root1.destroy()
            return zdg

        def bc(event):
            global zdg

            zdg = []

            root1.destroy()
            return zdg

        lz = Label(root1, text='请选择新增诊断名称')
        lz.pack()
        fz = Frame(root1)
        fz.pack()
        crollz = Scrollbar(fz, orient=VERTICAL)
        crollz.pack(side=RIGHT, fill=Y)
        zd.set(
            ('呼吸道感染',
             '化脓性扁桃体炎',
             '疱疹性咽峡炎',
             '口炎',
             '粒细胞减少症',
             'SIRS',
             '急性喉炎',
             '支气管炎',
             '肺炎',
             '毛细支气管炎',
             '喘息性支气管炎',
             '胃肠炎',
             '肠炎',
             '胃肠功能紊乱',
             '腹痛',
             '呕吐',
             '肠系膜淋巴结炎',
             '急性荨麻疹',
             '皮疹',
             '颅内感染',
             '头痛',
             '头晕',
             '植物神经功能紊乱',
             '泌尿系感染',
             '过敏性紫癜',
             '疫苗接种反应',
             '周围神经炎',
             '癫痫',
             '惊厥',
             '药物中毒',
             '农药中毒'))
        lb1 = Listbox(
            fz,
            listvariable=zd,
            selectmode=MULTIPLE,
            bd=1,
            relief='groove',
            font=(
                '微软雅黑',
                10),
            selectbackground='brown',
            width=25,
            height=32,
            yscrollcommand=crollz.set)
        lb1.focus_set()
        lb1.bind('<Button-3>', bb)

        lb1.bind('<Key-q>', bc)
        lb1.pack(side=RIGHT)
        crollz.config(command=lb1.yview)

        '''
        fz1 = Frame(root1)
        fz1.pack()
        bt1 = Button(fz1, text='确 定', command=bb, width=12, height=2,

                    activebackground='grey', relief='groove')
        bt1.pack(side=LEFT)
        bt2 = Button(fz1, text='取消', command=bc, width=12, height=2,

                    activebackground='grey', relief='groove')
        bt2.pack(side=RIGHT)
        '''
        root1.mainloop()
        if not zdg:
            pyautogui.press('esc')
            break

        else:

            for eachzd in zdg:
                clicks = loop_loc('.\\shot\\' + str(eachzd) + '.png')

                pyautogui.doubleClick(clicks)
                time.sleep(0.2)
            time.sleep(0.5)
            pyautogui.click(800, 700)
            pyautogui.hotkey('alt', 'c')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.press('enter')
            break

    while True:
        root1 = Tk()
        root_m = App(root1)
        root1.mainloop()
        m = root_m.ch_e

        if m == '血常规':
            pyautogui.click(471, 43)
            time.sleep(1.5)
            pyautogui.click(41, 317)
            pyautogui.doubleClick(140, 369)
            pyautogui.click(681, 235)
            pyautogui.press('esc')

        elif m == '胸片':
            rot = Tk()
            rot.title('书写病史')

            sw2 = rot.winfo_screenwidth()
            sh2 = rot.winfo_screenheight()

            ww2 = 600
            wh2 = 700
            rot.geometry(
                "%dx%d+%d+%d" %
                (ww2, wh2, (sw2 - ww2) / 2, (sh2 - wh2) / 2))

            zd = StringVar()

            lb1 = Label(rot, text='检查目的')
            lb1.pack()

            text1 = Text(rot)
            text1.pack()
            text1.delete(1.0, END)
            text1.insert(END, "了解肺部情况。")
            lb2 = Label(rot, text='现病史')
            lb2.pack()

            text2 = Text(rot)
            text2.pack()
            text2.delete(1.0, END)
            text2.insert(END, "反复咳嗽原因待查。肺部听诊：双肺呼吸音粗，闻及少量湿罗音。")

            def bxb():
                global bs1, bs2

                bs1 = text1.get(1.0, END)
                bs2 = text2.get(1.0, END)

                rot.destroy()
                return bs1, bs2

            def bxc():
                global bs1, bs2
                bs1 = bs2 = ''
                rot.destroy()
                return bs1, bs2

            fzx = Frame(rot)
            fzx.pack()
            bx1 = Button(fzx, text='确 定', command=bxb, width=12, height=2,

                         activebackground='grey', relief='groove')
            bx1.pack(side=LEFT)
            bx2 = Button(fzx, text='取 消', command=bxc, width=12, height=2,

                         activebackground='grey', relief='groove')
            bx2.pack(side=RIGHT)

            rot.mainloop()

            def paste1(aString):  # 写入剪切板
                w.OpenClipboard()
                w.EmptyClipboard()
                w.SetClipboardData(win32con.CF_TEXT, aString)
                w.CloseClipboard()
                pyautogui.hotkey('ctrl', 'v')

            def paste(foo):
                pyperclip.copy(foo)

                pyautogui.hotkey('ctrl', 'v')

            pyautogui.click(529, 43)  # 检验
            time.sleep(2)
            pyautogui.click(305, 813)  # 开单
            pyautogui.click(526, 89)  # 检查类型
            time.sleep(0.5)
            pyautogui.click(526, 89)
            pyautogui.press('home')
            pyautogui.click(526, 89)
            pyautogui.click(477, 153)  # 胸片
            pyautogui.click(732, 283)  # 检查目的
            paste(str(bs1))
            pyautogui.click(732, 343)  # 病历摘要
            paste(str(bs2))
            pyautogui.doubleClick(428, 391)  # 选择检查项目
            pyautogui.click(908, 224)  # 胸部正位
            pyautogui.click(930, 669)  # 确定
            pyautogui.click(685, 89)  # 保存
            pyautogui.click(805, 496)  # 确定
            pyautogui.click(770, 89)  # 打印
            pyautogui.click(882, 480)  # 确定
            pyautogui.click(865, 89)  # 退出

            # pyautogui.press('esc')

        elif m == '脑电图':
            rot = Tk()
            rot.title('书写病史')

            sw2 = rot.winfo_screenwidth()
            sh2 = rot.winfo_screenheight()

            ww2 = 600
            wh2 = 700
            rot.geometry(
                "%dx%d+%d+%d" %
                (ww2, wh2, (sw2 - ww2) / 2, (sh2 - wh2) / 2))
            rot.after(1, lambda: rot.focus_force())
            zd = StringVar()

            lb1 = Label(rot, text='检查目的')
            lb1.pack()

            text1 = Text(rot)
            text1.pack()
            text1.delete(1.0, END)
            text1.insert(END, "了解脑电情况。")
            lb2 = Label(rot, text='现病史')
            lb2.pack()

            text2 = Text(rot)
            text2.pack()
            text2.delete(1.0, END)
            text2.insert(END, "反复发热、头痛原因待查。查体：颈弱抵抗，腱反射活跃。")

            def bxb():
                global bs1, bs2

                bs1 = text1.get(1.0, END)
                bs2 = text2.get(1.0, END)

                rot.destroy()
                return bs1, bs2

            def bxc():
                global bs1, bs2
                bs1 = bs2 = ''
                rot.destroy()
                return bs1, bs2

            fzx = Frame(rot)
            fzx.pack()
            bx1 = Button(fzx, text='确 定', command=bxb, width=12, height=2,

                         activebackground='grey', relief='groove')
            bx1.pack(side=LEFT)
            bx2 = Button(fzx, text='取 消', command=bxc, width=12, height=2,

                         activebackground='grey', relief='groove')
            bx2.pack(side=RIGHT)

            rot.mainloop()

            def paste1(aString):  # 写入剪切板
                w.OpenClipboard()
                w.EmptyClipboard()
                w.SetClipboardData(win32con.CF_TEXT, aString)
                w.CloseClipboard()
                pyautogui.hotkey('ctrl', 'v')

            def paste(foo):
                pyperclip.copy(foo)

                pyautogui.hotkey('ctrl', 'v')

            pyautogui.click(529, 43)  # 检验
            time.sleep(2)
            pyautogui.click(305, 813)  # 开单
            pyautogui.click(526, 89)  # 检查类型
            time.sleep(0.5)
            pyautogui.click(526, 89)
            pyautogui.press('home')
            pyautogui.click(526, 89)
            pyautogui.click(470, 138)  # 脑电图
            pyautogui.click(732, 283)
            paste(str(bs1))
            pyautogui.click(732, 343)
            paste(str(bs2))
            pyautogui.doubleClick(405, 136)
            pyautogui.doubleClick(405, 192)
            pyautogui.doubleClick(405, 227)

            pyautogui.click(685, 89)  # 保存
            pyautogui.click(805, 496)  # 确定
            pyautogui.click(770, 89)  # 打印
            pyautogui.click(882, 480)  # 确定
            pyautogui.click(865, 89)  # 退出

            # pyautogui.press('esc')

        elif m == '腹部彩超':
            rot1 = Tk()
            rot1.title('书写病史')

            sw21 = rot1.winfo_screenwidth()
            sh21 = rot1.winfo_screenheight()

            ww21 = 600
            wh21 = 700
            rot1.geometry(
                "%dx%d+%d+%d" %
                (ww21, wh21, (sw21 - ww21) / 2, (sh21 - wh21) / 2))
            rot1.after(1, lambda: rot1.focus_force())

            lb11 = Label(rot1, text='检查目的')
            lb11.pack()

            text11 = Text(rot1)
            text11.pack()
            text11.delete(1.0, END)
            text11.insert(END, "除外阑尾炎；除外肠套叠或其他梗阻性疾病。")
            lb21 = Label(rot1, text='现病史')
            lb21.pack()

            text21 = Text(rot1)
            text21.pack()
            text21.delete(1.0, END)
            text21.insert(END, "腹痛、呕吐原因待查。")

            def bxb1():
                global bs11, bs21

                bs11 = text11.get(1.0, END)
                bs21 = text21.get(1.0, END)

                rot1.destroy()
                return bs11, bs21

            def bxc1():
                global bs11, bs21
                bs11 = bs21 = ''
                rot1.destroy()
                return bs11, bs21

            fzx1 = Frame(rot1)
            fzx1.pack()
            bx11 = Button(fzx1, text='确 定', command=bxb1, width=12, height=2,

                          activebackground='grey', relief='groove')
            bx11.pack(side=LEFT)
            bx21 = Button(fzx1, text='取消', command=bxc1, width=12, height=2,

                          activebackground='grey', relief='groove')
            bx21.pack(side=RIGHT)

            rot1.mainloop()

            def paste1(aString):  # 写入剪切板
                w.OpenClipboard()
                w.EmptyClipboard()
                w.SetClipboardData(win32con.CF_TEXT, aString)
                w.CloseClipboard()
                pyautogui.hotkey('ctrl', 'v')

            def paste(foo):
                pyperclip.copy(foo)

                pyautogui.hotkey('ctrl', 'v')

            pyautogui.click(529, 43)  # 检验
            time.sleep(2)
            pyautogui.click(305, 813)  # 开单
            pyautogui.click(526, 89)  # 检查类型
            time.sleep(0.5)
            pyautogui.click(526, 89)
            pyautogui.press('home')

            pyautogui.click(732, 283)
            paste(str(bs11))
            pyautogui.click(732, 343)
            paste(str(bs21))
            pyautogui.doubleClick(430, 460)
            pyautogui.press('pgdn')
            pyautogui.doubleClick(422, 552)
            pyautogui.press('pgdn')
            pyautogui.doubleClick(420, 558)
            pyautogui.click(685, 89)  # 保存
            pyautogui.click(805, 496)  # 确定
            pyautogui.click(770, 89)  # 打印
            pyautogui.click(882, 480)  # 确定
            pyautogui.click(865, 89)  # 退出

            # pyautogui.press('esc')
        elif m == '处方':
            class Rp(object):
                drug_seris = []
                drug_s = []
                dot = []
                drug_name = []
                code_input = []

                def __init__(self, root):
                    self.root = root
                    self.root.title('请输入患者信息')

                    sw = self.root.winfo_screenwidth()
                    sh = self.root.winfo_screenheight()

                    ww = 600
                    wh = 750
                    self.root.geometry(
                        "%dx%d+%d+%d" %
                        (ww, wh, (sw - ww) / 2, (sh - wh) / 2))
                    self.root.after(1, lambda: self.root.focus_force())
                    self.t_age = StringVar()
                    self.t_weight = StringVar()
                    self.t_input = StringVar()
                    self.t_input_show = StringVar()
                    self.val = StringVar()

                    self.t_age.set(para.age_s)
                    self.t_weight.set(para.weight_s)
                    self.t_input_show.set('请输入成套方案：')
                    l1 = Label(self.root, text='患者年龄:')
                    l1.pack()

                    self.e1 = Entry(
                        self.root,
                        textvariable=self.t_age,
                        width=10,
                        font=(
                            '微软雅黑',
                            10))
                    self.e1.focus_set()
                    self.e1.pack()

                    l2 = Label(self.root, text='患者体重:')
                    l2.pack()

                    self.e2 = Entry(
                        self.root,
                        textvariable=self.t_weight,
                        width=10,
                        font=(
                            '微软雅黑',
                            10))
                    self.e2.pack()

                    l3 = Label(self.root, textvariable=self.t_input_show)
                    l3.pack()

                    self.e3 = Entry(
                        self.root,
                        textvariable=self.t_input,
                        width=30,
                        font=(
                            '微软雅黑',
                            12))

                    self.e3.bind('<Key-Return>', self.e_get)
                    self.e3.bind('<Control-Key-e>', self.e_end)
                    self.e3.pack()

                    l4 = Label(self.root, text='已开药物及用量：')
                    l4.pack()

                    f = Frame(self.root)
                    f.pack()
                    croll = Scrollbar(f, orient=VERTICAL)
                    croll.pack(side=RIGHT, fill=Y)

                    self.lb = Listbox(
                        f,
                        listvariable=self.val,
                        selectmode=MULTIPLE,
                        bd=1,
                        font=(
                            '微软雅黑',
                            10),
                        selectbackground='brown',
                        width=50,
                        height=25,
                        yscrollcommand=croll.set)
                    self.lb.pack(side=RIGHT)
                    croll.config(command=self.lb.yview)

                    fdrug = Frame(self.root)
                    fdrug.pack()

                    btc = Button(
                        fdrug,
                        text='取 消',
                        command=self.bcc,
                        width=12,
                        height=2,
                        activebackground='grey',
                        relief='groove')
                    btc.pack()

                def bcc(self):
                    self.drug_seris = []
                    self.drug_s = []
                    self.dot = []
                    self.drug_name = []
                    self.code_input = []

                    self.root.destroy()

                def e_get(self, event):

                    para.age_s = self.e1.get()
                    para.weight_s = self.e2.get()

                    cmdc = self.e3.get()

                    if cmdc in para.drugs.keys():
                        for each in para.drugs[cmdc]:
                            self.drug_seris.append(each)

                    elif cmdc in para.drug.keys():
                        self.drug_seris.append(cmdc)

                    if self.drug_seris:
                        self.dot = []

                        patient = p(int(para.weight_s), int(para.age_s))
                        for eachcode in self.drug_seris:
                            dot = getattr(patient, eachcode)

                            self.dot.append(dot())

                        self.code_input = list(zip(self.drug_seris, self.dot))

                    self.lb.delete(0, END)
                    for eachx in self.code_input:
                        self.lb.insert(END, (para.drug[eachx[0]], eachx[1]))
                    self.e3.delete(0, 'end')

                def e_end(self, event):
                    self.root.destroy()

            root = Tk()
            rp = Rp(root)
            root.mainloop()

            if rp.code_input:

                area_c = (591, 20, 296, 300)
                list_c = ['.\\shot\\处方.png',
                          '.\\shot\\普通.png',
                          '.\\shot\\三角箭头.png',
                          '.\\shot\\儿童.png']
                j_p(list_c, area_c)
                result = rp.code_input

                for each1 in result:
                    code = each1[0]
                    dot = each1[1]

                    # 自动化

                    pyautogui.press('f7')
                    pyautogui.typewrite(str(code))
                    pyautogui.press('enter')
                    pyautogui.click(757, 496)
                    time.sleep(0.5)
                    pyautogui.typewrite(str(dot))

                pyautogui.alert(text='患儿年龄：{:0>2d} 岁\n患儿体重：{:0>2d} kg'.format(
                    int(para.age_s), int(para.weight_s)), title='请确认:！', button='OK')
                pyautogui.click(930, 730)  # 获得焦点

                pyautogui.press('f9')  # 保存
                time.sleep(0.5)
                pyautogui.press('enter')
                list_3 = ['.\\shot\\打印.png',
                          '.\\shot\\打印确认.png',
                          '.\\shot\\确定.png']
                area3 = (395, 23, 579, 812)
                j_p(list_3, area3)

        elif m == '1':
            '''
            pyautogui.click(640, 41)
            time.sleep(1)
            pyautogui.click(840, 129)
            pyautogui.click(792, 211)

            '''
            area1 = (591, 20, 296, 300)
            list_1 = ['.\\shot\\处方.png',
                      '.\\shot\\普通.png',
                      '.\\shot\\三角箭头.png',
                      '.\\shot\\儿童.png']
            j_p(list_1, area1)

        elif m == '2':
            pyautogui.click(930, 730)  # 获得焦点
            pyautogui.press('f9')
            time.sleep(0.5)
            pyautogui.press('enter')
            list_2 = ['.\\shot\\打印.png',
                      '.\\shot\\打印确认.png',
                      '.\\shot\\确定.png']
            area2 = (395, 23, 579, 812)
            j_p(list_2, area2)

        else:
            pyautogui.click(500, 500)
            pyautogui.press('esc')
            pyautogui.click(296, 40)
            time.sleep(0.5)
            pyautogui.click(624, 441)
            pyautogui.click(938, 450)
            Tv.end_t = time.time()
            Tv.dur_m += Tv.end_t - Tv.start_t
            Tv.dur_s = Tv.end_t - Tv.start_t
            break
