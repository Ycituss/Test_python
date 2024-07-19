#!/usr/bin/env python
# -*- coding: utf-8 -*-

from tkinter import *
import time
from tkinter import filedialog
from tkinter.font import Font

import pandas as pd


LOG_LINE_NUM = 0
pathA = "F:/共享/亚马逊/货师傅/发货订单/1111.xlsx"
pathB = "F:/共享/亚马逊/货师傅/发货订单/ImportMyOderD.xlsx"
pathC = "F:/共享/亚马逊/货师傅/发货订单/FaHuo.xlsx"
pathD = "F:/共享/亚马逊/货师傅/发货订单/HSFDaoChu.xlsx"

# SKU对应关系
sku_list = {
    '7901838385': '蚂蚁8',
    '5346822267': '蚂蚁4',
    '1005642201': '蚂蚁4',
    '1598435358': 'AC150',
    '8528352250': 'AC150',
    '1588181886': 'KnittedDoll5PCS',
    '1938881666': 'DishRack',
    '2206014831': 'X0047YAX8F',
    '2702999814': 'X0047YG8EN',
    '5539581531': '万紫千红',
    '9544198932': 'X0048FC813',
    '5185585871': 'X0048FC81D',
    '4759341561': 'X0048FC80T',
    '9773328250': 'X0048FC827',
    '9584073002': 'X0048FC81N',
    '8486112749': 'X0047SX9CD',
    '4609416887': 'X0048FC81X',
    '9930622132': 'X004869WLB',
    '4689636904': 'X0047T9MWX',
    '3598367395': 'EleFam3PCS',
    '3922779219': '提灯象',
    '6186035158': '情侣象',
    '1419226715': 'Lavender260PCS',
    '9803154652': '尤加利叶12',
    '3360709907': 'Youjialiye24PCS',
    '2529119064': 'Tulip20PCS',
    '6891024543': 'X0047YG8EN',
    '7456154194': 'X0047YAX8F',
    '9722031114': '万紫千红',
    '8484627031': 'X0047SX9CD',
    '3170836785': 'X0047T9MWX',
    '4395011099': 'X0048FC81X',
    '3062157200': 'X0048FC81N',
    '2186991813': 'X0048FC827',
    '9099181244': 'X0048FC81D',
    '6951960713': 'X0048FC813',
    '5303569445': 'X0048FC80T',
    '2303483163': 'X004869WLB',
    '1848137163': 'KnittedDUCK5PCS',
    '---temp---': '---temp---',
    '1161281136': 'X0047YAX8F'
}


class MY_GUI:
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name

    # 设置窗口
    def set_init_window(self, name):
        self.init_window_name.title(name)  # 窗口名
        self.init_window_name.geometry('500x600+800+400')
        # 标签
        self.log_label = Label(self.init_window_name, text="日志")
        self.log_label.grid(row=0, column=0)
        self.tips = Label(self.init_window_name, text="请打开从TEMU导出的订单", bg='red', font=Font(family="Helvetica", size=20))
        self.tips.grid(row=3, column=4)
        # 文本框
        self.log_data_Text = Text(self.init_window_name, width=66, height=9)  # 日志框
        self.log_data_Text.grid(row=1, column=0, columnspan=10)
        # 按钮
        self.main_button0 = Button(self.init_window_name, text="打开temu导出文件", bg="lightblue", width=14,
                                   command=self.get_file_path0)  # 调用内部方法  加()为直接调用
        self.main_button0.grid(row=2, column=4)

    # 获取temu导出文件路径
    def get_file_path0(self):
        global pathA
        file_path = filedialog.askopenfilename()
        print("Selected file path:", file_path)
        write_log_to_Text(self, file_path + '导入成功')
        self.main_button0.grid_forget()
        pathA = file_path
        temp_length = temu_to_hsf()
        self.tips.grid_forget()
        self.tips1 = self.tips = Label(self.init_window_name, text="转换中，请稍等", bg='red', font=Font(family="Helvetica", size=20))
        self.tips1.grid(row=3, column=4)
        time.sleep(6)

        self.main_button1 = Button(self.init_window_name, text="打开货师傅导出文件", bg="lightblue", width=14,
                                   command=self.get_file_path1)  # 调用内部方法  加()为直接调用
        self.main_button1.grid(row=2, column=4)
        self.tips1.grid_forget()
        self.tips2 = self.tips = Label(self.init_window_name, text="请将ImportMyOderD导入货师傅", bg='red', font=Font(family="Helvetica", size=20))
        self.tips2.grid(row=3, column=4)

    def get_file_path1(self):
        pass


    # # 获取当前时间
    # def get_current_time(self):
    #     current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    #     return current_time

    # # 日志动态打印
    # def write_log_to_Text(self, logmsg):
    #     global LOG_LINE_NUM
    #     current_time = self.get_current_time()
    #     logmsg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
    #     if LOG_LINE_NUM <= 7:
    #         self.log_data_Text.insert(END, logmsg_in)
    #         LOG_LINE_NUM = LOG_LINE_NUM + 1
    #     else:
    #         self.log_data_Text.delete(1.0, 2.0)
    #         self.log_data_Text.insert(END, logmsg_in)


# 获取当前时间
def get_current_time():
    current_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    return current_time


# 日志动态打印
def write_log_to_Text(self, logmsg):
    global LOG_LINE_NUM
    current_time = get_current_time()
    logmsg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
    if LOG_LINE_NUM <= 7:
        self.log_data_Text.insert(END, logmsg_in)
        LOG_LINE_NUM = LOG_LINE_NUM + 1
    else:
        self.log_data_Text.delete(1.0, 2.0)
        self.log_data_Text.insert(END, logmsg_in)


def temu_to_hsf():
    global pathA, pathB
    # 读取源 TEMU 文件
    temu_df = pd.read_excel(pathA)

    # 读取目标 货师傅 文件
    hsf_df = pd.read_excel(pathB)

    # 删除货师傅除第一行外的所有值
    hsf_df = hsf_df.iloc[0:0]

    lenth1 = len(temu_df)

    for j in range(len(temu_df)):
        i = j
        # 跳过已发货订单
        if str(temu_df.iloc[i, 1]) != '待发货':
            continue

        # 填写SKU信息
        nums = str(temu_df.iloc[i, 3])
        sku = str(temu_df.iloc[i, 4])
        hsf_sku = sku_list[sku]

        # 创建一个新的行数据并填入temu中对应值
        new_row = {'序号': i, '发货单': temu_df.iloc[i, 0], '收件人': temu_df.iloc[i, 9], '收件地址1': temu_df.iloc[i, 11],
                   '城市': temu_df.iloc[i, 15], '省/州': temu_df.iloc[i, 16], '邮编': temu_df.iloc[i, 17],
                   '电话': temu_df.iloc[i, 10], '签收签名服务': '否', 'SKU信息(编码*数量)': '', '收件地址2': ''}

        str1 = str(temu_df.iloc[i, 0])
        str2 = str(temu_df.iloc[i - 1, 0])

        # 判断是否一人买多个不同的产品
        if str(temu_df.iloc[i, 0]) == str(temu_df.iloc[i - 1, 0]) and i != 0:
            new_row['SKU信息(编码*数量)'] = sku_list[str(temu_df.iloc[i - 1, 4])] + '*' + str(
                temu_df.iloc[i - 1, 3]) + ';' + hsf_sku + '*' + nums
            hsf_df = hsf_df.iloc[0: len(hsf_df) - 1]
            hsf_df = hsf_df._append(new_row, ignore_index=True)
            continue

        # 地址二
        if str(temu_df.iloc[i, 12]) != '--':
            new_row['收件地址2'] = temu_df.iloc[i, 12]

        new_row['SKU信息(编码*数量)'] = hsf_sku + "*" + nums
        # 将新行添加到货师傅
        hsf_df = hsf_df._append(new_row, ignore_index=True)

    # 保存修改后的目标 Excel 文件
    hsf_df.to_excel(pathB, index=False)

    print('转换成功')

    return lenth1


def hsf_to_temu():
    global pathA, pathC, pathD
    temu_df = pd.read_excel(pathA)
    hsf_df = pd.read_excel(pathD)
    fh_df = pd.read_excel(pathC)

    # 删除发货模板除第一行外的所有值
    fh_df = fh_df.iloc[0:0]

    def check_column_B(df, str_to_check):
        if any(df.iloc[:, 1].astype(str).str.contains(str_to_check)):
            return True
        else:
            return False

    for i in range(len(temu_df)):
        new_row = {'订单号': '', '商品SKUID': '', '商品件数': '', '跟踪单号': '', '物流承运商': ''}

        order_number = str(temu_df.iloc[i, 0])

        if check_column_B(hsf_df, order_number):
            new_row['订单号'] = str(temu_df.iloc[i, 0])
            new_row['商品SKUID'] = str(temu_df.iloc[i, 2])
            new_row['商品件数'] = str(temu_df.iloc[i, 3])
            first_row = hsf_df[hsf_df.iloc[:, 1].astype(str).str.contains(order_number)].iloc[0]
            new_row['跟踪单号'] = str(first_row['物流单号'])

            if str(first_row['物流公司']) == 'UPS SLMI-OZ':
                new_row['物流承运商'] = 'UPS-MI'
            elif str(first_row['物流公司']) == 'UPS SLMI-LB':
                new_row['物流承运商'] = 'UPS-MI'
            else:
                new_row['物流承运商'] = 'USPS'
        fh_df = fh_df._append(new_row, ignore_index=True)

    fh_df.to_excel(pathC, index=False)

    print('发货转换成功')

def gui_start():
    init_window = Tk()  # 实例化出一个父窗口
    PORTAL = MY_GUI(init_window)
    # 设置根窗口默认属性
    PORTAL.set_init_window('货师傅发货')

    init_window.mainloop()


gui_start()
