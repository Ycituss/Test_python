import pandas as pd

pathA = "F:/共享/亚马逊/货师傅/发货订单/1111.xlsx"
pathD = "F:/共享/亚马逊/货师傅/发货订单/HSFDaoChu.xlsx"
pathC = "F:/共享/亚马逊/货师傅/发货订单/FaHuo.xlsx"

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

print('转换成功')
