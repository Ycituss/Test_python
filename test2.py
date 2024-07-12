import pandas as pd

pathA = "F:/共享/亚马逊/货师傅/发货订单/1111.xlsx"
pathB = "F:/共享/亚马逊/货师傅/发货订单/ImportMyOderD.xlsx"

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
    '2303483163': 'X004869WLB'
}

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
        new_row['SKU信息(编码*数量)'] = sku_list[str(temu_df.iloc[i - 1, 4])] + '*' + str(temu_df.iloc[i - 1, 3]) + ';' + hsf_sku + '*' + nums
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

print('success')
