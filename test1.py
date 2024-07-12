import requests

# 定义请求的 URL
url = 'http://www.kodatrading.com/Home/Api/GetWarehouse'

# 定义要发送的数据
data = {
    'PageSize': '10',
    'PageNumber': '1'
}

# 发送 POST 请求
headers = {'Content-Type': 'application/x-www-form-urlencoded'}
response = requests.post(url, data=data, headers=headers)

# 检查响应状态码
if response.status_code == 200:
    # 解析 JSON 响应数据
    try:
        json_response = response.json()
        print(json_response)
    except requests.exceptions.JSONDecodeError:
        print("响应不是有效的 JSON 格式")
else:
    print(f"请求失败，状态码: {response.status_code}")