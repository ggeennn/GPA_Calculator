import requests

def send_api_request():
    # API URL
    url = "https://your-api-url.com/endpoint"  # 修改为你的 API URL

    # 请求头信息（根据需要修改）
    headers = {
        'Accept': 'application/json, text/plain, */*',  
        'Accept-Encoding': 'gzip, deflate, br, zstd',        
        'Authorization': 'Bearer your_api_token',      # 如果需要身份验证，填入 token
        'X-Blackboard-XSRF': 'your_XSRF_token',        # 填入 XSRF token
        'Cookie': 'your_cookie_data',                  # 填入 Cookie 数据
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',  # 可根据需求修改
        'Content-Type': 'application/json',            # 可根据需求修改
    }

    try:
        # 发送 GET 请求
        response = requests.get(url, headers=headers)

        # 检查返回状态码
        if response.status_code == 200:
            print("请求成功!")
            print("返回数据:", response.json())  # 如果返回的是 JSON 数据
        else:
            print(f"请求失败! 状态码: {response.status_code}")
            print("错误信息:", response.text)

    except requests.exceptions.RequestException as e:
        print("请求时发生错误:", e)

# 调用函数进行请求
send_api_request()
