from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

options = Options()
# Service 对象代表 ChromeDriver 进程
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


# 打开 Google 测试
driver.get("https://learn.senecapolytechnic.ca/")
print("Browser launched successfully!")
time.sleep(1)

# 定位并点击“OK”按钮
try:
    button = driver.find_element(By.ID, "agree_button")  # 用 ID 定位按钮
    button.click()  # 点击按钮
    print("Clicked the 'OK' button successfully!")
except Exception as e:
    print(f"Failed to click the button: {e}")
time.sleep(1)

# 点击登录页面中的 Login 按钮
try:
    login_button = driver.find_element(By.ID, "bottom_Submit")  # 替换为实际登录按钮的 ID 或其他属性
    login_button.click()
    print("Clicked the 'Login' button successfully!")
except Exception as e:
    print(f"Failed to click 'Login' button: {e}")
time.sleep(1)


# 关闭浏览器
driver.quit()

'''
# WebSocket Communication in Selenium
- WebSocket is a network protocol for two-way communication between a client and a server.
- In Selenium, WebSocket is used for communication between the Selenium WebDriver and the browser (e.g., Chrome).
- When `webdriver.Chrome()` is called, a WebSocket connection is established between Selenium and ChromeDriver.
- The WebSocket interface enables real-time communication, allowing Selenium to send commands (like opening a page) to the browser and receive responses.

# ChromeDriver and Selenium
- Selenium uses ChromeDriver as a bridge to control the Chrome browser.
- When you initialize `webdriver.Chrome()`, it launches both ChromeDriver and the browser.
- The `DevTools listening on ws://...` message indicates that the WebSocket connection has been successfully established.
- Selenium sends commands to ChromeDriver through this WebSocket interface to perform actions like loading web pages.

# Opening Pages in Selenium
- To open a page, use `driver.get("url")`, which sends a command through the WebSocket to load the specified URL in the browser.
- WebSocket is always active during browser control, facilitating all interactions between Selenium and the browser.

'''