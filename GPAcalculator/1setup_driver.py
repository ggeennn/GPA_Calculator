from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


# 打开 Google 测试
driver.get("https://www.google.com")
print("Browser launched successfully!")

# 延迟 10 秒，以便查看打开的网页 
time.sleep(7)
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