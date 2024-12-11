from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

options = Options()

# Service 对象代表 ChromeDriver 进程
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
# driver.implicitly_wait(15)  #隐式等待元素加载（不建议和显式混用）
wait = WebDriverWait(driver, 15) # 显式等待元素加载

driver.get("https://learn.senecapolytechnic.ca/")
print("Browser launched successfully!")

# 定位并点击“OK”按钮
try:
    button = driver.find_element(By.ID, "agree_button") 
    button.click()  # 点击按钮
    print("Clicked the 'OK' button successfully!")
except Exception as e:
    print(f"Failed to click the button: {e}")

# 点击登录页面中的 Login 按钮
try:
    login_button = driver.find_element(By.ID, "bottom_Submit")  
    login_button.click()
    print("Clicked the 'Login' button successfully!")
except Exception as e:
    print(f"Failed to click 'Login' button: {e}")


try:
    email_input = wait.until(EC.element_to_be_clickable((By.ID, "i0116"))) 
    email_input.send_keys("ywang841@myseneca.ca")
    next_button = driver.find_element(By.ID, "idSIButton9")  # 定位 Next 按钮
    next_button.click()
    print("input email successfully!")
except Exception as e:
    print(f"Failed to : {e}")

try:
    password_input = wait.until(EC.element_to_be_clickable((By.ID, "i0118"))) 
    password_input.send_keys("Wyc406/304")
    signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))  # 显式等待元素加载（ID重名）
    signin_button.click()
    print("input password successfully!")
except Exception as e:
    print(f"Failed to : {e}")


time.sleep(3)
# 关闭浏览器
driver.quit()
