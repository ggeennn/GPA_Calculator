from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
#from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time

options = Options()
'''无法自动填充！！！
options.add_argument("--disable-blink-features=AutomationControlled")  # 防止Chrome的自动化标志
options.add_argument("--disable-save-password-bubble")  # 禁用保存密码的提示
options.add_argument("--autofill")  # 启用自动填充功能
'''
options.add_argument("--disable-blink-features=AutomationControlled")  # 禁用 Chrome 自动化标志
options.add_argument("--start-maximized")  # 启动时最大化浏览器

# Service 对象代表 ChromeDriver 进程
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get("https://learn.senecapolytechnic.ca/")
print("Browser launched successfully!")

# driver.implicitly_wait(15)  #隐式等待元素加载（不建议和显式混用）
wait = WebDriverWait(driver, 15) # 显式等待元素加载
success = False
# 你可以设置最大尝试次数，以避免死循环
max_attempts = 5
attempt = 0

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

try:
    # 等待并定位复选框
    checkbox = WebDriverWait(driver, 25).until(EC.element_to_be_clickable((By.ID, "KmsiCheckboxField")))
    if not checkbox.is_selected():  # 如果复选框未被选中
        checkbox.click()
        print("Checkbox selected successfully!")
    else:
        print("Checkbox was already selected!")

    # 等待并点击 "Yes" 按钮
    yes_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    yes_button.click()
    print("Clicked 'Yes' button successfully!")
except Exception as e:
    print(f"Error during checkbox selection or 'Yes' button click: {e}")

while not success and attempt < max_attempts:
    try:
        time.sleep(1)  # 可以保留一个短暂的等待，避免快速重试过于频繁
        grades_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//bb-base-navigation-button[8]//a')))
        grades_button.click()
        print("'Grades' clicked successfully!")
        success = True  # 成功点击后，退出循环
    except StaleElementReferenceException as e:
        print(f"Stale element found. Retrying... (Attempt {attempt + 1})")
        attempt += 1
    except Exception as e:
        print(f"Failed to click 'Grades': {e}")
        break  # 其他错误时退出循环

try:
    time.sleep(1)  # 等待内容加载
    main_container = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main-content-inner"]')))
    print("main_container located successfully.")
    # 滚动页面，确保懒加载的内容加载完成
    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", main_container)
except Exception as e:
    print(f"Failed to : {e}")

try: 
    #无法找到该元素？？
    IPC_all = wait.until(EC.element_to_be_clickable((By.XPATH, '(//*[@id="card__733205_1"]/div/div/div[2]/div[2]/div/a)[2]')))
    driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", IPC_all)
    time.sleep(1)  # 等待滑动
    IPC_all.click()
    print("'IPC' select successfully!")
except Exception as e:
    print(f"Failed to : {e}")

time.sleep(3)

 #关闭浏览器
driver.quit()
