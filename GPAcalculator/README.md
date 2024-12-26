# 环境配置步骤
1. 安装 Selenium 和 WebDriver Manager:
   ```bash
   pip install selenium webdriver-manager
   
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from collections import defaultdict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# 设置环境变量
在终端中输入以下命令，将 your_email 和 your_password 替换为你的实际邮箱和密码：
## Linux/macOS: 
	export SENECA_EMAIL="your_email"
	export SENECA_PASSWORD="your_password"
## Windows: 
	set SENECA_EMAIL=your_email
	set SENECA_PASSWORD=your_password

# 运行程序
semester 1:		s1.py
