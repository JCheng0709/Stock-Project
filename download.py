from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options          ### ← 新增
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pathlib, time                                           ### ← 新增
import os

# ─── 0. 設定「想存檔」的資料夾 ──────────────────────────────
target_dir = pathlib.Path(r"C:\Users\ivers\OneDrive\桌面\stock-project\database")    # ← 改成你的路徑
target_dir.mkdir(parents=True, exist_ok=True)                      # 若不存在就建立

# ─── 1. 建立 ChromeOptions，寫入下載偏好 ───────────────────
chrome_opts = Options()
prefs = {
    "download.default_directory": str(target_dir),   # 指定下載路徑
    "download.prompt_for_download": False,           # 不彈下載對話框
    "safebrowsing.enabled": True                     # 避免被阻擋
}
chrome_opts.add_experimental_option("prefs", prefs)
chrome_opts.add_argument("--log-level=3")            # 關閉雜訊日誌（可選）
# chrome_opts.add_argument("--headless=new")         # 要無頭可取消註解

# ─── 2. 啟動 Driver ───────────────────────────────────────
service  = Service(ChromeDriverManager().install())
browser  = webdriver.Chrome(service=service, options=chrome_opts)  ### ← 改

# ─── 3. 開頁、點下載 ──────────────────────────────────────
downloadURL = 'https://mopsov.twse.com.tw/nas/t21/sii/t21sc03_111_1_0.html'
browser.get(downloadURL)

button = browser.find_element(By.NAME, "download")
button.click()                                        # 下載 CSV

# ─── 4. 粗略等等檔案寫完（可用更嚴謹的輪詢） ──────────────
time.sleep(2)                                         # 視網速決定秒數

print("目錄內容：")
for p in target_dir.iterdir():
    print(" •", p.name)

# （如需後續解析）
# soup = BeautifulSoup(browser.page_source, 'lxml')

browser.quit()
