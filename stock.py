from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pathlib, time, os

def roc(year: int) -> int:
    return year - 1911

def build_url(year, month, market="sii") -> str:
    year_roc = roc(year)
    return f"https://mopsov.twse.com.tw/nas/t21/{market}/t21sc03_{year_roc}_{month}_0.html"

def download_monthly_income(url, target_dir):
    """
    下載月營收資料
    
    Args:
        url: 下載網址
        target_dir: 目標資料夾路徑
    
    Returns:
        pathlib.Path: 下載的檔案路徑，失敗則返回 None
    """
    try:
        # ─── 0. 設定「想存檔」的資料夾 ──────────────────────────────
        target_dir = pathlib.Path(target_dir)
        target_dir.mkdir(parents=True, exist_ok=True)

        # ─── 1. 建立 ChromeOptions，寫入下載偏好 ───────────────────
        chrome_opts = Options()
        prefs = {
            "download.default_directory": str(target_dir.resolve()),
            "download.prompt_for_download": False,
            "safebrowsing.enabled": True
        }
        chrome_opts.add_experimental_option("prefs", prefs)
        chrome_opts.add_argument("--log-level=3")
        chrome_opts.add_argument("--headless=new")  # 無頭模式
        chrome_opts.add_argument("--no-sandbox")
        chrome_opts.add_argument("--disable-dev-shm-usage")

        # ─── 2. 啟動 Driver ───────────────────────────────────────
        # service = Service(ChromeDriverManager().install())
        browser = webdriver.Chrome(options=chrome_opts)

        try:
            # ─── 3. 開頁、點下載 ──────────────────────────────────────
            print(f"正在訪問: {url}")
            browser.get(url)

            # 等待頁面載入並尋找下載按鈕
            wait = WebDriverWait(browser, 10)
            button = wait.until(EC.element_to_be_clickable((By.NAME, "download")))
            
            # 記錄下載前的檔案
            files_before = set(target_dir.glob("*"))
            
            button.click()
            print("已點擊下載按鈕...")

            # ─── 4. 等待檔案下載完成 ──────────────────────────────────
            max_wait_time = 30  # 最大等待30秒
            wait_time = 0
            
            while wait_time < max_wait_time:
                time.sleep(1)
                wait_time += 1
                
                # 檢查是否有新檔案
                files_after = set(target_dir.glob("*"))
                new_files = files_after - files_before
                
                if new_files:
                    # 找到新檔案，檢查是否下載完成（沒有 .crdownload 後綴）
                    for file_path in new_files:
                        if not file_path.name.endswith('.crdownload'):
                            print(f"✅ 檔案下載完成: {file_path.name}")
                            return file_path
                
                if wait_time % 5 == 0:
                    print(f"等待下載中... ({wait_time}s)")
            
            print("⚠️ 下載超時")
            return None
            
        except TimeoutException:
            print("❌ 頁面載入超時或找不到下載按鈕")
            return None
        except NoSuchElementException:
            print("❌ 找不到下載按鈕")
            return None
        finally:
            browser.quit()
            
    except Exception as e:
        print(f"❌ 下載過程發生錯誤: {e}")
        return None

def ask_int(prompt, valid_func):
    while True:
        try:
            value = int(input(prompt))
            if valid_func(value):
                return value
            print("輸入不在允許範圍，請重新輸入 (2013 - 2025)")
        except ValueError:
            print("請輸入整數")

def ask_user():
    ad_year = ask_int("請輸入 * 西元 * 年分 (2013 - 2025): ", lambda y: 2013 <= y <= 2025)
    month = ask_int("請輸入月份: ", lambda m: 1 <= m <= 12)
    return ad_year, month

def main():
    """
    主函數
    """
    print("=== 台股月營收資料下載工具 ===")
    print("=" * 50)
    
    # 取得使用者輸入
    ad_year, month = ask_user()
    
    # 建立 URL
    url = build_url(ad_year, month, "sii")
    print(f"準備下載: {ad_year}年{month}月 的營收資料")
    
    # 設定下載目錄（可以改為相對路徑或讓使用者選擇）
    download_dir = pathlib.Path("downloads")  # 改為相對路徑
    # 或者可以讓使用者選擇：
    # download_dir = pathlib.Path(input("請輸入下載目錄路徑: ") or "downloads")
    
    # 執行下載
    downloaded_file = download_monthly_income(url, download_dir)
    
    if downloaded_file:
        print(f"✅ 下載完成 → {downloaded_file.resolve()}")
        print(f"檔案大小: {downloaded_file.stat().st_size / 1024:.2f} KB")
    else:
        print("❌ 下載失敗")
        
    print("\n目錄內容:")
    for p in download_dir.iterdir():
        print(f" • {p.name}")

if __name__ == "__main__":
    main()
