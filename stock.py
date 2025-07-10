from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, WebDriverException, InvalidSessionIdException
import time
import os
import glob
import sys

def download_twse_csv_selenium(url, download_dir=None, retry_count=3):
    """
    使用Selenium自動化下載證交所CSV檔案 (改良版)
    
    Args:
        url: 目標網頁URL
        download_dir: 下載目錄，如果不指定則使用當前目錄
        retry_count: 重試次數
    """
    
    # 設置下載目錄
    if download_dir is None:
        download_dir = os.getcwd()
    
    # 確保下載目錄存在
    os.makedirs(download_dir, exist_ok=True)
    
    for attempt in range(retry_count):
        print(f"第 {attempt + 1} 次嘗試...")
        driver = None
        
        try:
            # 配置Chrome選項 (更保守的設置)
            chrome_options = Options()
            
            # 基本設置
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-plugins")
            chrome_options.add_argument("--disable-images")
            chrome_options.add_argument("--disable-javascript")  # 暫時禁用JS來避免衝突
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--headless")
            
            # 下載設置
            prefs = {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": False,
                "safebrowsing.disable_download_protection": True,
                "profile.default_content_setting_values.notifications": 2,
                "profile.default_content_settings.popups": 0,
                "profile.managed_default_content_settings.images": 2
            }
            chrome_options.add_experimental_option("prefs", prefs)
            
            # 反偵測設置
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # 初始化瀏覽器
            print("正在啟動瀏覽器...")
            driver = webdriver.Chrome(options=chrome_options)
            
            # 設置頁面載入超時
            driver.set_page_load_timeout(30)
            driver.implicitly_wait(10)
            
            # 移除webdriver屬性
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            print(f"正在訪問: {url}")
            
            # 訪問目標頁面
            driver.get(url)
            
            # 等待頁面基本載入
            print("等待頁面載入...")
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # 檢查會話是否仍然有效
            try:
                _ = driver.current_url
                print(f"當前頁面: {driver.current_url}")
            except InvalidSessionIdException:
                print("會話已失效，嘗試重新啟動...")
                continue
            
            # 等待更長時間讓頁面完全載入
            print("等待頁面完全載入...")
            time.sleep(5)
            
            # 重新啟用JavaScript (因為需要執行onclick)
            driver.execute_script("document.documentElement.style.display = 'block';")
            
            # 尋找下載按鈕 - 使用更簡單的方法
            print("正在尋找下載按鈕...")
            
            download_button = None
            button_found = False
            
            # 嘗試多種方式尋找按鈕
            selectors = [
                "input[name='download']",
                "input[value='另存CSV']",
                "input[onclick*='show_file']",
                "input[type='button'][value*='CSV']"
            ]
            
            for selector in selectors:
                try:
                    download_button = driver.find_element(By.CSS_SELECTOR, selector)
                    if download_button:
                        print(f"找到下載按鈕: {selector}")
                        button_found = True
                        break
                except:
                    continue
            
            if not button_found:
                print("無法找到下載按鈕，嘗試下一次...")
                continue
            
            # 檢查按鈕是否可見和可點擊
            if not download_button.is_displayed():
                print("按鈕不可見，嘗試滾動到按鈕位置...")
                driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
                time.sleep(2)
            
            # 記錄下載前的檔案
            csv_files_before = glob.glob(os.path.join(download_dir, "*.csv"))
            
            print("正在點擊下載按鈕...")
            
            # 使用JavaScript點擊 (更可靠)
            driver.execute_script("arguments[0].click();", download_button)
            
            # 等待下載完成
            print("等待下載完成...")
            download_completed = False
            max_wait_time = 30
            
            for i in range(max_wait_time):
                time.sleep(1)
                
                # 檢查會話是否仍然有效
                try:
                    _ = driver.current_url
                except InvalidSessionIdException:
                    print("會話在下載過程中失效")
                    break
                
                # 檢查新檔案
                csv_files_after = glob.glob(os.path.join(download_dir, "*.csv"))
                
                if len(csv_files_after) > len(csv_files_before):
                    new_files = [f for f in csv_files_after if f not in csv_files_before]
                    if new_files:
                        latest_file = max(new_files, key=os.path.getctime)
                        
                        # 檢查檔案是否完整
                        if os.path.getsize(latest_file) > 0:
                            # 等待檔案穩定
                            prev_size = 0
                            stable_count = 0
                            
                            for j in range(5):
                                time.sleep(1)
                                current_size = os.path.getsize(latest_file)
                                
                                if current_size == prev_size and current_size > 0:
                                    stable_count += 1
                                    if stable_count >= 2:
                                        download_completed = True
                                        break
                                else:
                                    stable_count = 0
                                
                                prev_size = current_size
                            
                            if download_completed:
                                print(f"下載完成！檔案: {latest_file}")
                                print(f"檔案大小: {os.path.getsize(latest_file)} bytes")
                                return latest_file
                
                print(f"等待中... ({i+1}/{max_wait_time})")
            
            if not download_completed:
                print("下載超時，嘗試下一次...")
                continue
                
        except TimeoutException:
            print("頁面載入超時，嘗試下一次...")
            continue
        except WebDriverException as e:
            print(f"WebDriver錯誤: {e}")
            continue
        except Exception as e:
            print(f"發生錯誤: {e}")
            continue
        finally:
            # 安全關閉瀏覽器
            if driver:
                try:
                    driver.quit()
                except:
                    pass
                time.sleep(2)  # 等待瀏覽器完全關閉
    
    print(f"嘗試 {retry_count} 次後仍然失敗")
    return None

def main():
    """
    主函數
    """
    url = "https://mopsov.twse.com.tw/nas/t21/sii/t21sc03_111_1_0.html"
    
    print("開始使用Selenium下載CSV檔案...")
    print("=" * 50)
    
    # 執行下載
    downloaded_file = download_twse_csv_selenium(url)
    
    if downloaded_file:
        print("=" * 50)
        print("下載成功！")
        print(f"檔案路徑: {downloaded_file}")
        
        # 顯示檔案前幾行來確認內容
        try:
            with open(downloaded_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()[:5]
                print("\n檔案前5行內容:")
                for i, line in enumerate(lines, 1):
                    print(f"{i}: {line.strip()}")
        except Exception as e:
            print(f"讀取檔案時發生錯誤: {e}")
    else:
        print("=" * 50)
        print("下載失敗！")

if __name__ == "__main__":
    main()