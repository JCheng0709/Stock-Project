from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pathlib, time
from dateutil.relativedelta import relativedelta
from datetime import date

def parse_ym(s: str) -> date:
    """'YYYY-MM' 轉 datetime.date（取該月 1 號）"""
    y, m = map(int, s.split("-"))
    return date(y, m, 1)

def ym_iter(start: date, end: date):
    """從 start 跑到 end（含），每次加一個月"""
    curr = start
    while curr <= end:
        yield curr
        curr += relativedelta(months=+1)

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

def ask_range():
    while True:
        raw = input("請輸入年月 (YYYY-MM) 或範圍 (YYYY-MM~YYYY-MM)：").strip()
        try:
            if "~" in raw:
                s, e = map(str.strip, raw.split("~"))
                start, end = parse_ym(s), parse_ym(e)
                if start > end:
                    raise ValueError("起始不得晚於結束")
            else:
                start = end = parse_ym(raw)

            if not (date(2013,1,1) <= start <= date(2025,12,1) and
                    date(2013,1,1) <= end   <= date(2025,12,1)):
                raise ValueError("僅支援 2013-01 ~ 2025-12")
            return start, end
        except Exception as e:
            print("❌", e, "‧ 請重新輸入")


def main():
    """
    主函數
    """
    print("=== 台股月營收資料下載工具 ===")
    print("=" * 50)
    
    # 取得使用者輸入
    start, end = ask_range()
    print(f"將下載 {start:%Y-%m} → {end:%Y-%m}...\n")
    
    # 建立 URL
    success, fail = 0, 0
    for d in ym_iter(start, end):
        url = build_url(d.year, d.month, "sii")
        download_dir = pathlib.Path(f"database/{d.year}")  # 改為相對路徑
        print(f"➜ {d:%Y-%m} ", end="")
        fp = download_monthly_income(url, download_dir)
        if fp:
            print("✔")
            success += 1
        else:
            print("✗")
            fail += 1
    print(f"\n=== 完成：成功 {success} 個月份，失敗 {fail} 個月份 ===")

if __name__ == "__main__":
    main()