from __future__ import annotations

import time
import pathlib
from typing import List, Tuple, Optional

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from pathlib import Path
import time

###############################################################################
# 解析輸入格式
###############################################################################

def parse_year_season(text: str) -> Tuple[int, Optional[int]]:
    """'110-01' -> (110, 1);  '110' -> (110, None)"""
    parts = text.split('-', 1)
    year = int(parts[0])
    season = int(parts[1].lstrip('0') or 0) if len(parts) == 2 else None
    return year, season


def ask_range() -> List[Tuple[int, Optional[int]]]:
    """互動輸入年份/範圍，回傳 [(year, season?)]"""
    while True:
        s = input("請輸入年份或年-季，可用 ~ 表示範圍：").strip()
        try:
            if "~" in s:
                left, right = s.split("~", 1)
                y1, q1 = parse_year_season(left)
                y2, q2 = parse_year_season(right)
                if (q1 is None) != (q2 is None):
                    raise ValueError("範圍兩端格式需一致")
                res: List[Tuple[int, Optional[int]]] = []
                for y in range(y1, y2 + 1):
                    if q1 is None:
                        res.append((y, None))
                    else:
                        start_q = q1 if y == y1 else 1
                        end_q = q2 if y == y2 else 4
                        res.extend((y, q) for q in range(start_q, end_q + 1))
                return res
            else:
                return [parse_year_season(s)]
        except ValueError:
            print("❌ 格式錯誤，請重新輸入（例如 113 或 110-01 或 102~109）。")

###############################################################################
# 主下載函式與輔助工具（增：防止 StaleElement 及允許多重下載）
###############################################################################

# -- 1️⃣ ChromeOptions: 允許一次按多顆下載按鈕 ------------------------------
def make_driver(download_dir: pathlib.Path) -> webdriver.Chrome:
    opts = Options()
    prefs = {
        "download.default_directory": str(download_dir.resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        # 允許同源連續自動下載 (Chrome 65+)
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--log-level=3")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=opts)

# -- 2️⃣ 點擊按鈕：自動重新定位，避開 stale element ---------------------------

def safe_click(driver: webdriver.Chrome, locator: Tuple[str, str], retries: int = 3):
    """嘗試重新定位元素並 click，避免 stale element reference"""
    by, sel = locator
    for i in range(retries):
        try:
            elem = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((by, sel))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
            time.sleep(0.3)
            driver.execute_script("arguments[0].click();", elem)
            return True
        except (StaleElementReferenceException, TimeoutException):
            if i == retries - 1:
                return False
            time.sleep(0.5)
    return False

# -- 3️⃣ 下載流程 --------------------------------------------------------------

def download_mops_data(year: int, market: str, season: Optional[int], out_dir: pathlib.Path):
    """簡化版：
    * 開網頁 → 選市場、年份、季別
    * 點查詢 → 進入 pop‑up
    * 依序點擊所有下載按鈕，檔名 market_year_Qx_idx.csv
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    browser = make_driver(out_dir)
    wait = WebDriverWait(browser, 20)
    
    try:
        browser.get("https://mops.twse.com.tw/mops/#/web/t163sb19")

        # ======== 填表單 ========
        sel_elem = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, "TYPEK"))
        )

        # 4️⃣  用 Select 操作市場別
        Select(sel_elem).select_by_value(market)   # sii / otc / rotc / pub


        wait.until(lambda d: d.find_element(By.NAME, "year")).send_keys(str(year))

        if season:
            sel = Select(browser.find_element(By.ID, "season"))  # <select id="season">
            sel.select_by_value(f"{season:02d}") 

        submit = browser.find_element(By.ID, "searchBtn")
        submit.click()
        print("已點擊提交按鈕")
        # safe_click(browser, (By.XPATH, "//input[@value='查詢']"))

        # ======== pop‑up ========
        wait.until(lambda d: len(d.window_handles) == 2)
        main_win, popup_win = browser.window_handles
        browser.switch_to.window(popup_win)

        # 下載按鈕們
        seen_filename : set[str] = set()
        buttons = browser.find_elements(By.CSS_SELECTOR, "button[onclick*='t105sb02']")
        # buttons = browser.find_elements(*buttons_locator)
        files_before = set(out_dir.iterdir())
        ok_cnt = 0
        for idx, btn in enumerate(buttons, 1):
            # 2-1 讀該按鈕所在 form 的隱藏欄位
            form = btn.find_element(By.XPATH, "./ancestor::form")
            # file_path = form.find_element(By.NAME, "filePath").get_attribute("value")
            file_name = form.find_element(By.NAME, "filename").get_attribute("value")

            if file_name in seen_filename:
                print(f"{file_name} 已經下載過")
                continue
            seen_filename.add(file_name)

            new_name = f"{market}_{year}_Q{season or 'all'}_{idx}.csv"

            # 2-2 點這顆實體 btn（不再用 locator）
            if not safe_click_elem(browser, btn):
                print(f"✗ 按鈕 {idx} 點擊失敗")
                continue

            # 2-3 等檔案寫完、重新命名
            if wait_for_download(out_dir, files_before, new_name):
                ok_cnt += 1
            files_before = set(out_dir.iterdir())
        return ok_cnt > 0

    finally:
        browser.quit()

def safe_click_elem(driver, elem, retries=3):
    for _ in range(retries):
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
            time.sleep(0.2)
            driver.execute_script("arguments[0].click();", elem)
            return True
        except StaleElementReferenceException:
            time.sleep(0.5)
    return False


def wait_for_download(target_dir: Path,
                      files_before: set[Path],
                      new_filename: str,
                      max_wait: int = 120) -> Path | None:
    """
    等下載完成並重新命名。若新檔名已存在則直接略過。
    """
    target_dir = Path(target_dir)
    new_path   = target_dir / new_filename

    # ❶ 檔案已存在 → 直接跳過
    if new_path.exists():
        print(f"⚠ {new_filename} 已存在，跳過下載。")
        return new_path                       # 或 return None 取決於你後續邏輯

    print(f"等待下載完成... (最多 {max_wait} 秒)")
    for i in range(max_wait):
        time.sleep(1)

        # ❷ 檢查有無新檔
        new_files = set(target_dir.glob("*")) - files_before
        for file_path in new_files:
            if file_path.suffix in {".crdownload", ".tmp", ".partial"}:
                continue                      # 還沒寫完
            if file_path.stat().st_size == 0:
                continue
            # ❸ 如目標檔仍存在衝突，就在尾巴加 (_1), (_2)…
            if new_path.exists():
                new_path = _auto_rename(new_path)
            file_path.rename(new_path)
            print(f"✅ 已重新命名為: {new_path.name}")
            return new_path

        if (i + 1) % 10 == 0:
            print(f"等待中... ({i + 1}/{max_wait}s)")

    print("❌ 下載超時")
    return None


def _auto_rename(path: Path) -> Path:
    """若 path 已存在，尾端自動加 (_1)…"""
    stem, suffix = path.stem, path.suffix
    for n in range(1, 100):
        candidate = path.with_name(f"{stem}({n}){suffix}")
        if not candidate.exists():
            return candidate
    raise FileExistsError("Too many name conflicts")

# ... 〈此處保留現有 download_mops_data / helper 函式原樣，可改型別註解〉 ...
# 由於篇幅限制，不貼重複程式；僅示範 main() 改動。
###############################################################################

def main():
    print("=== MOPS 公開資訊觀測站下載工具 ===")
    print("=" * 50)

    year_season_list = ask_range()   # List[(year, season?)]
    print(f"將下載 {len(year_season_list)} 個區段...\n")

    market = input("請輸入市場別 (sii=上市, otc=上櫃) [預設: sii]: ").lower() or "sii"
    target_root = pathlib.Path("EPS")

    success = fail = 0
    for year, season in year_season_list:
        dlabel = f"{year} 年" if season is None else f"{year} 年 Q{season}"
        print(f"\n▶ 下載 {dlabel} {'上市' if market=='sii' else '上櫃'}…")

        out_dir = target_root / f"{year}" / (f"Q{season}" if season else "all")
        result = download_mops_data(year, market, season or "all", out_dir)

        if result:
            success += 1
        else:
            fail += 1

    print(f"\n=== 完成：成功 {success} 筆，失敗 {fail} 筆 ===")

if __name__ == "__main__":
    main()