from pathlib import Path

root = Path(r"C:\Users\ivers\OneDrive\桌面\stock-project")
root.mkdir(parents=True, exist_ok=True)

for year in range(2016, 2026):                      # 2016 ~ 2025
    (root / str(year)).mkdir(exist_ok=True)
    print("✓ 建立 / 已存在：", root / str(year))