# 爬取周轉率.py
from DrissionPage import ChromiumPage, ChromiumOptions
import pandas as pd
from io import StringIO
import time
import os
from datetime import datetime, timedelta

# 設定 Edge 瀏覽器路徑
path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
ChromiumOptions().set_browser_path(path).save()


def getdahu(setsearchtime):
    # 初始化 ChromiumPage
    page = ChromiumPage()

    # 打開目標網站
    page.get("https://www.esunsec.com.tw/tw-rank/b2brwd/page/rank/priceVolume/0012")

    # 點擊「上市」按鈕
    listed_button = page.ele('xpath://button[text()="上市"]')
    if listed_button:
        listed_button.click()

        # 找到並點擊日期選擇按鈕
        date_button = page.ele('xpath://button[contains(@class, "select-button") and contains(@class, "btn-primary")]')
        if date_button:
            date_button.click()

            # 點擊日期選擇器的下拉按鈕來打開日曆
            calendar_icon = page.ele('xpath://span[contains(@class, "ant-picker-suffix")]')
            if calendar_icon:
                calendar_icon.click()

                # 選擇指定日期
                specific_date = page.ele(f'xpath://td[@title="{setsearchtime}"]')
                if specific_date:
                    specific_date.click()

                # 確認設置日期
                confirm_button = page.ele('xpath://button[contains(text(), "設定")]')
                if confirm_button:
                    confirm_button.click()

    # 手動等待數據加載完成
    time.sleep(5)  # 等待5秒，確保數據加載完成

    data_date_element = page.ele('xpath://div[contains(text(), "資料日期")]')
    if data_date_element:
        data_date_text = data_date_element.text.strip()

        # 嘗試將資料日期格式化為 YYYY-MM-DD 格式
        try:
            setsearchtime = datetime.strptime(data_date_text, "資料日期 : %Y/%m/%d").strftime('%Y-%m-%d')
            print(f"資料日期: {setsearchtime}")
        except ValueError:
            print(f"無法解析資料日期: {data_date_text}")
            return

    # 找到並提取表格元素
    table = page.ele('xpath://table[@id="DataTables_Table_0"]')
    if table:
        # 提取表格的 HTML 內容
        table_html = table.html

        # 使用 pandas 解析表格 HTML
        df = pd.read_html(StringIO(table_html))[0]

        # 檢查子資料夾 "原始輸出" 是否存在，如果不存在則創建它
        output_dir = "原始輸出"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 將數據保存到 CSV 文件，並放在 "原始輸出" 資料夾中
        output_file_path = os.path.join(output_dir, f"output_{setsearchtime}.csv")
        df.to_csv(output_file_path, index=False, encoding='utf-8-sig')
        filter_and_save_df(df, setsearchtime)

    # 最後關閉瀏覽器
    page.close()


def filter_and_save_df(df, setsearchtime):
    print("去除奇怪商品，檢查字串長度，保留長度為4的商品代號")
    print("保留漲跌幅大於等於3%的行")
    print("去除收盤欄位大於150的行")

    # 使用 '(', ')' 將商品欄位分割，並提取分割後的字串
    df['商品代號'] = df['商品'].str.extract(r'\((\w+)\)')

    # 檢查字串長度，保留長度為4的商品代號
    df_filtered = df[df['商品代號'].str.len() == 4].copy()

    # 使用 .loc 來避免 SettingWithCopyWarning
    df_filtered.loc[:, '漲跌幅'] = df_filtered['漲跌幅'].str.replace('%', '').astype(float)

    # 保留 "漲跌幅" 大於等於 3% 的行
    df_filtered = df_filtered[df_filtered['漲跌幅'] >= 3]
    # 去除 "收盤" 欄位大於 150 的行
    df_filtered = df_filtered[df_filtered['收盤'] <= 150]

    # 檢查子資料夾 "過濾輸出" 是否存在，如果不存在則創建它
    output_dir = "過濾輸出"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 將數據保存到 CSV 文件，並放在 "過濾輸出" 資料夾中
    output_file_path = os.path.join(output_dir, f"output_{setsearchtime}.csv")
    df_filtered.to_csv(output_file_path, index=False, encoding='utf-8-sig')

    # 顯示過濾後的數據
    print(df_filtered)


if __name__ == '__main__':
    # 如果需要自動捕捉今天日期，設定預設起始與結束日期
    today = datetime.today()
    #setsearchtime_start = '2024-09-27'
    #setsearchtime_end = '2024-09-27'

    setsearchtime_start = today.strftime('%Y-%m-%d') if 'setsearchtime_start' not in locals() else setsearchtime_start
    setsearchtime_end = today.strftime('%Y-%m-%d') if 'setsearchtime_end' not in locals() else setsearchtime_end

    start_date = datetime.strptime(setsearchtime_start, '%Y-%m-%d')
    end_date = datetime.strptime(setsearchtime_end, '%Y-%m-%d')

    current_date = start_date
    while current_date <= end_date:
        setsearchtime = current_date.strftime('%Y-%m-%d')
        getdahu(setsearchtime)
        current_date += timedelta(days=1)
        time.sleep(1)  # 確保數據加載完成
