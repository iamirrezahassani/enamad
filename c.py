import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

# تعداد صفحات برای بررسی
TOTAL_PAGES = 6667

# هدر برای شبیه‌سازی مرورگر
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
}

# حلقه برای بررسی هر صفحه
data = []
file_name = 'enamad_data.xlsx'

# ایجاد فایل اکسل در صورت عدم وجود
if not os.path.exists(file_name):
    df = pd.DataFrame(columns=['Rank', 'Domain', 'Business Title', 'Province', 'City'])
    df.to_excel(file_name, index=False)

for page in range(1, TOTAL_PAGES + 1):
    url = f"https://enamad.ir/DomainListForMIMT/Index/{page}"
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()  # بررسی وضعیت پاسخ
        soup = BeautifulSoup(response.text, 'html.parser')

        # پیدا کردن تمام دیوهای با کلاس row
        rows = soup.find_all('div', class_='row')

        for row in rows:
            cols = row.find_all('div', class_='col-sm-12')
            if len(cols) >= 5:  # بررسی تعداد ستون‌ها
                rank = cols[0].text.strip()
                domain = cols[1].find('a').text.strip() if cols[1].find('a') else ""
                business_title = cols[2].text.strip()
                province = cols[3].text.strip()
                city = cols[4].text.strip()

                # افزودن داده‌ها به لیست
                data.append([rank, domain, business_title, province, city])

        print(f"صفحه {page} با موفقیت بررسی شد.")
        time.sleep(2)  # توقف برای جلوگیری از بلاک شدن

        # ذخیره هر 10 صفحه در ادامه فایل قبلی
        if page % 10 == 0:
            df = pd.DataFrame(data, columns=['Rank', 'Domain', 'Business Title', 'Province', 'City'])
            with pd.ExcelWriter(file_name, mode='a', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name='Sheet1', startrow=writer.sheets['Sheet1'].max_row, header=False, index=False)
            print(f"ذخیره داده‌ها تا صفحه {page} انجام شد.")
            data = []  # ریست کردن داده‌ها برای 10 صفحه بعدی

    except Exception as e:
        print(f"خطا در صفحه {page}: {e}")

# ذخیره نهایی برای صفحات باقی‌مانده
if data:
    df = pd.DataFrame(data, columns=['Rank', 'Domain', 'Business Title', 'Province', 'City'])
    with pd.ExcelWriter(file_name, mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='Sheet1', startrow=writer.sheets['Sheet1'].max_row, header=False, index=False)

print("ذخیره اطلاعات به پایان رسید.")
