import re
import random
import string
import requests
import time
import os
import jdatetime
from openpyxl import Workbook
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

def random_string(length: int) -> str:
    """
    Generate a random string of specified length
    @param length: Length of the string to generate
    @return: Random string
    """
    characters = string.ascii_letters + string.digits
    return ''.join(random.choice(characters) for _ in range(length))


def convert_to_api_url(url: str) -> str | None:
    """
    تبدیل آدرس ورودی به فرمت API دیجی‌کالا
    @param url: Input URL
    @return: Final output or None if format is incorrect
    """
    try:
        # حذف query string
        parsed_url = urlparse(url)
        clean_url = parsed_url.path
        query_string = parsed_url.query

        # Remove "https://www.digikala.com" from the path
        clean_url = clean_url.replace(
            "/", "", 1) if clean_url.startswith("/") else clean_url

        parts = [part for part in clean_url.split("/") if part]

        if len(parts) < 2 or parts[0] != "search" or not parts[1].startswith("category-"):
            return None

        category = parts[1].replace("category-", "")
        rch = random_string(16)

        if len(parts) == 2:
            # حالت اول: فقط دسته‌بندی
            api_url = f"https://api.digikala.com/v1/categories/{category}/search/?page=1&_rch={rch}"
        elif len(parts) == 3 and parts[2]:
            # حالت دوم: دسته‌بندی + برند دلخواه
            api_url = f"https://api.digikala.com/v1/categories/{category}/brands/{parts[2]}/search/?page=1&_rch={rch}"
        else:
            return None

        if query_string:
            api_url += f"&{query_string}"

        return api_url

    except Exception:
        return None


def generate_user_agent() -> str:
    """
    Generate a random user agent string
    @return: Random user agent string
    """
    browsers = ['Chrome', 'Firefox', 'Edge', 'Safari', 'Opera']
    os_list = [
        'Windows NT 10.0', 'Windows NT 6.1', 'Macintosh; Intel Mac OS X 10_15_7',
        'X11; Linux x86_64', 'Android 10', 'iPhone; CPU iPhone OS 14_0 like Mac OS X'
    ]

    # Select browser based on probability
    rand = random.random()
    if rand < 0.6:
        browser = 'Chrome'
    elif rand < 0.8:
        browser = 'Firefox'
    else:
        # سایر مرورگرها
        others = ['Edge', 'Safari', 'Opera']
        browser = random.choice(others)

    # Select random OS
    os = random.choice(os_list)

    # Generate version based on browser
    if browser == 'Chrome':
        version = f"Chrome/{random.randint(70, 99)}.0.{random.randint(1000, 4999)}.100"
    elif browser == 'Firefox':
        version = f"Firefox/{random.randint(80, 99)}.0"
    elif browser == 'Edge':
        version = f"Edg/{random.randint(90, 119)}.0.{random.randint(1000, 1999)}.100"
    elif browser == 'Safari':
        version = f"Version/{random.randint(13, 17)}.0 Safari/605.1.15"
    elif browser == 'Opera':
        version = f"OPR/{random.randint(70, 99)}.0.{random.randint(1000, 4999)}.100"
    else:
        version = ''

    return f"Mozilla/5.0 ({os}) AppleWebKit/537.36 (KHTML, like Gecko) {version}"


def fetch_pagination_info(api_url: str) -> dict | None:
    """
    تابعی که درخواست GET به API می‌زند و اطلاعات صفحه‌بندی را برمی‌گرداند
    @param api_url: آدرس API
    @return: دیکشنری شامل current_page و total_pages یا None در صورت خطا
    """
    try:
        # ایجاد هدرهای درخواست
        headers = {
            'User-Agent': generate_user_agent(),
            'Accept': 'application/json',
            'Accept-Language': 'en-US,en;q=0.9,fa;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }

        # ارسال درخواست GET
        response = requests.get(api_url, headers=headers, timeout=10)
        response.raise_for_status()  # بررسی خطاهای HTTP

        # تبدیل پاسخ به JSON
        data = response.json()

        # استخراج اطلاعات صفحه‌بندی
        current_page = data.get('data', {}).get(
            'pager', {}).get('current_page')
        total_pages = data.get('data', {}).get('pager', {}).get('total_pages')

        if current_page is not None and total_pages is not None:
            return {
                'current_page': current_page,
                'total_pages': total_pages
            }
        else:
            print("اطلاعات صفحه‌بندی در پاسخ یافت نشد")
            return None

    except requests.exceptions.RequestException as e:
        print(f"خطا در ارسال درخواست: {e}")
        return None
    except ValueError as e:
        print(f"خطا در پردازش JSON: {e}")
        return None
    except Exception as e:
        print(f"خطای غیرمنتظره: {e}")
        return None


def generate_all_pages_for_fetch_data(api_url: str) -> list[str]:
    """
    تابعی که همه صفحات را برای دریافت داده‌ها پیمایش می‌کند
    @param api_url: آدرس API
    @return: لیست آدرس‌های API همه صفحات
    """
    all_pages = []
    current_page = 1
    total_pages = fetch_pagination_info(api_url)['total_pages']

    while current_page <= total_pages:

        if current_page > 10:
            break

        # تجزیه URL به اجزای مختلف
        parsed_url = urlparse(api_url)

        # تبدیل رشته پارامترها به دیکشنری
        # parse_qs مقادیر را به صورت لیست برمی‌گرداند (مثال: {'page': ['1']})
        query_params = parse_qs(parsed_url.query)

        # تغییر مقدار پارامتر 'page'
        query_params['page'] = [str(current_page)]

        # تبدیل دیکشنری جدید به رشته پارامترها
        new_query_string = urlencode(query_params, doseq=True)

        # ساختن دوباره URL با پارامترهای جدید
        # از متد _replace برای جایگزینی بخش query در URL تجزیه‌شده استفاده می‌کنیم
        new_url_parts = parsed_url._replace(query=new_query_string)

        # تبدیل اجزای جدید به یک رشته URL کامل
        final_url = urlunparse(new_url_parts)
        all_pages.append(final_url)
        current_page += 1
    return all_pages


def fetch_all_pages_data(total_api_urls: list[str]) -> list[dict]:
    """
    تابعی که برای هر آدرس API در لیست، درخواست GET می‌زند و پاسخ JSON را ذخیره می‌کند
    @param total_api_urls: لیست آدرس‌های API
    @return: لیست پاسخ‌های JSON
    """
    user_agent = generate_user_agent()
    all_data = []
    for idx, url in enumerate(total_api_urls, start=1):
        print(f"در حال دریافت اطلاعات صفحه {idx} از {len(total_api_urls)}")
        headers = {
            'User-Agent': user_agent,
            'Accept': 'application/json',
            'Accept-Language': 'en-US,en;q=0.9,fa;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }
        try:
            response = requests.get(url, headers=headers, timeout=60)
            response.raise_for_status()
            data = response.json()
            all_data.append(data)
        except requests.exceptions.RequestException as e:
            print(f"خطا در دریافت صفحه {idx}: {e}")
            all_data.append(None)
        except ValueError as e:
            print(f"خطا در پردازش JSON صفحه {idx}: {e}")
            all_data.append(None)
        time.sleep(1)
    return all_data


def extract_products_info(all_pages_data: list[dict]) -> list[dict]:
    """
    استخراج اطلاعات محصولات از لیست داده‌های صفحات
    @param all_pages_data: لیستی از داده‌های JSON صفحات
    @return: لیستی از دیکشنری‌های اطلاعات محصولات
    """
    products_info = []
    for page_data in all_pages_data:
        if not page_data:
            continue
        if page_data.get('status') != 200:
            continue
        products = page_data.get('data', {}).get('products', [])
        for product in products:
            info = {
                'id': product.get('id'),
                'title_fa': product.get('title_fa'),
                'rating_rate': None,
                'rating_count': None
            }
            rating = product.get('rating', {})
            info['rating_rate'] = rating.get('rate')
            info['rating_count'] = rating.get('count')
            products_info.append(info)
    return products_info


def sort_products_by_rating(products_info: list[dict]) -> list[dict]:
    """
    مرتب‌سازی لیست محصولات بر اساس rating_count و سپس rating_rate به صورت نزولی
    @param products_info: لیست اطلاعات محصولات
    @return: لیست مرتب‌شده
    """
    return sorted(
        products_info,
        key=lambda x: (
            x['rating_count'] if x['rating_count'] is not None else -1,
            x['rating_rate'] if x['rating_rate'] is not None else -1
        ),
        reverse=True
    )


def export_products_to_excel(sorted_products_info: list[dict], api_url: str):
    """
    ذخیره لیست محصولات مرتب‌شده در یک فایل اکسل در پوشه export-data
    نام فایل شامل مقدار کتگوری و تاریخ و ساعت شمسی است
    @param sorted_products_info: لیست محصولات مرتب‌شده
    @param api_url: آدرس API (برای استخراج category و نام فایل)
    """
    # ساخت پوشه اگر وجود ندارد
    export_dir = 'export-data'
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    # استخراج category از api_url
    path = urlparse(api_url).path
    parts = [part for part in path.split('/') if part]
    category = None
    if 'categories' in parts:
        idx = parts.index('categories')
        if idx + 1 < len(parts):
            category = parts[idx + 1]
    if not category:
        category = 'unknown_category'

    # ساخت نام فایل با تاریخ و ساعت شمسی
    now = jdatetime.datetime.now()
    date_str = now.strftime('%Y-%m-%d_%H-%M-%S')
    filename = f"{category}_{date_str}.xlsx"
    filepath = os.path.join(export_dir, filename)

    # ساخت فایل اکسل با openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"

    # نوشتن هدرها
    if sorted_products_info:
        headers = list(sorted_products_info[0].keys())
        headers.append("Link")
        ws.append(headers)
        # نوشتن داده‌ها
        for idx, item in enumerate(sorted_products_info, start=2):
            row = [item.get(h, "") for h in headers if h != "Link"]
            link_url = f"https://www.digikala.com/product/dkp-{item.get('id')}"
            row.append("مشاهده محصول")
            ws.append(row)
            # تنظیم هایپرلینک برای سلول آخر (ستون Link)
            link_cell = ws.cell(row=idx, column=len(headers))
            link_cell.hyperlink = link_url
            link_cell.style = "Hyperlink"
    else:
        ws.append(["No data"])

    wb.save(filepath)
    print(f"فایل اکسل با موفقیت ذخیره شد: {filepath}")


def main():
    url = "https://www.digikala.com/search/category-headphone/miscellaneous/?has_selling_stock=1&sort=21"
    api_url = convert_to_api_url(url)
    print(f"API URL: {api_url}")

    # Test the user agent generation
    user_agent = generate_user_agent()
    print(f"Generated User Agent: {user_agent}")

    # Test the pagination info fetch
    if api_url:
        total_api_urls = generate_all_pages_for_fetch_data(api_url)
        print(f"Total API URLs: {len(total_api_urls)}")
        all_pages_data = fetch_all_pages_data(total_api_urls)
        print(f"Collected {len(all_pages_data)} responses.")
        products_info = extract_products_info(all_pages_data)
        print(f"Collected {len(products_info)} products info.")
        sorted_products_info = sort_products_by_rating(products_info)
        print(f"Sorted products info: {sorted_products_info}")
        export_products_to_excel(sorted_products_info, api_url)


if __name__ == "__main__":
    main()
    print("This code runs when the script is executed directly.")
