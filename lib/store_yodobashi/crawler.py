#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ヨドバシ.com から販売履歴や購入履歴を収集します．

Usage:
  crawler.py [-c CONFIG] [-y YEAR] [-n ORDER_NO]
  crawler.py [-c CONFIG] -n ORDER_NO

Options:
  -c CONFIG     : CONFIG を設定ファイルとして読み込んで実行します．[default: config.yaml]
  -y YEAR       : 購入年．
  -n ORDER_NO   : 注文番号．
"""

import logging
import random
import math
import re
import datetime
import time
import traceback

from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import store_yodobashi.const
import store_yodobashi.handle

import local_lib.captcha
import local_lib.selenium_util

STATUS_ORDER_COUNT = "[collect] Count of year"
STATUS_ORDER_ITEM_ALL = "[collect] All orders"
STATUS_ORDER_ITEM_BY_YEAR = "[collect] Year {year} orders"

LOGIN_RETRY_COUNT = 2


def wait_for_loading(handle, xpath="//body", sec=1):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    wait.until(EC.visibility_of_all_elements_located((By.XPATH, xpath)))
    time.sleep(sec)


def parse_date(date_text):
    return datetime.datetime.strptime(date_text, "%Y年%m月%d日")


def gen_item_id_from_url(url):
    return re.match(r"https://www.yodobashi.com/product-detail/([^/]+)/", url).group(1)


def gen_item_id_from_thumb_url(url):
    return re.match(r".*/\d+/(\d+)_\d+\.", url).group(1)


def gen_order_url_from_no(no):
    return store_yodobashi.const.ORDER_URL_BY_NO.format(no=no)


def gen_status_label_by_year(year):
    return STATUS_ORDER_ITEM_BY_YEAR.format(year=year)


def visit_url(handle, url, xpath="//body"):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    driver.get(url)
    wait_for_loading(handle, xpath)


def save_thumbnail(handle, item, thumb_url):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    with local_lib.selenium_util.browser_tab(driver, thumb_url):
        png_data = driver.find_element(By.XPATH, "//img").screenshot_as_png

        with open(store_yodobashi.handle.get_thumb_path(handle, item), "wb") as f:
            f.write(png_data)


def fetch_item_detail(handle, item):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    wait_for_loading(handle)

    if local_lib.selenium_util.xpath_exists(driver, '//div[contains(@class, "notFoundMsg")]'):
        logging.info("{name}: 商品ページが削除されています".format(name=item["name"]))
        item["category"] = []
        return

    breadcrumb_list = driver.find_elements(
        By.XPATH,
        '//ul[@itemtype="http://schema.org/BreadcrumbList"]'
        + '/li[@itemtype="http://schema.org/ListItem"]/a[@itemprop="item"]',
    )
    category = list(map(lambda x: x.text, breadcrumb_list))

    category.pop(0)

    item["category"] = category


def parse_item(handle, item_xpath):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    thumb_url = driver.find_element(
        By.XPATH,
        item_xpath + '//td[contains(@class, "ecImgArea")]//img',
    ).get_attribute("src")

    if local_lib.selenium_util.xpath_exists(
        driver, item_xpath + '//td[contains(@class, "ecPriceArea")]/preceding-sibling::td/p/a'
    ):
        title = driver.find_element(
            By.XPATH, item_xpath + '//td[contains(@class, "ecPriceArea")]/preceding-sibling::td/p/a'
        )

        name = title.text.replace("\n", " ")
        url = title.get_attribute("href")
        item_id = gen_item_id_from_url(url)
    else:
        title = driver.find_element(
            By.XPATH, item_xpath + '//td[contains(@class, "ecPriceArea")]/preceding-sibling::td/p/strong'
        )
        name = title.text.replace("\n", " ")
        url = None
        item_id = gen_item_id_from_thumb_url(thumb_url)

    if local_lib.selenium_util.xpath_exists(
        driver, item_xpath + '//p/strong[contains(@class, "red")]/span[contains(text(), "キャンセル")]'
    ):
        return {"name": name, "cancel": True}

    price_text = driver.find_element(By.XPATH, item_xpath + '//td[contains(@class, "ecPriceArea")]/p').text
    price = int(re.match(r".*?(\d{1,3}(?:,\d{3})*)", price_text).group(1).replace(",", ""))

    count = int(
        driver.find_element(By.XPATH, item_xpath + '//td[contains(@class, "ecQuantityArea")]/span').text
    )

    item = {"name": name, "price": price, "count": count, "url": url, "id": item_id}

    save_thumbnail(handle, item, thumb_url)

    if item["url"] is not None:
        ActionChains(driver).key_down(Keys.COMMAND).click(title).key_up(Keys.COMMAND).perform()
        driver.switch_to.window(driver.window_handles[-1])
        fetch_item_detail(handle, item)
        driver.close()
        driver.switch_to.window(driver.window_handles[-1])
    else:
        logging.info("{name}: 商品ページが削除されています".format(name=item["name"]))
        item["category"] = []

    return item


def parse_order(handle, order_info):
    ITEM_XPATH = '//div[contains(@class, "orderDetailBlock")]'

    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    logging.info(
        "Parse order: {date} - {no}".format(
            date=order_info["date"].strftime("%Y-%m-%d"),
            no=order_info["no"],
        )
    )

    date_text = driver.find_element(
        By.XPATH,
        '//div[contains(@class, "ecOderStatus")]//li/strong[contains(text(), "注文日")]/..',
    ).text.split("：")[1]
    date = parse_date(date_text)

    no = driver.find_element(
        By.XPATH,
        '//div[contains(@class, "ecOderStatus")]//li/strong[contains(text(), "注文番号")]/..',
    ).text.split("：")[1]

    item_base = {"date": date, "no": no}

    is_unempty = False
    for i in range(len(driver.find_elements(By.XPATH, ITEM_XPATH))):
        item_xpath = "(" + ITEM_XPATH + ")[{index}]".format(index=i + 1)

        item = parse_item(handle, item_xpath)
        item |= item_base

        if "cancel" not in item:
            logging.info("{name} {price:,}円".format(name=item["name"], price=item["price"]))
            store_yodobashi.handle.record_item(handle, item)
        else:
            logging.info("{name}: キャンセルされました".format(name=item["name"]))

        is_unempty = True

    return is_unempty


def fetch_order_item_list_by_order_info(handle, order_info):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    with local_lib.selenium_util.browser_tab(driver, store_yodobashi.const.HIST_URL):
        keep_logged_on(handle)

        driver.find_element(By.XPATH, '//input[@id="orderNo"]').send_keys(order_info["no"])

        driver.find_element(
            By.XPATH, '//div[contains(@class, "piKwIpt")]//span[contains(@class, "yBtnInner")]/a'
        ).click()

        time.sleep(0.5)

        if not parse_order(handle, order_info):
            logging.warning("Failed to parse order of {no}".format(no=order_info["no"]))

    time.sleep(5)


def skip_order_item_list_by_year_page(handle, year, page):
    logging.info("Skip check order of {year} page {page} [cached]".format(year=year, page=page))
    incr_order = min(
        store_yodobashi.handle.get_order_count(handle, year)
        - store_yodobashi.handle.get_progress_bar(handle, gen_status_label_by_year(year)).count,
        store_yodobashi.const.ORDER_COUNT_PER_PAGE,
    )
    store_yodobashi.handle.get_progress_bar(handle, gen_status_label_by_year(year)).update(incr_order)
    store_yodobashi.handle.get_progress_bar(handle, STATUS_ORDER_ITEM_ALL).update(incr_order)

    return incr_order != store_yodobashi.const.ORDER_COUNT_PER_PAGE


def fetch_order_item_list_by_year_page(handle, year, page):
    ORDER_XPATH = '//div[contains(@class, "ecContainer")]/div[contains(@class, "orderList")]'

    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    total_page = math.ceil(
        store_yodobashi.handle.get_order_count(handle, year) / store_yodobashi.const.ORDER_COUNT_PER_PAGE
    )

    store_yodobashi.handle.set_status(
        handle,
        "注文履歴を解析しています... {year}年 {page}/{total_page} ページ".format(year=year, page=page, total_page=total_page),
    )

    logging.info(
        "Check order of {year} page {page}/{total_page}".format(year=year, page=page, total_page=total_page)
    )

    order_list = []
    for i in range(len(driver.find_elements(By.XPATH, ORDER_XPATH))):
        order_xpath = "(" + ORDER_XPATH + ")[{index}]".format(index=i + 1)

        date_text = driver.find_element(
            By.XPATH,
            order_xpath
            + '//ul[contains(@class, "hznList")]/li/strong[contains(text(), "注文日")]/following-sibling::span',
        ).text
        date = parse_date(date_text)

        no = driver.find_element(
            By.XPATH,
            order_xpath
            + '//ul[contains(@class, "hznList")]/li/strong[contains(text(), "注文番号")]/following-sibling::span',
        ).text

        order_list.append({"date": date, "no": no})

    for order_info in order_list:
        if not store_yodobashi.handle.get_order_stat(handle, order_info["no"]):
            fetch_order_item_list_by_order_info(handle, order_info)
            store_yodobashi.handle.store_order_info(handle)
        else:
            logging.info(
                "Done order: {date} - {no} [cached]".format(
                    date=order_info["date"].strftime("%Y-%m-%d"), no=order_info["no"]
                )
            )

        store_yodobashi.handle.get_progress_bar(handle, gen_status_label_by_year(year)).update()
        store_yodobashi.handle.get_progress_bar(handle, STATUS_ORDER_ITEM_ALL).update()

        if year == datetime.datetime.now().year:
            last_item = store_yodobashi.handle.get_last_item(handle, year)
            if (
                store_yodobashi.handle.get_year_checked(handle, year)
                and (last_item != None)
                and (last_item["no"] == order_info["no"])
            ):
                logging.info("Latest order found, skipping analysis of subsequent pages")
                for i in range(total_page):
                    store_yodobashi.handle.set_page_checked(handle, year, i + 1)

        time.sleep(3)

    return page == total_page


def fetch_order_item_list_by_year(handle, year):
    visit_order_list_by_year_page(handle, year)

    keep_logged_on(handle)

    year_list = store_yodobashi.handle.get_year_list(handle)

    logging.info(
        "Check order of {year} ({year_index}/{total_year})".format(
            year=year, year_index=year_list.index(year) + 1, total_year=len(year_list)
        )
    )

    store_yodobashi.handle.set_progress_bar(
        handle,
        gen_status_label_by_year(year),
        store_yodobashi.handle.get_order_count(handle, year),
    )

    page = 1
    while True:
        visit_order_list_by_year_page(handle, year, page)

        if not store_yodobashi.handle.get_page_checked(handle, year, page):
            is_last = fetch_order_item_list_by_year_page(handle, year, page)
            store_yodobashi.handle.set_page_checked(handle, year, page)
        else:
            is_last = skip_order_item_list_by_year_page(handle, year, page)

        store_yodobashi.handle.store_order_info(handle)

        if is_last:
            break

        page += 1

    store_yodobashi.handle.get_progress_bar(handle, gen_status_label_by_year(year)).update()
    store_yodobashi.handle.set_year_checked(handle, year)


def fetch_year_list(handle):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    visit_url(handle, store_yodobashi.const.HIST_URL)

    keep_logged_on(handle)

    year_list = list(
        sorted(
            map(
                lambda elem: int(elem.get_attribute("value")),
                driver.find_elements(
                    By.XPATH, '//select[@id="selectedPeriod"]/option[contains(@value, "20")]'
                ),
            )
        )
    )

    logging.info(year_list)

    store_yodobashi.handle.set_year_list(handle, year_list)

    return year_list


def visit_order_list_by_year_page(handle, year, page=1):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    driver.find_element(
        By.XPATH, '//select[@id="selectedPeriod"]/option[contains(@value, {year})]'.format(year=year)
    ).click()

    driver.find_element(
        By.XPATH,
        '//div[contains(@class, "ecHisOderHead")]//span[contains(@class, "yBtnInner")]/a[contains(text(), "検索")]',
    ).click()

    time.sleep(0.2)

    current_page = 1
    while current_page < page:
        driver.find_element(
            By.XPATH, '//ul[contains(@class, "hznList")]/li/a[span[contains(text(), "次のページ")]]'
        ).click()

        time.sleep(1)
        current_page += 1


def fetch_order_count_by_year(handle, year):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    store_yodobashi.handle.set_status(handle, "注文件数を調べています... {year}年".format(year=year))

    visit_order_list_by_year_page(handle, year)

    return int(
        driver.find_element(
            By.XPATH, '//div[contains(@class, "ecContainer")]/p/strong/span[contains(@class, "red")][1]'
        ).text
    )


def fetch_order_count(handle):
    year_list = store_yodobashi.handle.get_year_list(handle)

    logging.info("Collect order count")

    store_yodobashi.handle.set_progress_bar(handle, STATUS_ORDER_COUNT, len(year_list))

    total_count = 0
    for year in year_list:
        if year >= store_yodobashi.handle.get_cache_last_modified(handle).year:
            count = fetch_order_count_by_year(handle, year)
            store_yodobashi.handle.set_order_count(handle, year, count)
            logging.info("Year {year}: {count:4,} orders".format(year=year, count=count))
        else:
            count = store_yodobashi.handle.get_order_count(handle, year)
            logging.info("Year {year}: {count:4,} orders [cached]".format(year=year, count=count))

        total_count += count
        store_yodobashi.handle.get_progress_bar(handle, STATUS_ORDER_COUNT).update()

    logging.info("Total order is {total_count:,}".format(total_count=total_count))

    store_yodobashi.handle.get_progress_bar(handle, STATUS_ORDER_COUNT).update()
    store_yodobashi.handle.store_order_info(handle)


def fetch_order_item_list_all_year(handle):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    year_list = fetch_year_list(handle)
    fetch_order_count(handle)

    store_yodobashi.handle.set_progress_bar(
        handle, STATUS_ORDER_ITEM_ALL, store_yodobashi.handle.get_total_order_count(handle)
    )

    for year in year_list:
        if (
            (year == datetime.datetime.now().year)
            or (year == store_yodobashi.handle.get_cache_last_modified(handle).year)
            or (not store_yodobashi.handle.get_year_checked(handle, year))
        ):
            fetch_order_item_list_by_year(handle, year)
        else:
            logging.info(
                "Done order of {year} ({year_index}/{total_year}) [cached]".format(
                    year=year, year_index=year_list.index(year) + 1, total_year=len(year_list)
                )
            )
            store_yodobashi.handle.get_progress_bar(handle, STATUS_ORDER_ITEM_ALL).update(
                store_yodobashi.handle.get_order_count(handle, year)
            )

    store_yodobashi.handle.get_progress_bar(handle, STATUS_ORDER_ITEM_ALL).update()


def fetch_order_item_list(handle):
    store_yodobashi.handle.set_status(handle, "巡回ロボットの準備をします...")
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    store_yodobashi.handle.set_status(handle, "ウォームアップを行います...")
    warm_up(handle)

    store_yodobashi.handle.set_status(handle, "注文履歴の収集を開始します...")

    try:
        fetch_order_item_list_all_year(handle)
    except:
        local_lib.selenium_util.dump_page(
            driver, int(random.random() * 100), store_yodobashi.handle.get_debug_dir_path(handle)
        )
        raise

    store_yodobashi.handle.set_status(handle, "注文履歴の収集が完了しました．")


def warm_up(handle):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    # NOTE: ダメもとで，Akamai の EdgeSuite を翻弄してみる．

    logging.info("Dummy access to google.com")
    visit_url(handle, "https://www.google.com/")

    driver.find_element(By.XPATH, '//textarea[@title="検索"]').send_keys("ヨドバシ.com")
    driver.find_element(By.XPATH, '//textarea[@title="検索"]').send_keys(Keys.ENTER)

    time.sleep(1)

    driver.find_element(By.XPATH, '//a[contains(@href, "yodobashi.com")]').click()

    time.sleep(1)

    driver.back()

    time.sleep(1)

    logging.info("Dummy access to yahoo.co.jp")
    visit_url(handle, "https://www.yahoo.co.jp/")

    driver.find_element(By.XPATH, '//input[@name="p"]').send_keys("ヨドバシ.com")
    driver.find_element(By.XPATH, '//input[@name="p"]').send_keys(Keys.ENTER)

    time.sleep(1)

    driver.find_element(By.XPATH, '//a[contains(@href, "yodobashi.com")]').click()

    time.sleep(1)


def execute_login(handle):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    if local_lib.selenium_util.xpath_exists(driver, '//input[@id="memberId"]'):
        driver.find_element(By.XPATH, '//input[@id="memberId"]').send_keys(
            store_yodobashi.handle.get_login_user(handle)
        )
    driver.find_element(By.XPATH, '//input[@id="password"]').send_keys(
        store_yodobashi.handle.get_login_pass(handle)
    )

    local_lib.selenium_util.click_xpath(
        driver, '//div[contains(@class, "strcBtn30")]/a[span/strong[contains(text(), "ログイン")]]'
    )

    time.sleep(3)


def keep_logged_on(handle):
    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    wait_for_loading(handle)

    if not local_lib.selenium_util.xpath_exists(driver, '//div[contains(@class, "ecLogin")]'):
        return

    logging.info("Try to login")

    for i in range(LOGIN_RETRY_COUNT):
        if i != 0:
            logging.info("Retry to login")

        execute_login(handle)

        wait_for_loading(handle)

        if local_lib.selenium_util.xpath_exists(driver, '//h1[contains(text(), "Access Denied")]'):
            raise Exception("ロボットによるアクセスと判断され，ログインできませんでした．")
        if not local_lib.selenium_util.xpath_exists(driver, '//div[contains(@class, "ecLogin")]'):
            return

        logging.warning("Failed to login")

        local_lib.selenium_util.dump_page(
            driver,
            int(random.random() * 100),
            store_yodobashi.handle.get_debug_dir_path(handle),
        )

    logging.error("Give up to login")
    raise Exception("ログインに失敗しました．")


if __name__ == "__main__":
    from docopt import docopt

    import local_lib.logger
    import local_lib.config

    args = docopt(__doc__)

    local_lib.logger.init("test", level=logging.INFO)

    config = local_lib.config.load(args["-c"])
    handle = store_yodobashi.handle.create(config)

    driver, wait = store_yodobashi.handle.get_selenium_driver(handle)

    try:
        if args["-n"] is not None:
            no = args["-n"]
            fetch_year_list(handle)

            fetch_order_item_list_by_order_info(handle, {"date": datetime.datetime.now(), "no": no})
        else:
            fetch_order_item_list(handle)
    except:
        driver, wait = store_yodobashi.handle.get_selenium_driver(handle)
        logging.error(traceback.format_exc())

        local_lib.selenium_util.dump_page(
            driver,
            int(random.random() * 100),
            store_yodobashi.handle.get_debug_dir_path(handle),
        )
