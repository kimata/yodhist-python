#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ヨドバシ.com の購入履歴情報を収集して，Excel ファイルとして出力します．

Usage:
  yodhist.py [-c CONFIG] [-e]

Options:
  -c CONFIG    : CONFIG を設定ファイルとして読み込んで実行します．[default: config.yaml]
  -e           : データ収集は行わず，Excel ファイルの出力のみ行います．
"""

import logging
import random

import store_yodobashi.handle
import store_yodobashi.crawler
import store_yodobashi.order_history
import local_lib.selenium_util

NAME = "yodhist"
VERSION = "0.1.0"


def execute_fetch(handle):
    try:
        store_yodobashi.crawler.fetch_order_item_list(handle)
    except:
        driver, wait = store_yodobashi.handle.get_selenium_driver(handle)
        local_lib.selenium_util.dump_page(
            driver, int(random.random() * 100), store_yodobashi.handle.get_debug_dir_path(handle)
        )
        raise


def execute(config, is_export_mode=False):
    handle = store_yodobashi.handle.create(config)

    try:
        if not is_export_mode:
            execute_fetch(handle)
        store_yodobashi.order_history.generate_table_excel(handle, config["output"]["excel"]["table"])

        store_yodobashi.handle.finish(handle)
        input("完了しました．エンターを押すと終了します．")
    except:
        store_yodobashi.handle.set_status(handle, "エラーが発生しました", is_error=True)
        logging.error(traceback.format_exc())


######################################################################
if __name__ == "__main__":
    from docopt import docopt
    import traceback

    import local_lib.logger
    import local_lib.config
    import store_yodobashi.handle

    args = docopt(__doc__)

    local_lib.logger.init("yodhist", level=logging.INFO)

    config_file = args["-c"]
    is_export_mode = args["-e"]

    config = local_lib.config.load(args["-c"])

    execute(config, is_export_mode)
