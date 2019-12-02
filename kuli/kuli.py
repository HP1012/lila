# -*- coding: utf-8 -*-


__authors__ = 'Duy Nguyen'
__contact__ = 'duykn@gcs-vn.com'


import logging
import sys

import utils
from report import Report


def main():
    if len(sys.argv) == 1:
        utils.install()

    report = Report(sys.argv[1])

    utils.check_config()

    info = utils.get_task_info(report.data)
    target = utils.get_deliver_dir(info)

    msg = "\nDo you want to deliver test result to \n{0}".format(target)
    utils.confirm(msg)

    report.update_issue(info)

    report.deliver_result(target)

    report.xlsx_update()

    utils.tmp_print_info(report)


if __name__ == "__main__":
    try:
        utils.initial()
        main()

    except Exception as e:
        logging.exception(e)

    finally:
        utils.close()
