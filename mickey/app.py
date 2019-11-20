# -*- coding: utf-8 -*-

import logging

import eel

import lila.const as CONST
from lila import utils, web

logger = logging.getLogger(__name__)


def start_eel(mode='chrome-app'):
    '''Start Eel'''
    def on_close(page, sockets):
        pass

    eel.init('web')

    options = {
        'mode': mode,
        'host': 'localhost',
        'port': CONST.PORT
    }

    logger.debug("Start app in mode %s", mode)

    try:
        web.generate_html()
        eel.start('index.html', options=options, callback=on_close)

    except Exception as e:
        logger.exception(e)
        if mode == 'chrome-app':
            start_eel('edge')


if __name__ == "__main__":
    try:
        utils.merge_data()

        if utils.is_open_port() is True:
            utils.clean_port()

    except Exception as e:
        logger.exception(e)

    start_eel()
