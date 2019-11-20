# -*- coding: utf-8 -*-

import logging
from pathlib import Path

import win32com.client

import lila.const as CONST

logger = logging.getLogger(__name__)


class ExcelWin32(object):

    def __init__(self, xlsx):
        self.name = Path(xlsx).name
        self.excel = win32com.client.Dispatch('Excel.Application')
        # self.excel.Visible = False
        self.vba = self.excel.Workbooks.Open(CONST.VBA)
        self.wb = self.excel.Workbooks.Open(xlsx)

    def run(self, sheet, cmd, params):
        '''Execute macro command'''
        logger.debug("Execute macro %s %s", cmd, sheet)
        try:
            self.wb.Worksheets(sheet).Activate()
            cmd = '{0}!cmd.{1}'.format(Path(CONST.VBA).name, cmd)
            rst = self.excel.Application.Run(cmd, *tuple(params))
        except Exception as e:
            logger.exception(e)
            rst = "Exception {0}".format(str(e))
        finally:
            self.wb.Save()
            logger.debug("Run status %s", rst)
            return rst

    def close(self):
        '''Close all openfile'''
        logger.debug("Close excel file %s", self.name)
        try:
            self.vba.Close()
            self.wb.Close()
        except Exception as e:
            logger.exception(e)
