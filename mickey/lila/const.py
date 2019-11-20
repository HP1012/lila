# -*- coding: utf-8 -*-

import logging
from pathlib import Path

NAME = 'Lila'
CODE = 'Mickey'

PORT = 9892

HOME = Path.home().joinpath(NAME)
CONFIG = HOME.joinpath('config.json')
LOGS = HOME.joinpath('logs', 'messages')

DATA = HOME.joinpath('db')
WORKSPACE = DATA.joinpath('workspace.json')
PACKAGE = DATA.joinpath('package.json')
HISTORY = DATA.joinpath('history.json')

KEY = 'GSSIHSqMUt6eEwBmKCj6gJkjZ8Yr5Zq_z54b4eWNOEg='

ASSET = Path(__file__).parent.joinpath('assets')
SETTING = ASSET.joinpath('settings.json')
VBA = ASSET.joinpath('vba.xlsm')
VERSION = ASSET.joinpath('version.json')
SPEC = ASSET.joinpath('spec.xlsx')

WEB = Path(__file__).parent.parent.joinpath('web')
TEMPLATE = WEB.joinpath('templates')

SERVER_JIRA = 'https://gcsjira.cybersoft-vn.com/jira'

SERVER_BUILD = "https://ghosvn.cybersoft-vn.com/emb/19s.hics.1711.0/99_User/DuyNguyen/Mickey"