# -*- coding: utf-8 -*-


__authors__ = 'Duy Nguyen'
__contact__ = 'duykn@gcs-vn.com'


import ctypes
import getpass
import json
import logging
import shutil
import sys
import winreg
from pathlib import Path

import requests

import config as CONF

LOG_LEVEL = logging.INFO
FONT_NAME = 'Courier New'


FORMAT_DEBUG = '%(asctime)s [%(funcName)s] %(levelname)-5s %(message)s'
FORMAT_INFO = '%(asctime)s %(levelname)-5s %(message)s'

LF_FACESIZE = 32
STD_OUTPUT_HANDLE = -11


class COORD(ctypes.Structure):
    _fields_ = [("X", ctypes.c_short), ("Y", ctypes.c_short)]


class CONSOLE_FONT_INFOEX(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong),
                ("nFont", ctypes.c_ulong),
                ("dwFontSize", COORD),
                ("FontFamily", ctypes.c_uint),
                ("FontWeight", ctypes.c_uint),
                ("FaceName", ctypes.c_wchar * LF_FACESIZE)]


def initial():
    global LOG_LEVEL, FONT_NAME

    title = "Hi {0}, I'm {1}".format(getpass.getuser(), CONF.NAME)
    config = load_config()
    LOG_LEVEL = config.get('logging', LOG_LEVEL)
    FONT_NAME = config.get('font', FONT_NAME)

    if LOG_LEVEL == 10 and FONT_NAME == 'Courier New':
        FONT_NAME = 'NSimSun'

    set_console_title(title)
    set_console_log(LOG_LEVEL)
    set_consolet_font(FONT_NAME)

    banner = '{name} [Version {version}]\n' \
        'Do one thing and do it well\n' \
        .format(name=CONF.NAME, version=get_version())
    print(banner)


def install():
    '''Install Kuli to home directory'''
    msg = "Do you want to install {0}".format(CONF.NAME)
    confirm(msg)

    # Copy application file
    src = Path.cwd().joinpath(sys.argv[0])
    copy_file(src, CONF.FP_APP)

    # Copy config file
    logging.info("Start installing %s ...", CONF.NAME)
    load_config(CONF.FP_CONFIG, init=True)

    logging.info("Successfully installed")
    logging.info("Thank you for using %s", CONF.NAME)
    close()


def confirm(msg):
    '''Confirm to continue'''
    msg = '{0} (y/n)? '.format(msg)
    option = input(msg)
    if option.lower() not in ['y']:
        close()


def set_console_title(title):
    '''Set console title'''
    ctypes.windll.kernel32.SetConsoleTitleW(title)


def set_consolet_font(font_name=FONT_NAME):
    font = CONSOLE_FONT_INFOEX()
    font.cbSize = ctypes.sizeof(CONSOLE_FONT_INFOEX)
    font.nFont = 12
    font.dwFontSize.X = 11
    font.dwFontSize.Y = 16
    font.FontFamily = 54
    font.FontWeight = 400
    font.FaceName = font_name

    handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    ctypes.windll.kernel32.SetCurrentConsoleFontEx(
        handle, ctypes.c_long(False), ctypes.pointer(font))


def set_console_log(level=LOG_LEVEL):
    format = FORMAT_INFO if level > 10 else FORMAT_DEBUG
    logging.basicConfig(format=format, datefmt='%H:%M:%S', level=level)


def copy_file(src, dst):
    logging.debug("%s -> %s", src, dst)

    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    if Path(dst).absolute() != Path(src).absolute():
        shutil.copy2(src, dst)


def load_config(path=CONF.FP_CONFIG, init=False):
    try:
        config = check_config(path)

    except:
        config = {}
        if init == True:
            copy_file(CONF.FP_CONFIG_DF, CONF.FP_CONFIG)

    return config


def get_version(path=CONF.FP_VERSION):
    '''Get version from git log text file'''
    with open(path, errors='ignores') as fp:
        lines = fp.readlines()

    return '{0}.{1}'.format(float(int(lines[0])/10), lines[1][7:13])


def get_version_latest(api=CONF.API_VERSION):
    '''Get latest version from server'''
    url = get_url(api)
    logging.debug("Getting latest version from %s", url)

    r = requests.get(url)

    return r.json().get('version') if r.status_code == 200 else None


def download_build(path, api=CONF.API_DOWNLOAD):
    '''Download build from server to path'''
    url = get_url(api)
    logging.debug("Downloading %s -> %s", url, path)

    r = requests.get(url)
    if r.status_code == 200:
        with open(path, 'wb') as fp:
            fp.write(r.content)

    return path if r.status_code == 200 else None


def get_url(api):
    '''Get full api link'''
    config = load_config()
    server = config.get('server', CONF.SERVER)
    return 'http://{0}/{1}'.format(server, api)


def upgrade_build():
    '''Upgrade to latest version'''
    try:
        version = get_version_latest()
        if version != get_version() and version is not None:
            logging.debug("Upgrading to new version %s", version)
            path = CONF.FP_FILES.joinpath(version)
            path = download_build(path)
            if path is not None:
                copy_file(path, CONF.FP_APP)
                path.unlink()
    except Exception as e:
        if LOG_LEVEL < 20:
            logging.exception(e)


def close():
    '''Keep console'''
    upgrade_build()

    print('\nPress any key to exit')
    while True:
        input()


def format_string(string, tab=3):
    '''Modify string according to log level'''
    space = '' if tab == 0 else ' '*tab
    if isinstance(string, Path) and LOG_LEVEL == 10:
        return string
    else:
        text = string.name if isinstance(string, Path) else string
        return '{0}{1}'.format(space, text)


def read_file(path):
    with open(path, encoding='shift-jis', errors='ignore') as fp:
        return fp.readlines()


def check_config(path=CONF.FP_CONFIG):
    '''Check config file'''
    def load_json(jpath):
        with open(jpath) as fp:
            return json.load(fp)

    with open(path, encoding='shift-jis', errors='ignore') as fp:
        config = json.load(fp)

    default = load_json(path)
    config = load_json(path)
    [config[k] for k in default.keys()]

    return config


def get_task_info(data, api=CONF.API_TASK):
    '''Get task info from server'''
    logging.debug("Get task info %s", data)
    url = get_url(api)

    try:
        info = {
            'src_rel': data['src_rel'].replace('\\', '/'),
            'func': data['func'],
            'package': data['package'],
            'username': getpass.getuser()
        }

        try:
            r = requests.post(url, json=info)
            if r.status_code == 200:
                info = r.json()
            else:
                error = "Status code {0}".format(r.status_code)
                raise Exception(error)

        except Exception as e:
            error = "Unable to connect to server to get task info"
            logging.error(error)
            raise e

        if 'error' in info.keys():
            logging.error(info)
            raise Exception(info.get('error'))

        for key, value in info.items():
            if value is None:
                error = "Unable to get '{0} info from server".format(key)
                logging.error(error)
                raise Exception(error)

        return info

    except Exception as e:
        if LOG_LEVEL == 10:
            logging.exception(e)

        close()


def get_deliver_dir(info):
    '''Get deliver directory at GCS'''
    try:
        info.update({'jira_id': int(info['jira'][6:])})
        title = 'Task{jira_id:05d}_Group{group}_{src_name}_{fullname}'\
            .format(**info)

        target = load_config()['packages'][info['package']]
        target = Path(target).joinpath(title)

        return target

    except Exception as e:
        logging.exception(e)
        close()


def find_cell(ws, content):
    '''Find cell has content in excel'''
    logging.debug("Finding cell '%s'", content)
    row, col = 0, 0
    for r in ws.rows:
        for cell in r:
            if str(cell.value).strip() == content:
                row, col = cell.row, cell.column
                break

    return row, col


def collapse_list(lst):
    '''Collapse'''
    def text(a, b):
        return str(a) if a == b else '{0}~{1}'.format(a, b)
    rst = []
    t = 0
    for i in range(len(lst)):
        if i > 1 and lst[i] - lst[i-1] > 1:
            rst.append(text(lst[t], lst[i-1]))
            t = i
    rst.append(text(lst[t], lst[i]))

    return ', '.join(rst)


def tmp_print_info(report):
    # print label data
    print('\n\n------- Label -------')
    for label, value in report.tctbl.label_data.items():
        print(label)
        for cmt, lst in value.items():
            print(cmt)
            print(collapse_list(lst))
            print('-'*10)

    # print stub
    print('\n\n------- Stub -------')
    for i in report.data.get('stub', []):
        print(i)

    # print non-stub
    print('\n\n------- Non-stub -------')
    for i in report.data.get('non_stub', []):
        print(i)
