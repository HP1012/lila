# -*- coding: utf-8 -*-

import json
import logging
import os
import shutil
import signal
import socket
import sys
import uuid
from datetime import datetime
from pathlib import Path

import eel
import requests
import win32com.client
from cryptography.fernet import Fernet

import lila.const as CONST

logger = logging.getLogger(__name__)


def load(path, keys=''):
    '''Load data from json file'''
    logger.debug("Load data from %s", Path(path).name)

    def filter(data):
        for k in keys.split('.'):
            data = data.get(k, {}) if k != '' else data
        return data
    try:
        keys = '' if keys is None else keys.strip()

        with open(path, encoding='shift-jis', errors='ignore') as fp:
            data = json.load(fp)
    except Exception as e:
        data = {}
        if Path(path).is_file() is True:
            logger.exception(e)
    finally:
        return filter(data)


def write(data, path):
    '''Write dict to json'''
    logger.debug("Write data to %s", Path(path).name)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, encoding='shift-jis', errors='ignore', mode='w') as fp:
        json.dump(data, fp, indent=4, sort_keys=True)


def read_file(path):
    with open(path, encoding='shift-jis', errors='ignore') as fp:
        return fp.readlines()


def get_lang_data(key=None):
    '''Get language data'''
    lang = load(CONST.CONFIG).get('language', 'en')
    path = CONST.ASSET.joinpath('lang_{0}.json'.format(lang))
    data = load(path)
    return data if key is None else data.get(key, {})


def change_language(language):
    '''Update language'''
    logger.debug("Change language to %s", language)
    lang = 'jp' if language == 'Japan' else 'en'
    config = load(CONST.CONFIG)
    config.update({'language': lang})
    write(config, CONST.CONFIG)


def filter_tbl(data, opts):
    '''Filter rows from table data'''
    def match(row):
        dct = dict(zip(data[0], row))
        return all([dct.get(col) == value for col, value in opts.items()])

    logger.debug("Filter table by %s", opts)
    return [row for row in data if match(row)]


def fuzzy_find(data, func, src_rel):
    '''Find function by fuzzy'''
    def name(path, index):
        lst = Path(path).parts
        return lst[0] if index == 0 else '/'.join(lst[index:])

    def match(row, index=-1):
        dct = dict(zip(data[0], row))
        return name(src_rel, index) == name(dct.get('src_rel', ''), index)

    logger.debug("Fuzzy find %s %s", func, src_rel)

    lst = [row for row in filter_tbl(data, {'func': func}) if match(row)]
    index = -1
    while (len(lst) > 1):
        lst = [row for row in lst if match(row, index)]
        index -= 1

    return dict(zip(data[0], lst[0])) if len(lst) > 0 else {}


def get_auth_key():
    '''Get auth key'''
    return Fernet(CONST.KEY + str(uuid.getnode()))


def get_auth_info():
    '''Get auth info'''
    try:
        config = load(CONST.CONFIG)
        f = get_auth_key()
        text = f.decrypt(config['auth'].encode()).decode()
    except:
        text = 'unknown:unknown'
    finally:
        return tuple(text.split(':'))


def save_auth_info(username, password):
    '''Save auth info'''
    try:
        config = load(CONST.CONFIG)
        f = get_auth_key()
        text = '{0}:{1}'.format(username, password)
        config.update({
            'auth': f.encrypt(text.encode()).decode()
        })
        write(config, CONST.CONFIG)
    except:
        pass


def get_jira_server():
    '''Get jira server'''
    return load(CONST.CONFIG).get('server_jira', CONST.SERVER_JIRA)


def copy(src, dst):
    '''Copy file'''
    logger.debug("Copy %s to %s", Path(src).name, Path(dst).parent.name)
    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    if Path(dst).absolute() != Path(src).absolute():
        shutil.copy2(src, dst)


def download(url, target, auth):
    '''Download file'''
    try:
        logger.debug("Download %s", Path(url).name)
        r = requests.get(url, auth=auth, verify=False)
        if r.status_code == 200:
            with open(target, 'wb') as fp:
                fp.write(r.content)
            return True
    except Exception as e:
        logger.exception(e)


def update_file(path, target, date):
    '''Download/copy the latest file'''
    try:
        auth = get_auth_info()
        if path.startswith('http'):
            r = requests.head(path, auth=auth, verify=False)
            mdate = r.headers.get('last-modified')
        else:
            mdate = str(Path(path).stat().st_mtime)

        if mdate != date:
            date = mdate
            if path.startswith('http'):
                download(path, target, auth)
            else:
                copy(path, target)
        else:
            target = None
    except Exception as e:
        logger.exception(e)
        target = None
    finally:
        return target, date


def delete(filepath):
    '''Delete file if exist'''
    if Path(filepath).is_file():
        Path(filepath).unlink()


def get_label_list(labels=[]):
    '''Get label list'''
    logger.debug("Get label list")
    try:
        lst = load(CONST.SETTING).get('labelList')
        lst_count = [len(set(labels) & set(l)) for l in lst]
        index = lst_count.index(max(lst_count))
        data = lst[index]
    except Exception as e:
        logger.exception(e)
        data = []
    finally:
        return data


def collapse_list(lst):
    '''Collapse the list of number'''
    def text(a, b):
        return str(a) if a == b else '{0}~{1}'.format(a, b)

    if len(lst) == 0:
        return ''
    lst.sort()
    rst = []
    t = 0
    for i in range(len(lst)):
        if i > 1 and lst[i] - lst[i-1] > 1:
            rst.append(text(lst[t], lst[i-1]))
            t = i
    rst.append(text(lst[t], lst[i]))

    return ', '.join(rst)


def is_simulink(path):
    '''Check source code is Simulink model'''
    try:
        rst = False
        with open(path, encoding='shift-jis', errors='ignore') as fp:
            for line in fp.readlines()[:100]:
                if 'Simulink model' in line:
                    rst = True
                    break
    except:
        rst = None
    finally:
        return rst


def scan_files(directory, ext='.txt'):
    '''Scan all file that has extension in directory'''
    logger.debug("Scan directory %s %s", directory, ext)
    data = []
    latest = None
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.endswith(ext):
                filepath = Path(root).joinpath(filename)
                data.append(filepath)

                if (latest is None or
                        Path(latest).stat().st_mtime < filepath.stat().st_mtime):
                    latest = filepath

    return data, latest


def is_open_port(host='localhost', port=CONST.PORT):
    '''Check port is open or not'''
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex((host, port)) == 0
    except Exception as e:
        logger.exception(e)
        result = True
    finally:
        return result


def clean_port(port=CONST.PORT):
    '''Close all pid that listening on port'''
    def get_pid(line):
        try:
            pid = int(str(line).strip().split(' ')[-1])
        except:
            pid = None
        finally:
            return pid

    try:
        logger.debug("Clean program on port %s", port)

        try:
            cmd = 'netstat -ano -p tcp | find "{0}" | find "LISTENING"'
            cmd = cmd.format(port)

            filepath = CONST.ASSET.joinpath('pid')
            logger.debug("Execute command %s", cmd)
            os.system('{0} > {1}'.format(cmd, filepath))

            with open(filepath) as fp:
                lst = [get_pid(line)for line in fp.readlines()]

            if filepath.is_file():
                filepath.unlink()

            logger.debug("PIDs %s", lst)

        except Exception as e:
            logger.exception(e)
            lst = []

        lst = [p for p in lst if p != None]
        for pid in lst:
            try:
                logger.debug("Kill pid %s", pid)
                os.kill(pid, signal.SIGTERM)
            except Exception as e:
                logger.exception(e)

    except Exception as e:
        logging.exception(e)


def merge_data():
    '''Merge with previous version'''
    try:
        data = load(CONST.CONFIG)
        if data.get('code') != CONST.CODE:
            logger.debug("Clean database for running %s", CONST.CODE)

            delete(CONST.WORKSPACE)
            delete(CONST.PACKAGE)
            # shutil.rmtree(CONST.DATA, ignore_errors=True)

            data.update({'code': CONST.CODE})
            write(data, CONST.CONFIG)
    except:
        pass


def get_http_link(subpath):
    '''Generate link to svn for downloading file'''
    server = load(CONST.CONFIG).get('server', CONST.SERVER_BUILD)
    return '{0}/{1}'.format(server, subpath)


def get_http_date(url, auth):
    '''Get date of link'''
    try:
        r = requests.head(url, auth=auth, verify=False)
        return r.headers.get('last-modified')
    except Exception as e:
        logger.exception(e)


def sync_package_data():
    '''Sync package data only use at GCS'''
    logger.debug("Sync package data")
    try:
        auth = get_auth_info()

        # Download package config
        filename = 'package.json'
        url = get_http_link(filename)
        date = get_http_date(url, auth)

        history = load(CONST.HISTORY)

        if date != None and (date != history.get(filename) or CONST.PACKAGE.is_file() is False):

            download(url, CONST.PACKAGE, auth)
            if date != None:
                history.update({filename: date})
                write(history, CONST.HISTORY)

        # Update simulink data
        for package in load(CONST.PACKAGE).keys():
            if package != None and package.strip() != '':
                filename = '{0}.json'.format(package)
                url = get_http_link(filename)
                date = get_http_date(url, auth)
                target = CONST.DATA.joinpath(filename)

                if date != None and (date != history.get(filename) or target.is_file() is False):
                    download(url, target, auth)

                    history.update({filename: date})
                    write(history, CONST.HISTORY)

    except Exception as e:
        logger.exception(e)


def online_status(url=None):
    '''Check online status 
    401 Unauthorize
    404 File not found
    '''
    try:
        auth = get_auth_info()
        url = get_http_link('package.json') if url is None else url
        r = requests.head(url, auth=auth, verify=False)
        return r.status_code
    except:
        return 500


def is_expired(datetime_str):
    '''Check string is later than today'''
    try:
        date = datetime.strptime(datetime_str, '%Y-%m-%d')
        today = datetime.now()
        return today > date
    except Exception as e:
        logger.exception(e)
        return False


def upgrade_build(version):
    '''Upgrade to latest build'''
    try:
        logger.debug("Upgrade to build")
        status = False

        target = CONST.DATA.joinpath(version)
        exe_path = Path.cwd().joinpath(sys.argv[0])

        if str(exe_path).endswith('.exe'):
            url = get_http_link('Builds/{0}'.format(version))

            auth = get_auth_info()
            download(url, target, auth)

            # logger.debug("Exist program to upgrade to new version")

            # cmd = 'taskkill /pid {0} /f'.format(os.getpid())
            # cmd = '{0} && copy /y {1} {2}'.format(cmd, target, exe_path)

            # os.system(cmd)
            status = True

    except Exception as e:
        logger.exception(e)

    finally:
        delete(target)
        return status


def update_progress(msg, params=None, percent=0):
    '''Update report progress to web'''
    params = tuple(params) if isinstance(params, list) else tuple([params])
    msg = msg.format(*params)
    if msg.startswith('Error') or msg.startswith('Exception'):
        msg = "<code>{0}</code>".format(msg)
    eel.updateReportProgress(msg, percent)


def ask_login_info():
    eel.askLoginInfo()
