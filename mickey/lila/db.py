# -*- coding: utf-8 -*-

import logging
import time
from pathlib import Path

from jira import JIRA

import lila.const as CONST
from lila import parse, utils

logger = logging.getLogger(__name__)


class Package(object):

    def __init__(self, name, level=1):
        self.name = name
        self.json = CONST.DATA.joinpath('{0}.json'.format(name))
        self.info = utils.load(CONST.PACKAGE).get(name, {})
        self.data_pkg = utils.load(self.json)

        # Update data
        if self.name != None and self.info != {}:
            self.update_data(level)

    def update_data(self, level=1):
        '''Update data multiple level'''
        logger.debug("Update data level %s", level)
        try:
            if level > 0:
                self.update_data_xlsx('summary')

            if level > 1:
                self.update_data_xlsx('report')

            if level > 2:
                self.update_data_jira()

            if level > 3:
                self.update_data_simulink()

        except Exception as e:
            logger.exception(e)
        finally:
            logger.debug("Done")

    def update_data_xlsx(self, key):
        '''Update summary or report data'''
        logger.debug("Update data %s of package %s", key, self.name)
        try:
            key_date = '{0}_date'.format(key)
            params = [self.info.get(key, {}).get(k)
                      for k in ['xlsx', 'sheet', 'begin', 'end']]

            xlsx = CONST.DATA.joinpath('{0}.{1}.xlsx'.format(self.name, key))

            if str(params[0]).strip() != '' and str(params[1]).strip() != '':

                date = self.data_pkg.get(key_date)

                if params[0] is not None:
                    xlsx, date = utils.update_file(params[0], xlsx, date)

                if xlsx is not None and Path(xlsx).is_file():
                    params[0] = xlsx
                    params += [utils.load(CONST.SETTING, 'headerXlsx')]
                    data = parse.get_xlsx_raw(*tuple(params))
                    if data != {} and data != []:
                        self.data_pkg.update({
                            key: data,
                            key_date: date
                        })
                        utils.write(self.data_pkg, self.json)
                    utils.delete(xlsx)
        except Exception as e:
            logger.exception(e)

    def update_data_jira(self):
        '''Update jira data of package'''
        try:
            logger.debug("Update jira data of package %s", self.name)
            ticket = self.info.get('jira')
            if ticket is not None and ticket.strip() != '':
                options = {
                    'server': utils.get_jira_server(),
                    'verify': False
                }
                auth = utils.get_auth_info()
                jira = JIRA(options=options, auth=auth)
                issue = jira.issue(ticket)
                data = {t.key: t.fields.summary
                        for t in issue.fields.subtasks}
                if data != {} and data != self.data_pkg.get('jira', {}):
                    self.data_pkg.update({'jira': data})
                    utils.write(self.data_pkg, self.json)

        except Exception as e:
            logger.exception(e)

    def update_data_simulink(self):
        '''Update simulink data'''
        logger.debug("Update simulink data")
        try:
            data = {}
            directory = self.info.get('source')
            list_file, _ = utils.scan_files(directory, ext='.c')
            for filepath in list_file:
                data.update({
                    str(filepath): utils.is_simulink(filepath)
                })
            if data != {}:
                self.data_pkg.update({'simulink': data})
                utils.write(self.data_pkg, self.json)
        except Exception as e:
            logger.exception(e)
            data = {}
        finally:
            return data

    def find_ticket(self, info, count=0):
        '''Find ticket info'''
        try:
            src_name = Path(str(info.get('src_rel'))).name
            info.update({'src_name': src_name})
            title = 'Group{group}_{src_name}_{pic}'.format(**info)
            for key, value in self.data_pkg.get('jira', {}).items():
                if value.startswith(title):
                    ticket, title = key, value
                    break
            else:
                if count == 0:
                    self.update_data_jira()
                    ticket, title = self.find_ticket(info, count=1)
                else:
                    ticket, title = None, None
        except Exception as e:
            logger.exception(e)
            ticket, title = None, None
        finally:
            return ticket, title

    def get_func_info(self, func, src_rel, key='summary'):
        '''Get function info'''
        logger.debug("Get function info %s %s %s", func, src_rel, key)
        try:
            data = self.data_pkg.get(key)
            rst = {}
            if isinstance(data, list) and len(data) > 1:
                rst = utils.fuzzy_find(data, func, src_rel)

            if rst != {}:
                ticket, title = self.find_ticket(rst)
                if ticket is not None:
                    num = int(ticket.split('-')[-1])
                    task_title = 'Task{0:05d}_{1}'.format(num, title)
                    rst.update({
                        'jira': ticket,
                        'jira_title': title,
                        'task_title': task_title
                    })
        except Exception as e:
            logger.exception(e)
            rst = rst if isinstance(rst, dict) else {}
        finally:
            return rst

    def is_simulink(self, src_full):
        '''Check source is simulink model or not'''
        def name(path, index):
            lst = Path(path).parts
            return lst[0] if index == 0 else '/'.join(lst[index:])

        def match(path, index=-1):
            return name(src_full, index) == name(path, index)

        try:
            logger.debug("Check is simulink %s", src_full)

            data = self.data_pkg.get('simulink', {})
            dct = {path: data.get(path) for path in data.keys()
                   if Path(path).name == Path(src_full).name}

            lst = list(set(dct.values()))

            if len(lst) == 0:
                result = None
            elif len(lst) == 1:
                result = lst[0]
            else:
                lst = dct.keys()
                index = -1
                while len(lst) > 1:
                    lst = [path for path in lst if match(path, src_full)]
                    index -= 1

                if len(lst) == 1:
                    result = data.get(lst[0])
                else:
                    result = None

        except Exception as e:
            logger.exception(e)
            result = None
        finally:
            return result


def update_workspace(info, action, filepath=CONST.WORKSPACE):
    '''Update workspace'''
    data = utils.load(filepath)
    name = info.get('name')

    if action == 'delete':
        if name in data.keys():
            logger.debug("Delete workspace %s", name)
            del data[name]

    else:
        dct = data.get(name, {})

        action = "Add new" if dct == {} else "Update"
        logger.debug("%s workspace %s", action, info)

        info.update({'stamp': time.time()})

        dct.update(info)
        data.update({name: dct})

    # Write changes to file
    utils.write(data, filepath)


def update_package(info, action, path=CONST.PACKAGE):
    '''Update package'''
    data = utils.load(path)
    name = info.get('name')

    if action == 'delete':
        if name in data.keys():
            logger.debug("Delete package %s", name)
            del data[name]

    else:
        action = "Add new" if name in data.keys() else "Update"
        logger.debug("%s package %s", action, info)

        data.update({name: info})

    # Write changes to file
    utils.write(data, path)


def get_workspace_data():
    '''Get workspace data'''
    logger.debug("Get workspace data")
    data_wsp = utils.load(CONST.WORKSPACE)
    if data_wsp == {}:
        return {}

    list_wsp = [[name, dct.get('path'), dct.get('package'), dct.get('stamp', 0)]
                for name, dct in data_wsp.items()]
    list_wsp.sort(key=lambda x: x[0])

    # Latest workspace
    last_wsp_name = sorted(list_wsp, key=lambda x: x[3])[-1][0]
    last_wsp = data_wsp.get(last_wsp_name, {})

    # Scan testlog in directory
    dir_log = last_wsp.get('path')
    list_log, latest_log = utils.scan_files(dir_log)
    list_log = [[Path(f).name, f] for f in list_log]
    list_log.sort(key=lambda x: x[0])

    # Latest function
    last_log = Path(str(last_wsp.get('function', latest_log)))
    if str(last_log).startswith(str(dir_log)) is False or last_log.is_file() is False:
        last_log = latest_log

    return {
        'list_wsp': list_wsp,
        'last_wsp_name': last_wsp_name,
        'list_log': list_log,
        'last_log': last_log if last_log is None else last_log.name,
        'latest_log': latest_log if latest_log is None else latest_log.name,
        'last_pkg_name': last_wsp.get('package')
    }


def get_func_info(testlog, package, level=1):
    '''Get function info'''
    logger.debug("Get function info %s %s", Path(testlog).name, package)
    info = parse.parse_testlog(testlog)
    pkg = Package(package, level)
    data = pkg.get_func_info(info.get('func'), info.get('src_full'))
    data.update({'package': package})

    for key, value in info.items():
        if key not in data.keys():
            data.update({key: value})

    return data


def get_list_package():
    '''Get list package'''
    def is_package(filepath):
        data = utils.load(filepath)
        return 'summary' in data.keys() and 'simulink' in data.keys()

    try:
        data = list(utils.load(CONST.PACKAGE).keys())

        lst, _ = utils.scan_files(CONST.DATA, '.json')
        for filepath in lst:
            try:
                if is_package(filepath):
                    filename = Path(filepath).name
                    name = filename.replace('.json', '')
                    if name not in data:
                        data.append(name)
            except Exception as e:
                logger.exception(e)

    except Exception as e:
        logger.exception(e)
        data = data if isinstance(data, list) else []
    finally:
        return data
