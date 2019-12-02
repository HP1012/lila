# -*- coding: utf-8 -*-

import logging
import sys
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askdirectory, askopenfilename

import eel
from jinja2 import Environment, FileSystemLoader, Template

import lila.const as CONST
from lila import db, parse, utils
from lila.ams import Report

logger = logging.getLogger(__name__)


@eel.expose
def change_language(language):
    '''Change language'''
    logger.debug("Request change language to %s", language)
    try:
        utils.change_language(language.strip())
        generate_html()
    except Exception as e:
        logger.exception(e)


@eel.expose
def get_language_text():
    '''Get language text'''
    try:
        data = utils.get_lang_data('text')
    except Exception as e:
        logger.exception(e)
        data = {}
    finally:
        return data


@eel.expose
def update_workspace(info, action):
    '''Update workspace'''
    try:
        db.update_workspace(info, action)
    except Exception as e:
        logger.exception(e)


@eel.expose
def get_workspace_data():
    '''Get list of space and testlog'''
    wsp_options = '''
        {% for i in list_wsp %}
        <option path="{{i[1]}}" package="{{i[2]}}" data-tokens="{{i[0]}}">{{i[0]}}</option>
        {% endfor %}
        '''
    log_options = '''
        {% for i in list_log %}
        <option path="{{i[1]}}" data-tokens="{{i[0]}}">{{i[0]}}</option>
        {% endfor %}
        '''
    pkg_options = '''
        {% for pkg in list_pkg %}
        <option data-tokens="{{pkg}}">{{pkg}}</option>
        {% endfor %}
        '''
    try:
        info = db.get_workspace_data()

        data = {
            'workspace': render(info, wsp_options, mode="string"),
            'function': render(info, log_options, mode="string"),
            'lastWsp': info.get('last_wsp_name'),
            'lastFunc': info.get('last_log'),
            'latestFunc': info.get('latest_log')
        }

        list_pkg = db.get_list_package()
        data.update({
            'pkgOptions': render({'list_pkg': list_pkg}, pkg_options, mode="string")
        })

    except Exception as e:
        logger.exception(e)
        data = data if isinstance(data, dict) else {}
    finally:
        return data


@eel.expose
def get_workspace_summary(check_all=False):
    '''Get list of workspace and testlog'''
    def check(testlog, package):
        try:
            report = Report(testlog)
            report.check(package)

            checklist = report.get_checklist()
            issue = [i for i in checklist
                     if i[2] != True and i[0] != 'AMSTB_SrcFile.c']

            if len(issue) > 0:
                status = '<span class="label label-danger">NG {0}</span>' \
                    .format(len(issue))

            else:
                status = '<span class="label label-success">OK {0}</span>' \
                    .format(len(checklist))

        except Exception as e:
            logger.exception(e)
            status = '<span class="label label-warning">NC</span>'
        finally:
            return status

    wsp_options = '''
        {% for i in list_wsp %}
        <option path="{{i[1]}}" package="{{i[2]}}" data-tokens="{{i[0]}}">{{i[0]}}</option>
        {% endfor %}
        '''
    sync_btn = '<button path="{testlog}" type="button" class="btn btn-sm btn-icon btn-pure btn-outline sync" data-toggle="tooltip" data-original-title="Sync"><i class="mdi mdi-sync" aria-hidden="true"></i></button>'

    try:
        logger.debug("Request summary data")
        data_wsp = db.get_workspace_data()

        data = {
            'workspace': render(data_wsp, wsp_options, mode="string"),
            'lastWsp': data_wsp.get('last_wsp_name')
        }

        if check_all is True:
            eel.updateProgress(data)

            pkg_name = data_wsp.get('last_pkg_name')
            package = db.Package(pkg_name)

            sum_header = package.data_pkg.get('summary', [[]])[-1]
            sum_header_key = package.data_pkg.get('summary', [[]])[0]
            dct_header = dict(zip(sum_header, sum_header_key))

            tbl_header = ['No.', 'Function', 'Result',
                          'C0', 'C1', 'MC/DC', 'Lila Check', 'Actions']
            all_header = [h for h in sum_header
                          if h not in tbl_header and h != None]
            all_header = tbl_header + all_header + ['Testlog']

            table = generate_sum_header(all_header, tbl_header)

            tbody = ''
            count = 1
            total = len(data_wsp.get('list_log', []))
            for _, testlog in data_wsp.get('list_log', []):
                info = parse.parse_testlog(testlog)
                summary = db.get_func_info(testlog, pkg_name, level=0)
                if 'src_rel' in summary:
                    info.update({'src_rel': summary.get('src_rel')})
                summary.update(info)

                status = check(testlog, pkg_name)

                row = ''
                for header in all_header:
                    key = dct_header.get(header)
                    value = summary.get(key)

                    if header == 'Actions':
                        value = sync_btn.format(testlog=testlog)
                    elif header == 'Testlog':
                        value = testlog
                    elif header == 'Result':
                        if value == 'OK':
                            value = '<span class="label label-success">OK</span>'
                        elif value == 'NG':
                            value = '<span class="label label-danger">NG</span>'
                        else:
                            value = '<span class="label label-warning">NC</span>'

                    elif header == 'Lila Check':
                        value = status

                    cell = '<td>{0}</td>'.format(value)
                    row += cell

                row = '<tr>{0}</tr>'.format(row)
                tbody += row

                percent = int(count*100 / total)
                eel.updateProgress({'percent': percent})

                count += 1

            tbody = '<tbody>{0}</tbody>'.format(tbody)
            table += tbody

            data = {'table': table}
    except Exception as e:
        logger.exception(e)
        data = data if isinstance(data, dict) else {}
    finally:
        eel.updateProgress({'percent': 100})
        return data


def generate_sum_header(lst_all, lst_show):
    text = ''
    for header in lst_all:
        if header in lst_show:
            th = '<th> {0} </th>'.format(header)
        else:
            th = '<th data-hide="all" class="lang"> {0} </th>'.format(header)
        text += th
    return '<thead><tr>{0}</tr></thead>'.format(text)


@eel.expose
def select_folder():
    '''Ask the user to select a folder'''
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    directory = askdirectory(parent=root)

    return str(Path(directory))


@eel.expose
def select_file(extension):
    """ Ask the user to select a file """
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    if extension == 'excel':
        file_types = [('Excel files', '*.xls;*.xlsx')]
    else:
        file_types = [('All files', '*')]
    file_path = askopenfilename(parent=root, filetypes=file_types)
    root.update()

    return str(Path(file_path))


@eel.expose
def get_coverage_report(testlog):
    '''Get coverage report from testlog'''
    return parse.parse_testlog(testlog)


@eel.expose
def get_function_info(testlog, package):
    '''Get function more info'''
    logger.debug("Request function info %s %s", Path(testlog).name, package)
    try:
        # Get function info
        data = db.get_func_info(testlog, package)
    except Exception as e:
        logger.exception(e)
        data = {}
    finally:
        return data


@eel.expose
def sync_package_data():
    '''Sync package data from server'''
    try:
        is_online = utils.online_status()
        if is_online == 401:
            utils.ask_login_info()

        elif is_online == 200:
            utils.sync_package_data()

    except Exception as e:
        logger.exception(e)


@eel.expose
def get_report_info(testlog, workspace):
    '''Get report info'''
    logger.debug("Get report info of testlog %s %s", testlog, workspace)
    try:
        package = utils.load(CONST.WORKSPACE).get(workspace, {}).get('package')
        data = get_function_info(testlog, package)
        func = data.get('func', '')
        task = data.get('task_title', '')

        deliver = utils.load(CONST.WORKSPACE).get(workspace).get('deliver')
        deliver = Path(deliver)

        deliver = deliver.joinpath(task)

        dirTarget = utils.load(CONST.SETTING, 'jpDict.dirresult')
        dirSpec = utils.load(CONST.SETTING, 'jpDict.dirspec')

        data.update({
            'dirTarget': str(deliver.joinpath(dirTarget, func)),
            'dirSpec': str(deliver.joinpath(dirSpec)),
            'dirTargetGcs': str(deliver.joinpath(dirTarget, task, func)),
            'func_no': data.get('func_no')
        })

    except Exception as e:
        logger.exception(e)
        data = {}
    finally:
        return data


@eel.expose
def get_workspace_info(workspace):
    '''Get workspace info'''
    logger.debug("Get workspace info %s", workspace)
    pkg_options = '''
        {% for pkg in list_pkg %}
        <option data-tokens="{{pkg}}">{{pkg}}</option>
        {% endfor %}
        '''
    try:
        data = utils.load(CONST.WORKSPACE).get(workspace)
        list_pkg = db.get_list_package()
        data.update({
            'pkgOptions': render({'list_pkg': list_pkg}, pkg_options, mode="string")
        })
    except Exception as e:
        logger.exception(e)
        data = data if isinstance(data, dict) else {}
    finally:
        return data


@eel.expose
def check_testlog(testlog, package):
    '''Check testlog base on checklist'''
    logger.debug("Request check testlog %s", testlog)
    try:
        report = Report(testlog)
        report.check(package)

    except Exception as e:
        logger.exception(e)
    finally:
        data = utils.get_lang_data('ui')
        template = 'table.checklist.html'

        checklist = report.get_checklist()
        data.update({'data': checklist})
        tbl_checklist = render(data, template)

        warninglist = [i for i in checklist if i[2] != True]
        data.update({'data': warninglist})
        tbl_warnlist = render(data, template)

        return {
            'checklist': tbl_checklist,
            'warninglist': tbl_warnlist
        }


@eel.expose
def deliver_report(info):
    '''Deliver report'''
    logger.debug("Deliver report")
    try:
        testlog = info.get('testlog')
        report = Report(testlog)

        # Deliver test result file
        target = info.get('target')
        report.deliver_files(target)

        # Deliver test spec
        report.generate_spec(info,update_spec = False)
    except Exception as e:
        logger.exception(e)

@eel.expose
def deliver_report_update(info):
    '''Deliver report'''
    logger.debug("Deliver report")
    try:
        testlog = info.get('testlog')
        report = Report(testlog)

        # Deliver test result file
        target = info.get('target')
        report.deliver_files(target)

        # Deliver test spec
        report.generate_spec(info,update_spec = True)
    except Exception as e:
        logger.exception(e)


@eel.expose
def update_package(info, action):
    '''Update package'''
    logger.debug("Request update package")
    try:
        db.update_package(info, action)
    except Exception as e:
        logger.exception(e)


@eel.expose
def update_package_data(name):
    '''Update package data'''
    logger.debug("Request update package data")
    try:
        package = db.Package(name)
        package.update_data(level=4)
    except Exception as e:
        logger.exception(e)


@eel.expose
def get_package_info(name=None):
    '''Get package info'''
    logger.debug("Request package info")
    try:
        data = utils.load(CONST.PACKAGE)
        info = data if name is None else data.get(name, {})
    except Exception as e:
        logger.exception(e)
        info = {}
    finally:
        return info


@eel.expose
def update_auth(username, password):
    '''Update auth info'''
    try:
        logger.debug("Update auth info")
        utils.save_auth_info(username, password)
        return utils.online_status()
    except:
        pass


@eel.expose
def get_messages():
    '''Get messages from server'''
    try:
        msg = '''<a id="new-build" href={URL}>
                    <div class="mail-contnet">
                        <h5>{title}</h5> <h6>{content}</h6> <h6>{content1}</h6>
                    </div>
                </a>
        '''
        html = ''

        is_online = utils.online_status()
        if is_online == 200:
            auth = utils.get_auth_info()

            # Get version
            version = utils.load(CONST.VERSION, 'version')
            
            filename = 'messages.json'
            url = utils.get_http_link(filename)
            filepath = CONST.DATA.joinpath(filename)
            utils.download(url, filepath, auth)

            info = utils.load(filepath)
            latest = info.get('version')
            link_down = info.get('link')
            descrip = info.get('descrip')

            if latest != None and version != latest:
                html += msg.format(URL = link_down,title='New Build Available',
                                   content='Version {0}'.format(latest),content1 = descrip )

            # Get messages
            for title, body in info.items():
                if title == 'version' or title == 'link' or title == 'descrip':
                    continue

                try:
                    content, tilldate = tuple(body)
                    if utils.is_expired(tilldate) is False:
                        html += msg.format(URL = "",title=title, content = content, content1 = "")
                except Exception as e:
                    logger.exception(e)

            # Delete file
            utils.delete(filepath)

        data = {'html': html} if html.strip() != '' else {}

    except Exception as e:
        logger.exception(e)
        data = {}

    finally:
        return data


@eel.expose
def upgrade_build(version):
    '''Upgrade build'''
    try:
        logger.debug("Request upgrade to build")
        status = utils.upgrade_build(version)
    except Exception as e:
        logger.exception(e)
        status = False
    finally:
        return status


@eel.expose
def close():
    '''Close app'''
    logger.debug("Exit app")
    sys.exit()


def generate_html():
    '''Generate html files'''
    data = utils.get_lang_data('ui')
    data.update(utils.load(CONST.VERSION))
    lst_page = [
        ('workspace.html', 'index.html'),
        ('package.html', 'package.html'),
        ('summary.html', 'summary.html')
    ]

    for page in lst_page:
        render(data, *page)


def render(data, template, path=None, mode='template'):
    '''Render template and write to file'''
    if mode == 'template':
        logger.debug("Render template %s", template)
        env = Environment(loader=FileSystemLoader(str(CONST.TEMPLATE)))
        tmp = env.get_template(template)
        content = tmp.render(**data)
    else:
        content = Template(template).render(**data)

    content = clean_html(content)

    if path is not None:
        path = CONST.WEB.joinpath(path)
        with open(path, encoding='shift-jis', errors='ignore', mode='w') as fp:
            logger.debug("Write to file %s", path)
            fp.write(content)

    return content


def clean_html(content):
    '''Clean html code'''
    data = [l for l in [k.strip() for k in content.split('\n')]
            if (l.startswith('<!--') and l.endswith('-->')) is False]
    return ''.join(data)
