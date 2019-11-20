# -*- coding: utf-8 -*-

import logging
from pathlib import Path
from unicodedata import normalize

import lxml.html

import lila.const as CONST
from lila import db, parse, utils
from lila.excel import ExcelWin32
from lila.utils import update_progress as wlogger

logger = logging.getLogger(__name__)


class Base(object):

    def __init__(self, path):
        self.path = Path(path)


class Table(Base):

    def __init__(self, path):
        super().__init__(path)
        self.doc = lxml.html.parse(str(path))

    def get_tag(self, tag, index=0):
        '''Get normalized text of tag base on index of tag'''
        node = [e for e in self.doc.iterfind('.//{0}'.format(tag))][index]
        return self.get_node(node)

    def get_node(self, node, form='NFKD'):
        '''Get normalized text of node'''
        return normalize(form, node.text_content())

    def get_table(self, index=0):
        '''Get virtual table from html base on table tag'''
        logger.debug("Get virtual table at index %s", index)

        node = [t for t in self.doc.iterfind('.//table')][index]
        lst = [int(c.get('colspan', 1)) for c in list(node)[0]]
        table = [[None for i in range(sum(lst))]
                 for j in range(len(list(node)))]

        row = 0
        for tr in list(node):
            col = 0
            for td in list(tr):
                colspan = int(td.get('colspan', 1))
                rowspan = int(td.get('rowspan', 1))

                while table[row][col] is not None:
                    col += 1

                td.set('origin', ','.join([str(row), str(col)]))
                for r in range(row, row + rowspan):
                    for c in range(col, col + colspan):
                        table[r][c] = td

            row += 1

        return table

    def get_table_raw(self, table):
        '''Get table raw data'''
        row = [None for c in range(len(table[0]))]
        data = [row[:] for r in range(len(table))]

        buff = 0
        for r in range(len(table)):
            rowspan = 1
            for c in range(len(table[0])):
                cell = table[r][c]
                text = self.get_node(cell)
                if (r, c) != self.get_origin(cell):
                    continue
                string = str(lxml.html.tostring(cell))
                if '<br>' in string:
                    string = string.replace('<br>', '\n')[2:-6]
                    string = ''.join(lxml.html.fromstring(string).itertext())
                    lst = string.splitlines()
                    if len(lst) > rowspan:
                        for _ in range(rowspan, len(lst)):
                            data.append(row[:])
                        rowspan = len(lst)
                    for k in range(len(lst)):
                        data[r+buff+k][c] = lst[k]
                else:
                    data[r+buff][c] = text
            buff += (rowspan - 1)

        r0 = row[:]
        r0[0] = self.get_tag('h4', 0)
        data = [r0, row[:]] + data

        return data

    def get_origin(self, cell):
        '''Get origin (row, col) of cell in table'''
        return tuple([int(v) for v in cell.get('origin').split(',')])


class FileCollection(Base):

    collection = {
        # 'testlog': '',
        'csv': (3, 'TestCsv/{func}.csv'),
        'ini': (3, 'TestCsv/{func}.ini'),
        'xeat': (3, 'TestCsv/{func}.xeat'),
        'xtct': (3, 'TestCsv/{func}.xtct'),
        'ie': (3, 'TestCsv/{func}_IE.html'),
        'io': (3, 'TestCsv/{func}_IO.html'),
        'oe': (3, 'TestCsv/{func}_OE.html'),
        'tc': (3, 'TestCsv/{func}_TC.html'),
        'info': (2, '{func}_Info.html'),
        'table': (2, '{func}_Table.html'),
        'report_html': (2, 'TestReport.htm'),
        'report_csv': (2, 'TestReport.csv'),
        'stub': (3, 'AMSTB_SrcFile.c'),
        'xlsx': (2, '{dirspec}/{func}.xlsx'),
        'report_html_jp': (2, '{testreport}.htm'),
        'report_csv_jp': (2, '{testreport}.csv')
    }

    def __init__(self, path):
        super().__init__(path)
        self.is_ams = True

    def get_files(self, info):
        '''Collect files'''
        def get_path(level, filename):
            '''Generate filepath from base line'''
            filename = filename.format(**info)
            if self.is_ams is False and filename.endswith('.xlsx') is False:
                level = 0
                filename = Path(filename).name
            return list(self.path.parents)[level].joinpath(filename)

        # JP name of TestReport
        info['testreport'] = utils.load(CONST.SETTING, 'jpDict.testreport')
        info['dirspec'] = utils.load(CONST.SETTING, 'jpDict.dirspec')

        self.is_ams = (info.get('src_name') == self.path.parent.name)
        logger.debug("Collect files from winAMS %s", self.is_ams)

        result = {'testlog': self.path}
        result.update({
            key: get_path(*value)
            for key, value in self.collection.items()
        })

        # Remove backup file
        lst = [
            ('report_html', 'report_html_jp'),
            ('report_csv', 'report_csv_jp')
        ]
        for origin, backup in lst:
            if result[origin].is_file() is False and result[backup].is_file():
                result[origin] = result[backup]

            del result[backup]

        return result


class FileCsv(Base):

    def __init__(self, path):
        super().__init__(path)
        self.content = self.load_file()
        self.info = self.get_func_info()
        self.init_vars, self.init_data = self.get_initial_vars()
        self.io_vars = self.get_io_vars()
        self.stub_info = self.get_stub_info()

    def check_2_desc(self, src_full, package=None):
        '''Check description'''
        try:
            result, exp = True, []

            func = Path(self.info['func_full']).name
            src_full = Path(src_full)

            desc = self.info.get('desc')
            simulink = 'Simulink model'
            intro = "Actual vs Expected"

            is_simulink = utils.is_simulink(src_full)
            if is_simulink is None:
                pkg = db.Package(package, level=0)
                is_simulink = pkg.is_simulink(src_full)

            # Check description
            if desc is None or str(desc).strip() == '':
                result = False
                exp = ["<code>Missing description</code>"]

            elif is_simulink is True and desc != simulink:
                result = False
                msg = "<code>{0} != {1}</code>".format(desc, simulink)
                exp = [intro, msg]

            elif is_simulink is False and desc != func:
                result = False
                msg = "<code>{0} != {1}</code>".format(desc, func)
                exp = [intro, msg]

            elif is_simulink is None and desc not in [func, simulink]:
                result = False
                msg = "<code>{0}</code> not in [{1}, {2}]" \
                    .format(desc, func, simulink)
                exp = [msg]

            elif is_simulink is None:
                result, exp = None, ["<code>Not found simulink info</code>"]

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]

        finally:
            return (result, '<br>'.join(exp))

    def check_3_init_opt(self):
        '''Check initial option whenerver call'''
        try:
            result, exp = True, ''
            if self.init_vars.get('#InitWheneverCall', '1') != '1':
                result, exp = False, '<code>InitWheneverCall is False</code>'
        except Exception as e:
            logger.exception(e)
            result, exp = None, str(e)
        finally:
            return (result, exp)

    def check_4_init_var(self):
        '''Not init input variable.
        Init all output variale'''
        try:
            result, exp = True, []

            # Should not init input variable
            tmp = [var for var in self.init_vars
                   if var in self.io_vars.get('input')]

            if len(tmp) > 0:
                result = False
                msg = '<code>{0}</code>'.format('<br>'.join(tmp))
                exp += ['Initial input variables:', msg]

            # Should init all output variable
            tmp = [var for var in self.io_vars.get('output')
                   if var not in self.io_vars.get('input') and
                   '@@' not in var and
                   var not in self.init_vars]

            if len(tmp) > 0:
                result = False
                msg = '<code>{0}</code>'.format('<br>'.join(tmp))
                exp += ['Missing init output variables:', msg]

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]
        finally:
            return (result, '<br>'.join(exp))

    def load_file(self):
        '''Load csv file'''
        with open(self.path, encoding='shift-jis', errors='ignore') as fp:
            return [line.strip() for line in fp.readlines()]

    def modify_value(self, value):
        return value[1:-1] if value.startswith('"') and value.endswith('"') else value

    def get_func_info(self):
        '''Get function info'''
        lst = [l for l in self.content if l.startswith('mod')][0].split(',')
        return {
            'func_full': self.modify_value(lst[1]),
            'desc': self.modify_value(lst[2]),
            'num_input': int(lst[3]),
            'num_output': int(lst[4])
        }

    def get_initial_vars(self):
        '''Get initial info'''
        variable = []
        value = []
        for i in range(len(self.content)):
            if self.content[i].startswith('#InitWheneverCall'):
                variable = [self.modify_value(v)
                            for v in self.content[i].split(',')]
                value = self.content[i+1].split(',')
                break

        init_data = []
        for i in range(len(variable)):
            init_data.append([variable[i], value[i]])

        return dict(zip(variable, value)), init_data

    def get_io_vars(self):
        '''Get input and output variable'''
        variables = {}
        for line in self.content:
            if line.startswith('#COMMENT'):
                lst = [self.modify_value(v) for v in line.split(',')]
                variables = {
                    'input': lst[1: (self.info['num_input'] + 1)],
                    'output': lst[self.info['num_input'] + 1:]
                }
                break

        return variables

    def get_stub_info(self):
        '''Get stub info'''
        stub = []
        non_stub = []

        for line in self.content:
            if line.startswith('%'):
                lst = [self.modify_value(v) for v in line.split(',')]
                if lst[1] == '':
                    non_stub.append(lst[1:])
                else:
                    stub.append(lst[1:])

        return {
            'stub': stub,
            'non_stub': non_stub
        }


class FileTxt(Base):
    '''.txt'''

    def __init__(self, path):
        super().__init__(path)
        self.info = parse.parse_testlog(self.path)

    def check_13_parse(self):
        '''Check able to parse testlog to get info or not'''
        try:
            result, exp = True, ''
            if self.info == {}:
                result, exp = False, 'Unable to parse testlog to get info'
            else:
                for key in ['c0', 'c1', 'mcdc']:
                    int(self.info[key][:-1])
        except Exception as e:
            logger.exception(e)
            result, exp = False, str(e)
        finally:
            return (result, exp)

    def check_14_func(self):
        '''Check testlog file is enough or not'''
        try:
            result, exp = False, '<code>Missing function in testlog</code>'
            lst = utils.read_file(self.path)
            lst_char = ['+', '-', '/', '=']
            func1 = ' ' + self.info['func']
            func2 = '*' + self.info['func']
            for line in lst[6:]:
                func = func1
                if func1 not in line and func2 not in line:
                    continue

                if func2 in line:
                    func = func2

                tmp = [block.strip() for block in line.split(func)]
                if tmp[1].startswith('(') is False:
                    continue
                tmp2 = [i.strip() for i in tmp[0].split(' ')
                        if i.strip() != '']

                if set(lst_char) - set(tmp2) == set(lst_char):
                    result, exp = True, ''
                    break

        except Exception as e:
            logger.exception(e)
            result, exp = None, str(e)
        finally:
            return (result, exp)


class FileTable(Table):
    # _Table.html

    class_no = [
        'data-no',
        'data-no-last',
        'data-commentout-center-left-right'
    ]
    class_cmt = [
        'data-commentout-center-left-right',
        'data-commentout-right'
    ]
    class_input = ['head-input-no', 'head-input-no-last']
    class_output = ['head-output-no', 'head-output-no-last']

    def __init__(self, path):
        super().__init__(path)
        self.table = self.get_table(0)

        self.ino = self.get_index_header('no')
        self.icf = self.get_index_header('confirmation')
        self.iid = self.get_index_header('id')
        self.iai = self.get_index_header('item')
        self.icm = self.get_index_header('comment')

    def get_number(self, cell):
        number = None
        if cell.get('class') in self.class_no:
            number = abs(int(self.get_node(cell)))

        return number

    def get_index_header(self, col):
        '''Get index of header'''
        logger.debug("Get index of column %s", col)
        lst_header = utils.load(CONST.SETTING, 'headerTable.{0}'.format(col))
        tbl_header = [self.get_node(cell) for cell in self.table[0]]

        headers = [h for h in lst_header if h in tbl_header]

        return tbl_header.index(headers[0]) if len(headers) > 0 else None

    def get_testcase_data(self):
        '''Get testcase data from testcase table'''
        logger.debug("Get testcase data")
        data = {}
        for row in self.table:
            no = self.get_number(row[self.ino])
            if no is None or row[self.ino].get('class') in self.class_cmt:
                continue
            item = self.get_node(row[self.iai])

            tmp = data.get(item, [])
            tmp.append(no)
            tmp.sort()
            data.update({item: tmp})

        return data

    def get_io_vars(self):
        '''Get input/output variable in _Table.html'''
        logger.debug("Get input/output variables")
        data = {
            'input': [],
            'output': []
        }
        for cell in self.table[1]:
            if cell.get('class') in self.class_input:
                data['input'].append(self.get_node(cell))

            elif cell.get('class') in self.class_output:
                data['output'].append(self.get_node(cell))

        return data

    def get_confirm(self):
        '''Get confirm'''
        try:
            result = None
            for row in self.table:
                no = self.get_number(row[self.ino])
                confirm = self.get_node(row[self.icf])

                if row[self.icf].get('class') in self.class_cmt or no is None:
                    continue

                if confirm == 'Fault':
                    result = confirm
                    break

        except Exception as e:
            logger.exception(e)
            result = None

        finally:
            return result

    def check_15_io_var(self, data):
        '''Check input/output variable'''
        def is_same_var(var1, var2):

            result = var1 == var2

            # Argument
            if result is False:
                result = var1.startswith('@') and var2.endswith(var1)

            # Pointer
            if result is False:
                if var1.startswith('$') and var2.startswith('$'):
                    result = is_same_var(var1[1:], var2[2:])

            return result

        logger.debug("Check 15_io_var")
        try:
            result, exp = True, []
            io_vars = self.get_io_vars()

            for key in ['input', 'output']:
                # Check length of two list
                if len(io_vars[key]) != len(data[key]):
                    result = False
                    msg = "{0} length: <code>{1} != {2}</code>"\
                        .format(key.title(), len(io_vars[key]), len(data[key]))
                    exp.append(msg)

                # Check each variable
                for i in range(len(io_vars[key])):
                    if is_same_var(io_vars[key][i], data[key][i]) is False:
                        result = False
                        msg = "{0} index.{1} <code>{2} != {3}</code>" \
                            .format(key.title(), i, io_vars[key][i], data[key][i])
                        exp.append(msg)

            if result == False:
                exp = ["_Table.html vs CSV"] + exp

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]
        finally:
            return (result, '<br>'.join(exp))

    def check_7_title(self, func):
        '''Check title'''
        logger.debug("Check 7_title")
        try:
            result, exp = True, ''
            expected = '{0}.csv'.format(func)
            actual = self.get_tag('a', 0)
            result = expected == actual
            if result is False:
                exp = "Title diff <code>{0} != {1}</code>"\
                    .format(expected, actual)

        except Exception as e:
            logger.exception(e)
            result, exp = None, str(e)
        finally:
            return (result, exp)

    def check_8_index(self):
        '''Check index'''
        logger.debug("Check 8_index")
        try:
            result, exp = True, ''
            lst = [self.get_number(row[self.ino]) for row in self.table]
            lst = [n for n in lst if n is not None]

            # Index hopping
            prev = 0
            for tcno in lst:
                if prev != 0 and tcno != (prev + 1):
                    result = False
                    exp = "Index hopping <code>No.{0}->{1}</code>"\
                        .format(prev, tcno)
                    break
                prev = tcno

            # Index start at 1
            if lst[0] != 1:
                result = False
                exp = "Index start at <code>No.{0}</code>".format(lst[0])

        except Exception as e:
            logger.exception(e)
            result, exp = None, str(e)
        finally:
            return (result, exp)

    def check_9_confirm(self):
        '''Check confirm'''
        logger.debug("Check 9_confirm")
        try:
            result, exp = True, ''
            lst = []
            for row in self.table:
                no = self.get_number(row[self.ino])
                confirm = self.get_node(row[self.icf])

                if row[self.icf].get('class') in self.class_cmt or no is None:
                    continue

                if confirm not in ['OK', 'Fault']:
                    lst.append(no)

            if lst != []:
                result = False
                exp = "Missing confirmation <code>No.{0}</code>" \
                    .format(utils.collapse_list(lst))
        except Exception as e:
            logger.exception(e)
            result, exp = None, str(e)
        finally:
            return (result, exp)

    def check_10_header(self):
        '''Check header'''
        logger.debug("Check 10_header")
        try:
            result, exp = True, ''
            actual = [self.get_node(cell) for cell in self.table[0]]
            lst = utils.load(CONST.SETTING, 'headerHtmlTable')

            diff = [list(set(expected) - set(actual)) for expected in lst]

            if len(diff[0]) > 0 and len(diff[1]) > 0:
                result = False
                diff = diff[1] if len(diff[0]) > len(diff[1]) else diff[0]
                exp = 'Missing column:<br><code>{0}</code>' \
                    .format('<br>'.join(diff))

        except Exception as e:
            logger.exception(e)
            result, exp = None, str(e)
        finally:
            return (result, exp)

    def check_11_analysis(self, data):
        '''Check test analysis item mapping with data in _IE.html'''
        def is_condition(key):
            key = '{0}-'.format(key)
            lst = [k for k in data.keys() if k.startswith(key)]
            return len(lst) > 0

        logger.debug("Check 11_analysis")
        try:
            result, exp = True, []
            for i in range(len(self.table)):
                row = self.table[i]
                no = self.get_number(row[self.ino])
                if no is None:
                    continue

                item = self.get_node(row[self.iai])
                cid = self.get_node(row[self.iid])
                comment = self.get_node(row[self.icm])

                if item not in data.keys():
                    # Ignore condition analysis item
                    if is_condition(item) is True:
                        continue
                    result = False
                    msg = 'Not found test analysis item <code>{0}</code>'\
                        .format(item)
                    exp.append(msg)

                # Check ID column
                if cid != data.get(item).get('id'):
                    result = False
                    msg = 'No.{0} ID diff <code>{1} != {2}</code>' \
                        .format(no, data.get('item').get('id'), cid)
                    exp.append(msg)

                # Check comment column
                r, _ = self.get_origin(row[self.iai])
                if r == i and comment != data.get(item).get('comment'):
                    result = False
                    msg = 'No.{0} Comment diff <code>{1} != {2}</code>' \
                        .format(no, data.get(item).get('comment'), comment)
                    exp.append(msg)

                if r != i and comment.strip() != '':
                    result = False
                    msg = 'No.{0}. Comment should be empty'.format(no)
                    exp.append(msg)

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]
        finally:
            return (result, '<br>'.join(exp))


class FileIE(Table):
    # _IE.html

    class_uniqid = ['input-tp-uniqid-w', 'input-tp-uniqid-s']
    class_uniqid_sub = ['input-tp-uniqid-s']
    class_comment = ['input-tp-comment']
    class_id = ['input-tp-id']
    class_kind = ['title-input-kind', 'title-input-kind-r']
    class_name = ['title-input-name', 'title-input-name-r']

    def __init__(self, path):
        super().__init__(path)
        self.table = self.get_table(0)
        self.input_vars = self.get_vars()
        self.data_analysis = self.get_analysis_item()
        self.lst_label = utils.get_label_list(self.get_labels())

    def get_analysis_item(self):
        '''Get data analysis item'''
        logger.debug("Get analysis item from %s", self.path.name)
        data = {}
        item = {}
        for row in self.table:
            if row[0].get('class') not in self.class_uniqid:
                continue

            sub = row[0].get('class') in self.class_uniqid_sub

            if row[1].get('class') in self.class_id:
                item = {
                    'item': self.get_node(row[0]),
                    'id': self.get_node(row[1]),
                    'sub': sub
                }

            elif row[1].get('class') in self.class_comment:
                item.update({
                    'comment': self.get_node(row[1])
                })
                data.update({
                    item.get('item'): item
                })
        return data

    def get_vars(self):
        '''Get input variables [classification, name, vartype]'''
        logger.debug("Get variables from %s", self.path.name)
        data = []

        for i in range(len(self.table[1])):
            if (1, i) != self.get_origin(self.table[1][i]):
                continue
            if self.table[1][i].get('class') not in self.class_kind:
                continue

            classification = self.get_node(self.table[1][i])
            name = self.get_node(self.table[2][i])
            vartype = self.get_node(self.table[3][i])

            data.append([classification, name, vartype])

        return data

    def get_labels(self, char=','):
        '''Get set of label'''
        logger.debug('Get set of label from %s', self.path.name)
        lst = []
        for info in self.data_analysis.values():
            tmp = [l.strip() for l in info.get('id').replace(';', ',').split(char)
                   if l.strip() != '']
            lst.extend(tmp)
        return list(set(lst))

    def check_5_input_var(self, lst):
        '''Check variables'''
        def is_same_var(var1, var2):

            result = var1 == var2

            # Argument
            if result is False:
                result = var1.startswith('@') and var2.endswith(var1)

            # Pointer
            if result is False:
                if var1.startswith('$') and var2.startswith('$'):
                    result = is_same_var(var1[1:], var2[2:])

            return result

        logger.debug("Check 5_input_var")
        try:
            result, exp = True, []
            var = [i[1] for i in self.input_vars]
            if len(var) != len(lst):
                result = False
                msg = 'Length diff <code>{0} != {1}<code>'\
                    .format(len(var), len(lst))
                exp.append(msg)
            else:
                for i in range(len(var)):
                    if is_same_var(var[i], lst[i]) is False:
                        result = False
                        msg = 'Index.{0} diff <code>{1} != {2}</code>' \
                            .format(i+1, var[i], lst[i])
                        exp.append(msg)

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]

        finally:
            return (result, '<br>'.join(exp))

    def check_6_label(self, char=','):
        '''Check label'''
        logger.debug("Check 6_label")
        try:
            result, exp = True, []

            for item, info in self.data_analysis.items():
                idstr = info.get('id').replace(';', ',')
                lst = [l.strip() for l in idstr.split(char)
                       if l.strip() != '']
                for label in lst:
                    if label not in self.lst_label:
                        result = False
                        exp = ['Unknown labels:'] if exp == [] else exp
                        msg = "<code>{0}</code> in item <code>{1}</code>" \
                            .format(label, item)
                        exp.append(msg)

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]

        finally:
            return (result, '<br>'.join(exp))


class FileReport(Table):
    # TestReport.html

    def __init__(self, path):
        super().__init__(path)
        self.info = self.get_info()

    def check_12_entire(self, info):
        '''Check Entire Information table compare with testlog file'''
        logger.debug("Check 12_entire")
        try:
            result, exp = True, []

            if self.info == {}:
                result = None
                exp = ['Unable to get info from {0}'.format(self.path.name)]

            else:

                info.update({
                    'csv': '{func}.csv'.format(**info),
                    'title': info.get('func')
                })

                dct = {
                    'c0': 'C0',
                    'c1': 'C1',
                    'mcdc': 'MC/DC',
                    'csv': 'Top CSV Filename',
                }

                for key, value in self.info.items():
                    if value != info.get(key):
                        result = False
                        msg = '{0}: <code>{1} ! {2}</code>' \
                            .format(dct.get(key), info.get(key), value)
                        exp.append(msg)

                if result is False:
                    exp = ["Testlog vs TestReport.html"] + exp

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]
        finally:
            return (result, '<br>'.join(exp))

    def get_info(self):
        '''Get info from tables'''
        try:
            data = {}
            tbl_0 = self.get_table(0)
            tbl_2 = self.get_table(2)

            data.update({
                'csv': Path(self.get_node(tbl_0[0][1]).strip()).name,
                # 'tile': self.get_node(tbl_0[1][1]).strip(),
                # 'func_full': self.get_node(tbl_2[1][0]).strip(),
                'c0': self.get_node(tbl_2[1][1]).strip(),
                'c1': self.get_node(tbl_2[1][2]).strip(),
                'mcdc': self.get_node(tbl_2[1][3]).strip()
            })

        except Exception as e:
            logger.exception(e)
            data = data if isinstance(data, dict) else {}

        finally:
            return data


class FileIO(Table):
    # _IO.html
    class_input = ['title-input-kind', 'title-input-kind-r']
    class_output = ['title-output-kind', 'title-output-kind-r']
    class_io_tp = ['io-input-tp', 'io-input-tp-r']

    def __init__(self, path):
        super().__init__(path)
        self.table = self.get_table(0)

    def get_io_vars(self):
        '''Get input/ouput variable'''
        data = {
            'input': [],
            'output': []
        }

        for i in range(len(self.table[2])):
            r, c = self.get_origin(self.table[2][i])
            if (r, c) != (2, i):
                continue

            if self.table[2][i].get('class') in self.class_input:
                data['input'].append([
                    self.get_node(self.table[2][i]),
                    self.get_node(self.table[3][i]),
                    self.get_node(self.table[4][i]),
                    i
                ])

            elif self.table[2][i].get('class') in self.class_output:
                data['output'].append([
                    self.get_node(self.table[2][i]),
                    self.get_node(self.table[3][i]),
                    self.get_node(self.table[4][i]),
                    i
                ])

        return data

    def get_var_ai(self, col):
        '''Get variable was used in test analysis item'''
        data = []
        for row in self.table:
            if row[col].get('class') not in self.class_io_tp:
                continue
            text = self.get_node(row[col])
            tmp = [i.strip() for i in text.split(',') if i.strip() != '']
            data += tmp
        data = sorted(list(set(data)))
        return data

    def check_16_io_var(self, data):
        '''Check input/output variable in _IO.html vs _Table.html'''
        logger.debug("Check 16_io_var")
        try:
            result, exp = True, []
            io_vars = self.get_io_vars()

            for key in ['input', 'output']:
                lst = [i[1] for i in io_vars[key]]
                lst_all = lst + data[key]

                diff = [i for i in lst_all
                        if i not in lst or i not in data[key]]

                if diff != []:
                    result = False
                    msg = '{0} diff:<br> <code>{1}</code>'\
                        .format(key.title(), '<br>'.join(diff))
                    exp.append(msg)

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]
        finally:
            return (result, '<br>'.join(exp))


class FileOE(Table):
    # _OE.html

    def __init__(self, path):
        super().__init__(path)
        self.table = self.get_table(0)


class FileXlsx(Base):
    # .xlsx

    def __init__(self, path):
        super().__init__(path)
        self.list_sheet = parse.get_xlsx_sheets(self.path)
        self.summary = self.get_test_result()

    def check_html(self, sheet, data):
        '''Compare data in excel with html'''
        def mod(value, form='NFKD'):
            try:
                value = float(value)
                value = 0 if value == 0 else value
            except:
                value = value

            value = '' if value is None else str(value)
            value = value.strip() if value.strip() == '' else value
            return normalize(form, value)

        logger.debug("Check sheet %s", sheet)
        try:
            result, exp = True, []

            if sheet not in self.list_sheet:
                result = False
                msg = "Not found sheet <code>{0}</code>".format(sheet)
                exp.append(msg)

            else:
                dxlsx = parse.get_xlsx_raw(self.path, sheet)

                # Compare size
                size_html = '[{0}x{1}]'.format(len(data), len(data[0]))
                size_xlsx = '[{0}x{1}]'.format(len(dxlsx), len(dxlsx[0]))
                msg = "Size: <code>{0} != {1}</code>" \
                    .format(size_html, size_xlsx)
                if len(dxlsx) < len(data) or len(dxlsx[0]) < len(data[0]):
                    result = False
                    exp.append(msg)

                # Compare cell by cell
                for i in range(min(len(data), len(dxlsx))):
                    for j in range(min(len(data[i]), len(dxlsx[i]))):
                        cd = mod(data[i][j])
                        cx = mod(dxlsx[i][j])
                        if cd == cx:
                            continue
                        else:
                            result = False
                            msg = "Cell [{0}x{1}] <code>{2} != {3}</code>" \
                                .format(i+1, j+1, data[i][j], dxlsx[i][j])
                            exp.append(msg)
                            if len(exp) > 5:
                                break

                    if len(exp) > 5:
                        break

                if result is False:
                    exp = ['HTML vs XLSX'] + exp

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [str(e)]
        finally:
            return (result, '<br>'.join(exp))

    def check_21_sheet_testlog(self, sheet, testlog):
        '''Compare testlog sheet vs testlog txt'''
        def mod(line):
            line = '' if line is None else line
            line = str(line).replace('\n', 'N').rstrip()
            if line.strip().isdigit():
                line = line.strip()
            return line

        def clean(data):
            '''Remove last empty line'''
            while len(data) > 0 and \
                    (data[-1] is None or str(data[-1]).strip() == ''):
                data = data[:-1]
            return data

        logger.debug("Check sheet %s", sheet)
        try:
            result, exp = True, []

            if sheet not in self.list_sheet:
                result = False
                msg = "Not found sheet <code>{0}</code>".format(sheet)
                exp.append(msg)

            else:
                dxlsx = []
                for line in parse.get_xlsx_raw(self.path, sheet):
                    if len(line) > 0:
                        dxlsx.append(line[0])
                    else:
                        dxlsx.append('')
                dxlsx = clean(dxlsx)

                data = [l.replace('\n', '') for l in utils.read_file(testlog)]
                intro = utils.load(CONST.SETTING, 'jpDict.testlog_intro')
                data = [intro, '', '', ''] + data
                data = clean(data)

                # Compare lines
                if len(dxlsx) < len(data):
                    result = False
                    msg = "Number of lines <code>{0} != {1}</code>" \
                        .format(len(data), len(dxlsx))
                    exp.append(msg)

                # Compare line by line
                for i in range(min(len(dxlsx), len(data))):
                    if mod(dxlsx[i]) == mod(data[i]):
                        continue

                    result = False
                    msg = "Row.{0}<br><code>{1}<br>{2}</code>" \
                        .format(i+1, data[i].rstrip(), dxlsx[i])
                    exp.append(msg)
                    break

                if result == False:
                    exp = ['TXT vs XLSX'] + exp

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [(str)]
        finally:
            return (result, '<br>'.join(exp))

    def check_test_result(self, info, keyword='Summary'):
        '''Compare test result with testlog and summary'''
        def mod(value):
            if isinstance(value, int) or isinstance(value, float):
                value = float(value)
            return str(value).strip()

        logger.debug("Check table 1.1 in sheet")

        try:
            result, exp = True, []

            sheet = utils.load(CONST.SETTING, 'specSheets.spec')
            if sheet not in self.list_sheet:
                result = False
                msg = "Not found sheet <code>{0}</code>".format(sheet)
                exp.append(msg)

            else:
                for key, value in info.items():

                    if key == 'issue' and (value is None or value.strip() == ''):
                        value = utils.load(CONST.SETTING, 'jpDict.noissue')

                    strkey = key.title() if key != 'mcdc' else 'MC/DC'

                    if mod(value) != mod(self.summary.get(key, '')):
                        result = False
                        msg = "{0}: <code>{1} != {2}</code>" \
                            .format(strkey, value, self.summary.get(key))
                        exp.append(msg)

                if result == False:
                    exp = ['{0} vs XLSX'.format(keyword)] + exp

        except Exception as e:
            logger.exception(e)
            if result is True:
                result, exp = None, [(str)]
        finally:
            return (result, '<br>'.join(exp))

    def get_test_result(self):
        '''Get test result from test spec'''
        try:
            logger.debug("Get test result from table 1.1")
            sheet = utils.load(CONST.SETTING, 'specSheets.spec')
            list_cell = ['F8', 'F9', 'F10', 'F11', 'F12']
            info = parse.get_xlsx_cells(self.path, sheet, list_cell)

            # Replace JP char
            colon = utils.load(CONST.SETTING, 'jpDict.colon')
            f12 = info.get('F12')
            for char, newchar in [(colon, ':'), ('%_x000D_', '%'), ('_x000D_', '')]:
                f12 = f12.replace(char, newchar)

            lines = [i.strip() for i in f12.split('\n')]
            lst = ['result', 'c0', 'c1', 'mcdc', 'issue']

            data = {lst[i]: lines[i][lines[i].index(':')+1:].strip()
                    for i in range(len(lst))}

            data.update({
                'src_rel_dir': info.get('F8'),
                'src_name': info.get('F9'),
                'func': info.get('F10'),
                'csv': info.get('F11')
            })

        except Exception as e:
            logger.exception(e)
            data = {}
        finally:
            return data


class Report(Base):

    def __init__(self, path):
        super().__init__(path)

        self.info = parse.parse_testlog(self.path)
        self.collection = FileCollection(self.path)

        self.files = self.collection.get_files(self.info)

        self.checklist = {}
        self.chklist = utils.get_lang_data().get('checklist', {})
        self.check_files_exist()

        self.csv = self.init(FileCsv, 'csv')
        if self.csv != None and self.csv.stub_info.get('stub', []) == []:
            self.remove_key('stub')

        if self.collection.is_ams is True:
            self.remove_key('xlsx')

    def init(self, clsname, keyword):
        '''Init object'''
        try:
            filepath = self.files.get(keyword)
            if filepath != None and Path(filepath).is_file():
                return clsname(filepath)
        except Exception as e:
            logger.exception(e)

    def check(self, package):
        '''Generate checklist'''
        # Check .txt
        try:
            self.txt = self.init(FileTxt, 'testlog')

            if self.txt is not None:
                rst = self.txt.check_13_parse()
                self.update_checklist('testlog', '13_parse', *rst)

                rst = self.txt.check_14_func()
                self.update_checklist('testlog', '14_func', *rst)
        except Exception as e:
            logger.exception(e)

        # Check .csv
        try:
            if self.csv != None:
                rst = self.csv.check_2_desc(self.info.get('src_full'), package)
                self.update_checklist('csv', '2_desc', *rst)

                rst = self.csv.check_3_init_opt()
                self.update_checklist('csv', '3_init_opt', *rst)

                rst = self.csv.check_4_init_var()
                self.update_checklist('csv', '4_init_var', *rst)
        except Exception as e:
            logger.exception(e)

        # Check _Table.html
        try:
            self.tctbl = self.init(FileTable, 'table')
            if self.tctbl != None:
                rst = self.tctbl.check_7_title(self.info.get('func'))
                self.update_checklist('table', '7_title', *rst)

                rst = self.tctbl.check_8_index()
                self.update_checklist('table', '8_index', *rst)

                rst = self.tctbl.check_9_confirm()
                self.update_checklist('table', '9_confirm', *rst)

                rst = self.tctbl.check_10_header()
                self.update_checklist('table', '10_header', *rst)
        except Exception as e:
            logger.exception(e)

        # Check _IE.html
        try:
            self.ietbl = self.init(FileIE, 'ie')

            if self.ietbl != None:

                lst = self.csv.io_vars.get('input', [])
                rst = self.ietbl.check_5_input_var(lst)
                self.update_checklist('ie', '5_input_var', *rst)

                rst = self.ietbl.check_6_label()
                self.update_checklist('ie', '6_label', *rst)

        except Exception as e:
            logger.exception(e)

        # Check TestReport.htm
        try:
            self.trtbl = self.init(FileReport, 'report_html')

            if self.trtbl != None:
                rst = self.trtbl.check_12_entire(self.info)
                self.update_checklist('report_html', '12_entire', *rst)
        except Exception as e:
            logger.exception(e)

        # Check Test Analysis Item _Table.html mapping with _IE.html
        try:
            if self.ietbl != None and self.tctbl != None:
                data_analysis = self.ietbl.data_analysis
                rst = self.tctbl.check_11_analysis(data_analysis)
                self.update_checklist('table', '11_analysis', *rst)
        except Exception as e:
            logger.exception(e)

        # Check input/output variable in _Table.html mapping with .csv
        try:
            if self.tctbl != None:
                rst = self.tctbl.check_15_io_var(self.csv.io_vars)
                self.update_checklist('table', '15_io_var', *rst)
        except Exception as e:
            logger.exception(e)

        # Check input/output variable in _Table.html mapping with _IO.html
        try:
            self.iotbl = self.init(FileIO, 'io')

            if self.iotbl != None:
                rst = self.iotbl.check_16_io_var(self.tctbl.get_io_vars())
                self.update_checklist('io', '16_io_var', *rst)
        except Exception as e:
            logger.exception(e)

        # Check .xlsx
        spec_sheets = utils.load(CONST.SETTING, 'specSheets')
        self.xlsx = self.init(FileXlsx, 'xlsx')
        self.oetbl = self.init(FileOE, 'oe')

        lst = [
            ('io', self.iotbl, '17_sheet_io'),
            ('table', self.tctbl, '18_sheet_table'),
            ('oe', self.oetbl, '19_sheet_oe'),
            ('ie', self.ietbl, '20_sheet_ie')
        ]

        if self.xlsx != None:

            for key, obj, item in lst:
                logger.debug("Check %s %s", key, item)
                try:
                    filepath = self.files.get(key)
                    if filepath.is_file() is False:
                        rst = None, "File not found {0}".format(filepath.name)
                    else:
                        data = obj.get_table_raw(obj.table)
                        rst = self.xlsx.check_html(spec_sheets.get(key), data)

                    self.update_checklist('xlsx', item, *rst)
                except Exception as e:
                    logger.exception(e)

            # Check testlog sheet
            try:
                sheet = spec_sheets.get('testlog')
                rst = self.xlsx.check_21_sheet_testlog(
                    sheet, self.files.get('testlog'))
                self.update_checklist('xlsx', '21_sheet_testlog', *rst)
            except Exception as e:
                logger.exception(e)

            # Check table 1.1 vs summary
            try:
                logger.debug("Check %s", '22_result_summary')
                summary = db.get_func_info(self.files.get('testlog'), package)

                if 'func_no' not in summary.keys():
                    rst = (None, "Unable to get function info from summary")

                else:
                    src_rel_dir = '/'.join(Path(summary['src_rel']).parts[:-1])
                    summary.update({'src_rel_dir': src_rel_dir})

                    # Update c0, c1, mcdc
                    for key in ['c0', 'c1', 'mcdc']:
                        value = summary.get(key, 0)
                        if isinstance(value, int) or isinstance(value, float):
                            value = float(value)*100
                            value = '{0}%'.format(int(value))
                            summary.update({key: value})

                    lst = ['c0', 'c1', 'mcdc', 'result', 'issue',
                           'src_rel_dir', 'src_name', 'func']
                    info = {key: summary[key] for key in lst}

                    if self.xlsx.summary != {}:
                        rst = self.xlsx.check_test_result(info, 'Summary')
                    else:
                        rst = None, "Unable to parse table 1.1 to get info"

                self.update_checklist('xlsx', '22_result_summary', *rst)
            except Exception as e:
                logger.exception(e)

            # Check table 1.1 vs testlog
            try:
                try:
                    num_issue = summary.get('issue_num')
                    if num_issue is None or str(num_issue) == '':
                        num_issue = 0

                    num_issue = int(num_issue)
                except Exception as e:
                    logger.exception(e)
                    num_issue = 1

                if num_issue == 0:
                    for key in ['c0', 'c1', 'mcdc']:
                        if self.info[key] != '100%':
                            num_issue = 1

                if num_issue == 0 and self.tctbl != None \
                        and self.tctbl.get_confirm() == 'Fault':
                    num_issue = 1

                if num_issue > 0:
                    result = 'NG'
                    lst = ['{0}.{1}'.format(summary.get('func_no', 0), i+1)
                           for i in range(num_issue)]
                    issuestr = ', '.join(lst)
                    xlsxissuestr = utils.load(CONST.SETTING, 'jpDict.issue')
                    issuestr = '{0}_{1}No{2}'\
                        .format(summary.get('package'), xlsxissuestr, issuestr)
                else:
                    result = 'OK'
                    issuestr = utils.load(CONST.SETTING, 'jpDict.noissue')

                logger.debug("Check %s", '23_result_testlog')

                self.info.update({
                    'csv': '{func}.csv'.format(**self.info),
                    'result': result,
                    'issue': issuestr
                })

                lst = ['c0', 'c1', 'mcdc', 'result', 'issue',
                       'src_name', 'func', 'csv']
                info = {key: self.info[key] for key in lst}

                if self.xlsx.summary != {}:
                    rst = self.xlsx.check_test_result(info, 'Testlog')
                else:
                    rst = None, "Unable to parse table 1.1 to get info"

                self.update_checklist('xlsx', '23_result_testlog', *rst)
            except Exception as e:
                logger.exception(e)

    def update_checklist(self, ftype, item, value, explain=''):
        '''Update checklist'''
        dct = self.checklist.get(ftype, {})
        dct.update({
            item: [value, explain]
        })
        self.checklist.update({ftype: dct})

    def check_files_exist(self):
        '''Check files exist or not'''
        for key, filepath in self.files.items():
            if Path(filepath).is_file() is True:
                self.update_checklist(key, '1_exist', True)
            else:
                msg = "File not found <code>{0}</code>" \
                    .format(Path(filepath).name)

                for item in self.chklist.get(key):
                    result = False if item[0] == '1_exist' else None
                    self.update_checklist(key, item[0], result, msg)

    def remove_key(self, key):
        '''Remove key from checklist'''
        for dct in [self.files, self.checklist, self.chklist]:
            if key in dct:
                del dct[key]

    def deliver_files(self, target):
        '''Deliver files'''
        logger.debug("Deliver files to %s", target)
        wlogger("Delivering test result files to {0}", Path(target).name, 1)
        progress = 1
        for key, src in self.files.items():
            try:
                # Do not deliver excel file
                if key == 'xlsx':
                    continue

                dst = Path(target).joinpath(src.name)
                utils.copy(src, dst)

                wlogger("{0}", dst.name, progress)
            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

            progress += 1

    def generate_spec(self, options):
        '''Generate test spec'''
        logger.debug("Generate test spec")
        try:
            # Copy template
            filename = '{func}.xlsx'.format(**self.info)
            filespec = Path(options.get('spec')).joinpath(filename)

            template = Path(options.get('template'))
            if template is None or Path(template).is_file() is False:
                template = CONST.SPEC
            wlogger("Generating test spec {0}", filename, 15)
            utils.copy(template, filespec)

            # Copy .html file to excel
            spec_sheets = utils.load(CONST.SETTING, 'specSheets')
            excel = ExcelWin32(filespec)
            vba_err = "Error VBA {0}"

            progress = 20
            for key in ['testlog', 'table', 'io', 'oe', 'ie']:
                try:
                    sheet = spec_sheets[key]
                    cell = 'A5' if key == 'testlog' else 'A1'
                    cmd = 'Import_HTML_File'
                    if key == 'testlog':
                        cmd = 'Import_Coverage_File'
                    path = self.files.get(key)
                    paras = [filename, sheet, cell, str(path)]

                    wlogger("Copying {0} to sheet {1}", [
                            path.name, sheet], progress)
                    if path.is_file():
                        rst = excel.run(sheet, cmd, paras)
                        wlogger(vba_err, rst) if rst != 'vba_ok' else None
                    else:
                        wlogger("Error file not found {0}", path.name)

                    cmd = 'Fmt_correction'
                    if key == 'table':
                        paras = [filename, sheet, 'TC']
                        rst = excel.run(sheet, cmd, paras)
                        wlogger(vba_err, rst) if rst != 'vba_ok' else None
                    elif key == 'ie':
                        paras = [filename, sheet, 'IE']
                        rst = excel.run(sheet, cmd, paras)
                        wlogger(vba_err, rst) if rst != 'vba_ok' else None
                    #phi update 14/11/2019
                    cmd = 'Remove_Row_Description'
                    if key == 'io' or key == 'oe' or key == 'ie' :
                        paras = [filename, sheet]
                        rst = excel.run(sheet, cmd, paras)
                        wlogger(vba_err, rst) if rst != 'vba_ok' else None


                except Exception as e:
                    logger.exception(e)
                    wlogger("Exception {0}", str(e))

                progress += 10

            # Update test spec sheet
            sheet = spec_sheets.get('spec')

            wlogger("Updating sheet {0} in {1}", [sheet, filename])

            # Update table 1.5
            cmd = 'Import_StubFile'
            try:
                if self.csv.stub_info.get('stub') != []:
                    wlogger("Table {0}: Updating", '1.5', 65)
                    filepath = self.files.get('stub')
                    paras = [filename, sheet, 'B54', str(filepath)]
                    if filepath.is_file():
                        rst = excel.run(sheet, cmd, paras)
                        wlogger(vba_err, rst) if rst != 'vba_ok' else None
                    else:
                        wlogger(
                            "Error file not found {0}", (filepath.name, ''))
                else:
                    wlogger("Table {0}: Default", '1.5')
            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

            # Update table 1.4
            cmd = 'Fill_Table_1_4'
            try:
                stub_data = self.csv.stub_info.get('stub')[:]
                stub_data += self.csv.stub_info.get('non_stub')
                data = []
                for lst in stub_data:
                    lst.reverse()
                    if lst[1] == '':
                        lst[1] = '-'
                    data.append(lst)

                if data != []:
                    paras = [filename, sheet, 'B47', data]
                    wlogger("Table {0}: Updating", '1.4', 70)
                    rst = excel.run(sheet, cmd, paras)
                    wlogger(vba_err, rst) if rst != 'vba_ok' else None
                else:
                    wlogger("Table {0}: Default", '1.4')
            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

            # Update table 1.3
            cmd = 'Fill_Table_1_3'
            try:
                init_data = self.csv.init_data
                if len(init_data) > 1:
                    init_data = init_data[1:]
                    paras = [filename, sheet, 'B40', init_data]
                    wlogger("Table {0}: Updating", '1.3', 75)
                    rst = excel.run(sheet, cmd, paras)
                    wlogger(vba_err, rst) if rst != 'vba_ok' else None
                else:
                    wlogger("Table {0}: Default", '1.3')
            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

            # Update table 1.2
            cmd = 'Fill_Table_1_2'
            try:
                wlogger("Table {0}: Updating", '1.2', 80)

                label_list = FileIE(self.files.get('ie')).lst_label
                data_spec = self.get_spec_data(label_list)

                for i in range(len(label_list)):
                    label = label_list[i]
                    data_point = data_spec.get(label, [])
                    if data_point != []:
                        paras = [filename, sheet, 'B17', i+1, data_point]
                        rst = excel.run(sheet, cmd, paras)
                        wlogger(vba_err, rst) if rst != 'vba_ok' else None

            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

            # Update table 1.1
            cmd = 'Fill_cell'
            try:
                wlogger("Table {0}: Updating", '1.1', 85)
                data_sum = self.get_summary_data(options)

                for cell, value in data_sum.items():
                    paras = [filename, sheet, cell, value]
                    rst = excel.run(sheet, cmd, paras)
                    wlogger(vba_err, rst) if rst != 'vba_ok' else None

            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

            # Update format
            cmd = 'Fill_fault_conditional_format_n_reset_pointer'
            try:
                wlogger("Formatting {0}", filename, 90)
                paras = [filename]
                rst = excel.run(sheet, cmd, paras)
                wlogger(vba_err, rst) if rst != 'vba_ok' else None
            except Exception as e:
                logger.exception(e)
                wlogger("Exception {0}", str(e))

        except Exception as e:
            logger.exception(e)
            wlogger("Exception {0}", str(e))
        finally:
            wlogger("Done !!! Please double check !!!", '', 100)
            excel.close()

    def get_spec_data(self, label_list):
        '''Get data to fill table 1.2 test spec'''
        def get_var_tc_data(index, label):
            '''Get testcase of variable'''
            lst_item = iotbl.get_var_ai(index)
            lst_tc = []
            for item in set(lst_item):
                if label in data_ai.get(item).get('id'):
                    lst_tc += data_tc.get(item)
            return lst_tc

        def get_var_str(var, default=None):
            '''Generate var data to write to table 1.2'''
            var[1] = var[1].replace('AMSTB_SrcFile.c/', '')
            src_name = self.info.get('src_name', '')
            if var[1].startswith(src_name):
                var[1] = var[1].replace('{0}/'.format(src_name), '')
            str_var = '{0} {1}'.format(var[2], var[1])
            str_tc = utils.collapse_list(var[-1])
            if default != None and str_tc == '':
                str_tc = default
            return [str_var, str_tc]

        logger.debug("Get data to fill table 1.2 %s", label_list)
        try:
            data = {}

            iotbl = FileIO(self.files.get('io'))
            ietbl = FileIE(self.files.get('ie'))
            tctbl = FileTable(self.files.get('table'))

            data_tc = tctbl.get_testcase_data()
            data_ai = ietbl.get_analysis_item()
            data_var = iotbl.get_io_vars()

            # Get all label spec info
            data_label = {}
            for item, info in data_ai.items():
                idstr = info.get('id').replace(';', ',')
                lst = [l.strip() for l in idstr.split(',')
                       if l.strip() != '']
                for label in lst:
                    tmp = [info['comment'], utils.collapse_list(
                        data_tc.get(item, []))]
                    tmp2 = data_label.get(label, [])
                    tmp2.append(tmp)
                    data_label.update({label: tmp2})

            for key, value in data_label.items():
                lst = [c for c in value if c[1] != '']
                data.update({key: lst})

            # Get input variable and AMIN (point 1 and point 2)
            lst_input = []
            point_1 = label_list[0]
            point_2 = label_list[1]
            for var in data_var.get('input'):
                if var[2] == 'Number of elements':
                    continue
                label = point_1 if '@AMIN' not in var[1] else point_2
                tmp = [int(i) for i in get_var_tc_data(var[3], label)]
                var.append(tmp)

                lst_input.append(var)

            data_var_spec = [get_var_str(var) for var in lst_input]
            data.update({
                point_1: [var for var in data_var_spec if '@AMIN' not in var[0]],
                point_2: [var for var in data_var_spec if '@AMIN' in var[0]]
            })

            # Get amount (point 10)
            point_10 = label_list[9]
            data_p10 = data.get(point_10, [])
            data_p10 += [get_var_str(var + [[]], 'PLEASE UPDATE') for var in data_var.get('output')
                         if '@AMOUT' in var[1]]

            data.update({
                point_10: data_p10
            })

        except Exception as e:
            logger.exception(e)
        finally:
            return data

    def get_summary_data(self, options):
        '''Update table 1.1'''
        logger.debug("Get summary data")

        package = options.get('package')
        summary = db.get_func_info(self.files.get('testlog'), package)

        src_rel = summary.get('src_rel', self.info.get('src_rel'))
        data = {
            'F8': '/'.join(Path(src_rel).parts[:-1]),
            'F9': self.info['src_name'],
            'F10': self.info['func'],
            'F11': '{func}.csv'.format(**self.info)
        }

        result = utils.load(CONST.SETTING, 'jpDict.result')
        num_issue = int(options.get('issue', '0'))
        if num_issue > 0:
            confirm = 'NG'
            lst = ['{0}.{1}'.format(options.get('func_no', 0), i+1)
                   for i in range(num_issue)]
            issue = ', '.join(lst)
            xlsxissuestr = utils.load(CONST.SETTING, 'jpDict.issue')
            issue = '{0}_{1}No{2}'\
                .format(summary.get('package'), xlsxissuestr, issue)
        else:
            confirm = 'OK'
            issue = utils.load(CONST.SETTING, 'jpDict.noissue')

        self.info.update({
            'confirm': confirm,
            'issue': issue
        })
        result = result.format(**self.info)
        data.update({'F12': result})

        return data

    def get_checklist(self):
        '''Get checklist'''
        logger.debug("Get checklist")
        chk = {}
        tooltip = utils.get_lang_data('tooltip')
        for file_type, list_item in self.chklist.items():
            name = self.files.get(file_type).name
            dct = self.checklist.get(file_type, {})
            lst = []
            for item, desc in list_item:
                exp = tooltip.get(item, desc)
                tmp = [desc] + dct.get(item, [None, '']) + [exp]
                lst.append(tmp)
            chk.update({name: lst})

        rst = []
        for filename, lst_item in chk.items():
            for item in lst_item:
                rst.append([filename] + item)

        return rst

    def get_label_data(self):
        '''Get lable data'''
        logger.debug("Get label data")
        tbl = FileTable(self.files.get('table'))
        data_tc = tbl.get_testcase_data()

        tbl = FileIE(self.files.get('ie'))
        data_ie = tbl.data_analysis

        data = {}
        for _, dct in data_ie.items():
            idstr = dct.get('id').replace(';', ',')
            lst = [l.strip() for l in idstr.split(',')
                   if l.strip() != '']
            dct.update({
                'tc': data_tc.get(dct.get('item', []))
            })
            for label in lst:
                tmp = data.get(label, [])
                tmp.append(dct)
                data.update({label: tmp})

        return data
