# -*- coding: utf-8 -*-


__authors__ = 'Duy Nguyen'
__contact__ = 'duykn@gcs-vn.com'


import logging
from unicodedata import normalize

import cssutils
import lxml.html

import config as CONF
import utils
from style import Style
from utils import format_string as fstr

cssutils.log.setLevel(logging.CRITICAL)


class Base(object):

    def __init__(self, path):
        self.path = path
        self.errors = {}

    def add_error(self, err, desc):
        logging.debug("%s %s", err, desc)

        lst = self.errors.get(err, [])
        lst.append(desc)
        self.errors[err] = lst

    def print_error(self):
        '''Print errors'''
        logging.error("Found %s errors in '%s'",
                      len(self.errors), fstr(self.path, 0))

        for err, dsc in self.errors.items():
            err = '{0:20}'.format(err)
            if len(dsc) == 1 and isinstance(dsc[0], list):
                dsc = dsc[0]
            logging.error("%s %s", fstr(err), dsc)


class Table(Base):

    def __init__(self, path):
        super().__init__(path)
        self.doc = lxml.html.parse(str(path))
        self.table = self.get_table()

        self.css = self.get_css()

    def get_title(self):
        '''Get title of page'''
        h4 = self.doc.find('body/h4')
        return self.get_text(h4)

    def get_text(self, node, form='NFKD'):
        '''Normalize text'''
        return normalize(form, node.text_content())

    def get_table(self):
        node = self.doc.find('body/table')
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

    def get_origin(self, cell):
        '''Get origin row, col of cell'''
        return tuple([int(v) for v in cell.get('origin').split(',')])

    def get_header(self):
        return [self.get_text(cell) for cell in self.table[0]]

    def get_index_header(self, header):
        return self.get_header().index(header)

    def check_header(self, lst):
        logging.info(fstr('Checking header'))

        diff = set(lst) - set(self.get_header())
        if len(diff) > 0:
            dsc = ', '.join(diff)
            self.add_error('Missing column', dsc)

    def get_css(self):
        '''Convert css rule to dict'''
        def parse_line(text, char=':'):
            lst = text.split(char)
            return lst[0].strip(), char.join(lst[1:]).strip()

        def parse_rule(text, char=';'):
            result = {}
            for rule in text.split(char):
                key, value = parse_line(rule)
                result[key] = value
            return result

        try:
            node = self.doc.find('head/style')
            csstext = node.text_content()
            lst_rule = cssutils.parseString(csstext)

            data = {}
            for rule in lst_rule:
                selector = rule.selectorText
                styles = rule.style.cssText
                data[selector.strip().lower()] = parse_rule(styles)
        except Exception as e:
            logging.exception(e)
            # if utils.LOG_LEVEL < 20:
            #     logging.exception(e)

            data = {}

        return data

    def get_css_td(self, node):
        '''Get style of node'''
        tdstyle = 'td.{0}'.format(node.get('class'))

        style = self.css.get('body')
        style.update(self.css.get('table'))
        style.update(self.css.get('td'))
        style.update(self.css.get(tdstyle))

        return style

    def write_xlsx(self, ws, row=3):
        '''Write html table to excel'''

        for r in range(len(self.table)):
            for c in range(len(self.table[r])):
                node = self.table[r][c]
                if (r, c) != self.get_origin(node):
                    continue

                self.write_xlsx_cell(ws, node, row)
                print('.', end='')
        else:
            print('.')

    def write_xlsx_cell(self, ws, node, row):
        '''Write cell to excel'''
        r, c = self.get_origin(node)
        row = r + row
        col = c + 1

        cell = ws.cell(row=row, column=col)
        try:
            value = int(self.get_text(node))
        except:
            value = self.get_text(node)
            if node.get('class') in CONF.TCCLS_HEAD and '/' in value:
                value = value.replace('/', '\n/')

        cell.value = value

        # Apply style
        cell_style = Style(self.get_css_td(node))
        cell.font = cell_style.font
        cell.fill = cell_style.fill
        cell.border = cell_style.border
        cell.alignment = cell_style.alignment

        # Merge cells
        end_row = row + int(node.get('rowspan', 1)) - 1
        end_col = col + int(node.get('colspan', 1)) - 1
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=end_row, end_column=end_col)

        # Column width
        width = node.get('width')
        if width is not None:
            width = int(width) * 8.11 / 52
            ws.column_dimensions[cell.column_letter].width = width


class Testcase(Table):

    def __init__(self, path, func):
        super().__init__(path)
        self.func = func
        self.confirm = None
        self.label_data = self.get_label_data()

        self.check_title()
        self.check_header(CONF.TC_HEADER)
        self.check_confirmation()
        self.check_index()
        self.check_label()

    def check_title(self):
        logging.info(fstr('Checking title'))

        title = 'Test Case [{0}.csv] [Test CSV Information]' \
            .format(self.func)

        h4 = self.get_title()
        if title != h4:
            dsc = '{0} != {1}'.format(h4, title)
            self.add_error('Mismatch title', dsc)

    def check_index(self, header='No.'):
        logging.info(fstr('Checking index'))

        index = self.get_index_header(header)
        lst = [self.get_number(row[index]) for row in self.table]
        lst = [n for n in lst if n is not None]

        # (Index) start at 1
        if lst[0] != 1:
            dsc = 'No.{0}'.format(lst[0])
            self.add_error('Index start at', dsc)

        # (Index) Continuous
        prev = 0
        for tcno in lst:
            if prev != 0 and tcno != (prev + 1):
                dsc = 'No.{0} -> {1}'.format(prev, tcno)
                self.add_error('Index hopping', dsc)
            prev = tcno

    def check_confirmation(self, header='Confirmation', header_no='No.'):
        logging.info(fstr('Checking confirmation'))

        index_no = self.get_index_header(header_no)
        index = self.get_index_header(header)
        dct = {self.get_number(row[index_no]): self.get_text(row[index])
               for row in self.table
               if row[index_no].get('class') not in CONF.TCCLS_CMT}

        lsterr = [k for k, v in dct.items()
                  if k is not None and v not in ['OK', 'Fault']]

        if len(lsterr) > 0:
            dsc = 'No.{0}'.format(', '.join([str(n) for n in lsterr]))
            self.add_error('Missing confirmation', dsc)
        else:
            self.confirm = 'NG' if 'Fault' in dct.values() else 'OK'

    def check_label(self):
        '''Checking label in list'''
        logging.info(fstr('Checking label'))

        diff = set(self.label_data.keys()) - set(CONF.XLS_LABEL.keys())
        if len(diff) > 0:
            self.add_error('Unknown label', list(diff))

    def get_number(self, cell):
        number = None
        if cell.get('class') in CONF.TCCLS_NO:
            number = abs(int(self.get_text(cell)))

        return number

    def get_label_data(self):
        '''Get label data'''
        ino = self.get_index_header('No.')
        iitem = self.get_index_header('Test Analysis Item')
        iid = self.get_index_header('ID')
        icomment = self.get_index_header('Comment')

        data = {}
        for row in self.table:
            tcno = self.get_number(row[ino])
            if tcno is None:
                continue
            elif row[ino].get('class') in CONF.TCCLS_CMT:
                continue

            cellid = row[iid]
            labels = self.get_text(cellid)
            if labels.strip() == '':
                continue

            r, _ = self.get_origin(row[iitem])
            cellcmt = self.table[r][icomment]
            comment = self.get_text(cellcmt)

            for label in labels.split(','):
                label = label.strip()
                lst = data.get(label, {}).get(comment, [])
                lst.append(tcno)

                d1 = data.get(label, {})
                d1.update({comment: lst})

                data.update({label: d1})

        return data


class AmsCsv(Base):

    def __init__(self, path, data):
        super().__init__(path)
        self.data = data

        self.check_description()

    def check_description(self):
        '''Get description'''
        logging.info(fstr('Checking description'))
        data = self.data

        def is_simulink():
            try:
                path = data['src_full']
                result = any(['Simulink model' in line
                              for line in utils.read_file(path)[0:100]])

            except FileNotFoundError:
                result = None

            except Exception:
                result = False

            finally:
                return result

        try:
            lst = [l for l in utils.read_file(self.path)
                   if l.startswith('mod')]

            lst = lst[0].split(',')
            func_full = lst[1].strip()[1:-1]
            description = lst[2].strip()[1:-1]

            lst_dsc = [data['func'], 'Simulink model']

            if func_full != data['func_full']:
                desc = 'Inconsistent function name {0} != {1}' \
                    .format(func_full, data['func_full'])
                self.add_error('Function name', desc)

            if is_simulink() is True and description != 'Simulink model':
                desc = "The description should be 'Simulink model'"
                self.add_error('Description', desc)

            elif is_simulink() is False and description != data['func']:
                desc = "The description should be '{0}'".format(data['func'])
                self.add_error('Description', desc)

            elif is_simulink() is None and description not in lst_dsc:
                desc = "The description should be in '{0}'"\
                    .format(str(lst_dsc))
                self.add_error('Description', desc)

        except Exception as e:
            logging.exception(e)
            self.add_error('Parse csv', str(e))

    def get_stub_info(self):
        '''Get stub and non-stub info'''
        logging.debug("Getting stub and non-stub info")

        lst_line = [l for l in utils.read_file(self.path)
                    if l.startswith('%')]

        lst_stub = []
        lst_non_stub = []
        for line in lst_line:
            try:
                lst = [l.strip() for l in line.split(',')]
                func = lst[2][1:-1]
                if lst[1] != '':
                    lst_stub.append(func)
                else:
                    lst_non_stub.append(func)
            except:
                pass
        
        return lst_stub, lst_non_stub
