# -*- coding: utf-8 -*-


__authors__ = 'Duy Nguyen'
__contact__ = 'duykn@gcs-vn.com'


import logging


from openpyxl.styles import (Alignment, Border, Font, PatternFill, Protection,
                             Side)
from openpyxl.styles.fills import FILL_SOLID


class Style(object):

    def __init__(self, cell_css):
        info = {}
        for key in cell_css:
            value = cell_css[key].replace(', ', ',')
            lst = value.split(' ')

            ignore = ['width', 'word-break', 'border-collapse', 'border-spacing',
                      'border-top-style', 'padding', 'margin']
            if key in ignore:
                continue

            elif key == 'background-color':
                info['start_color'] = self.rgb2hex(value)

            elif key == 'text-align':
                info['horizontal'] = value

            elif key == 'vertical-align':
                if value == 'middle':
                    value = 'center'
                info['vertical'] = value

            elif key == 'font':
                info['fname'] = lst[4][1:-1]
                info['fsize'] = int(lst[3][:-2])
                if lst[2] == 'bold':
                    info['fbold'] = True

            elif key == 'border-left':
                lst[1] = self.convert_boder_style(lst[1], lst[0])
                info['blbs'] = lst[1]
                info['blc'] = self.rgb2hex(lst[2])

            elif key == 'border-right':
                lst[1] = self.convert_boder_style(lst[1], lst[0])
                info['brbs'] = lst[1]
                info['brc'] = self.rgb2hex(lst[2])

            elif key == 'border-top':
                lst[1] = self.convert_boder_style(lst[1], lst[0])
                info['btbs'] = lst[1]
                info['btc'] = self.rgb2hex(lst[2])

            elif key == 'border-bottom':
                lst[1] = self.convert_boder_style(lst[1], lst[0])
                info['bbbs'] = lst[1]
                info['bbc'] = self.rgb2hex(lst[2])

            elif key == 'color':
                info['fc'] = self.rgb2hex(value)

            else:
                logging.debug("Unknown css attribute '%s'", key)

        self.font = Font(name=info.get('fname', 'Courier'),
                         size=info.get('fsize', 10),
                         bold=info.get('fbold', False),
                         italic=False,
                         vertAlign=None,
                         underline='none',
                         strike=False,
                         color=info.get('fc', 'FF000000'))
        self.fill = PatternFill(fill_type=FILL_SOLID,
                                start_color=info.get(
                                    'start_color', 'FFFFFFFF'),
                                end_color='FF000000')
        self.border = Border(left=Side(border_style=info.get('blbs', None),
                                       color=info.get('blc', 'FF000000')),
                             right=Side(border_style=info.get('brbs', None),
                                        color=info.get('brc', 'FF000000')),
                             top=Side(border_style=info.get('btbs', None),
                                      color=info.get('btc', 'FF000000')),
                             bottom=Side(border_style=info.get('bbbs', None),
                                         color=info.get('bbc', 'FF000000')),
                             diagonal=Side(border_style=None,
                                           color='FF000000'),
                             diagonal_direction=0,
                             outline=Side(border_style=None,
                                          color='FF000000'),
                             vertical=Side(border_style=None,
                                           color='FF000000'),
                             horizontal=Side(border_style=None,
                                             color='FF000000')
                             )
        self.alignment = Alignment(horizontal=info.get('horizontal', 'general'),
                                   vertical=info.get('vertical', 'bottom'),
                                   text_rotation=0,
                                   wrap_text=True,
                                   shrink_to_fit=True,
                                   indent=0)
        self.number_format = 'General'
        self.protection = Protection(locked=False,
                                     hidden=False)

    def rgb2hex(self, string):
        char = ','
        rgb = [int(s) for s in string[4:-1].split(char)]
        return '00{:02x}{:02x}{:02x}'.format(*tuple(rgb))

    def convert_boder_style(self, style, width):
        if style == 'solid':
            style = 'thin'
        return style
