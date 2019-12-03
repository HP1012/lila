# -*- coding: utf-8 -*-


__authors__ = 'Duy Nguyen'
__contact__ = 'duykn@gcs-vn.com'


from pathlib import Path

NAME = 'Kuli'


SERVER = 'duynguyen3116.cybersoft.vn:7890'
API_VERSION = 'api/build/version'
API_DOWNLOAD = 'api/build/download'
API_TASK = 'api/task/info'


HOME = Path.home().joinpath(NAME)
FP_CONFIG = HOME.joinpath('config.json')
FP_APP = HOME.joinpath('{0}.exe'.format(NAME))

FP_FILES = Path(__file__).parent.joinpath('files')
FP_VERSION = FP_FILES.joinpath('version.txt')
FP_CONFIG_DF = FP_FILES.joinpath('config.json')
FP_TEMPLATE = FP_FILES.joinpath('template.xlsx')

# Directory
DIR_SPEC = '単体テスト仕様書'
DIR_RESULT = '単体テスト結果'

# Testcase table
TC_HEADER = [
    'No.', 'Test Analysis Item', 'Decision',
    'Attribute', 'ID', 'Comment',
    'Input Value', 'OutputValue', 'Check',
    'Match', 'Confirmation'
]

TCCLS_NO = [
    'data-no',
    'data-no-last',
    'data-commentout-center-left-right'
]

TCCLS_CMT = [
    'data-commentout-center-left-right'
]

TCCLS_HEAD = [
    'head-input-no',
    'head-input-no-last',
    'head-output-no',
    'head-output-no-last'
]

# Unit test specification
XLS_WS = [
    'フォーマット変更来歴',
    '単体テスト仕様',
    '入力データ分析表',
    '出力データ分析表',
    '入出力データ分析表',
    'テストケース結果表',
    'カバレッジ結果'
]

# Label
XLS_LABEL = {
    'coverage_p1': '未到達コード',
    'div_zero_p2': 'ゼロ割り',
    'overflow_p3': 'オーバーフロー(加算/乗算)',
    'casting_p4': 'オーバーフロー(小さな型への代入)',
    'underflow_p5': 'アンダーフロー(減算)',
    'array_p6': 'メモリ破壊(配列)',
    'pointer_p7': 'メモリ破壊(ポインタ)'
}
