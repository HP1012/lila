# -*- coding: utf-8 -*-


import json
import sys
from pathlib import Path


def load(filepath):
    '''Load json to dict'''
    try:
        with open(filepath, encoding='shift-jis', errors='ignore') as fp:
            data = json.load(fp)
    except:
        data = {}
    finally:
        return data


def write(data, filepath):
    '''Write dict to json'''
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, encoding='shift-jis', errors='ignore', mode='w') as fp:
        json.dump(data, fp, indent=4, sort_keys=True)


# Update hash to version
verpath = './mickey/lila/assets/version.json'
data = load(verpath)

write(data, 'lila.version')

with open('lila.git') as fp:
    lines = fp.readlines()
    githash = lines[0][-7:].strip()
    gitcount = lines[-1].strip()
    gitcount = int(gitcount) + 239

version = '{0}.{1}.{2}'.format(data.get('version'), gitcount, githash)
data.update({'version': version})
write(data, verpath)


eeljs = None
for path in sys.path:
    path = Path(path).joinpath('eel', 'eel.js')
    if path.is_file():
        eeljs = path
        break

data = {
    "pathex": str(Path.cwd()).replace('\\', '\\\\'),
    "eeljs": str(eeljs).replace('\\', '\\\\'),
    "name": "Lila.{0}".format(version)
}

spec = Path.cwd().joinpath('lila.spec')

with open(spec) as fp:
    text = fp.read()

text = text.format(**data)

with open('build.spec', 'w') as fp:
    fp.write(text)
