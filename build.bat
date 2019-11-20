REM Get latest log hash
git log -1 > lila.git
git rev-list --all --count >> lila.git

REM Generate build.spec
python build.py

REM Build exe file
pyinstaller build.spec

REM Restore old version.json
copy /y lila.version mickey\\lila\\assets\\version.json

REM Clean
if exist build.spec del build.spec
if exist lila.git del lila.git
if exist lila.version del lila.version

pause