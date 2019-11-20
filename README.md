# Resolve crash issue

## Solution 1:

https://github.com/samuelhwilliams/Eel/issues/44

```
Change sleep(1.0) to sleep(10.0) in function _websocket_close
C:\Users\Administrator\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\eel\__init__.py
```

## Solution 2:

https://github.com/samuelhwilliams/Eel/issues/54

# Check process id that locking port

```
netstat -ano -p tcp |find "9892"

netstat -ano -p tcp | find "9892" | find "LISTENING"

import subprocess
pids = subprocess.check_output(command, shell=True)
```

# Get current process

```
import os
os.getpid()
```

# Kill process

```
import os
import signal

os.kill(pid, signal.SIGTERM) #or signal.SIGKILL
```

# Remove all pip

```
pip freeze > requirements_temp.txt
pip uninstall -y -r requirements_temp.txt
```# lila # lila # lila
# lila
