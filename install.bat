@echo off
echo pip...
venv\Scripts\python -m pip install --upgrade pip -i https://mirrors.aliyun.com/pypi/simple/

echo install...
venv\Scripts\pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

echo finish！
pause