@echo off
py -m venv image2ppt_venv
cd image2ppt_venv
call Scripts\activate
cd ..
pip install -r requirements.txt
cd Image2ppt
cd src
pyinstaller main.spec
pause

