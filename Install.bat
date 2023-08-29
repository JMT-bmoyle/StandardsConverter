@echo off

echo Installing Python...
echo --------------------

start /wait Q:\AZURE_HVMD\00-Reality_Capture\Sandbox\bmoyle\TopoDot_Notifier\python-3.10.7-amd64.exe /quiet InstallAllUsers=0 InstallLauncherAllUsers=0 PrependPath=1

echo Updating PIP...
echo --------------------

start /wait %localappdata%\Programs\Python\Launcher\py.exe -m pip install --upgrade pip

echo Grabbing Packages...
echo --------------------

start /wait %localappdata%\Programs\Python\Launcher\py.exe -m pip install openpyxl

start /wait %localappdata%\Programs\Python\Launcher\py.exe -m pip install tqdm