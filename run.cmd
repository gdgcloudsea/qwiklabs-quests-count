@echo off

IF EXIST venv (
    REM Activate virtual env
    call venv\Scripts\activate
) ELSE (
    echo -------------
    echo First time running this script may take time as it initializes env
    echo Make sure you have PYTHON3 in your PATH (tested with python 3.7 only)
    echo Please wait a little bit ...
    echo -------------
    
    REM Install virtualenv
    pip install virtualenv
    REM Create virtual env
    virtualenv venv
    REM Activate virtual env
    call venv\Scripts\activate
    REM Install dependencies
    pip install -r requirements.txt
)

python main.py %*
