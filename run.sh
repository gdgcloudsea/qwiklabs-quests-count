#!/bin/bash

if [ -d venv ]; then
    echo "Activate virtual env"
    source venv/bin/activate
else
    echo "-------------"
    echo "First time running this script may take time as it initializes env"
    echo "Make sure you have PYTHON3 in your PATH (tested with python 3.7 only)"
    echo "Please wait a little bit ..."
    echo "-------------"
    
    echo "Install virtualenv"
    pip install virtualenv
    echo "Create virtual env"
    virtualenv venv --python=python3
    echo "Activate virtual env"
    source venv/bin/activate
    echo "Install dependencies"
    pip install -r requirements.txt
fi

python main.py $@
