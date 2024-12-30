# csv 2 excel 

# intro

This is a simple tool which I build to help a person to convert .csv files to .xlsx files.

## local dev

    python -m venv venv
    
    .\venv\Scripts\activate
    
    pip install -r requirements.txt

## build

    pyinstaller --onefile --noconsole --name csv2xlsx --icon=icon.ico -F main.py --additional-hooks-dir=.
