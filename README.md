# Anti Money Laundering (AML) Project

## Purpose

There're two tools included in this project:

1. Tool 8
2. Tool plug-ins

"Tool 8" is the main processor to analysis cash flows from PCMS output. It's usually as a PDF format.Due to Tool 8 is designed to available executing by RPA robot, for now we only simple design allowing one file as input.
"Tool plug-ins" can say that be equal to an extendable Tool 8. It includes many experimental features, something like keywords extraction and analysis from "comment" column in raw data from cash flow、 Multiple cash files input, from a folder or a zip file, etc..
Another reason to separate two tools is keyword extraction could be extremely slow. To avoid to impact performance of Tool 8 in RPA processes, to ensure good producing, we separate them for now.

## How to run

### Before execute

Set up your environment with pip: pip install -r docker/requirements.txt

### Run up with Python command line

1. Tool 8:
   python main.py

2. Tool 8 plug-in:
   a. Non-Windows:
     - export PULGIN_UI=True; python main.py
     - PULGIN_UI=True python main.py
   b. Windows:
     - $Env:PULGIN_UI="True"; python main.py

### Run up with Pyinstaller

Option: pyi-makespec -D/-F main.py

### Build to One-File

1. Tool 8:
   pyinstaller onefile.spec

2. Tool 8 plug-in:
   pyinstaller plugin_onefile.spec

### Build to One-Directory

1. Tool 8:
   pyinstaller onedir.spec

2. Tool 8 plug-in:
   pyinstaller plugin_onedir.spec
