# excelkalender
Python tool for creating an excel file as calendar overview. It contains all month of the specified year and each month before and after the year.

The script can be configured by the configuration file config.yml. If you want a different file name, specify it in this file (variable config_name).

Holidays can be fetched from openapi.org via web request, see options oh_api_* in configuration file.


## Usage
To use the script simply call:

```
python3 kalender.py
```

As output you get an excel file with this result:
!(assets/images/excelkalender-output.png)


## Configuration
For more informations about configuring the script, see configuration file itself [config.yml].


## Build
If you want to build or execute the script you have to install python3 and some libraries.

On Ubuntu 22.04 you have to install the following packages:
* python3-openpyxl
* python3-bs4
* pyyaml

Tested with Python 3.10.12.