
from distutils.core import setup
import py2exe

APP_NAME = 'VRS'
APP = ['VRS.py']
DATA_FILES = [
    ('resourses',['resourses/vrs_details.xlsx']),
    ('resourses',['resourses/icon.ico'])
]
OPTIONS = {
    'iconfile': 'icon.ico',
    'argv_emulation': True
}

setup (
    app=APP,
    name=APP_NAME,
    data_files=DATA_FILES,
    options=OPTIONS,
    console=['py2exe'],
)