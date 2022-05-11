from sys import argv
from setuptools import setup

APP = ['GUI_project.py']
OPTIONS = {
    'argv_emulation' : True,
}

setup(
    app = APP,
    options={'py2app' : OPTIONS},
    setup_requires = ['py2app']
)