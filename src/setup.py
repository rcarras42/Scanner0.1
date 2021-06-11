import sys
from cx_Freeze import setup, Executable

options = {
    'build_exe': {
        'excludes': ['collections.sys','collections._weakref']
    }
}

setup(
	name='Simple Scan app',
	version='0.1',
	description="Application permettant la numerisation d'images a partir d'un scanner twain",
	options=options,
	executables=[Executable('pyscan.py',base='Win32GUI')]
	)