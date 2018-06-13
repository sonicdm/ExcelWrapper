try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

config = dict(
    description='Excel Wrapper',
    author='Allan Barcellos',
    version='0.1a',
    packages=['ExcelWrapper'],
    install_requires=['openpyxl', 'xlrd'],
    name='ExcelWrapper',
)

setup(**config)
