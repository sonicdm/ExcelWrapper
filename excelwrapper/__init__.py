from __future__ import print_function
import os

from excelwrapper.Workbook import Workbook
from excelwrapper.Workbook import Xls
from excelwrapper.Workbook import Xlsx

handlers = {
    'xls': Xls(),
    'xlsx': Xlsx(),
}


def read_workbook(path):
    print(path)
    filename = os.path.basename(path)
    fileformat = filename.split('.')[-1]
    reader = handlers.get(fileformat, None)
    if reader:
        return reader.read(path)
    else:
        print('Unsupported File Type {0} - File: {1}'.format(fileformat, filename))
        exit()


def write_workbook(wb, style_data, filetype='xlsx'):
    try:
        return handlers[filetype].write(wb, style_data)
    except KeyError as e:
        print('Unsupported File Type {0}\n{1}'.format(filetype, e))
        exit()
