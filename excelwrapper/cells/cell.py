import datetime
from decimal import Decimal

from excelwrapper.cells.styles import Style

TIME_TYPES = (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)
STRING_TYPES = (str, unicode, bytes)
NUMERIC_TYPES = (int, float, long, Decimal)
KNOWN_TYPES = NUMERIC_TYPES + TIME_TYPES + STRING_TYPES + (bool, type(None))


class Cell(object):
    def __init__(self, value=None, styleobj=None, x=0, y=0, **kwargs):

        # init private vars
        self._fill = None
        self._value = None
        self._format_string = None
        self._style = None
        self._parent = None
        self.x = x
        self.y = y
        # set properties
        self.value = value
        if styleobj:
            self.style = styleobj
        else:
            self.style = Style()

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        if type(value) in KNOWN_TYPES:
            self._value = value
        else:
            raise ValueError("Cannot convert {0!r} to Excel".format(type(value)))

        self._value = value

    @property
    def style(self):
        return self._style

    @style.setter
    def style(self, styleobj):
        if styleobj is None:
            return
        if isinstance(styleobj, Style):
            self._style = styleobj
        else:
            raise TypeError('must be an instance of {0!r} not {1}'.format(Style, type(styleobj)))

    def __repr__(self):
        return "Cell {0!r}".format(type(self.value), self.x)

    def __str__(self):
        return "<Cell {}>".format(self.value)
