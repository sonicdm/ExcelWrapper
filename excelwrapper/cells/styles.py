from __future__ import print_function
import string


def is_hexcolor(s):
    """ 
    Check if a string is a valid HEX RGB color.
    >>> is_hexcolor('ff55ee')
    True
    >>> is_hexcolor('123')
    True
    >>> is_hexcolor('abab')
    False
    >>> is_hexcolor('FFEERR')
    False
    """
    if is_hex(s):
        if len(s) == 3 or len(s) == 6:
            return True
        else:
            return False
    else:
        return False


def is_hex(s):
    # type: (str) -> bool
    """
    Check if a string is hexadecimal
    :param s: string
    :type s: str
    :return: bool
    :rtype: bool
    >>> is_hex('aavv22')
    False
    >>> is_hex('44dFa5')
    True
    """
    hex_digits = set(string.hexdigits)
    return all(c in hex_digits for c in s)


class Style(object):
    """
    Style Object for defining styles within a cell. Easier to copy to other cells as an object.
    """

    def __init__(self, size=11, color='fff', bold=False, italic=False,
                 font='Calibri', strike=False, underline=False, fill=None):
        self._size = size
        self._font = font
        self._color = color
        self._bold = bold
        self._italic = italic
        self._strike = strike
        self._underline = underline
        self._fill = fill
        self._merged = None
        self._numberformat = None

    @property
    def size(self):
        return self._size

    @size.setter
    def size(self, size):
        if type(size) == int:
            self._size = size
        else:
            raise TypeError("size must be an integer")

    @property
    def color(self):
        return self._color

    @color.setter
    def color(self, color):
        if type(color) is not str:
            raise TypeError('expected string or buffer, got {} instead'.format(type(color)))

        elif is_hexcolor(color):
            self._color = color
        else:
            raise ValueError('value must be a hex color value')

    @property
    def fill(self):
        return self._fill

    @fill.setter
    def fill(self, color):
        if type(color) is not str:
            raise TypeError('expected string or buffer, got {} instead'.format(type(color)))

        elif is_hexcolor(color):
            self._fill = color
        else:
            raise ValueError('value must be a hex color value')

    @property
    def font(self):
        return self._font

    @font.setter
    def font(self, font):
        if isinstance(font, str):
            self._font = font
        else:
            raise TypeError('expected string or buffer got {} instead'.format(type(font)))

    @property
    def bold(self):
        return self._bold

    @bold.setter
    def bold(self, v):
        if isinstance(v, bool):
            self._bold = v
        else:
            raise TypeError("expected type 'bool' got {} instead".format(type(v)))

    @property
    def italic(self):
        return self._italic

    @italic.setter
    def italic(self, v):
        if isinstance(v, bool):
            self._italic = v
        else:
            raise TypeError("expected type 'bool' got {} instead".format(type(v)))

    @property
    def strike(self):
        return self._strike

    @strike.setter
    def strike(self, v):
        if isinstance(v, bool):
            self._strike = v
        else:
            raise TypeError("expected type 'bool' got {} instead".format(type(v)))

    @property
    def underline(self):
        return self._underline

    @underline.setter
    def underline(self, v):
        if isinstance(v, bool):
            self._underline = v
        else:
            raise TypeError("expected type 'bool' got {} instead".format(type(v)))

    @property
    def number_format(self):
        return self._numberformat

    @number_format.setter
    def number_format(self, formatobject):
        self._numberformat = formatobject


class NumberFormat(object):
    def __init__(self, formatstr=None):
        self._numberformat = formatstr
