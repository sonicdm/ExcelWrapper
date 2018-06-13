from __future__ import print_function
from past.builtins import xrange
import unittest
from datetime import datetime

from excelwrapper.cells.cell import Cell
from excelwrapper.cells.styles import Style
from excelwrapper.rows import Row


def dummy_func():
    pass


class StyleTestCase(unittest.TestCase):
    def test_set_a_correct_font(self):
        font = "calibri.ttf"
        styleobj = Style()
        styleobj.font = font
        self.assertEqual(font, styleobj.font)

    def test_set_an_incorrect_font(self):
        font = 123
        styleobj = Style()
        with self.assertRaises(TypeError):
            styleobj.font = font
            self.assertEqual(font, styleobj.font)

    def test_set_bold(self):
        styleobj = Style()
        styleobj.bold = True
        self.assertTrue(styleobj.bold)
        styleobj.bold = False
        self.assertFalse(styleobj.bold)

    def test_set_bad_bold(self):
        styleobj = Style()
        with self.assertRaises(TypeError):
            styleobj.bold = 1
        with self.assertRaises(TypeError):
            styleobj.bold = "True"

    def test_set_fill(self):
        styleobj = Style()
        styleobj.fill = 'fff'
        self.assertEqual(styleobj.fill, 'fff')

    def test_fail_set_bad_fill(self):
        styleobj = Style()
        with self.assertRaises(ValueError):
            styleobj.fill = 'abz171'
        with self.assertRaises(TypeError):
            styleobj.fill = False

    def test_set_italic(self):
        styleobj = Style()
        styleobj.italic = True
        self.assertTrue(styleobj.italic)

    def test_fail_set_italic(self):
        styleobj = Style()
        with self.assertRaises(TypeError):
            styleobj.italic = 1

    def test_set_strike(self):
        styleobj = Style()
        styleobj.strike = True
        self.assertTrue(styleobj.strike)

    def test_fail_set_strike(self):
        styleobj = Style()
        with self.assertRaises(TypeError):
            styleobj.strike = 1

    def test_set_underline(self):
        styleobj = Style()
        styleobj.underline = True
        self.assertTrue(styleobj.underline)

    def test_fail_set_underline(self):
        styleobj = Style()
        with self.assertRaises(TypeError):
            styleobj.underline = 1
        with self.assertRaises(TypeError):
            styleobj.underline = "True"

    def test_set_size(self):
        styleobj = Style()
        styleobj.size = 11
        self.assertEqual(styleobj.size, 11)

    def test_set_bad_size_value(self):
        styleobj = Style()
        with self.assertRaises(TypeError):
            styleobj.size = 'a'
        with self.assertRaises(TypeError):
            styleobj.size = False

    def test_set_styles_from_signature(self):
        styleobj = Style(12, 'abc', True, True, 'test', True, True, 'aabbcc')
        self.assertTrue(styleobj.bold)
        self.assertTrue(styleobj.strike)
        self.assertTrue(styleobj.underline)
        self.assertTrue(styleobj.italic)
        self.assertEqual(styleobj.fill, 'aabbcc')
        self.assertEqual(styleobj.color, 'abc')
        self.assertEqual(styleobj.font, 'test')


class CellTestCase(unittest.TestCase):
    def test_set_value(self):
        # test for success
        cell = Cell('test')
        self.assertEqual(cell.value, 'test')
        cell = Cell()
        cell.value = 'blah'
        self.assertEqual(cell.value, 'blah')

        # test for failure
        obj = object

        failobjects = [dummy_func, obj, {1: 2}, [1]]
        for thing in failobjects:
            with self.assertRaises(ValueError):
                cell = Cell()
                cell.value = thing
            with self.assertRaises(ValueError):
                cell = Cell(thing)

    def test_set_cell_style(self):
        styleobj = Style()
        objects = [[1], {1: 2}, (1,), ' ', -1]
        cell = Cell('test', styleobj)
        cell = Cell()
        cell.style = styleobj
        for object in objects:
            with self.assertRaises(TypeError):
                cell = Cell('test', object)


class RowTestCase(unittest.TestCase):
    def test_fetch_cell_values(self):
        cells = [Cell(x) for x in xrange(1, 10)]
        cell_vals = [x for x in xrange(1, 10)]
        row = Row()
        for cell in cells:
            row.append(cell)
        for idx, val in enumerate(cells):
            self.assertEqual(val.value, row.cell_values[idx])

        for idx, val in enumerate(cell_vals):
            self.assertEqual(val, row.cell_values[idx])

    def test_insert_cell(self):

        row = Row()
        values = ["Text", datetime.today(), 100, 50.2, u'hi', 12387123981723918273123712930, None, True]
        for idx, value in enumerate(values):
            row.insert_cell(value, idx)
        for idx, cell in enumerate(row.cells):
            self.assertEqual(cell.value, values[idx])

        row = Row()
        values = [1, 2, 3, 4, 5, 6, 7, 8]
        for value in values:
            row.append(value)

        row.insert_cell('lala', 4)
        self.assertEqual(row[4].value, 'lala')

        row.insert_cell(Cell('haha'), 2)
        self.assertEqual(row[2].value, 'haha')

    def test_append_cell(self):
        row = Row()
        values = ["Text", datetime.today(), 100, 50.2, u'hi', 12387123981723918273123712930, None, True]
        for value in values:
            row.append(value)
        cell_vals = [v.value for v in row.cells]
        for value in values:
            self.assertIn(value, cell_vals)

    def test_access_cell(self):
        row = Row()
        row[0] = 1
        self.assertEqual(row[0], 1)


if __name__ == '__main__':
    unittest.main()
