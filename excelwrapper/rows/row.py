from excelwrapper.cells.cell import Cell


class Row(object):
    def __init__(self, rownum=0, length=1, fromlist=None, parent=None):
        self._rownum = rownum
        self._length = length
        self._cells = {}
        self._parent = None
        # create with default styles from a list
        self._fromlist = fromlist

    def insert_cell(self, v, idx=0, **kwargs):
        """
        Take a new cell and shift cells around idx to insert it.
        :param v: cell value or Cell object.
        :param idx: 
        :param kwargs: 
        :return: 
        """

        for pos, cell in self._cells.items():
            if cell:
                if cell.x < idx:
                    continue
                if cell.x >= idx:
                    cell.x += 1

        if isinstance(v, Cell):
            v.x = idx
            self._cells[idx] = v
        else:
            newcell = Cell(v, x=idx)
            self._cells[idx] = newcell

    def append(self, v):
        end = len(self._cells) + 1
        if isinstance(v, Cell):
            v.x = end
            self._cells[end] = v
        else:
            newcell = Cell(v, x=end)
            self._cells[end] = newcell

    @property
    def cells(self):
        cells = [c for c in self._cells.values()]
        return cells

    @cells.setter
    def cells(self, values):
        pass

    @property
    def cell_values(self):
        output = []
        for cell in self._cells.values():
            output.append(cell.value)

        return output

    def create_cell(self, val, x=0, y=0, **kwargs):
        newcell = Cell(val, **kwargs)
        return newcell

    def __getitem__(self, key):
        return self._cells[key]

    def __setitem__(self, key, value):
        self._cells[key] = value
