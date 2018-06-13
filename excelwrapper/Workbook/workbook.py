class Workbook(object):
    def __init__(self):
        self.filename = None
        self.name = None
        self.raw_catalog = None
        self._sheets = []
        self.loaded = False

    def create(self, name):
        pass

    def save(self, filepath, filetype):
        pass

    def add_sheet(self, name):
        pass

    def remove_sheet(self, id):
        pass


class _WorkbookChild(Workbook):
    _parent = None

    pass
