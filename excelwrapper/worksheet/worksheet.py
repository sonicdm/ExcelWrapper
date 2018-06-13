class Worksheet:
    def __init__(self, name, rows=None):
        self._rows = []

    def add_row(self, row):
        pass

    def insert_row(self, idx=0):
        pass

    def remove_row(self, idx):
        pass

    def move_row(self, idx):
        pass

    def __getitem__(self, item):
        yield self._rows[item]

    def row(self, idx):
        yield self._rows[idx]
