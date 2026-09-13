from SheetCell import SheetCell


class SheetsManager:
    _sheet_dict = {}
    manager = None

    def __init__(self):
        pass

    def instantiate_sheet(self, sheet_name):
        sheet = Sheet(sheet_name)
        SheetsManager._sheet_dict[sheet_name] = sheet
        return sheet

    def set_sheet(self, sheet_name, sheet_data):
        SheetsManager._sheet_dict[sheet_name] = sheet_data

    def get_sheet(self, sheet_name):
        return SheetsManager._sheet_dict.get(sheet_name, None)


class Sheet:
    def __init__(self, name=""):
        self.sheet_name = name
        self.elements = {}

    def get_cell(self, cell_id):
        return self.elements.get(cell_id, None)

    def set_cell(self, cell_id, cell):
        self.elements[cell_id] = cell
        cell.set_coordinate(cell_id)
        cell.set_sheet(self.sheet_name)

    def __str__(self):
        return "\n".join(
            f"{element_id}: {element}" for element_id, element in self.elements.items()
        )
