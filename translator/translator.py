from Sheet import SheetsManager, Sheet
from SheetCell import SheetCell
from operations.operation_factory import OperationTypes
from operations.operations import Value
from translator.AstProcessor import OperationProcessor
import ast
import re
import sys

sheet_manager = SheetsManager()
astProcessor = OperationProcessor()

sheet_simbol_map = {
    "!": "___ref___",
    ":": "___range___",
    "&": "+",
    "%": "+ ___percentage___",
    ",,": ", None,",
    "if(": "IF(",
    "and(": "AND(",
}


def normalize_sheet_name(name):
    """Sheet names reach ast.parse as identifiers, so spaces have to go.

    'CASH FLOW'!A1 would not parse; CASHFLOW!A1 does. Sheets are registered under the
    normalized name, so look them up the same way.
    """
    return name.replace(" ", "")


def translate(workbook):
    creation_switch = {
        "n": generate_value_cell,
        "s": generate_value_cell,
        "f": generate_formula_cell,
        "b": generate_value_cell,
        "d": generate_value_cell,
        "e": generate_value_cell,
    }
    for sheet_name in workbook.sheetnames:
        sheet_manager.instantiate_sheet(normalize_sheet_name(sheet_name))
    for raw_name in workbook.sheetnames:
        sheet = workbook[raw_name]
        sheet_name = normalize_sheet_name(raw_name)

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                creation_switch.get(cell.data_type, None)(cell, sheet_name)
    return sheet_manager


def generate_value_cell(cell, sheet_name):
    sheet = sheet_manager.get_sheet(sheet_name)
    sheet_cell = sheet.get_cell(cell.coordinate)
    if sheet_cell == None:
        sheet_cell = SheetCell(OperationTypes.Value, cell.value)
        sheet_cell.set_coordinate(cell.coordinate)
        sheet_cell.set_sheet(sheet_name)

    else:
        sheet_cell.init_operation(OperationTypes.Value, cell.value)
    sheet.set_cell(cell.coordinate, sheet_cell)


formule_generators = {}


def generate_countif():
    return OperationTypes.Value, 0.000101


def generate_if():
    return OperationTypes.Value, 0.000101


def generate_mround():
    return OperationTypes.Value, 0.000101


def generate_rounddown():
    return OperationTypes.Value, 0.000101


def generate_roundup():
    return OperationTypes.Value, 0.000101


def generate_sum():
    return OperationTypes.Value, 0.000101


def generate_today():
    return OperationTypes.Value, 0.000101


def generate_iferror():
    return OperationTypes.Value, 0.000101


def generate_text():
    return OperationTypes.Value, 0.000101


formule_generators = {
    "COUNTIF": generate_countif,
    "IF": generate_if,
    "MROUND": generate_mround,
    "ROUNDDOWN": generate_rounddown,
    "ROUNDUP": generate_roundup,
    "SUM": generate_sum,
    "TODAY": generate_today,
    "IFERROR": generate_iferror,
    "TEXT": generate_text,
}


def generate_formula_cell(cell, sheet_name):
    current_sheet: Sheet = sheet_manager.get_sheet(sheet_name)

    # remove '='
    cell_formula = str(cell.value)
    if cell_formula.find("=#REF!") != -1:
        current_cell = SheetCell()
        operation = Value("#REF!")
        current_cell.set_operation(operation)
        current_sheet.set_cell(cell.coordinate, current_cell)
        return
    if cell_formula.startswith("="):
        cell_formula = cell_formula[1:]

    # reemplazos simboilos invalidos en python para permitir parsear, en el astprocessor se revertiran
    for simbol, replacement in sheet_simbol_map.items():
        # Todo Eventualmente resolver esto de otra manera
        # caso particular donde la sheet esta referenciada como 'sheet name', esto rompe el parser
        if simbol == "!":
            # 'Sheet Name'!A1 -> SheetName!A1, matching normalize_sheet_name()
            pattern = r'["\']([\w\d\s]+)["\']!'
            cell_formula = re.sub(
                pattern, lambda m: normalize_sheet_name(m.group(1)) + "!", cell_formula
            )
        cell_formula = cell_formula.replace(simbol, replacement)

    # TODO reemplazo de = en condiciones logicas por ==. Ver si no interfiere con otra cosa
    pattern = r"(?<![<>])="
    cell_formula = re.sub(pattern, r"==", cell_formula)

    tree = ast.parse(cell_formula)
    current_cell = current_sheet.get_cell(cell.coordinate)

    if current_cell is None:
        current_cell = SheetCell()

        current_sheet.set_cell(cell.coordinate, current_cell)
    # TODO astprocesor ya tiene la expresion para parsear, deberá retornar la operacion a asignar en la celda
    astProcessor.set_current_cell(current_cell)
    astProcessor.set_current_sheet(current_sheet)

    operation = astProcessor.visit(tree)

    current_cell.set_operation(operation)


def have_operators(expression):
    pattern = r"[+\-*/()<> =]"
    return bool(re.search(pattern, expression))


def generate_expression_operation(expression: str, sheet_name):
    sheet_referenced_name = sheet_name
    cell_name = expression

    if expression.find("!") != -1:
        sheet_referenced_name, cell_name = expression.split("!")

    sheet = sheet_manager.get_sheet(sheet_referenced_name)
    cell_referenced = sheet.get_cell(cell_name)
    if cell_referenced is None:
        cell_referenced = SheetCell()
        sheet.set_cell(cell_name, cell_referenced)
    return OperationTypes.Reference, cell_referenced
