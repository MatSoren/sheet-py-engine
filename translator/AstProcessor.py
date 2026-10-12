import logging
import ast
import re
from operations.operation_factory import OperationTypes, OperationFactory
from SheetCell import SheetCell
from Sheet import SheetsManager
import openpyxl.utils.cell as utils
from openpyxl.formula.tokenizer import Tokenizer
from operations.operations import ExcelError

sheet_manager = SheetsManager()
bin_ope_dict = {
    "ADD": OperationTypes.BinaryAdd,
    "SUB": OperationTypes.BinarySub,
    "MULT": OperationTypes.BinaryMul,
    "DIV": OperationTypes.BinaryDiv,
    "POW": OperationTypes.BinaryPow,
    "BITOR": OperationTypes.Concat,
}

REF_MARKER = "___ref___"
RANGE_MARKER = "___range___"
ERROR_MARKER = "___error___"
error_codes = {re.sub(r"\W", "", code): code for code in Tokenizer.ERROR_CODES}  # 'NA': '#N/A'


def normalize_sheet_name(name):
    """The name a sheet is registered under in the SheetsManager.

    Every non-alphanumeric ASCII char becomes '_' ('Cost Report' -> 'Cost_Report'), so it can
    live inside a Python identifier. ASCII only on purpose: ast NFKC-normalizes identifiers,
    which could make a unicode name stop matching its registry key.
    """
    return re.sub(r"\W", "_", name, flags=re.ASCII)


def encode_ref(sheet, ref):
    """('2024 Ventas', 'A1:B2') -> '_2024_Ventas___ref___A1___range___B2', a valid identifier.

    The leading '_' covers sheet names starting with a digit. decode_ref() undoes this.
    """
    ref = ref.replace(":", RANGE_MARKER)
    if sheet is None:
        return ref
    return f"_{normalize_sheet_name(sheet)}{REF_MARKER}{ref}"


def encode_error(code):
    """'#N/A' -> '___error___NA'; visit_Name turns it back into an ExcelError."""
    return ERROR_MARKER + re.sub(r"\W", "", code)


def decode_ref(identifier):
    """Inverse of encode_ref(): returns (normalized sheet name or None, 'A1' or 'A1:B2')."""
    sheet, _, ref = identifier.rpartition(REF_MARKER)
    return (sheet[1:] if sheet else None), ref.replace(RANGE_MARKER, ":")


formula_dict = {
    "COUNTIF": OperationTypes.CountIf,
    "IF": OperationTypes.If,
    "MROUND": OperationTypes.MRound,
    "ROUNDDOWN": OperationTypes.RoundDown,
    "ROUNDUP": OperationTypes.RoundUp,
    "SUM": OperationTypes.Sum,
    "HLOOKUP": OperationTypes.Hlookup,
    "TODAY": OperationTypes.Today,
    "IFERROR": OperationTypes.IfError,
    "TEXT": OperationTypes.Text,
    "OR": OperationTypes.OrCondition,
    "AND": OperationTypes.AndCondition,
}


class OperationProcessor(ast.NodeVisitor):
    def visit_BinOp(self, node):
        operation_name = type(node.op).__name__

        left = super().visit(node.left)

        right = super().visit(node.right)

        operation_type = bin_ope_dict.get(operation_name.upper())
        return OperationFactory.get_instance(operation_type, left, right)

    def visit_Expr(self, node):
        return super().visit(node.value)

    def visit_Name(self, node):
        if node.id.startswith(ERROR_MARKER):
            code = error_codes[node.id[len(ERROR_MARKER):]]
            return OperationFactory.get_instance(OperationTypes.Value, ExcelError(code))
        cell: SheetCell = None
        sheet_name, cell_name = decode_ref(node.id)
        current_sheet = self.current_sheet
        if sheet_name is not None:
            current_sheet = sheet_manager.get_sheet(sheet_name)
        else:
            sheet_name = ""

        if ":" in cell_name:
            initial, final = cell_name.split(":")
            initial_row, initial_col = utils.coordinate_to_tuple(initial)
            final_row, final_col = utils.coordinate_to_tuple(final)
            cells_to_attach = []
            for row in range(initial_row, final_row + 1):
                for col in range(initial_col, final_col + 1):
                    cell = current_sheet.get_cell(
                        utils.get_column_letter(col) + str(row)
                    )
                    if cell is None:
                        cell = SheetCell()
                        current_sheet.set_cell(
                            utils.get_column_letter(col) + str(row), cell
                        )
                    cells_to_attach.append(cell)

            for cell_to_attach in cells_to_attach:
                cell_to_attach.attach(self.current_cell)
            # built after the missing cells exist, or Range would hold None for them
            return OperationFactory.get_instance(
                OperationTypes.Range, cell_name, current_sheet
            )

        cell = current_sheet.get_cell(cell_name)

        if cell is None:
            cell = SheetCell()
            cell.set_coordinate(cell_name)
            cell.set_sheet(sheet_name)
            current_sheet.set_cell(cell_name, cell)

        cell.attach(self.current_cell)
        return OperationFactory.get_instance(OperationTypes.Reference, cell)

    def visit_Compare(self, node):
        left = super().visit(node.left)

        op_type = super().visit(node.ops[0])

        comparator = super().visit(node.comparators[0])

        return OperationFactory.get_instance(op_type, left, comparator)

    def visit_UnaryOp(self, node):
        operand = super().visit(node.operand)
        if isinstance(node.op, ast.USub):
            return OperationFactory.get_instance(
                OperationTypes.BinaryMul,
                OperationFactory.get_instance(OperationTypes.Value, -1),
                operand,
            )
        if isinstance(node.op, ast.UAdd):
            return operand
        raise ValueError(f"unsupported unary operator: {type(node.op).__name__}")

    def visit_Eq(self, node):
        return OperationTypes.EqualCondition

    def visit_NotEq(self, node):
        return OperationTypes.DiferentCondition

    def visit_Lt(self, node):
        return OperationTypes.LessCondition

    def visit_LtE(self, node):
        return OperationTypes.LessOrEqualCondition

    def visit_Gt(self, node):
        return OperationTypes.GratherCondition

    def visit_GtE(self, node):
        return OperationTypes.GratherOrEqualCondition

    def visit_Constant(self, node):
        return OperationFactory.get_instance(OperationTypes.Value, node.value)

    def visit_Call(self, node):
        args = []
        for arg in node.args:
            args.append(super().visit(arg))

        func_name = node.func.id
        func_type = formula_dict.get(func_name.upper())
        if func_type == None:
            logging.warning("unrecognised function: %s", func_name)
        return OperationFactory.get_instance(func_type, *args)

    def visit_Module(self, node):
        if node is None:
            return None
        node = node.body[0]
        return super().visit(node)

    def visit(self, node):
        if node is None:
            return None
        return super().visit(node)

    def __init__(self):
        self.current_cell = None
        self.current_sheet = None

    def set_current_cell(self, cell):
        self.current_cell = cell

    def set_current_sheet(self, sheet):
        self.current_sheet = sheet
