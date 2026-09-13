import logging
import ast
from operations.operation_factory import OperationTypes, OperationFactory
from SheetCell import SheetCell
from Sheet import SheetsManager
import openpyxl.utils.cell as utils

sheet_manager = SheetsManager()
bin_ope_dict = {
    "ADD": OperationTypes.BinaryAdd,
    "SUB": OperationTypes.BinarySub,
    "MULT": OperationTypes.BinaryMul,
    "DIV": OperationTypes.BinaryDiv,
}


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
        cell: SheetCell = None
        cell_name = node.id
        sheet_name = ""
        current_sheet = self.current_sheet

        if node.id.find("___range___") != -1:
            if node.id.find("___ref___") != -1:
                sheet_name = str(node.id).split("___ref___")[0]
                node.id = node.id.replace("___range___", ":")
                node.id = str(node.id).replace(f"{sheet_name}___ref___", "")
                current_sheet = sheet_manager.get_sheet(sheet_name)
            range_operation = OperationFactory.get_instance(
                OperationTypes.Range, node.id.replace("___range___", ":"), current_sheet
            )

            initial, final = node.id.replace("___range___", ":").split(":")
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
            return range_operation
        else:
            if node.id.find("___ref___") != -1:
                sheet_name, cell_name = str(node.id).split("___ref___")
                current_sheet = sheet_manager.get_sheet(sheet_name)

                cell = current_sheet.get_cell(cell_name)
            else:
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

    def visit_Eq(self, node):
        return OperationTypes.EqualCondition

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
