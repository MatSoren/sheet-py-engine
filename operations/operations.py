import logging
from abc import ABC, abstractmethod

from datetime import date, datetime, timedelta

import openpyxl.utils.cell as utils

import re

import math

from collections.abc import Iterable

"""
Operaciones:
- COUNTIF
- IF *
- MROUND *
- ROUNDDOWN *
- ROUNDUP *
- SUM
- TODAY*
- iferror
- text *
- +
- -
- *
- /
"""


class Operation(ABC):
    def __str__(self):
        return "Operation"

    @abstractmethod
    def operate(self):
        pass


class NoneOperation(Operation):
    def __str__(self):
        return "NoneOperation"

    def __init(self):
        pass

    def operate(self):
        return None


class Value(Operation):
    def __str__(self):
        return "Value"

    def __init__(self, value):
        self.value = value

    def operate(self):
        return self.value


class Reference(Operation):
    def __str__(self):
        return "Reference"

    def __init__(self, cell):
        self.cell = cell

    def operate(self):
        return self.cell.calculated_value


class Range(Operation):
    def __str__(self):
        return "Range"

    def __init__(self, sheet_range, current_sheet):
        initial, final = sheet_range.split(":")
        initial_row, initial_col = utils.coordinate_to_tuple(initial)
        final_row, final_col = utils.coordinate_to_tuple(final)

        self.cell_range = [
            current_sheet.get_cell(utils.get_column_letter(col) + str(row))
            for row in range(initial_row, final_row + 1)
            for col in range(initial_col, final_col + 1)
        ]

    def operate(self):
        return self.cell_range


class Today(Operation):
    def __str__(self):
        return "Today"

    def __init__(self):
        pass

    def operate(self):
        return datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)


class EqualCondition(Operation):
    def __str__(self):
        return "EqualCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        op_1 = self.operand_1.operate()
        op_2 = self.operand_2.operate()
        if op_1 is None or op_2 is None:
            return None
        
        if not (isinstance(op_1, str) and isinstance(op_2, str)):
            if isinstance(op_1, str):
                try:
                    op_1 = float(op_1)
                except:
                    op_1 = 0

            if isinstance(op_2, str):
                try:
                    op_2 = float(op_1)
                except:
                    op_2 = 0

            if isinstance(op_1, (date, datetime)):
                base_date = datetime(1899, 12, 31)
                op_1 = (op_1 - base_date).days

            if isinstance(op_2, (date, datetime)):
                base_date = datetime(1899, 12, 31)

                op_2 = (op_2 - base_date).days
        else:
            op_1 = op_1.lower()
            op_2 = op_2.lower()
        return op_1 == op_2


class DiferentCondition(Operation):
    def __str__(self):
        return "DiferentCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        op_1 = self.operand_1.operate()
        op_2 = self.operand_2.operate()
        if op_1 is None or op_2 is None:
            return None

        if isinstance(op_1, str):
            try:
                op_1 = float(op_1)
            except:
                op_1 = 0

        if isinstance(op_2, str):
            try:
                op_2 = float(op_1)
            except:
                op_2 = 0

        if isinstance(op_1, (date, datetime)):
            base_date = datetime(1899, 12, 31)
            op_1 = (op_1 - base_date).days

        if isinstance(op_2, (date, datetime)):
            base_date = datetime(1899, 12, 31)

            op_2 = (op_2 - base_date).days

        return op_1 != op_2


class GratherCondition(Operation):
    def __str__(self):
        return "GratherCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        op_1 = self.operand_1.operate()
        op_2 = self.operand_2.operate()
        if op_1 is None or op_2 is None:
            return None

        if isinstance(op_1, str):
            try:
                op_1 = float(op_1)
            except:
                op_1 = 0

        if isinstance(op_2, str):
            try:
                op_2 = float(op_1)
            except:
                op_2 = 0

        if isinstance(op_1, (date, datetime)):
            base_date = datetime(1899, 12, 31)
            op_1 = (op_1 - base_date).days

        if isinstance(op_2, (date, datetime)):
            base_date = datetime(1899, 12, 31)

            op_2 = (op_2 - base_date).days

        return op_1 > op_2


class GratherOrEqualCondition(Operation):
    def __str__(self):
        return "GratherOrEqualCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        op_1 = self.operand_1.operate()
        op_2 = self.operand_2.operate()
        if op_1 is None or op_2 is None:
            return None

        if isinstance(op_1, str):
            try:
                op_1 = float(op_1)
            except:
                op_1 = 0

        if isinstance(op_2, str):
            try:
                op_2 = float(op_1)
            except:
                op_2 = 0

        if isinstance(op_1, (date, datetime)):
            base_date = datetime(1899, 12, 31)
            op_1 = (op_1 - base_date).days

        if isinstance(op_2, (date, datetime)):
            base_date = datetime(1899, 12, 31)

            op_2 = (op_2 - base_date).days

        return op_1 >= op_2


class LessCondition(Operation):
    def __str__(self):
        return "LessCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        op_1 = self.operand_1.operate()
        op_2 = self.operand_2.operate()
        if op_1 is None or op_2 is None:
            return None

        if isinstance(op_1, str):
            try:
                op_1 = float(op_1)
            except:
                op_1 = 0

        if isinstance(op_2, str):
            try:
                op_2 = float(op_1)
            except:
                op_2 = 0

        if isinstance(op_1, (date, datetime)):
            base_date = datetime(1899, 12, 31)
            op_1 = (op_1 - base_date).days

        if isinstance(op_2, (date, datetime)):
            base_date = datetime(1899, 12, 31)

            op_2 = (op_2 - base_date).days

        return op_1 < op_2


class LessOrEqualCondition(Operation):
    def __str__(self):
        return "LessOrEqualCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        op_1 = self.operand_1.operate()
        op_2 = self.operand_2.operate()
        if op_1 is None or op_2 is None:
            return None

        if isinstance(op_1, str):
            try:
                op_1 = float(op_1)
            except:
                op_1 = 0

        if isinstance(op_2, str):
            try:
                op_2 = float(op_1)
            except:
                op_2 = 0

        if isinstance(op_1, (date, datetime)):
            base_date = datetime(1899, 12, 31)
            op_1 = (op_1 - base_date).days

        if isinstance(op_2, (date, datetime)):
            base_date = datetime(1899, 12, 31)

            op_2 = (op_2 - base_date).days

        return op_1 <= op_2


class BlankCondition(Operation):
    def __str__(self):
        return "BlankCondition"

    def __init__(self, cell):
        self.cell = cell

    def operate(self):
        if self.cell.operate() is None:
            return True
        return False


class AndCondition(Operation):
    def __str__(self):
        return "AndCondition"

    def __init__(self, *args, **kargs):
        self.operands = list(args)

    def operate(self):
        acumulated_bool = True
        for op in self.operands:
            operated = op.operate()
            if operated is None:
                return None

            acumulated_bool = acumulated_bool and operated
            if not acumulated_bool:
                return acumulated_bool

        return acumulated_bool


class OrCondition(Operation):
    def __str__(self):
        return "OrCondition"

    def __init__(self, *args, **kargs):
        self.operands = list(args)

    def operate(self):

        acumulated_bool = False
        for op in self.operands:
            operated = op.operate()
            if operated is None:
                return None

            acumulated_bool = acumulated_bool or operated

        return acumulated_bool


class NotCondition(Operation):
    def __str__(self):
        return "NotCondition"

    def __init__(self, operand):
        self.operand = operand

    def operate(self):
        return not self.operand.operate()


class If(Operation):
    def __str__(self):
        return "If"

    def __init__(self, _condition, _then, _else=Value(False)):
        self._condition = _condition
        self._then = _then
        self._else = _else

    def operate(self):
        if self._condition.operate():
            return self._then.operate()
        return self._else.operate()


class CountIf(Operation):
    def __str__(self):
        return "CountIf"

    _condition_operation = {
        "<=": lambda left, right: left <= right,
        ">=": lambda left, right: left >= right,
        "<>": lambda left, right: left != right,
        "=": lambda left, right: left == right,
        "<": lambda left, right: left < right,
        ">": lambda left, right: left > right,
    }

    def __init__(self, cell_range, condition):
        self.condition, self.right_side = self._split_condition(condition)
        self.range = cell_range

    def operate(self):
        operation = self._condition_operation.get(self.condition, None)
        count = 0
        if operation is None:
            logging.warning("condition did not load correctly")
            return count
        if self.right_side is None:
            return None
        for cell in self.range.operate():
            cell_value = cell.calculated_value
            if operation(cell_value, self.right_side):
                count += 1
        return count

    def _split_condition(self, criteria):
        criteria = criteria.operate()

        if criteria is None:
            return "=", None

        pattern = r"(<=|>=|<>|=|<|>)(.*)"
        match = re.match(pattern, criteria)

        if match:
            return match.group(1), match.group(2).strip()
        else:
            return "=", criteria.strip()


class MRound(Operation):
    def __str__(self):
        return "MRound"

    def __init__(self, number, multiple):
        self.number = number
        self.multiple = multiple

    def operate(self):
        number = self.number.operate()
        mutiple = self.multiple.operate()
        if mutiple is None or number is None:
            return None
        if mutiple == 0:
            return 0
        # Excel rounds halves away from zero; Python's round() goes to the even neighbour.
        # ponytail: 12 decimals absorb float noise (0.15/0.1 = 1.4999999999999998)
        quotient = round(number / mutiple, 12)
        return math.copysign(math.floor(abs(quotient) + 0.5), quotient) * mutiple


class RoundDown(Operation):
    def __str__(self):
        return "RoundDown"

    def __init__(self, number, num_digits=Value(0)):
        self.number = number
        self.num_digits = num_digits

    def operate(self):
        number = self.number.operate()
        if self.num_digits.operate() is None or number is None:
            return None
        factor = 10 ** self.num_digits.operate()
        return math.floor(number * factor) / factor


class RoundUp(Operation):
    def __str__(self):
        return "RoundUp"

    def __init__(self, number, num_digits=Value(0)):
        self.number = number
        self.num_digits = num_digits

    def operate(self):
        number = self.number.operate()
        if self.num_digits.operate() is None or number is None:
            return None
        factor = 10 ** self.num_digits.operate()
        return math.ceil(number * factor) / factor


class BinaryOperation(Operation, ABC):
    def __str__(self):
        return "BinaryAdd"

    def __init__(self, left, right):
        self.left = left
        self.right = right


class BinaryAdd(BinaryOperation):
    def operate(self):
        left = self.left.operate()
        right = self.right.operate()
        if left is not None and right is not None:
            if isinstance(left, datetime) or isinstance(right, datetime):
                return self.add_dates(left, right)
            try:
                left = float(left)
                right = float(right)
            except:
                left = str(left)
                right = str(right)
            return left + right

    def add_dates(self, left, right):
        number = ""
        date = ""

        if isinstance(left, (int, float)):
            number = left
            date = right

        if isinstance(right, (int, float)):
            number = right
            date = left

        return date + timedelta(days=number)


class BinarySub(BinaryOperation):
    def operate(self):
        left = self.left.operate()
        right = self.right.operate()
        if left is not None and right is not None:
            return left - right


class BinaryMul(BinaryOperation):
    def operate(self):
        left = self.left.operate()
        right = self.right.operate()
        if left is not None and right is not None:
            if isinstance(left, str):
                left = float(left)

            if isinstance(right, str):
                right = float(right)

            return left * right


class BinaryDiv(BinaryOperation):
    def operate(self):
        left = self.left.operate()
        right = self.right.operate()
        if left is not None and right is not None:
            if right == 0:
                return None

            if isinstance(left, str):
                left = float(left)

            if isinstance(right, str):
                right = float(right)

            return left / right


class BinaryPow(BinaryOperation):
    def operate(self):
        left = self.left.operate()
        right = self.right.operate()
        if left is not None and right is not None:
            if isinstance(left, str):
                left = float(left)

            if isinstance(right, str):
                right = float(right)

            return left**right


def to_excel_text(value):
    """How Excel shows a value inside a concatenation: 1.0 -> '1', True -> 'TRUE'."""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return format(value, ".15g")
    return str(value)


class Concat(BinaryOperation):
    def operate(self):
        return to_excel_text(self.left.operate()) + to_excel_text(self.right.operate())


class Sum(Operation):
    def __str__(self):
        return "Sum"

    def __init__(self, cell_range):
        self.range = cell_range

    def operate(self):
        sum_calc = 0
        _range = self.range.operate()
        if not isinstance(_range, Iterable):
            return _range
        for cell in _range:
            if cell is None:
                continue
            cell_value = cell.calculated_value
            if isinstance(cell_value, (int, float, complex)):
                sum_calc += cell_value
        return sum_calc


class Hlookup(Operation):
    def __str__(self):
        return "Hlookup"

    def __init__(self, lookup_value, cell_range, row_index, approximated=True):
        self.lookup_value = lookup_value
        self.cell_range = cell_range
        self.row_index = row_index
        self.approximated = approximated

    def operate(self):
        first_row_indicator = self._get_indicator()
        return self._match(first_row_indicator)

    def _get_indicator(self):
        first_row_indicator = ""
        first_element = self.cell_range.operate()[0]

        if first_element is None:
            return None

        first_char = first_element[0]
        for char in first_char:
            first_row_indicator += char
        return first_row_indicator

    def _match(self, first_row_indicator):
        row_lenght = 100000
        column_index = None
        _range = self.cell_range.operate()
        for index, cell in enumerate(_range):
            if cell is None:
                continue
            value_type = str(type(self.lookup_value))
            same_type = str(type(cell.calculated_value)) != value_type
            if column_index is not None and same_type:
                if not approximated:
                    if self.lookup_value == cell.calculated_value:
                        column_index = index
                else:
                    if self.lookup_value < cell.calculated_value:
                        column_index = index

            if cell.coordinate.count(first_row_indicator) == 0:
                row_lenght = index + 1
                break

        if column_index is None:
            return None
        lookup_value_index = row_lenght * row_index + column_index
        if lookup_value_index > len(_range):
            logging.warning("index out of range")
            return None
        return _range[lookup_value_index]


class Text(Operation):
    def __str__(self):
        return "Text"

    def __init__(self, element, _format):
        self.element = element.operate()
        # print(self.element)
        self.format = _format.operate()

    def operate(self):
        if isinstance(self.element, (int, float)):
            return format(self.element, self.format)

        elif isinstance(self.element, (date, datetime)):
            formated_date = (
                self.format.replace("YYYY", "%Y")
                .replace("YY", "%y")
                .replace("MM", "%m")
                .replace("DDDD", "%A")
                .replace("DDD", "%a")
                .replace("DD", "%d")
                .replace("hh", "%H")
                .replace("mm", "%M")
                .replace("ss", "%S")
            )
            return self.element.strftime(formated_date)
        return self.element


class IfError(Operation):
    def __str__(self):
        return "IfError"

    def __init__(self, _then, _else):
        self._then = _then
        self._else = _else

    def operate(self):
        try:
            return self._then.operate()
        except:
            try:
                return self._else.operate()
            except:
                return None
