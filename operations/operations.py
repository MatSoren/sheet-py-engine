import logging
from abc import ABC, abstractmethod

from datetime import date, datetime, timedelta

import openpyxl.utils.cell as utils

import re

import math

from decimal import Decimal, ROUND_HALF_UP

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
        self.width = final_col - initial_col + 1

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


def excel_comparable(left, right):
    """Two values -> two keys Python compares the way Excel does.

    numbers < text < TRUE/FALSE, text is case-insensitive, dates are serial numbers, and a
    blank cell counts as 0, "" or FALSE depending on the other side.
    """
    if left is None:
        left = _blank_like(right)
    if right is None:
        right = _blank_like(left)
    return _comparison_key(left), _comparison_key(right)


def _blank_like(other):
    if isinstance(other, bool):
        return False
    if isinstance(other, str):
        return ""
    return 0


def _comparison_key(value):
    if isinstance(value, bool):
        return (2, value)
    if isinstance(value, str):
        return (1, value.lower())
    if isinstance(value, datetime):
        return (0, (value - datetime(1899, 12, 30)).total_seconds() / 86400)
    if isinstance(value, date):
        return (0, (value - date(1899, 12, 30)).days)
    return (0, value)


class EqualCondition(Operation):
    def __str__(self):
        return "EqualCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        left, right = excel_comparable(self.operand_1.operate(), self.operand_2.operate())
        return left == right


class DiferentCondition(Operation):
    def __str__(self):
        return "DiferentCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        left, right = excel_comparable(self.operand_1.operate(), self.operand_2.operate())
        return left != right


class GratherCondition(Operation):
    def __str__(self):
        return "GratherCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        left, right = excel_comparable(self.operand_1.operate(), self.operand_2.operate())
        return left > right


class GratherOrEqualCondition(Operation):
    def __str__(self):
        return "GratherOrEqualCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        left, right = excel_comparable(self.operand_1.operate(), self.operand_2.operate())
        return left >= right


class LessCondition(Operation):
    def __str__(self):
        return "LessCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        left, right = excel_comparable(self.operand_1.operate(), self.operand_2.operate())
        return left < right


class LessOrEqualCondition(Operation):
    def __str__(self):
        return "LessOrEqualCondition"

    def __init__(self, operand_1, operand_2):
        self.operand_1 = operand_1
        self.operand_2 = operand_2

    def operate(self):
        left, right = excel_comparable(self.operand_1.operate(), self.operand_2.operate())
        return left <= right


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
        self.range = cell_range
        self.condition = condition

    def operate(self):
        # evaluated here, not in __init__: the criterion may live in a cell loaded later
        operator, target = self._split_condition(self.condition.operate())
        operation = self._condition_operation[operator]
        values = self.range.operate()
        if not isinstance(values, list):  # COUNTIF(A1, ...) gets a Reference, not a Range
            values = [values]
        else:
            values = [cell.calculated_value for cell in values]
        return sum(1 for value in values if self._matches(value, operator, operation, target))

    def _split_condition(self, criteria):
        """'>=10' -> ('>=', 10.0), 'abc' -> ('=', 'abc'), 5 -> ('=', 5); empty counts as 0."""
        if criteria is None:
            return "=", 0
        if not isinstance(criteria, str):
            return "=", criteria
        operator, target = re.match(r"(<=|>=|<>|=|<|>)?(.*)", criteria, re.DOTALL).groups()
        try:
            target = float(target)
        except ValueError:
            pass
        return operator or "=", target

    @staticmethod
    def _matches(value, operator, operation, target):
        # ponytail: no wildcards (* ? ~) and numbers stored as text don't match numeric criteria
        if isinstance(target, str):
            value = "" if value is None else value
            if not isinstance(value, str):
                return operator == "<>"
            return operation(value.lower(), target.lower())
        if value is None or isinstance(value, str) or isinstance(value, bool) != isinstance(target, bool):
            return operator == "<>"
        return operation(value, target)


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


def round_magnitude(number, num_digits, rounding):
    """ROUNDUP/ROUNDDOWN: apply floor/ceil to the absolute value, so the sign doesn't flip
    the direction. 12 decimals absorb float noise first (1.1*3 = 3.3000000000000003)."""
    factor = 10**num_digits
    result = math.copysign(rounding(round(abs(number) * factor, 12)) / factor, number)
    return result or 0.0  # no -0.0


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
        return round_magnitude(number, self.num_digits.operate(), math.floor)  # toward zero


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
        return round_magnitude(number, self.num_digits.operate(), math.ceil)  # away from zero


class BinaryOperation(Operation, ABC):
    def __str__(self):
        return "BinaryAdd"

    def __init__(self, left, right):
        self.left = left
        self.right = right

    def operands(self):
        """Both sides evaluated; a blank cell counts as 0 in arithmetic, like in Excel.

        ponytail: errors are None too, so an error feeding arithmetic also reads as 0;
        a real error value would fix it.
        """
        return [0 if value is None else value for value in (self.left.operate(), self.right.operate())]

    @staticmethod
    def as_number(value):
        return float(value) if isinstance(value, str) else value


class BinaryAdd(BinaryOperation):
    def operate(self):
        left, right = self.operands()
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
        left, right = self.operands()
        return left - right


class BinaryMul(BinaryOperation):
    def operate(self):
        left, right = map(self.as_number, self.operands())
        return left * right


class BinaryDiv(BinaryOperation):
    def operate(self):
        left, right = map(self.as_number, self.operands())
        if right == 0:
            return None  # Excel: #DIV/0!
        return left / right


class BinaryPow(BinaryOperation):
    def operate(self):
        left, right = map(self.as_number, self.operands())
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

    def __init__(self, lookup_value, cell_range, row_index, approximated=Value(True)):
        self.lookup_value = lookup_value
        self.cell_range = cell_range
        self.row_index = row_index
        self.approximated = approximated

    def operate(self):
        lookup = self.lookup_value.operate()
        row_index = self.row_index.operate()
        if lookup is None or row_index is None:
            return None
        cells = self.cell_range.operate()
        width = self.cell_range.width
        first_row = [cell.calculated_value for cell in cells[:width]]
        column = self._find_column(first_row, lookup, self.approximated.operate())
        index = (int(row_index) - 1) * width + (column or 0)
        if column is None or not 0 <= index < len(cells):
            return None  # Excel: #N/A / #REF!
        return cells[index].calculated_value

    @staticmethod
    def _find_column(first_row, lookup, approximated):
        """Exact: first match. Approximate: last value <= lookup, assuming an ascending row.

        Like Excel, values of the other kind (text vs number) are skipped, text is
        case-insensitive.
        """
        is_text = isinstance(lookup, str)
        key = lookup.lower() if is_text else lookup
        found = None
        for column, value in enumerate(first_row):
            if value is None or isinstance(value, str) != is_text:
                continue
            value = value.lower() if is_text else value
            if value == key:
                return column
            if approximated:
                if value > key:
                    break
                found = column
        return found if approximated else None


class Text(Operation):
    def __str__(self):
        return "Text"

    def __init__(self, element, _format):
        self.element = element
        self.format = _format

    def operate(self):
        # evaluated here, not in __init__, so TEXT follows changes to its source cell
        element = self.element.operate()
        _format = self.format.operate()
        if isinstance(element, (int, float)) and not isinstance(element, bool):
            if not re.search(r"[YMDHS]", _format, re.IGNORECASE):
                return self._number_text(element, _format)
            element = datetime(1899, 12, 30) + timedelta(days=element)  # a date is a serial number

        if isinstance(element, (date, datetime)):
            formated_date = (
                _format.replace("YYYY", "%Y")
                .replace("YY", "%y")
                .replace("MM", "%m")
                .replace("DDDD", "%A")
                .replace("DDD", "%a")
                .replace("DD", "%d")
                .replace("hh", "%H")
                .replace("mm", "%M")
                .replace("ss", "%S")
            )
            return element.strftime(formated_date)
        return element

    @staticmethod
    def _number_text(value, _format):
        # ponytail: 0 / 0.00 / #,##0.00 / 0% style formats only; no sections, no literals
        percent = "%" in _format
        if percent:
            value *= 100
        decimals = len(_format.split(".")[1].rstrip("%")) if "." in _format else 0
        rounded = Decimal(repr(value)).quantize(Decimal(1).scaleb(-decimals), ROUND_HALF_UP)
        return f"{rounded:{',' if ',' in _format else ''}f}" + ("%" if percent else "")


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
