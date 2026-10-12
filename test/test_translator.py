"""Tests for the formula -> Python translation and its end-to-end evaluation.

    python -m unittest test.test_translator -v

Two layers:
- formula_to_python() is pure (string -> string): checked against the exact expected
  expression, and that ast.parse accepts it. No workbook needed.
- End to end: a workbook built in memory with openpyxl goes through translator.translate()
  and the computed cell values are compared with what Excel gives.
"""

import ast
import collections
import unittest
from datetime import datetime

from openpyxl import Workbook

from Sheet import SheetsManager
from SheetCell import SheetCell
from operations.operations import Value
from translator import translator
from translator.AstProcessor import decode_ref, encode_ref, normalize_sheet_name
from translator.translator import formula_to_python


class TranslationCase(unittest.TestCase):
    def check(self, cases):
        for formula, expected in cases:
            with self.subTest(formula=formula):
                result = formula_to_python(formula)
                self.assertEqual(result, expected)
                ast.parse(result)


class TestReferences(TranslationCase):
    def test_plain_and_absolute(self):
        self.check([
            ("=A1", "A1"),
            ("=AA100", "AA100"),
            ("=XFD1048576", "XFD1048576"),
            ("=$A$1", "A1"),
            ("=$A1", "A1"),
            ("=A$1", "A1"),
            ("=$A$1+1", "A1+1"),
        ])

    def test_ranges(self):
        self.check([
            ("=SUM(A1:B2)", "SUM(A1___range___B2)"),
            ("=SUM($A$1:$B$2)", "SUM(A1___range___B2)"),
            ("=SUM(A$1:$B2)", "SUM(A1___range___B2)"),
        ])

    def test_other_sheet(self):
        self.check([
            ("=Hoja!A1", "_Hoja___ref___A1"),
            ("=Hoja!$A$1", "_Hoja___ref___A1"),
            ("=Hoja!A1:B2", "_Hoja___ref___A1___range___B2"),
            ("='Mi Hoja'!A1", "_Mi_Hoja___ref___A1"),
            ("='Mi Hoja'!$A$1:$B$2", "_Mi_Hoja___ref___A1___range___B2"),
            ("='2024 Ventas'!A1", "_2024_Ventas___ref___A1"),
            ("='2024'!A1", "_2024___ref___A1"),
            ("='Costos-Fijos'!A1", "_Costos_Fijos___ref___A1"),
            ("='Costos-Fijos'!A1:B2", "_Costos_Fijos___ref___A1___range___B2"),
            ("='Año 1'!A1", "_A_o_1___ref___A1"),
            ("='O''Brien'!A1", "_O_Brien___ref___A1"),
            ("='Hoja.1'!A1", "_Hoja_1___ref___A1"),
            ("='a ref b'!A1", "_a_ref_b___ref___A1"),
            ("=Hoja!A1+'Otra Hoja'!B2", "_Hoja___ref___A1+_Otra_Hoja___ref___B2"),
            # sheet repeated on both ends of the range
            ("=SUM(Hoja!A1:Hoja!B2)", "SUM(_Hoja___ref___A1___range___B2)"),
            ("=SUM(Hoja!$A$1:Hoja!$B$2)", "SUM(_Hoja___ref___A1___range___B2)"),
            ("=SUM('Mi Hoja'!A1:'Mi Hoja'!B2)", "SUM(_Mi_Hoja___ref___A1___range___B2)"),
        ])

    def test_ref_error(self):
        self.check([
            ("=#REF!", "___error___REF"),
            ("=Hoja!#REF!", "___error___REF"),
            ("=#REF!+1", "___error___REF+1"),
            ("=SUM(#REF!)", "SUM(___error___REF)"),
        ])


class TestTextOperands(TranslationCase):
    def test_contents_untouched(self):
        self.check([
            ('="hola"', "'hola'"),
            ('=""', "''"),
            ('="a:b"', "'a:b'"),
            ('="x=y"', "'x=y'"),
            ('="a<>b"', "'a<>b'"),
            ('="Hoja!A1"', "'Hoja!A1'"),
            ('="$A$1"', "'$A$1'"),
            ('="#N/A"', "'#N/A'"),
            ('="TRUE"', "'TRUE'"),
            ('="50%"', "'50%'"),
            ('="1^2"', "'1^2'"),
            ('="a&b"', "'a&b'"),
            ('="a,b"', "'a,b'"),
            ('="(x)"', "'(x)'"),
            ('="{1,2}"', "'{1,2}'"),
            ('="  dos  espacios  "', "'  dos  espacios  '"),
            ('="ñandú"', "'ñandú'"),
            ('="it\'s"', '"it\'s"'),
            ('="back\\slash"', "'back\\\\slash'"),
            ('="if(and("', "'if(and('"),
        ])

    def test_escaped_quotes(self):
        self.check([
            ('="dijo ""hola"""', "'dijo \"hola\"'"),
            ('=""""', "'\"'"),
            ('="a""b"', "'a\"b'"),
        ])

    def test_text_inside_condition(self):
        self.check([
            ('=IF(A1="a:b","x","y")', "IF(A1=='a:b','x','y')"),
            ('=IF(A1="=","eq","ne")', "IF(A1=='=','eq','ne')"),
            ('=COUNTIF(A1:A3,">=10")', "COUNTIF(A1___range___A3,'>=10')"),
        ])


class TestLogicalNumberErrorOperands(TranslationCase):
    def test_logical(self):
        self.check([
            ("=TRUE", "True"),
            ("=FALSE", "False"),
            ("=IF(B1=TRUE,1,0)", "IF(B1==True,1,0)"),
            ("=AND(TRUE,FALSE)", "AND(True,False)"),
        ])

    def test_numbers(self):
        self.check([
            ("=1", "1"),
            ("=1.5", "1.5"),
            ("=0.25", "0.25"),
            ("=1E+3", "1E+3"),
            ("=1.5E-3", "1.5E-3"),
        ])

    def test_errors(self):
        self.check([
            ("=#N/A", "___error___NA"),
            ("=#DIV/0!", "___error___DIV0"),
            ("=#VALUE!", "___error___VALUE"),
            ("=#NAME?", "___error___NAME"),
            ("=#NUM!", "___error___NUM"),
            ("=#NULL!", "___error___NULL"),
            ("=#GETTING_DATA", "___error___GETTING_DATA"),
            ("=IFERROR(A1,#N/A)", "IFERROR(A1,___error___NA)"),
            ("=IF(A1>0,#DIV/0!,1)", "IF(A1>0,___error___DIV0,1)"),
            ('="#N/A"', "'#N/A'"),  # text that looks like an error stays text
        ])


class TestOperators(TranslationCase):
    def test_infix(self):
        self.check([
            ("=1+2", "1+2"),
            ("=1-2", "1-2"),
            ("=1*2", "1*2"),
            ("=1/2", "1/2"),
            ("=2^3", "2**3"),
            ('="a"&"b"', "'a'|'b'"),
            ("=1+2&3", "1+2|3"),
            ("=A1=B1", "A1==B1"),
            ("=A1<>B1", "A1!=B1"),
            ("=A1<B1", "A1<B1"),
            ("=A1>B1", "A1>B1"),
            ("=A1<=B1", "A1<=B1"),
            ("=A1>=B1", "A1>=B1"),
            ("=IF(A1<>0,1,2)", "IF(A1!=0,1,2)"),
        ])

    def test_prefix(self):
        self.check([
            ("=-A1", "(-A1)"),
            ("=+A1", "(+A1)"),
            ("=-1", "(-1)"),
            ("=--A1", "(-(-A1))"),
            ("=-(A1+1)", "(-(A1+1))"),
            ("=-SUM(A1:A2)", "(-SUM(A1___range___A2))"),
            ("=1--1", "1-(-1)"),
            ("=A1*-1", "A1*(-1)"),
            ("=SUM(-A1,2)", "SUM((-A1),2)"),
        ])

    def test_power_precedence_and_associativity(self):
        self.check([
            ("=-2^2", "(-2)**2"),
            ("=-A1^2", "(-A1)**2"),
            ("=2^-1", "2**(-1)"),
            ("=2^3^2", "(2**3)**2"),
            ("=2^3^2^2", "((2**3)**2)**2"),
            ("=1+2^3^2", "1+(2**3)**2"),
            ("=2^3*2^2", "2**3*2**2"),
            ("=(2^3)^2", "(2**3)**2"),
            ("=2^(3^2)", "2**(3**2)"),
            ("=SUM(2^3^2,2^2)", "SUM((2**3)**2,2**2)"),
            ("=-SUM(A1)^2", "(-SUM(A1))**2"),
            ("=2^3%^2", "(2**(3/100))**2"),
        ])

    def test_percent(self):
        self.check([
            ("=50%", "(50/100)"),
            ("=A1%", "(A1/100)"),
            ("=1/50%", "1/(50/100)"),
            ("=A1*10%", "A1*(10/100)"),
            ("=(1+1)%", "((1+1)/100)"),
            ("=SUM(A1:A2)%", "(SUM(A1___range___A2)/100)"),
            ("=-50%", "((-50)/100)"),
            ("=50%%", "((50/100)/100)"),
            ("=2^50%", "2**(50/100)"),
            ("=IF(A1>50%,1,0)", "IF(A1>(50/100),1,0)"),
            ("=Hoja!A1%", "(_Hoja___ref___A1/100)"),
        ])


class TestFunctionsAndArguments(TranslationCase):
    def test_case_insensitive_names(self):
        # if/and/or/not are Python keywords: lowercase would not even parse
        self.check([
            ("=if(A1>0,1,2)", "IF(A1>0,1,2)"),
            ("=If(A1>0,1,2)", "IF(A1>0,1,2)"),
            ("=and(A1,B1)", "AND(A1,B1)"),
            ("=or(A1,B1)", "OR(A1,B1)"),
            ("=not(A1)", "NOT(A1)"),
            ("=sum(A1:A3)", "SUM(A1___range___A3)"),
            ("=Sum(A1:A3)", "SUM(A1___range___A3)"),
        ])

    def test_nesting(self):
        self.check([
            ("=IF(AND(A1>0,B1<>0),1,0)", "IF(AND(A1>0,B1!=0),1,0)"),
            ("=ROUNDUP(SUM(A1:A3)*0.21,2)", "ROUNDUP(SUM(A1___range___A3)*0.21,2)"),
            ("=TODAY()", "TODAY()"),
            ("=IF(A1,IF(B1,1,2),3)", "IF(A1,IF(B1,1,2),3)"),
        ])

    def test_empty_arguments(self):
        self.check([
            ("=IF(A1,,1)", "IF(A1,None,1)"),
            ("=IF(A1,1,)", "IF(A1,1,None)"),
            ("=IF(A1,,)", "IF(A1,None,None)"),
            ("=ROUND(,2)", "ROUND(None,2)"),
            ("=IF(A1,(1),)", "IF(A1,(1),None)"),
        ])

    def test_whitespace_is_dropped(self):
        self.check([
            ("= A1", "A1"),
            ("=A1 + 1", "A1+1"),
            ("=SUM( A1:A3 )", "SUM(A1___range___A3)"),
            ('=IF(A1 = 1, "a b", "c")', "IF(A1==1,'a b','c')"),
            ("=- A1", "(-A1)"),
            ("=A1\n+1", "A1+1"),
        ])


class TestUnsupported(unittest.TestCase):
    def test_intersection_operator(self):
        for formula in ["=A1:A3 B1:B3", "=SUM(A1:B2 B1:C3)", "=(A1:A3) (B1:B3)", "=A1 B1"]:
            with self.subTest(formula=formula):
                with self.assertRaisesRegex(ValueError, "intersection"):
                    formula_to_python(formula)

    def test_3d_references(self):
        for formula in ["=SUM(Hoja1:Hoja3!A1)", "=SUM('Hoja 1:Hoja 3'!A1:B2)"]:
            with self.subTest(formula=formula):
                with self.assertRaisesRegex(ValueError, "3D"):
                    formula_to_python(formula)

    def test_array_constants(self):
        for formula in ["={1,2,3}", "=SUM({1,2;3,4})"]:
            with self.subTest(formula=formula):
                with self.assertRaisesRegex(ValueError, "array"):
                    formula_to_python(formula)


class TestSheetNameEncoding(unittest.TestCase):
    def test_normalize(self):
        cases = [
            ("Items", "Items"),
            ("Cost Report", "Cost_Report"),
            ("2024 Ventas", "2024_Ventas"),
            ("Costos-Fijos", "Costos_Fijos"),
            ("Año 1", "A_o_1"),
            ("O'Brien", "O_Brien"),
        ]
        for name, expected in cases:
            with self.subTest(name=name):
                self.assertEqual(normalize_sheet_name(name), expected)
                self.assertEqual(normalize_sheet_name(expected), expected)

    def test_round_trip(self):
        cases = [
            (None, "A1"),
            (None, "A1:B2"),
            ("Items", "A1"),
            ("2024 Ventas", "B7:C9"),
            ("a ref b", "A1"),
            ("x___ref___y", "A1:B2"),
        ]
        for sheet, ref in cases:
            with self.subTest(sheet=sheet, ref=ref):
                identifier = encode_ref(sheet, ref)
                self.assertTrue(identifier.isidentifier(), identifier)
                expected_sheet = None if sheet is None else normalize_sheet_name(sheet)
                self.assertEqual(decode_ref(identifier), (expected_sheet, ref))


def load(sheets):
    """{sheet name: {coordinate: value or formula}} -> SheetsManager, via translate()."""
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, cells in sheets.items():
        worksheet = workbook.create_sheet(name)
        for coordinate, value in cells.items():
            worksheet[coordinate] = value
    # ponytail: the registry is class-level (one workbook per process), so reset it per load
    SheetsManager._sheet_dict.clear()
    return translator.translate(workbook)


def value(manager, sheet, coordinate):
    return manager.get_sheet(normalize_sheet_name(sheet)).get_cell(coordinate).calculated_value


DATA = {
    "A1": 10, "A2": 20, "A3": 30,
    "B1": True, "B2": False,
    "C1": "abc", "C2": "5",
    "D1": 2.5, "D2": 0, "D3": 2.0,
    "E1": "#N/A",  # openpyxl stores error codes as error cells
    # horizontal lookup table, H10:K12
    "H10": "code", "I10": 1, "J10": 2, "K10": 3,
    "H11": "name", "I11": "uno", "J11": "dos", "K11": "tres",
    "H12": "amount", "I12": 100, "J12": 200, "K12": 300,
}

OTHER_SHEETS = {
    "2024 Ventas": {"A1": 100, "A2": 200},
    "Costos-Fijos": {"A1": 1, "A2": 2, "A3": 3},
    "Año 1": {"A1": 7},
    "O'Brien": {"A1": 5},
    "Cost Report": {"A1": 11},
}

# (formula, value Excel gives)
EVALUATION_CASES = [
    # references
    ("=A1", 10),
    ("=$A$1", 10),
    ("=$A$1+1", 11),
    ("=A$1*$A2", 200),
    ("=SUM($A$1:$A$3)", 60),
    # arithmetic and precedence
    ("=1+2*3", 7),
    ("=(1+2)*3", 9),
    ("=2^3", 8),
    ("=2*3^2", 18),
    ("=A1^2", 100),
    ("=D1^2", 6.25),
    ("=4^0.5", 2.0),
    # comparisons
    ("=A1=10", True),
    ("=A1<>10", False),
    ("=A1<>0", True),
    ("=A1<20", True),
    ("=A1>20", False),
    ("=A1<=10", True),
    ("=A1>=11", False),
    ("=IF(A1<>0,1,2)", 1),
    ("=IF(D2<>0,1,2)", 2),
    ('=C1="ABC"', True),
    # a blank cell counts as 0 in arithmetic
    ("=Z99*A1", 0),
    ("=A1*Z99", 0),
    ("=Z99+1", 1),
    ("=A1-Z99", 10),
    ("=Z99-A1", -10),
    ("=Z99/A1", 0.0),
    ("=A1/Z99", "#DIV/0!"),
    ('=A1/"0"', "#DIV/0!"),
    ("=Z99^2", 0),
    ("=-Z99", 0),
    # comparisons follow Excel: numbers < text < logicals, text case-insensitive
    ('="1"<"12"', True),
    ('="b">"A"', True),
    ('="abc"<>"ABC"', False),
    ('="a"<>"b"', True),
    ('=C1<>"x"', True),
    ('="junior"<2', False),
    ('="5"=5', False),
    ("=C2>1000", True),
    ("=TRUE>1000", True),
    ('=TRUE>"z"', True),
    ("=A2<=A1", False),
    ("=TODAY()>A1", True),
    # a blank cell compares as 0, "" or FALSE
    ("=Z99=0", True),
    ('=Z99=""', True),
    ("=Z99<1", True),
    ("=Z99=FALSE", True),
    ('=Z99<"a"', True),
    ("=Z99=Z98", True),
    ('=IF(Z99<"12","NONE","x")', "NONE"),
    # concatenation
    ("=1+2&3", "33"),
    ('="a"&"b"', "ab"),
    ('="a"&1+2', "a3"),
    ('="n="&A1', "n=10"),
    ('="v"&D1', "v2.5"),
    ('="v"&D3', "v2"),
    ('=TRUE&"x"', "TRUEx"),
    ('=B2&"x"', "FALSEx"),
    ('="x"&Z99', "x"),
    ('="Total: "&SUM(A1:A3)', "Total: 60"),
    ("=A1&A2&A3", "102030"),
    # text untouched
    ('="a:b"', "a:b"),
    ('="x=y"', "x=y"),
    ('="Hoja!A1"', "Hoja!A1"),
    ('="dijo ""hola"""', 'dijo "hola"'),
    ('="it\'s"', "it's"),
    ('="ñandú"', "ñandú"),
    ('="  dos  espacios  "', "  dos  espacios  "),
    ('=IF(C1="a:b","x","y")', "y"),
    ('=IF("a:b"="a:b","x","y")', "x"),
    # logical
    ("=TRUE", True),
    ("=FALSE", False),
    ("=IF(B1=TRUE,1,0)", 1),
    ("=IF(B2=TRUE,1,0)", 0),
    ("=IF(B2=FALSE,1,0)", 1),
    ("=AND(TRUE,B1)", True),
    ("=OR(FALSE,B2)", False),
    # errors
    ("=#N/A", "#N/A"),
    ("=#REF!", "#REF!"),
    ("=IFERROR(A1,#N/A)", 10),
    ("=IFERROR(C1*1,#N/A)", "#N/A"),
    ("=IFERROR(C2*1,#N/A)", 5.0),
    # errors propagate until IFERROR catches them
    ("=1/0", "#DIV/0!"),
    ("=1/0+1", "#DIV/0!"),
    ('="x"&1/0', "#DIV/0!"),
    ("=IFERROR(1/0,\"div\")", "div"),
    ('=C1*1', "#VALUE!"),
    ("=E1", "#N/A"),
    ("=E1+1", "#N/A"),
    ("=-E1", "#N/A"),
    ("=E1=E1", "#N/A"),
    ("=SUM(E1:E2)", "#N/A"),
    ("=IF(E1>0,1,2)", "#N/A"),
    ("=IF(TRUE,1,E1)", 1),
    ("=IFERROR(E1,0)", 0),
    ("=IFERROR(E1,1/0)", "#DIV/0!"),
    ("=COUNTIF(E1:E2,1)", 0),
    ("=HLOOKUP(E1,I10:K12,2)", "#N/A"),
    ('=IFERROR(HLOOKUP(9,I10:K12,2,FALSE),"none")', "none"),
    ('=IFERROR("#N/A","not an error")', "#N/A"),
    ('=IFERROR(#REF!,"ref")', "ref"),
    ("=ROUNDUP(1/0,1)", "#DIV/0!"),
    ('=TEXT(1/0,"0")', "#DIV/0!"),
    # percent
    ("=50%", 0.5),
    ("=1/50%", 2.0),
    ("=A1*10%", 1.0),
    ("=(1+1)%", 0.02),
    ("=SUM(A1:A2)%", 0.3),
    ("=-50%", -0.5),
    ("=50%%", 0.005),
    ("=2^50%", 1.4142135624),
    # unary
    ("=-A1", -10),
    ("=-1", -1),
    ("=+A1", 10),
    ("=--A1", 10),
    ("=-(A1+1)", -11),
    ("=-SUM(A1:A2)", -30),
    ("=1--1", 2),
    ("=A1*-1", -10),
    ("=-C2", -5.0),
    # functions, any case, empty arguments
    ("=if(A1>5,\"big\",\"small\")", "big"),
    ("=Sum(A1:A3)", 60),
    ("=and(B1,TRUE)", True),
    ("=IF(B2,,2)", 2),
    ("=IF(B1,,2)", None),
    ("=ROUNDUP(SUM(A1:A3)*0.21,1)", 12.6),
    # COUNTIF
    ("=COUNTIF(A1:A3,\">15\")", 2),
    ("=COUNTIF(A1:A3,\"<=20\")", 2),
    ("=COUNTIF(A1:A3,20)", 1),
    ("=COUNTIF(A1:A3,\"20\")", 1),
    ("=COUNTIF(A1:A3,\"=20\")", 1),
    ("=COUNTIF(A1:A3,\"<>20\")", 2),
    ("=COUNTIF(A1:A3,D3*10)", 1),
    ("=COUNTIF(C1:C2,\"ABC\")", 1),
    ("=COUNTIF(C1:C2,\"<>abc\")", 1),
    ("=COUNTIF(C1,\"abc\")", 1),
    ("=COUNTIF(A1:C1,TRUE)", 1),
    ("=COUNTIF(A1:A3,Z99)", 0),
    ("=COUNTIF(Z1:Z3,\"\")", 3),
    ("=COUNTIF('Costos-Fijos'!A1:A3,\">1\")", 2),
    # HLOOKUP
    ("=HLOOKUP(2,I10:K12,2)", "dos"),
    ("=HLOOKUP(2,I10:K12,3,FALSE)", 200),
    ("=HLOOKUP(2.5,I10:K12,3)", 200),
    ("=HLOOKUP(2.5,I10:K12,3,FALSE)", "#N/A"),
    ("=HLOOKUP(99,I10:K12,2)", "tres"),
    ("=HLOOKUP(0,I10:K12,2)", "#N/A"),
    ("=HLOOKUP(2,H10:K12,3)", 200),
    ("=HLOOKUP(\"DOS\",I11:K12,2,FALSE)", 200),
    ("=HLOOKUP(\"CODE\",H10:K12,3,FALSE)", "amount"),
    ("=HLOOKUP(A1/5,$I$10:$K$12,3)", 200),
    ("=HLOOKUP(2,I10:K12,1)", 2),
    ("=HLOOKUP(2,I10:K12,4)", "#REF!"),
    # TEXT
    ('=TEXT(1234.5,"0")', "1235"),
    ('=TEXT(2.5,"0")', "3"),
    ('=TEXT(D1,"0.00")', "2.50"),
    ('=TEXT(1234.567,"#,##0.00")', "1,234.57"),
    ('=TEXT(0.125,"0%")', "13%"),
    ('=TEXT(0.125,"0.0%")', "12.5%"),
    ('=TEXT(-1.005,"0.00")', "-1.01"),
    ('=TEXT(45292,"DD/MM/YYYY")', "01/01/2024"),
    ('=TEXT(45292,"DDDD")', "Monday"),
    ('=TEXT(C1,"0")', "abc"),
    # ROUNDUP goes away from zero, ROUNDDOWN toward zero, both ignoring float noise
    ("=ROUNDUP(1.21,1)", 1.3),
    ("=ROUNDUP(-1.21,1)", -1.3),
    ("=ROUNDUP(-0.1,0)", -1.0),
    ("=ROUNDUP(1.1*3,1)", 3.3),
    ("=ROUNDUP(1234,-2)", 1300.0),
    ("=ROUNDUP(2,0)", 2.0),
    ("=ROUNDUP(D1)", 3.0),
    ("=ROUNDDOWN(1.29,1)", 1.2),
    ("=ROUNDDOWN(-1.29,1)", -1.2),
    ("=ROUNDDOWN(0.3*3,1)", 0.9),
    ("=ROUNDDOWN(1299,-2)", 1200.0),
    ('="x"&ROUNDDOWN(-0.1,0)', "x0"),
    # MROUND rounds halves away from zero, like Excel (Python's round() would give 10 / 300)
    ("=MROUND(12.5,5)", 15),
    ("=MROUND(325,50)", 350),
    ("=MROUND(12.4,5)", 10),
    ("=MROUND(7.5,5)", 10),
    ("=MROUND(-12.5,-5)", -15),
    ("=MROUND(0.15,0.1)", 0.2),
    ("=MROUND(A1,0)", 0),
    # other sheets
    ("='2024 Ventas'!A1", 100),
    ("='2024 Ventas'!$A$2", 200),
    ("=SUM('2024 Ventas'!A1:A2)", 300),
    ("='Costos-Fijos'!A2", 2),
    ("=SUM('Costos-Fijos'!A1:A3)", 6),
    ("=SUM('Costos-Fijos'!A1:'Costos-Fijos'!A3)", 6),
    ("='Año 1'!A1", 7),
    ("='O''Brien'!A1", 5),
    ("='Cost Report'!A1", 11),
    ("='2024 Ventas'!A1+'Costos-Fijos'!A1&\"x\"", "101x"),
    # whitespace
    ("= A1 + 1", 11),
    ("=SUM( A1:A3 )", 60),
]


class TestEvaluation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        main = dict(DATA)
        for row, (formula, _) in enumerate(EVALUATION_CASES, start=1):
            main[f"F{row}"] = formula
        cls.manager = load({"Main": main, **OTHER_SHEETS})

    def test_cases(self):
        for row, (formula, expected) in enumerate(EVALUATION_CASES, start=1):
            with self.subTest(formula=formula):
                self.assertEqual(value(self.manager, "Main", f"F{row}"), expected)

    def test_sheets_registered_under_normalized_name(self):
        for name in OTHER_SHEETS:
            with self.subTest(name=name):
                self.assertIsNotNone(self.manager.get_sheet(normalize_sheet_name(name)))


class TestPropagation(unittest.TestCase):
    def test_cross_sheet_reference(self):
        manager = load({
            "Main": {
                "A1": "='2024 Ventas'!A1*2",
                "A2": "=SUM('Costos-Fijos'!A1:A3)",
                "A3": '="total "&A1+A2',
                "A4": "=A1^2",
                "A5": "=IF(A1<>200,\"cambió\",\"igual\")",
                "A6": "=-A1",
                "A7": "=A1*10%",
            },
            "2024 Ventas": {"A1": 100},
            "Costos-Fijos": {"A1": 1, "A2": 2, "A3": 3},
        })
        self.assertEqual(value(manager, "Main", "A3"), "total 206")
        self.assertEqual(value(manager, "Main", "A5"), "igual")

        manager.get_sheet(normalize_sheet_name("2024 Ventas")).get_cell("A1").set_operation(Value(150))
        self.assertEqual(value(manager, "Main", "A1"), 300)
        self.assertEqual(value(manager, "Main", "A3"), "total 306")
        self.assertEqual(value(manager, "Main", "A4"), 90000)
        self.assertEqual(value(manager, "Main", "A5"), "cambió")
        self.assertEqual(value(manager, "Main", "A6"), -300)
        self.assertEqual(value(manager, "Main", "A7"), 30.0)

        manager.get_sheet(normalize_sheet_name("Costos-Fijos")).get_cell("A3").set_operation(Value(10))
        self.assertEqual(value(manager, "Main", "A2"), 13)
        self.assertEqual(value(manager, "Main", "A3"), "total 313")

    def test_range_over_cells_loaded_later(self):
        # the formula is translated before the cells it sums exist (later row / later sheet)
        manager = load({
            "Main": {
                "A1": "=SUM(A5:A6)",
                "A2": "=SUM(Later!A1:A2)",
                "A3": "=COUNTIF(A5:A6,Later!B1)",
                "A5": 1,
                "A6": 2,
            },
            "Later": {"A1": 10, "A2": 20, "B1": 2},
        })
        self.assertEqual(value(manager, "Main", "A1"), 3)
        self.assertEqual(value(manager, "Main", "A2"), 30)
        self.assertEqual(value(manager, "Main", "A3"), 1)
        manager.get_sheet("Later").get_cell("A2").set_operation(Value(5))
        self.assertEqual(value(manager, "Main", "A2"), 15)
        # the criterion is re-read on every update, not frozen at load time
        manager.get_sheet("Later").get_cell("B1").set_operation(Value(">=1"))
        self.assertEqual(value(manager, "Main", "A3"), 2)

    def test_text_follows_its_source(self):
        manager = load({
            "Main": {
                "A1": "=TEXT(Later!A1,\"DDD\")",
                "A2": "=TEXT(Later!A1+1,\"DD/MM/YYYY\")",
            },
            "Later": {"A1": datetime(2024, 1, 1)},  # a Monday, loaded after Main
        })
        self.assertEqual(value(manager, "Main", "A1"), "Mon")
        self.assertEqual(value(manager, "Main", "A2"), "02/01/2024")
        manager.get_sheet("Later").get_cell("A1").set_operation(Value(datetime(2024, 1, 3)))
        self.assertEqual(value(manager, "Main", "A1"), "Wed")
        self.assertEqual(value(manager, "Main", "A2"), "04/01/2024")

    def test_load_evaluates_each_cell_once(self):
        # a chain written bottom-up: every loaded cell would re-propagate the whole chain
        size = 200
        cells = {f"A{row}": f"=A{row + 1}+1" for row in range(1, size)}
        cells[f"A{size}"] = 1
        original_update, original_recalculate = SheetCell.update, SheetCell.recalculate
        updates, recalculations = [], collections.Counter()
        SheetCell.update = lambda cell: updates.append(cell) or original_update(cell)
        SheetCell.recalculate = lambda cell: recalculations.update([id(cell)]) or original_recalculate(cell)
        try:
            manager = load({"Chain": cells})
        finally:
            SheetCell.update, SheetCell.recalculate = original_update, original_recalculate
        self.assertEqual(value(manager, "Chain", "A1"), size)
        self.assertEqual(updates, [])
        # once when the cell is built, once in the dependency-ordered pass
        self.assertLessEqual(max(recalculations.values()), 2)
        # after loading, changes still propagate
        manager.get_sheet("Chain").get_cell(f"A{size}").set_operation(Value(10))
        self.assertEqual(value(manager, "Chain", "A1"), size + 9)

    def test_circular_reference_loads_with_a_warning(self):
        with self.assertLogs(level="WARNING") as logs:
            manager = load({"Loop": {"A1": "=B1+1", "B1": "=A1+1", "C1": 5, "D1": "=C1*2"}})
        self.assertIn("circular reference", logs.output[0])
        self.assertEqual(value(manager, "Loop", "D1"), 10)

    def test_error_clears_when_the_cause_is_fixed(self):
        manager = load({"Main": {"A1": "=1/B1", "A2": "=A1*2", "A3": "=IFERROR(A2,-1)", "B1": 0}})
        self.assertEqual(value(manager, "Main", "A2"), "#DIV/0!")
        self.assertEqual(value(manager, "Main", "A3"), -1)
        manager.get_sheet("Main").get_cell("B1").set_operation(Value(4))
        self.assertEqual(value(manager, "Main", "A2"), 0.5)
        self.assertEqual(value(manager, "Main", "A3"), 0.5)

    def test_colliding_sheet_names_fail(self):
        with self.assertRaisesRegex(ValueError, "collide"):
            load({"A-B": {"A1": 1}, "A B": {"A1": 2}})


class TestPowerPrecedence(unittest.TestCase):
    """Excel binds unary minus tighter than ^, and ^ is left-associative."""

    def test_values(self):
        cases = [("=-2^2", 4), ("=2^3^2", 64), ("=-2^3", -8), ("=2^-1", 0.5), ("=0-2^2", -4),
                 ("=1+2^3^2", 65), ("=2^3^2^0.5", 8.0), ("=2^(3^2)", 512), ("=-(2^2)", -4)]
        cells = {f"A{row}": formula for row, (formula, _) in enumerate(cases, start=1)}
        manager = load({"S": cells})
        for row, (formula, expected) in enumerate(cases, start=1):
            with self.subTest(formula=formula):
                self.assertEqual(value(manager, "S", f"A{row}"), expected)


if __name__ == "__main__":
    unittest.main()
