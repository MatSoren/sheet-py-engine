# sheet-engine

A spreadsheet calculation engine in Python. It reads an `.xlsx` file, parses every
formula into an expression tree, and builds a **reactive dependency graph** of cells:
change one value and everything that depends on it recalculates on its own, across
sheets.

In other words — the part of Excel that actually computes, without Excel.

```console
$ python main.py examples/example.xlsx 'Items!D5' 'Summary!B3'
Items!D5 = 301.0
Summary!B3 = 364.21

$ python main.py examples/example.xlsx --set 'Items!B2=10' 'Items!D5' 'Summary!B3'
set Items!B2 = 10
Items!D5 = 451.0
Summary!B3 = 545.71
```

Nothing recomputed the whole workbook there. Setting `Items!B2` notified only the cells
observing it, which notified theirs, and so on until the change reached `Summary!B3` on
another sheet.

## How it works

```
.xlsx ──openpyxl──> formula tokens ──translate──> Python-parseable expressions
                                                          │
                                                    ast.parse
                                                          │
                                              OperationProcessor
                                          (an ast.NodeVisitor walk)
                                                          │
                                                 Operation tree
                                       (Value, Reference, Range, Sum, If, …)
                                                          │
                                                     SheetCell
                                           (Observer + Subject + Operation)
```

1. **`translator/translator.py`** walks the workbook. `formula_to_python()` runs each formula
   through openpyxl's formula tokenizer and translates it token by token, so strings,
   references and operators never get mixed up: `'Mi Hoja'!$A$1:B2` becomes one identifier,
   `<>` becomes `!=`, `^` becomes `**`, `&` becomes `|` (same precedence slot as in Excel),
   `"a:b"` stays a string. It is a pure string → string function, testable without a workbook.
2. **`ast.parse`** — Python's own parser, from the standard library — turns the rewritten
   formula into a syntax tree. No hand-written expression parser.
3. **`translator/AstProcessor.py`** visits that tree and emits an `Operation` tree.
4. **`operations/`** holds the operation types. Each one knows how to `operate()` on its
   operands, recursively.
5. **`SheetCell`** is simultaneously an `Operation`, an `Observer` and a `Subject`. When a
   cell is built from a formula it attaches itself to every cell it references. Writing to
   a cell calls `update()`, which recomputes and then `notify()`s its observers — the
   change propagates outward through the graph, touching only what actually depends on it.

Using `ast` for the parsing is the trick that keeps this small: operator precedence,
parentheses, nesting and function calls come free from CPython's own grammar.

## Install

```sh
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

The only dependency is `openpyxl`. Everything else is the standard library.

## Usage

```sh
python main.py FILE.xlsx [SHEET!CELL ...] [--set SHEET!CELL=VALUE] [-v]
```

| Flag | Meaning |
|---|---|
| `SHEET!CELL` | print the computed value of that cell |
| `--set SHEET!CELL=VALUE` | overwrite a cell first; recalculation propagates. Repeatable. |
| `-v` | log the propagation as it happens |
| `--recursion-limit N` | deep formula chains recurse; raise it if you hit the limit |

Run with no cells to list the sheets.

As a library:

```python
from openpyxl import load_workbook
from translator import translator
from operations.operations import Value

manager = translator.translate(load_workbook("book.xlsx", data_only=False))
sheet = manager.get_sheet("Summary")

sheet.get_cell("B3").calculated_value          # read

manager.get_sheet("Items").get_cell("B2").set_operation(Value(10))
sheet.get_cell("B3").calculated_value          # already recalculated
```

## Supported

**Operators** — `+ - * / ^`, comparisons (`= <> < <= > >=`), `&` (concatenation, Excel
text conversion: `1` → `"1"`, `TRUE` → `"TRUE"`), postfix `%`, unary `-` / `+`

**Operands** — numbers, strings (with `""` escapes), `TRUE`/`FALSE`, error literals (`#N/A`,
`#REF!`, `#DIV/0!`, … evaluate to that string), empty arguments (`IF(A1,,1)`)

**References** — same sheet (`A1`), absolute (`$A$1`), cross-sheet (`Sheet!A1`,
`'2024 Ventas'!A1`), ranges (`A1:B4`, `'Costos-Fijos'!A1:B4`)

**Functions** — `SUM`, `IF`, `AND`, `OR`, `COUNTIF`, `IFERROR`, `TEXT`, `TODAY`, `MROUND`,
`ROUNDUP`, `ROUNDDOWN`, `HLOOKUP`, in any case (`if(`, `Sum(`)

**Any sheet name** — sheets are registered under `normalize_sheet_name(name)` (every
non-alphanumeric ASCII char becomes `_`): `Cost Report` → `Cost_Report`, `Costos-Fijos` →
`Costos_Fijos`, `Año 1` → `A_o_1`. Formulas use the same function, so they always match; the
CLI accepts the real name (`'Cost Report!B1'`). Two sheets that normalize to the same name
make `translate()` fail instead of mixing them.

## Testing

`test/test.py` is a differential check: Excel caches the last value it computed for every
formula cell, so loading a file both ways gives a reference to compare against.

```sh
python -m test.test path/to/book.xlsx
```

It prints every cell where the engine disagrees with Excel.

`test/test_translator.py` covers the translation itself (every token type, exact output plus
`ast.parse`) and end-to-end evaluation on workbooks built in memory, including propagation:

```sh
python -m unittest test.test_translator -v
```

> **Note:** this only means something for files saved by Excel or LibreOffice. Workbooks
> written by openpyxl — including `examples/example.xlsx` — carry no cached values, so the
> check silently passes over every formula. Open the example in a spreadsheet program and
> save it to get a real comparison.

## Known limitations

- **`COUNTIF` has no wildcards** (`*`, `?`, `~`), and numbers stored as text don't match a
  numeric criterion.
- **One workbook per process.** `SheetsManager._sheet_dict` is class-level state and
  `translator` holds a module-level manager, so loading a second workbook writes into the
  same registry as the first.
- **Excel precedence quirks of `^`.** Python's precedence and associativity are used, so
  `=-2^2` gives `-4` (Excel: `4`, unary minus binds tighter) and `=2^3^2` gives `512`
  (Excel: `64`, left-associative). Parenthesize to be safe. `%` combined with `^` works
  (`=2^50%` is `2^0.5`), since `%` wraps the operand right before it.
- **Intersection operator, array constants and 3D references** (`=A1:A3 B1:B3`,
  `={1,2,3}`, `=SUM(Sheet1:Sheet3!A1)`) are not
  supported; they raise a `ValueError` naming the formula.
- **Errors have no value of their own**: `#N/A` evaluates to the string `"#N/A"`, while
  computed errors (`/0`, a failed `HLOOKUP`) are `None`, the same as a blank cell, so an
  error feeding arithmetic reads as 0. `IFERROR` only catches Python exceptions.
- **No cycle detection.** Circular references recurse until Python's recursion limit.
- **`detach()` is a no-op**, so observers are never removed.
- Dates, formatting, charts, pivot tables and macros are out of scope. This computes
  values.

## Why

It started as a translation layer that had to evaluate spreadsheet logic outside of Excel,
and turned into a small reactive calculation engine. It is now the subject of my
undergraduate thesis in Systems Engineering at UNICEN.

## License

MIT — see [LICENSE](LICENSE).
