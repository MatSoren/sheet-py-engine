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
.xlsx ──openpyxl──> formula strings ──rewrite──> Python-parseable expressions
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

1. **`translator/translator.py`** walks the workbook. Excel syntax that Python's parser
   would reject gets rewritten first — `Sheet!A1` becomes a marker token, `A1:B4` another,
   `&` becomes `+`, and so on.
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

**Operators** — `+ - * /`, comparisons (`= < <= > >=`), `&` (concatenation), `%`

**References** — same sheet (`A1`), cross-sheet (`Sheet!A1`), ranges (`A1:B4`)

**Functions** — `SUM`, `IF`, `AND`, `COUNTIF`, `IFERROR`, `TEXT`, `TODAY`, `MROUND`,
`ROUNDUP`, `ROUNDDOWN`

**Sheet names with spaces** — formulas become Python identifiers before parsing, so spaces
are stripped from sheet names. A sheet called `Cost Report` is registered as `CostReport`
and looked up that way; `'Cost Report'!B1` inside a formula resolves fine either way.

## Testing

`test/test.py` is a differential check: Excel caches the last value it computed for every
formula cell, so loading a file both ways gives a reference to compare against.

```sh
python -m test.test path/to/book.xlsx
```

It prints every cell where the engine disagrees with Excel.

> **Note:** this only means something for files saved by Excel or LibreOffice. Workbooks
> written by openpyxl — including `examples/example.xlsx` — carry no cached values, so the
> check silently passes over every formula. Open the example in a spreadsheet program and
> save it to get a real comparison.

## Known limitations

- **`COUNTIF` with a numeric criterion is broken.** `"=COUNTIF(A1:A9,75)"` raises a
  `TypeError`, and `">75"` raises another one comparing `int` to `str`: the criterion is
  never coerced to a number. Only string criteria work today.
- **One workbook per process.** `SheetsManager._sheet_dict` is class-level state and
  `translator` holds a module-level manager, so loading a second workbook writes into the
  same registry as the first.
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
