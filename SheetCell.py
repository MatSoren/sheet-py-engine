import graphlib
import logging
from abc import ABC, abstractmethod
from operations.operation_factory import OperationFactory, OperationTypes
from operations.operations import FORMULA_FAILURES, ExcelError, Operation

# Todo Operaciones recursivas
# Todo Attach de todas las celdas en la traduccion


def evaluate_in_dependency_order(cells):
    """Recomputes every cell once, after the cells it depends on.

    Used when loading (translate() only assigns operations) and on every update(): letting
    each cell notify its observers recursively would recompute a cell once per path that
    reaches it, which grows exponentially with chained diamonds.

    `cells` must hold every observer of each of its cells, as both callers guarantee.
    Kahn's algorithm by hand: graphlib does the same with more bookkeeping per edge, and a
    large range is one edge per cell in it.
    """
    pending_inputs = dict.fromkeys(cells, 0)
    for cell in cells:
        for observer in cell.observers:
            pending_inputs[observer] += 1
    ready = [cell for cell, count in pending_inputs.items() if count == 0]
    order = []
    while ready:
        cell = ready.pop()
        order.append(cell)
        for observer in cell.observers:
            pending_inputs[observer] -= 1
            if pending_inputs[observer] == 0:
                ready.append(observer)
    if len(order) < len(pending_inputs):
        logging.warning("circular reference, values may be wrong: %s", find_cycle(cells))
        order = cells  # ponytail: one pass in load order; Excel would iterate or show 0
    for cell in order:
        cell.recalculate()


def find_cycle(cells):
    """'Main!A1 -> Main!A2 -> Main!A1', for the warning; only runs when there is one."""
    sorter = graphlib.TopologicalSorter()
    for cell in cells:
        sorter.add(cell)
        for observer in cell.observers:
            sorter.add(observer, cell)
    try:
        sorter.prepare()
    except graphlib.CycleError as error:
        return " -> ".join(f"{cell.sheet}!{cell.coordinate}" for cell in error.args[1])


class Observer(ABC):

    @abstractmethod
    def update(self):
        pass


class Subject(ABC):

    @abstractmethod
    def notify(self):
        pass

    @abstractmethod
    def attach(self, observer):
        pass

    @abstractmethod
    def detach(self, observer):
        pass


class SheetCell(Observer, Subject, Operation):
    def __init__(self, operation=OperationTypes.NoneOperation, *args, **kwargs):
        self.observers = {}  # dict as an ordered set: attach() checks membership in O(1)
        self.operation = OperationFactory.get_instance(operation, *args)
        self.recalculate()
        self.coordinate = ""
        self.sheet = ""

    def set_operation(self, operation):
        self.operation = operation
        self.update()

    def set_coordinate(self, coordinate):
        self.coordinate = coordinate
    def set_sheet(self, sheet):
        self.sheet = sheet        
    def attach(self, observer):
        if observer not in self.observers:
            self.observers[observer] = None
        else:
            logging.debug("observer already attached to %s", self)

    # TODO hacer el detach por las dudas
    def detach(self, observer):
        pass

    def notify(self):
        logging.debug("notifying %d observers of %s", len(self.observers), self)
        for observer in self.observers:
            observer.update()

    def get_value(self):
        pass

    def update(self):
        """Recomputes this cell and everything that depends on it, each one once, in order."""
        affected, pending = {self: None}, [self]  # dict as an ordered set
        while pending:
            for observer in pending.pop().observers:
                if observer not in affected:
                    affected[observer] = None
                    pending.append(observer)
        logging.debug("%s!%s changed, recomputing %d dependent cells", self.sheet, self.coordinate, len(affected) - 1)
        evaluate_in_dependency_order(list(affected))

    def recalculate(self):
        """Recomputes this cell's value without notifying its observers."""
        try:
            operation_result = self.operation.operate()
        except FORMULA_FAILURES as error:
            operation_result = ExcelError.from_exception(error)
        if isinstance(operation_result,float):
            operation_result = round(operation_result,10)
        self.calculated_value = operation_result 

    def operate(self):
        return self.calculated_value

    def __str__(self):
        return f"coord: {self.coordinate}, sheet: {self.sheet}, {self.calculated_value}, operation_type: {self.operation}"
