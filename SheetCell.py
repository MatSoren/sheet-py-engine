import logging
from abc import ABC, abstractmethod
from operations.operation_factory import OperationFactory, OperationTypes
from operations.operations import Operation

# Todo Operaciones recursivas
# Todo Attach de todas las celdas en la traduccion


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
        self.observers = []
        self.operation = OperationFactory.get_instance(operation, *args)
        self.calculated_value = self.operation.operate()
        self.coordinate = ""
        self.sheet = ""

    def init_operation(self, operation_type, *args, **kwargs):
        self.operation = OperationFactory.get_instance(operation_type, *args)
        self.update()

    def set_operation(self, operation):
        self.operation = operation
        self.update()

    def set_coordinate(self, coordinate):
        self.coordinate = coordinate
    def set_sheet(self, sheet):
        self.sheet = sheet        
    def attach(self, observer):
        if self.observers.count(observer) == 0:
            self.observers.append(observer)
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
        operation_result = self.operation.operate()
        if isinstance(operation_result,float):
            operation_result = round(operation_result,10)
        self.calculated_value = operation_result 
        self.notify()

    def operate(self):
        return self.calculated_value

    def __str__(self):
        return f"coord: {self.coordinate}, sheet: {self.sheet}, {self.calculated_value}, operation_type: {self.operation}"
