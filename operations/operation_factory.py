from enum import Enum
from operations.operations import (
    AndCondition,
    BinaryAdd,
    BinarySub,
    BinaryMul,
    BinaryDiv,
    BlankCondition,
    CountIf,
    DiferentCondition,
    EqualCondition,
    If,
    IfError,
    GratherCondition,
    GratherOrEqualCondition,
    Hlookup,
    LessCondition,
    LessOrEqualCondition,
    MRound,
    NoneOperation,
    NotCondition,
    OrCondition,
    Range,
    Reference,
    RoundDown,
    RoundUp,
    Sum,
    Text,
    Today,
    Value,
)


class OperationTypes(Enum):
    AndCondition = ("AndCondition",)
    BinaryAdd = ("BinaryAdd",)
    BinarySub = ("BinarySub",)
    BinaryMul = ("BinaryMul",)
    BinaryDiv = ("BinaryDiv",)
    BlankCondition = ("BlankCondition",)
    CountIf = ("CountIf",)
    DiferentCondition = ("DiferentCondition",)
    EqualCondition = ("EqualCondition",)
    If = ("If",)
    IfError = ("IfError",)
    GratherCondition = ("GratherCondition",)
    GratherOrEqualCondition = ("GratherOrEqualCondition",)
    Hlookup = ("Hlookup",)
    LessCondition = ("LessCondition",)
    LessOrEqualCondition = ("LessOrEqualCondition",)
    MRound = ("MRound",)
    NoneOperation = ("NoneOperation",)
    NotCondition = ("NotCondition",)
    OrCondition = ("OrCondition",)
    Range = ("Range",)
    Reference = ("Reference",)
    RoundDown = ("RoundDown",)
    RoundUp = ("RoundUp",)
    Sum = ("Sum",)
    Text = ("Text",)
    Today = ("Today",)
    Value = ("Value",)


class OperationFactory:
    _instanciator = {
        OperationTypes.AndCondition: AndCondition,
        OperationTypes.BinaryAdd: BinaryAdd,
        OperationTypes.BinarySub: BinarySub,
        OperationTypes.BinaryMul: BinaryMul,
        OperationTypes.BinaryDiv: BinaryDiv,
        OperationTypes.BlankCondition: BlankCondition,
        OperationTypes.CountIf: CountIf,
        OperationTypes.DiferentCondition: DiferentCondition,
        OperationTypes.EqualCondition: EqualCondition,
        OperationTypes.If: If,
        OperationTypes.IfError: IfError,
        OperationTypes.GratherCondition: GratherCondition,
        OperationTypes.GratherOrEqualCondition: GratherOrEqualCondition,
        OperationTypes.Hlookup: Hlookup,
        OperationTypes.LessCondition: LessCondition,
        OperationTypes.LessOrEqualCondition: LessOrEqualCondition,
        OperationTypes.MRound: MRound,
        OperationTypes.NoneOperation: NoneOperation,
        OperationTypes.NotCondition: NotCondition,
        OperationTypes.OrCondition: OrCondition,
        OperationTypes.Range: Range,
        OperationTypes.Reference: Reference,
        OperationTypes.RoundDown: RoundDown,
        OperationTypes.RoundUp: RoundUp,
        OperationTypes.Sum: Sum,
        OperationTypes.Text: Text,
        OperationTypes.Today: Today,
        OperationTypes.Value: Value,
    }

    @classmethod
    def get_instance(self, operation_type, *args, **kwargs):
        operation = self._instanciator.get(operation_type, None)
        if operation is None:
            return None

        return operation(*args)
