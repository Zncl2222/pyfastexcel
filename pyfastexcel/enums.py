from enum import Enum


class BaseEnum(Enum):
    @classmethod
    def get_enum(cls, name: str):
        # Use casefold for case-insensitive comparison
        name_casefolded = name.casefold()
        for enum in cls:
            if enum.name.casefold() == name_casefolded:
                return enum
        raise ValueError(f'{name} is not a valid {cls.__name__}')


class ChartType(BaseEnum):
    Area = 0
    AreaStacked = 1
    AreaPercentStacked = 2
    Area3D = 3
    Area3DStacked = 4
    Area3DPercentStacked = 5
    Bar = 6
    BarStacked = 7
    BarPercentStacked = 8
    Bar3DClustered = 9
    Bar3DStacked = 10
    Bar3DPercentStacked = 11
    Bar3DConeClustered = 12
    Bar3DConeStacked = 13
    Bar3DConePercentStacked = 14
    Bar3DPyramidClustered = 15
    Bar3DPyramidStacked = 16
    Bar3DPyramidPercentStacked = 17
    Bar3DCylinderClustered = 18
    Bar3DCylinderStacked = 19
    Bar3DCylinderPercentStacked = 20
    Col = 21
    ColStacked = 22
    ColPercentStacked = 23
    Col3D = 24
    Col3DClustered = 25
    Col3DStacked = 26
    Col3DPercentStacked = 27
    Col3DCone = 28
    Col3DConeClustered = 29
    Col3DConeStacked = 30
    Col3DConePercentStacked = 31
    Col3DPyramid = 32
    Col3DPyramidClustered = 33
    Col3DPyramidStacked = 34
    Col3DPyramidPercentStacked = 35
    Col3DCylinder = 36
    Col3DCylinderClustered = 37
    Col3DCylinderStacked = 38
    Col3DCylinderPercentStacked = 39
    Doughnut = 40
    Line = 41
    Line3D = 42
    Pie = 43
    Pie3D = 44
    PieOfPie = 45
    BarOfPie = 46
    Radar = 47
    Scatter = 48
    Surface3D = 49
    WireframeSurface3D = 50
    Contour = 51
    WireframeContour = 52
    Bubble = 53
    Bubble3D = 54


class ChartDataLabelPosition(BaseEnum):
    Unset = 0
    BestFit = 1
    Below = 2
    Center = 3
    InsideBase = 4
    InsideEnd = 5
    Left = 6
    OutsideEnd = 7
    Right = 8
    Above = 9


class ChartLineType(BaseEnum):
    Unset = 0
    Solid = 1
    NONE = 2
    Automatic = 3


class MarkerSymbol(BaseEnum):
    Circle = 'circle'
    Dash = 'dash'
    Diamond = 'diamond'
    Dot = 'dot'
    NONE = 'none'
    Picture = 'picture'
    Plus = 'plus'
    Square = 'square'
    Star = 'star'
    Triangle = 'triangle'
    X = 'x'
    Auto = 'auto'


class PivotSubTotal(BaseEnum):
    Average = 'Average'
    Count = 'Count'
    CountNums = 'CountNums'
    Max = 'Max'
    Min = 'Min'
    Product = 'Product'
    StdDev = 'StdDev'
    StdDevp = 'StdDevp'
    Sum = 'Sum'
    Var = 'Var'
    Varp = 'Varp'
