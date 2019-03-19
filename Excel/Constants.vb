Imports SpreadsheetLight
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace Excel
    Public Enum Constants
        xlSolid = PatternValues.Solid
        xlCenter = SLHorizontalTextAlignmentValues.Center
        xlBottom = VerticalAlignmentValues.Bottom
        xlContext = SLAlignmentReadingOrderValues.ContextDependent
        xlNone = PatternValues.None
    End Enum

    Public Enum XlBordersIndex
        xlEdgeTop
        xlEdgeBottom
        xlEdgeLeft
        xlEdgeRight
        xlInsideHorizontal
        xlInsideVertical
    End Enum

    Public Enum XlBorderWeight
        xlThin = BorderStyleValues.Thin
    End Enum

    Public Enum XlPasteType
        xlPasteValues = SLPasteTypeValues.Values
    End Enum
End Namespace
