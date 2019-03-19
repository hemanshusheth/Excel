Imports DocumentFormat.OpenXml.Spreadsheet
Imports SpreadsheetLight

Namespace Excel
    Public Class Interior
        Private _cellValue As String
        Private _wb As Workbook
        Private _colorIndex As Integer
        Private _pattern As Constants

        Public Sub New(ByVal wb As Workbook, ByVal cellValue As String)
            _cellValue = cellValue
            _wb = wb
        End Sub
        Public Property ColorIndex As Integer
            Get
                Return _colorIndex
            End Get
            Set(value As Integer)
                Dim style As SLStyle = _wb.SLDoc.CreateStyle()
                style.Fill.SetPatternBackgroundColor(value)
                _wb.SLDoc.SetCellStyle(_cellValue, style)
                _colorIndex = value
            End Set
        End Property

        Public Property Pattern As Constants
            Get
                Return _pattern
            End Get
            Set(value As Constants)
                Dim style As SLStyle = SetFillStyle(value)
                _wb.SLDoc.SetCellStyle(_cellValue, style)
                _pattern = value
            End Set
        End Property
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function SetFillStyle(value As Constants)
            Dim style As SLStyle = _wb.SLDoc.CreateStyle()
            If value = Constants.xlSolid Then
                style.Fill.SetPatternType(PatternValues.Solid)
            End If
            Return style
        End Function
    End Class
End Namespace

