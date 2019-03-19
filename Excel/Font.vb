Imports SpreadsheetLight

Namespace Excel
    Public Class Font
        Private _cellValue As String
        Private _wb As Workbook
        Private _colorIndex As Integer
        Private _fontStyle As String

        Public Sub New(ByVal wb As Workbook, ByVal cellValue As String)
            _cellValue = cellValue
            _wb = wb
            _fontStyle = ""
            _colorIndex = 0
        End Sub

        Public Property ColorIndex As Integer
            Get
                Return _colorIndex
            End Get
            Set(value As Integer)
                Dim style As SLStyle = _wb.SLDoc.CreateStyle()
                style.Font.SetFontThemeColor(value)
                If Not _cellValue Is Nothing Then
                    _wb.SLDoc.SetCellStyle(_cellValue, style)
                End If
                _colorIndex = value
            End Set
        End Property

        Public Property FontStyle As String
            Get
                Return _fontStyle
            End Get
            Set(value As String)
                Dim style As SLStyle = _wb.SLDoc.CreateStyle()
                If value Is "Bold" Then
                    style.SetFontBold(True)
                    If Not _cellValue Is Nothing Then
                        _wb.SLDoc.SetCellStyle(_cellValue, style)
                    End If
                    _fontStyle = value
                End If
            End Set
        End Property
    End Class
End Namespace