Namespace Excel
    Public Class Columns
        Private _wb As Workbook
        Private _cellValue As String
        Public Sub New(ByVal wb As Workbook, ByVal cellValue As String)
            _wb = wb
            _cellValue = cellValue
        End Sub

        Public Sub AutoFit()
            _wb.SLDoc.AutoFitColumn(_cellValue)
        End Sub
    End Class
End Namespace
