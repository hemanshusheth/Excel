
Namespace Excel
    Public Class Application
        Private wb As Workbook
        Private _activeWorksheet As WorkSheet
        Private _selectedRange As Range

        ''' <summary>
        ''' Default Constructor
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()

        End Sub

        ''' <summary>
        '''  Gets sets active workbook for the application
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ActiveWorkBook() As Workbook
            Get
                Return wb
            End Get
            Set(value As Workbook)
                wb = value
            End Set
        End Property
        ''' <summary>
        ''' Gets list of all worksheets
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property WorkSheets() As List(Of WorkSheet)
            Get
                Return wb.SLWorkSheets
            End Get
        End Property

        Public Function Sheets(i As Integer) As WorkSheet
            _activeWorksheet = WorkSheets(i)
            Return _activeWorksheet
        End Function

        ''' <summary>
        ''' Gets worksheet from selected by its name
        ''' </summary>
        ''' <param name="sheetName"></param>
        ''' <returns></returns>
        ''' <remarks>Selected worksheet will be set as activeworksheet</remarks>
        Public Function Sheets(ByVal sheetName As String) As WorkSheet
            For Each workSheet As WorkSheet In WorkSheets
                If workSheet.Name.Equals(sheetName) Then
                    _activeWorksheet = workSheet
                    wb.SLDoc.SelectWorksheet(sheetName)
                    Return workSheet
                End If
            Next
            Throw New Exception("Worksheet" + sheetName + " not found")
        End Function

        ''' <summary>
        ''' provides the selected range
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Selection As Range
            Get
                _selectedRange = _activeWorksheet.SelectedRange
                Return _selectedRange
            End Get
            Set(value As Range)
                _selectedRange = value
                _activeWorksheet.SelectedRange = value
            End Set
        End Property

        Public Sub Quit()
            wb = Nothing
            _activeWorksheet = Nothing
            _selectedRange = Nothing
        End Sub
    End Class
End Namespace
