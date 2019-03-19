
Namespace Excel
    Public Class WorkSheet
        Private _name As String
        Private _wb As Workbook
        Private _selectedRange As Range
        Private _copiedRange As Range

        ''' <summary>
        ''' Creates a new worksheet. 
        ''' </summary>
        ''' <param name="name">Name of the worksheet</param>
        ''' <param name="wb">Name of the workbook</param>
        ''' <remarks>To use this constructor make you have created a workbook</remarks>
        Public Sub New(ByVal name As String, ByVal wb As Workbook)
            _name = name
            _wb = wb
            wb.SLWorkSheets.Add(Me)
            'Sheet1 is added by default
            If name <> "Sheet1" Then
                wb.SLDoc.AddWorksheet(name)
            End If
        End Sub
        ''' <summary>
        ''' Name of the worksheet
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        Public Property SelectedRange As Range
            Get
                Return _selectedRange
            End Get
            Set(value As Range)
                _selectedRange = value
            End Set
        End Property

        Public Property CopiedRange As Range
            Get
                Return _copiedRange
            End Get
            Set(value As Range)
                _copiedRange = value
            End Set
        End Property
        ''' <summary>
        ''' Selects the range of cell provided by row and column index
        ''' </summary>
        ''' <param name="row">row index of the cell</param>
        ''' <param name="column">column index of the cell</param>
        ''' <returns>The Range of row and column selected</returns>
        ''' <remarks></remarks>
        Public Function Cells(ByVal row As Integer, ByVal column As Integer) As Range
            Dim range As Range = New Range(row, column, _wb)
            _selectedRange = range
            Return range
        End Function
        ''' <summary>
        ''' Deletes a worksheet if it exists
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Delete()
            If Not _wb Is Nothing Then
                _wb.DeleteWorkSheet(_name)
            End If
        End Sub
        ''' <summary>
        ''' Defines a new Range 
        ''' </summary>
        ''' <param name="cellValue">Range of the cell presented as string</param>
        ''' <returns>a new Range</returns>
        ''' <remarks></remarks>
        Public Function Range(ByVal cellValue As String) As Range
            Dim newrange As Range = New Range(cellValue, _wb)
            _selectedRange = newrange
            Return newrange
        End Function

    End Class
End Namespace
