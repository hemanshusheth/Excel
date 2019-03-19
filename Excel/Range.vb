Imports SpreadsheetLight

Namespace Excel
    Public Class Range
        Private _wb As Workbook
        Private _cellvalue As String
        Private _value As String
        Private _columnFormula As String
        Private _numberFormat As String
        Private _interior As Interior
        Private _font As Font
        Private _startCellValue As String
        Private _endCellValue As String
        Private _horizontalAlignment As Constants
        Private _style As SLStyle
        Private _verticalAlignment As Constants
        Private _isWraptext As Boolean
        Private _columns As Columns
        Private _hasCopy As Boolean

        Public Sub New()

        End Sub
        Public Sub New(ByVal row As Integer, ByVal column As Integer, ByVal wb As Workbook)
            Me.New(SLConvert.ToCellReference(row, column), wb)
        End Sub

        Public Sub New(ByVal startRow As Integer, ByVal startColumn As Integer, ByVal endRow As Integer, ByVal endColumn As Integer, ByVal wb As Workbook)
            Me.New(SLConvert.ToCellRange(startRow, startColumn, endRow, endColumn), wb)
        End Sub

        Public Sub New(ByVal cellValue As String, ByVal wb As Workbook)
            _cellvalue = cellValue
            _wb = wb
            _interior = New Interior(wb, cellValue)
            _font = New Font(wb, cellValue)
            _columns = New Columns(_wb, _cellvalue)
            _style = wb.SLDoc.CreateStyle()
            _hasCopy = False
            DecomposeCellValue(cellValue)
        End Sub

        Public Property Value2 As String
            Get
                Return _value
            End Get
            Set(value As String)
                Dim intValue As Integer
                If Integer.TryParse(value, intValue) Then
                    Wb.SLDoc.SetCellValue(_cellvalue, intValue)
                Else
                    Wb.SLDoc.SetCellValue(_cellvalue, value)
                End If
                _value = value
            End Set
        End Property

        Public Property NumberFormat As String
            Get
                Return _numberFormat
            End Get
            Set(value As String)
                _style.FormatCode = value
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _numberFormat = value
            End Set
        End Property

        Public Property Formula As String
            Get
                Return _columnFormula
            End Get
            Set(value As String)
                Wb.SLDoc.SetCellValue(_cellvalue, value)
                _columnFormula = value
            End Set
        End Property

        Public Property HorizontalAlignment As Constants
            Get
                Return _horizontalAlignment
            End Get
            Set(value As Constants)
                _style.SetHorizontalAlignment(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _horizontalAlignment = value
            End Set
        End Property

        Public Property VerticalAlignment As Constants
            Get
                Return _verticalAlignment
            End Get
            Set(value As Constants)
                _style.SetVerticalAlignment(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _verticalAlignment = value
            End Set
        End Property

        Public Property WrapText As Boolean
            Get
                Return _isWraptext
            End Get
            Set(value As Boolean)
                _style.SetWrapText(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _isWraptext = value
            End Set
        End Property
        Public Property Orientation As Integer
            Get
                Return _isWraptext
            End Get
            Set(value As Integer)

            End Set
        End Property
        Public Property ShrinkToFit As Boolean
            Get
                Return _isWraptext
            End Get
            Set(value As Boolean)
                Orientation = value
                _style.SetWrapText(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _isWraptext = value
            End Set
        End Property
        Public Property IndentLevel As Integer
            Get
                Return _isWraptext
            End Get
            Set(value As Integer)
                _style.SetWrapText(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _isWraptext = value
            End Set
        End Property
        Public Property AddIndent As Boolean
            Get
                Return _isWraptext
            End Get
            Set(value As Boolean)
                _style.SetWrapText(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _isWraptext = value
            End Set
        End Property
        Public Property MergeCells As Boolean
            Get
                Return _isWraptext
            End Get
            Set(value As Boolean)
                _style.SetWrapText(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _isWraptext = value
            End Set
        End Property
        Public Property ReadingOrder As Constants
            Get
                Return _isWraptext
            End Get
            Set(value As Constants)
                _style.SetWrapText(value)
                Wb.SLDoc.SetCellStyle(_cellvalue, _style)
                _isWraptext = value
            End Set
        End Property

        Public ReadOnly Property Interior As Interior
            Get
                Return _interior
            End Get
        End Property

        Public ReadOnly Property Wb As Workbook
            Get
                Return _wb
            End Get
        End Property

        Public ReadOnly Property Font As Font
            Get
                Return _font
            End Get
        End Property

        Public ReadOnly Property EntireColumn As Columns
            Get
                Return _columns
            End Get
        End Property

        Private Sub DecomposeCellValue(cellValue As String)
            cellValue = cellValue.ToUpper()
            Dim ranges As String()
            ranges = cellValue.Split(":")
            If ranges.Length = 2 Then
                _startCellValue = ranges(0)
                _endCellValue = ranges(1)
            Else
                _startCellValue = _endCellValue = ranges(0)
            End If
        End Sub

        'Private Sub ConvertToIndexes(range As String)
        '    Dim charArray = range.ToCharArray()
        '    Dim columnBuilder As New StringBuilder
        '    Dim rowBuilder As New StringBuilder
        '    Dim columnIndex As Integer
        '    For Each c As Char In charArray
        '        If (Char.IsLetter(c)) Then
        '            columnBuilder.Append(c)
        '        Else
        '            rowBuilder.Append(c)
        '        End If
        '    Next
        '    Dim rowIndex As Integer
        '    If Integer.TryParse(rowBuilder.ToString(), rowIndex) Then
        '        columnIndex = ConvertStringToIntColumn(columnBuilder.ToString())
        '    Else
        '        Throw New Exception("Invalid row value in the cell " + range)
        '    End If


        'End Sub

        Public Function Borders(xlBordersIndex As XlBordersIndex) As Border
            Return New Border(_wb, _cellvalue, xlBordersIndex)
        End Function

        Public Sub Merge()
            Wb.SLDoc.MergeWorksheetCells(_startCellValue, _endCellValue)
        End Sub

        Public Sub [Select]()
            Wb.SLDoc.SetActiveCell(_cellvalue)
            Wb.CurrentWorkSheet.SelectedRange = Me
        End Sub

        Sub Copy()
            Dim rangeCopy As Range = MemberwiseClone()
            Wb.CurrentWorkSheet.CopiedRange = rangeCopy
            _hasCopy = True
        End Sub

        Sub PasteSpecial(Paste As XlPasteType, Operation As Constants, SkipBlanks As Boolean, Transpose As Boolean)
            If _hasCopy = True Then
                Dim selectedRange As Range = Wb.CurrentWorkSheet.SelectedRange
                Dim copiedRange As Range = Wb.CurrentWorkSheet.CopiedRange
                Wb.SLDoc.CopyCell(selectedRange._cellvalue, copiedRange._cellvalue, PasteOption:=Paste)
                If Transpose = True Then
                    Wb.SLDoc.CopyCell(selectedRange._cellvalue, copiedRange._cellvalue, SLPasteTypeValues.Transpose)
                End If
                If SkipBlanks = True Then
                    'to do
                End If
            Else
                Throw New Exception("First copy something to be pasted")
            End If

        End Sub

    End Class
End Namespace

