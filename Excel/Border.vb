Imports SpreadsheetLight

Namespace Excel
    Public Class Border
        Private _wb As Workbook
        Private _bordersIndex As XlBordersIndex
        Private _weight As XlBorderWeight
        Dim _cellValue As String

        Public Sub New(ByVal wb As Workbook, ByVal cellValue As String, ByVal bordersIndex As XlBordersIndex)
            _cellValue = cellValue
            _wb = wb
            _bordersIndex = bordersIndex
        End Sub

        Public Property Weight As XlBorderWeight
            Get
                Return _weight
            End Get
            Set(value As XlBorderWeight)
                _wb.SLDoc.SetCellStyle(_cellValue, GenerateStyle(value))
                _weight = value
            End Set
        End Property

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GenerateStyle(value As XlBorderWeight) As SLStyle
            Dim style As SLStyle = _wb.SLDoc.CreateStyle()
            If (_bordersIndex.Equals(XlBordersIndex.xlEdgeTop)) Then
                style.Border.TopBorder.BorderStyle = value
            ElseIf (_bordersIndex.Equals(XlBordersIndex.xlEdgeBottom)) Then
                style.Border.BottomBorder.BorderStyle = value
            ElseIf (_bordersIndex.Equals(XlBordersIndex.xlEdgeRight)) Then
                style.Border.RightBorder.BorderStyle = value
            ElseIf (_bordersIndex.Equals(XlBordersIndex.xlEdgeLeft)) Then
                style.Border.LeftBorder.BorderStyle = value
            ElseIf (_bordersIndex.Equals(XlBordersIndex.xlInsideHorizontal)) Then
                style.Border.HorizontalBorder.BorderStyle = value
            ElseIf (_bordersIndex.Equals(XlBordersIndex.xlInsideVertical)) Then
                style.Border.VerticalBorder.BorderStyle = value
            End If
            Return style
        End Function

    End Class
End Namespace