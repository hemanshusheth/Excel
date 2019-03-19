Imports SpreadsheetLight

Namespace Excel
    Public Class Workbook
        Private _slDoc As SLDocument
        Private _slWorkSheets As List(Of WorkSheet)
        Private _currentWorkSheet As WorkSheet

        ''' <summary>
        ''' Constructor for a WorkBook
        ''' </summary>
        ''' <remarks>A workbook contains List of WorkSheets.
        ''' A default Worksheet Sheet1 is created </remarks>
        ''' 
        Public Sub New()
            _slDoc = New SLDocument()
            _slWorkSheets = New List(Of WorkSheet)()
            ' The first worksheet of a new spreadsheet is named "Sheet1",
            Dim workSheet As WorkSheet = New WorkSheet("Sheet1", Me)
            _currentWorkSheet = workSheet
        End Sub

        Public Property Saved As Boolean

        ReadOnly Property SLWorkSheets As List(Of WorkSheet)
            Get
                Return _slWorkSheets
            End Get
        End Property

        ReadOnly Property SLDoc As SLDocument
            Get
                Return _slDoc
            End Get
        End Property

        Public ReadOnly Property CurrentWorkSheet As WorkSheet
            Get
                Return _currentWorkSheet
            End Get
        End Property

        ''' <summary>
        ''' Deletes a workSheet if found else throws an exception
        ''' </summary>
        ''' <param name="sheetName">Name of the worksheet to be deleted</param>
        ''' <remarks>A deleted worksheet is also removed from the List Sheets</remarks>
        Public Sub DeleteWorkSheet(ByVal sheetName As String)
            Dim found As Boolean = False
            For Each workSheet As WorkSheet In SLWorkSheets
                If workSheet.Name.Equals(sheetName) Then
                    SLWorkSheets.Remove(workSheet)
                    SLDoc.DeleteWorksheet(sheetName)
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Throw New Exception("Worksheet " + sheetName + " not found")
            End If
            SetDefault()
        End Sub

        Public Overloads Function Sheets(ByVal sheetName As String) As WorkSheet
            For Each worksheet As WorkSheet In SLWorkSheets
                If (worksheet.Name.Equals(sheetName)) Then
                    _currentWorkSheet = worksheet
                    Return worksheet
                End If
            Next
            Throw New Exception("Requested sheet " + sheetName + " not found")
        End Function
        ''' <summary>
        ''' Saves the workbook at the given path
        ''' </summary>
        ''' <param name="strPath">path of the file to be saved</param>
        ''' <remarks>once workbook is saved the Property Saved is true</remarks>
        Public Sub SaveAs(ByVal strPath As String)
            If Not String.IsNullOrEmpty(strPath) Then
                SLDoc.SaveAs(strPath)
                Saved = True
            Else
                Saved = False
                Throw New Exception("Path " + strPath + " not found")
            End If
        End Sub

        ''' <summary>
        ''' Selects the first worksheet as default
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub SetDefault()
            If SLWorkSheets.Count > 0 Then
                _currentWorkSheet = SLWorkSheets(0)
            End If
        End Sub

    End Class
End Namespace
