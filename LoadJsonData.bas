Sub LoadJsonData()
    Dim wb As Workbook
    Dim LoadDataSheet As Worksheet
    Dim QueryName As String
    Dim url As String
    Dim conn As WorkbookConnection

    ' Set error handler
    On Error GoTo ErrorHandler

    ' Set workbook and sheet
    Set wb = ThisWorkbook
    Set LoadDataSheet = wb.Sheets("DATAUSER")

    ' Clear existing data
    LoadDataSheet.Cells.Clear

    ' Configurable settings
    QueryName = "MyJsonQuery"
    url = "https://login.arvib.workers.dev/boskin?npsn=12345"

    ' Add new query
    wb.Queries.Add Name:=QueryName, Formula:= _
        "let" & vbCrLf & _
        "    Source = Json.Document(Web.Contents(""" & url & """))," & vbCrLf & _
        "    RecordAsTable = Record.ToTable(Source{0})," & vbCrLf & _
        "    PromotedHeaders = Table.PromoteHeaders(Table.Transpose(RecordAsTable), [PromoteAllScalars=true])" & vbCrLf & _
        "in" & vbCrLf & _
        "    PromotedHeaders"

    ' Load data to "Data" sheet
    LoadQuery QueryName, LoadDataSheet

    ' Delete query after loading
    wb.Queries(QueryName).Delete

    ' Delete all connections
    For Each conn In wb.Connections
        conn.Delete
    Next conn

    ' End procedure successfully
    Exit Sub

ErrorHandler:
    ' Display error message if an error occurs
    MsgBox "Error occurred: " & Err.Description, vbCritical, "Error"
    
    ' Log error to Immediate Window (Ctrl+G to view)
    Debug.Print "Error in LoadJsonData: " & Err.Description
    
    ' Clean up connections if error occurred
    On Error Resume Next
    For Each conn In wb.Connections
        conn.Delete
    Next conn
End Sub

Private Sub LoadQuery(ByVal QueryName As String, ByVal LoadDataSheet As Worksheet)
    On Error GoTo QueryErrorHandler

    ' Add query to sheet and load data
    With LoadDataSheet.ListObjects.Add(SourceType:=0, source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName & ";Extended Properties=""""" _
        , Destination:=LoadDataSheet.Range("$A$1")).queryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & QueryName & "]")
        .BackgroundQuery = True
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    Exit Sub

QueryErrorHandler:
    ' Display error message if query loading fails
    MsgBox "Error occurred while loading data: " & Err.Description, vbCritical, "Error"
    
    ' Log error to Immediate Window (Ctrl+G to view)
    Debug.Print "Error in LoadQuery: " & Err.Description
End Sub
