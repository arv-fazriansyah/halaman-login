Sub DataUpdate()
    Dim wsData As Worksheet
    Dim UrlData As String
    Dim conn As WorkbookConnection

    ' Set error handler
    On Error GoTo ErrorHandler

    ' Set worksheet and URL
    Set wsData = ThisWorkbook.Sheets("DATAUSER")
    UrlData = "https://login.arvib.workers.dev/boskin?npsn=12345"

    ' Clear existing data in the worksheet
    wsData.Cells.Clear

    ' Set up QueryTable to fetch data from the URL
    With wsData.QueryTables.Add(Connection:="URL;" & UrlData, Destination:=wsData.Range("A1"))
        .BackgroundQuery = True
        .AdjustColumnWidth = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    ' Delete all data connections in the workbook
    For Each conn In ThisWorkbook.Connections
        conn.Delete
    Next conn

    Exit Sub

ErrorHandler:
    ' Display error message if an error occurs
    MsgBox "Error occurred: " & Err.Description, vbCritical, "Error"

    ' Log error to Immediate Window (Ctrl+G to view)
    Debug.Print "Error in DataUpdate: " & Err.Description
End Sub
