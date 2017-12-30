Option Explicit
Option Base 1

Public Function OpenDB(FileName As String) As Object

Dim Res As Object

Set Res = CreateObject("Excel.Application")

With Res
    .Workbooks.Open FileName
    .Visible = False
End With

Set OpenDB = Res

End Function

Public Function GetDBSource(DB As Object, DBName As String) As Worksheet

Set GetDBSource = DB.ActiveWorkbook.Worksheets(DBName)

End Function

Public Sub CloseDB(DBSource As Object, DBName As String)

DBSource.Application.DisplayAlerts = False
DBSource.ActiveWorkbook.Save
DBSource.Application.DisplayAlerts = True

DBSource.Quit

End Sub

Public Sub CloseDBNoSave(DBSource As Object, DBName As String)

DBSource.Quit

End Sub
