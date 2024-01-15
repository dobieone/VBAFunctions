Attribute VB_Name = "Functions"
Option Explicit

Public Function Inc(ByRef Value As Variant, _
                    Optional ByVal Amount As Variant = 1) As Variant
    Value = Value + Amount
End Function

Public Function Dec(ByRef Value As Variant, _
                    Optional ByVal Amount As Variant = -1) As Variant
    If Amount > 0 Then
        Amount = 0 - Amount
    End If
    Value = Value + Amount
End Function

Public Function MakeID() As String
    MakeID = Format(Now(), "YYMMDDHHMMSS")
End Function

Public Function IDToDate(ByVal ID As String) As Date
    IDToDate = Format(Mid(ID, 5, 2) & "/" & _
                      Mid(ID, 3, 2) & "/" & _
                      Mid(ID, 1, 2) & "/" & _
                      Mid(ID, 7, 2) & "/" & _
                      Mid(ID, 9, 2) & "/" & _
                      Mid(ID, 11, 2), "dd-mm-yy hh:mm:ss")
End Function

Public Function LoadTextFile(ByVal FileName As String, _
                             ByVal FilePath As String)

    Dim fso As New FileSystemObject
    Dim TextFile As File
    Dim ts As TextStream
    
    If Right(FilePath, 1) <> "\" Then
        FilePath = FilePath & "\"
    End If
    
    Set TextFile = fso.GetFile(FilePath & FileName)
    Set ts = TextFile.OpenAsTextStream(ForReading)
    
    LoadTextFile = ts.ReadAll
    
    ts.Close

End Function

Public Function BrowseFolder(ByVal Title As String, _
                             Optional ByVal InitialFolder As String = vbNullString, _
                             Optional ByVal InitialView As Office.MsoFileDialogView = msoFileDialogViewList)
    Dim Folder As Variant
    Dim InitFolder As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        Folder = .SelectedItems(1)
        If Err.Number <> 0 Then
            Folder = vbNullString
        End If
    End With
    BrowseFolder = Folder
End Function

Public Function BrowseFile(ByVal Title As String, _
                           Optional ByVal InitialFolder As String = vbNullString, _
                           Optional ByVal InitialView As Office.MsoFileDialogView = msoFileDialogViewList)
    Dim File As Variant
    Dim InitFolder As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = Title
        .InitialView = InitialView
        If Len(InitialFolder) > 0 Then
            If Dir(InitialFolder, vbDirectory) <> vbNullString Then
                InitFolder = InitialFolder
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        File = .SelectedItems(1)
        If Err.Number <> 0 Then
            File = vbNullString
        End If
    End With
    BrowseFile = File
End Function

Public Function LastRow(ByVal ws As Worksheet)
    On Error Resume Next
    LastRow = ws.Cells.Find("*", ws.Range("A1"), xlFormula, xlPart, xlByRows, xlPrevious, False).Row
    On Error GoTo 0
End Function

Public Function LastCol(ByVal ws As Worksheet)
    On Error Resume Next
    LastCol = ws.Cells.Find("*", ws.Range("A1"), xlFormula, xlPart, xlByColumns, xlPrevious, False).Column
    On Error GoTo 0
End Function

Public Function CreateNamedRange(ByVal RangeName As String, _
                                 ByVal SheetName As String, _
                                 ByVal Range As String)
    Dim FullRangeName As String
    On Error Resume Next
    SheetName = "'" & SheetName & "'"
    FullRangeName = SheetName & "!" & RangeName
    With ActiveWorkbook
        .Sheets(SheetName).Select
        .Names(RangeName).Delete
        .Names.Add FullRangeName, "=" & Range
    End With
    On Error GoTo 0
End Function

Public Function NamedRangeExists(ByVal RangeName As String, _
                                 ByVal SheetName As String) As Boolean
    Dim RangeCheck As Range
    On Error Resume Next
    SheetName = "'" & SheetName & "'"
    Set RangeCheck = Range(SheetName & "!" & RangeName)
    On Error GoTo 0
    
    NamedRangeExists = Not RangeCheck Is Nothing
End Function

Public Function UpdateChartAxis(ByVal SheetName As String, _
                                ByVal ChartName As String, _
                                ByVal MinValue As Double, _
                                ByVal MaxValue As Double, _
                                Optional ByVal Axis As Long = xlPrimary) ' xlSecondary
    Dim cht As ChartObject
    Set cht = Worksheets(SheetName).ChartObjects(ChartName)
    
    On Error Resume Next
    cht.Chart.Axes(xlValue, Axis).MinimumScale = MinValue
    cht.Chart.Axes(xlValue, Axis).MaximumScale = MaxValue
    On Error GoTo 0
    
End Function


















