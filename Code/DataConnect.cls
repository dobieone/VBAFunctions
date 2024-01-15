VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pDBName As String
Private pDBPath As String
Private pSQL As String

Private pQueryPath As String
Private pQueryFileName As String
Private pNumberOfRecords As Long

Private pReturnResults As Boolean
Private pResultsTab As String
Private pRowStart As Long
Private pColStart As Long
Private pClearSheet As Boolean
Private pOutputHeadings As Boolean

Private CN As ADODB.Connection
Private rs As ADODB.Recordset
Private KeyValue As Dictionary

Private Sub Class_Initialize()
    Call Reset
End Sub

Public Property Get Databasename() As String
    Databasename = pDBName
End Property

Public Property Let Databasename(ByVal FileName As String)
    pDBName = FileName
End Property

Public Property Get DatabasePath() As String
    DatabasePath = pDBPath
End Property

Public Property Let DatabasePath(ByVal FilePath As String)
    If Right(FilePath, 1) <> "\" And Len(QueryPath) > 0 Then
        FilePath = FilePath & "\"
    End If
    pDBPath = FilePath
End Property

Public Property Get SQL() As String
    SQL = pSQL
End Property

Public Property Let SQL(ByVal SQL As String)
    pSQL = SQL
End Property

Public Property Get QueryPath() As String
    QueryPath = pQueryPath
End Property

Public Property Let QueryPath(ByVal QueryPath As String)
    If Right(QueryPath, 1) <> "\" And Len(QueryPath) > 0 Then
        QueryPath = QueryPath & "\"
    End If
    pQueryPath = QueryPath
End Property

Public Property Get QueryFileName() As String
    QueryFileName = pQueryFileName
End Property

Public Property Let QueryFileName(ByVal QueryFileName As String)
    pQueryFileName = QueryFileName
End Property

Public Property Get ReturnResults() As Boolean
    ReturnResults = pReturnResults
End Property

Public Property Let ReturnResults(ByVal ReturnResults As Boolean)
    pReturnResults = ReturnResults
End Property

Public Property Get ResultsTab() As String
    ResultsTab = pResultsTab
End Property

Public Property Let ResultsTab(ByVal ResultsTab As String)
    pResultsTab = ResultsTab
End Property

Public Property Get RowStart() As Long
    RowStart = pRowStart
End Property

Public Property Let RowStart(ByVal RowStart As Long)
    pRowStart = RowStart
End Property

Public Property Get ColStart() As Long
    ColStart = pColStart
End Property

Public Property Let ColStart(ByVal ColStart As Long)
    pColStart = ColStart
End Property

Public Property Get ClearSheet() As Boolean
    ClearSheet = pClearSheet
End Property

Public Property Let ClearSheet(ByVal ClearSheet As Boolean)
    pClearSheet = ClearSheet
End Property

Public Property Get OutputHeadings() As Boolean
    OutputHeadings = pOutputHeadings
End Property

Public Property Let OutputHeadings(ByVal OutputHeadings As Boolean)
    pOutputHeadings = OutputHeadings
End Property

Public Property Get NumberOfRecords() As String
    NumberOfRecords = pNumberOfRecords
End Property

Public Function AddKey(ByVal Key As String, ByVal Value As Variant)
    If KeyValue.Exists(Key) = False Then
        KeyValue.Add Key, CStr(Value)
    End If
End Function

Public Function Connect() As Boolean
    Dim FileType As String
    Dim ConnectionString As String

    FileType = Right(Databasename, Len(Databasename) - InStrRev(Databasename, "."))

    Select Case LCase(FileType)
        Case "accdb", "mdb", "accdr"
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & DatabasePath & Databasename & ";" & _
                                "Persist Security Info=False;"
        Case "txt", "csv"
            ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
                                "Dbq=" & DatabasePath & ";" & _
                                "Extensions=csv,txt;"
        Case "xlsx", "xlsm", "xlsb"
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & DatabasePath & Databasename & ";" & _
                                "Extended Properties=""Excel 12.0;HDR=YES"";"
        Case "xls"
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                "Data Source=" & DatabasePath & Databasename & ";" & _
                                "Extended Properties=""Excel 8.0;HDR=YES"";"
    End Select

    Set CN = New ADODB.Connection
    CN.Open ConnectionString

    If CN.State = 0 Then
        Connect = False
    Else
        Connect = True
    End If
    
End Function

Public Sub Disconnect()
    CN.Close
    Set CN = Nothing
End Sub

' Returns rs as array if ReturnResults = True
' Returns True if ReturnResults = False
Public Function Run() As Variant

    Dim Count As Long

    BuildQuery

    Set rs = New Recordset
    rs.Open SQL, CN, adOpenStatic

    Select Case rs.State
        Case adStateOpen
            If ReturnResults Then
                Run = Application.Transpose(rs.GetRows)
            Else
                If ClearSheet Then
                    Worksheets(ResultsTab).Cells.Clear
                End If
                If OutputHeadings Then
                    For Count = 0 To rs.Fields.Count - 1
                        Worksheets(ResultsTab).Cells(RowStart, ColStart + Count) = _
                                rs.Fields(Count).Name
                    Next Count
                End If
                Worksheets(ResultsTab).Cells(RowStart + 1, ColStart).CopyFromRecordset rs

                Run = True
            End If

            pNumberOfRecords = rs.RecordCount
            rs.Close

        Case adStateClosed
            Run = True
    End Select

    Set rs = Nothing

End Function

Public Sub BuildQuery()
    Dim Key As Variant

    If QueryPath <> "" And QueryFileName <> "" Then
        SQL = LoadTextFile(QueryFileName, QueryPath)
    End If

    For Each Key In KeyValue
        SQL = Replace(SQL, "{" & Key & "}", KeyValue.Item(Key))
    Next Key

End Sub

Public Sub Reset()
    
    Set KeyValue = New Dictionary

    Databasename = ""
    DatabasePath = ""

    QueryFileName = ""
    QueryPath = ""
    SQL = ""

    ReturnResults = False

    ResultsTab = "Sheet1"
    RowStart = 1
    ColStart = 1
    ClearSheet = True
    OutputHeadings = True

End Sub

