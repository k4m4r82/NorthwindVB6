VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetOpenConnection() As ADODB.Connection
    Dim conn    As ADODB.Connection
    Dim strCon  As String
    Dim dbPath  As String
    
    On Error GoTo errHandle
    
    dbPath = App.Path & "\db\Northwind.db"
    strCon = "Driver=SQLite3 ODBC Driver;Database=" & dbPath
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = strCon
    conn.Open
    
    Set GetOpenConnection = conn
    
    Exit Function
errHandle:
    ' do nothing
End Function

Public Sub CloseConnection(ByVal conn As ADODB.Connection)
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub

Public Function OpenRecordset(ByVal query As String, ByVal conn As ADODB.Connection) As ADODB.Recordset
    Dim obj As ADODB.Recordset
                    
    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
        
    obj.Open query, conn, adOpenForwardOnly, adLockReadOnly
    
    Set OpenRecordset = obj
End Function

Public Sub CloseRecordset(ByVal rs As ADODB.Recordset)
    On Error Resume Next
    
    If Not (rs Is Nothing) Then
        If rs.State = adStateOpen Then rs.Close
    End If
    
    Set rs = Nothing
End Sub

Public Function GetLastId(ByVal conn As ADODB.Connection) As Long
    Dim sql As String
    Dim rs  As ADODB.Recordset
    
    sql = "SELECT LAST_INSERT_ROWID()"
    Set rs = OpenRecordset(sql, conn)
    
    If Not rs.EOF Then
        GetLastId = rs.Fields(0).Value
    End If
End Function

