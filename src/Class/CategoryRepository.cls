VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CategoryRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dbConn  As DbConnection
Private conn    As ADODB.Connection
Private strSql  As String

Private Sub Class_Initialize()
    Set dbConn = New DbConnection
    Set conn = dbConn.GetOpenConnection
End Sub

Private Function MappingRecordToObject(ByVal rs As ADODB.Recordset) As Category
    Dim cat As Category
    
    Set cat = New Category
    cat.categoryId = IIf(IsNull(rs("CategoryID").Value), 0, rs("CategoryID").Value)
    cat.CategoryName = IIf(IsNull(rs("CategoryName").Value), "", rs("CategoryName").Value)
    cat.Description = IIf(IsNull(rs("Description").Value), "", rs("Description").Value)
    
    Set MappingRecordToObject = cat
End Function

Public Function Save(ByVal cat As Category) As Integer
    Dim cmd             As ADODB.Command
    Dim key             As Variant
    
    On Error GoTo errHandle

    strSql = "INSERT INTO Categories (CategoryName, Description) " & _
             "VALUES (?, ?)"
             
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = strSql
        .Prepared = True
        
        .Parameters(0) = cat.CategoryName
        .Parameters(1) = cat.Description

        .Execute
    End With
    Set cmd = Nothing
    
    cat.categoryId = dbConn.GetLastId(conn)
    Save = 1
    
    Exit Function

errHandle:
    Save = 0
End Function

Public Function Update(ByVal cat As Category) As Integer
    Dim cmd             As ADODB.Command
    Dim key             As Variant
    
    On Error GoTo errHandle

    strSql = "UPDATE Categories SET CategoryName = ?, Description = ? " & _
             "WHERE CategoryID = ?"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = strSql
        .Prepared = True
        
        .Parameters(0) = cat.CategoryName
        .Parameters(1) = cat.Description
        .Parameters(2) = cat.categoryId
        
        .Execute
    End With
    Set cmd = Nothing

    Update = 1
    
    Exit Function

errHandle:
    Update = 0
End Function

Public Function Delete(ByVal categoryId As Long) As Integer
    Dim cmd             As ADODB.Command
    Dim key             As Variant
    
    On Error GoTo errHandle

    strSql = "DELETE FROM Categories " & _
             "WHERE CategoryID = ?"
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdText
        .CommandText = strSql
        .Prepared = True
        
        .Parameters(0) = categoryId
        
        .Execute
    End With
    Set cmd = Nothing

    Delete = 1
    
    Exit Function

errHandle:
    Delete = 0
End Function

Public Function GetById(ByVal categoryId As Long) As Category
    Dim cmd     As ADODB.Command
    Dim rs      As ADODB.Recordset
    Dim cat     As Category
    
    On Error GoTo errHandler
    
    strSql = "SELECT CategoryID, CategoryName, Description " & _
             "FROM Categories " & _
             "WHERE CategoryID = " & categoryId
    Set rs = dbConn.OpenRecordset(strSql, conn)
    If Not rs.EOF Then
        Set cat = MappingRecordToObject(rs)
    End If
    Set cmd = Nothing
    Call dbConn.CloseRecordset(rs)
    
    Set GetById = cat
    
    Exit Function
errHandler:
End Function

Public Function GetAll() As Scripting.Dictionary
    Dim rs              As ADODB.Recordset
    Dim listOfCategory  As Scripting.Dictionary
    Dim cat             As Category
    
    On Error GoTo errHandler
    
    strSql = "SELECT CategoryID, CategoryName, Description " & _
             "From Categories " & _
             "ORDER BY CategoryName"
    Set rs = dbConn.OpenRecordset(strSql, conn)
    If Not rs.EOF Then
        Set listOfCategory = New Scripting.Dictionary
        
        Do While Not rs.EOF
        
            Set cat = MappingRecordToObject(rs)
            listOfCategory.Add cat.categoryId, cat
            
            rs.MoveNext
        Loop
    End If
    Call dbConn.CloseRecordset(rs)
    
    Set GetAll = listOfCategory
    
    Exit Function
errHandler:
End Function

Private Sub Class_Terminate()
    Call dbConn.CloseConnection(conn)
End Sub


