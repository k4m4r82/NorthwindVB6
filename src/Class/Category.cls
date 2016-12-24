VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCategoryID     As Integer
Private mCategoryName   As String
Private mDescription    As String

Public Property Let CategoryID(ByVal vData As Long)
    mCategoryID = vData
End Property
Public Property Get CategoryID() As Long
    CategoryID = mCategoryID
End Property

Public Property Let CategoryName(ByVal vData As String)
    mCategoryName = vData
End Property
Public Property Get CategoryName() As String
    CategoryName = mCategoryName
End Property

Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property
Public Property Get Description() As String
    Description = mDescription
End Property


