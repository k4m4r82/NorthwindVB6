VERSION 5.00
Begin VB.Form frmAddEditCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Category"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtCategoryName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdSelesai 
      Caption         =   "Selesai"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1005
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Category Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddEditCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mode     As ACTION
Public cat      As Category
Public isSimpan As Boolean

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    Dim repo    As CategoryRepository
    
    If mode = ADD_DATA Then
        Set cat = New Category
    End If
    
    cat.CategoryName = txtCategoryName.Text
    cat.Description = txtDescription
    
    Dim result As Integer
    Set repo = New CategoryRepository
    
    If mode = ADD_DATA Then
        result = repo.Save(cat)
    Else
        result = repo.Update(cat)
    End If
    
    isSimpan = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    isSimpan = False
    
    If mode = EDIT_DATA Then
        txtCategoryName.Text = cat.CategoryName
        txtDescription.Text = cat.Description
    End If
End Sub
