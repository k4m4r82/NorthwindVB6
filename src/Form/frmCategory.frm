VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategory 
   Caption         =   "Data Category"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   2790
      TabIndex        =   3
      Top             =   5775
      Width           =   1215
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "Perbaiki"
      Height          =   495
      Left            =   1455
      TabIndex        =   2
      Top             =   5775
      Width           =   1215
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5775
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsvCategory 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9763
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private repo            As CategoryRepository
Private listOfCategory  As Scripting.Dictionary
Private cat             As Category

Private Sub cmdHapus_Click()
    If Not (lsvCategory.SelectedItem.Index < 0) Then
        Set cat = listOfCategory.Items()(lsvCategory.SelectedItem.Index - 1)
    
        If Not (cat Is Nothing) Then
            Dim result As Integer
            
            If MsgBox("Apakah data ini ingin di hapus ???", vbExclamation + vbYesNo, "Konfirmasi") = vbYes Then
                result = repo.Delete(cat.categoryId)
                
                If result > 0 Then
                    Call LoadCategory
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdPerbaiki_Click()
    If Not (lsvCategory.SelectedItem.Index < 0) Then
        Set cat = listOfCategory.Items()(lsvCategory.SelectedItem.Index - 1)
    
        If Not (cat Is Nothing) Then
            With frmAddEditCategory
                .mode = EDIT_DATA
                Set .cat = cat
                
                .Show vbModal
                
                If .isSimpan Then
                    Call LoadCategory
                End If
            End With
        End If
    End If
End Sub

Private Sub cmdTambah_Click()
    With frmAddEditCategory
        .mode = ADD_DATA
        .Show vbModal
        
        If .isSimpan Then
            Call LoadCategory
        End If
    End With
End Sub

Private Sub Form_Load()
    Call InisialisasiListView
    Call LoadCategory
End Sub

Private Sub InisialisasiListView()
    With lsvCategory
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        
        .ColumnHeaders.Add , , "No.", 500
        .ColumnHeaders.Add , , "CategoryName", 2000
        .ColumnHeaders.Add , , "Description", 4000
    End With
End Sub

Private Sub LoadCategory()
    Dim key             As Variant
    Dim cat             As Category
    Dim noUrut          As Integer
    
    Set repo = New CategoryRepository
    Set listOfCategory = repo.GetAll
    
    noUrut = 1
    lsvCategory.ListItems.Clear
    For Each key In listOfCategory
        
        Set cat = listOfCategory.Item(key) ' ekstrak objek category
        
        lsvCategory.ListItems.Add , , noUrut
        lsvCategory.ListItems(noUrut).SubItems(1) = cat.CategoryName
        lsvCategory.ListItems(noUrut).SubItems(2) = cat.Description
        
        noUrut = noUrut + 1
    Next
End Sub
