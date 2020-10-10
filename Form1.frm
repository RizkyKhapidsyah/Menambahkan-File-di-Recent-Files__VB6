VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menambahkan File di Recent Files"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)

Private Sub Command1_Click()
  'Ganti "C:\My Documents\Test.txt" dengan nama file
  'yang ingin Anda tambahkan ke Recent Files (Documents
  'menu).
  'Agar file ini bisa dipanggil nantinya dari Documents
  'menu tersebut, sebaiknya file ini harus ada di dalam
  'C:\My Documents. Jika tidak, file ini tetap
  'ditambahkan ke dalam Documents menu, tetapi tidak
  'dapat dibuka.
  Call SHAddToRecentDocs(2, "C:\My Documents\Test.txt")
  MsgBox "Lihat hasilnya dari menu Start->Documents", vbInformation, "Sukses Tambah Recent Files"
End Sub


Private Sub Form_Load()
    Command1.Caption = "Tambahkan"
End Sub
