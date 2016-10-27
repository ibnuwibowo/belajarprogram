VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "Sipusta"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8805
   LinkTopic       =   "Form4"
   ScaleHeight     =   6495
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Form"
      Height          =   1455
      Left            =   4440
      TabIndex        =   12
      Top             =   720
      Width           =   4215
      Begin VB.CommandButton formpengembalian 
         Caption         =   "Form Pengembalian"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3975
      End
      Begin VB.CommandButton Formpeminjaman 
         Caption         =   "Form Peminjaman"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton editdatabuku 
      Caption         =   "Edit Data Buku"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   3855
   End
   Begin VB.CommandButton exit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton inputdatabuku 
      Caption         =   "Input Data Buku"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Inputdatapengarang 
      Caption         =   "Input Data Pengarang"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton inputjenisbuku 
      Caption         =   "Input Jenis Buku"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Data"
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4095
      Begin VB.CommandButton inputdataanggota 
         Caption         =   "Input Data Anggota"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Edit Data"
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   4095
      Begin VB.CommandButton editdataanggota 
         Caption         =   "Edit Data Anggota"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CommandButton editdatajenisbuku 
         Caption         =   "Edit Data Jenis Buku"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   3855
      End
      Begin VB.CommandButton editdatapengarang 
         Caption         =   "Edit Data Pengarang"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sistem Informasi Perpustakaan"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub editdataanggota_Click()
    MsgBox ("untuk saat ini menu belum tersedia")
End Sub

Private Sub editdatabuku_Click()
    Me.Hide
    Load Form4
    Form4.Show
End Sub

Private Sub editdatajenisbuku_Click()
MsgBox ("untuk saat ini menu belum tersedia")
End Sub

Private Sub editdatapengarang_Click()
    MsgBox ("untuk saat ini menu belum tersedia")
End Sub

Private Sub exit_Click()
Dim msg
msg = MsgBox("Anda yakin ingin keluar ?", vbYesNo + vbInformation, "Informasi")

If msg = vbYes Then
    End
End If
End Sub

Private Sub Formpeminjaman_Click()
    Me.Hide
    Load Form5
    Form5.Show
End Sub

Private Sub formpengembalian_Click()
    Me.Hide
    Load Form6
    Form6.Show
End Sub

Private Sub inputdataanggota_Click()
    Me.Hide
    Load Form3
    Form3.Show
End Sub

Private Sub inputdatabuku_Click()
    Me.Hide
    Load inputdatabuku
    inputdatabuku.Show
End Sub

Private Sub Inputdatapengarang_Click()
    Me.Hide
    Load pengarang
    pengarang.Show
    
End Sub

Private Sub inputjenisbuku_Click()
    Me.Hide
    Load jenisbuku
    jenisbuku.Show
End Sub
