VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Edit Data"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3210
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton menu 
      Caption         =   "Menu"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtstok 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtpengarang 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtjudul 
      Height          =   405
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   2655
   End
   Begin VB.ComboBox combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Stok"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Pengarang"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Judul Buku"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Buku"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cb As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim xx As New ADODB.Recordset
Dim koneksi As ADODB.Connection

Private Sub cancel_Click()
kosong
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rs.CursorLocation = adUseClient
rs.Open "select * from tblbuku where kode_buku ='" & combo1.Text & "'", koneksi, adOpenStatic, adLockOptimistic
  If rs.RecordCount <> 0 Then
    txtjudul.Text = rs.Fields("judul").Value
    txtpengarang.Text = rs.Fields("pengarang").Value
    txtstok.Text = rs.Fields("stok").Value
    rs.Close
  Else
    MsgBox "data tidak ada"
    save.SetFocus
  End If
   
End If
End Sub

Private Sub Command1_Click()
    Hide Me
    Load menu
    menu.Show
End Sub

Private Sub Form_Load()
kosong

Set koneksi = New ADODB.Connection
koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"
        combo1.Clear
        Set cb = New ADODB.Recordset
        cb.Open "select kode_buku from tblBuku", koneksi, adOpenStatic, adLockOptimistic
        Do Until cb.EOF
            combo1.AddItem cb!kode_buku
            cb.MoveNext
        Loop
        
End Sub
Private Sub kosong()
    txtjudul.Text = ""
    txtpengarang.Text = ""
    txtstok.Text = ""
End Sub
Private Sub kunci(a As Boolean)
    txtjudul.Enabled = a
    txtpengarang.Enabled = a
    txtstok.Enabled = a
End Sub

Private Sub menu_Click()
    Me.Hide
    Load menu
    menu.Show
End Sub

Private Sub save_Click()
Dim psn
    psn = MsgBox("Apakah anda yakin data akan di simpan ?", vbYesNo + vbInformation, "Informasi")
    
    xx.CursorLocation = adUseClient
    xx.Open "select * from tblBuku where kode_buku ='" & combo1.Text & "'", koneksi, adOpenStatic, adLockOptimistic

    If psn = vbYes Then
        
        With xx
            .Update
            .Fields("judul").Value = txtjudul.Text
            .Fields("pengarang").Value = txtpengarang.Text
            .Fields("stok").Value = txtstok.Text
            .Update
        End With
        xx.Close
   
    Else
        combo1.SetFocus
        kosong
    End If
End Sub
