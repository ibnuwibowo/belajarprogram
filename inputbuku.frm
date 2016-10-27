VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Inputdatabuku 
   Caption         =   "Input Data Buku"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10350
   LinkTopic       =   "Form4"
   ScaleHeight     =   5355
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton menu 
      Caption         =   "Menu"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtstok 
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox txtkodepengarang 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtjudul 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1320
      Width           =   6735
   End
   Begin VB.ComboBox txtkodejenis 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtkodebuku 
      Height          =   405
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Stok Buku"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Kode Pengarang"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Judul Buku"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Jenis"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Buku"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Inputdatabuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim cb As New ADODB.Recordset
Dim kp As New ADODB.Recordset
Dim va As String
Dim koneksi As ADODB.Connection

Private Sub Form_Load()
kosong
Set koneksi = New ADODB.Connection
koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"

        txtkodejenis.Clear
        Set cb = New ADODB.Recordset
        cb.Open "select kode_jenis from tblJenisbuku", koneksi, adOpenDynamic, adLockOptimistic
        Do Until cb.EOF
            txtkodejenis.AddItem cb!kode_jenis
            cb.MoveNext
        Loop
        cb.Close
        
        txtkodepengarang.Clear
        Set kp = New ADODB.Recordset
        kp.Open "select kode_pengarang from tblPengarang", koneksi, adOpenDynamic, adLockOptimistic
        Do Until kp.EOF
            txtkodepengarang.AddItem kp!kode_pengarang
            kp.MoveNext
        Loop
        kp.Close
        
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select * from tblBuku order by [kode_buku]", koneksi, adOpenStatic, adLockOptimistic
        Set DataGrid1.DataSource = rs
End Sub
        
Private Sub kosong()
    txtkodebuku.Text = ""
    txtkodejenis.Text = "Pilih jenis buku"
    txtjudul.Text = ""
    txtkodepengarang.Text = "Pilih pengarang"
    txtstok.Text = ""
End Sub

Private Sub kunci(a As Boolean)
    txtkodejenis.Enabled = a
    txtjudul.Enabled = a
    txtkodepengarang.Enabled = a
    txtstok.Enabled = a
End Sub

Private Sub menu_Click()
    Me.Hide
    Load menu
    menu.Show
End Sub

Private Sub txtkodejenis_Click()
Set koneksi = New ADODB.Connection
koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"

    cb.Open "select * from tblJenisbuku where kode_jenis ='" & Left(txtkodejenis.Text, 2) & "'", koneksi, adOpenDynamic, adLockOptimistic
    cb.Requery
    With cb
        If .EOF And .BOF Then
            MsgBox "data tidak di temukan", vbOKOnly + vbCritical, "Eror"
        Exit Sub
        Else
            txtkodejenis.Text = !kode_jenis
            List1.Text = !keterangan
        End If
    End With
    cb.Close
End Sub

Private Sub txtstok_Change()
    If txtstok.Text = "" Then
        save.Enabled = False
    Else
        save.Enabled = True
    End If
End Sub

Private Sub cancel_Click()
    kosong
End Sub

Private Sub save_Click()
    Dim psn
    psn = MsgBox("Apakah anda yakin data akan di simpan ?", vbYesNo + vbInformation, "Informasi")
    If psn = vbYes Then
        With rs
            .AddNew
            .Fields("kode_buku").Value = txtkodebuku.Text
            .Fields("kode_jenis").Value = txtkodejenis.Text
            .Fields("judul").Value = txtjudul.Text
            .Fields("kode_pengarang").Value = txtkodepengarang.Text
            .Fields("stok").Value = txtstok.Text
            .Update
        End With
        save.Enabled = False
        txtkodebuku.SetFocus
        kosong
    Else
        txtkodebuku.SetFocus
        kosong
    End If
    
End Sub

Private Sub txtkodebuku_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtkodebuku.Text) <= 5 Then
            txtkodejenis.Enabled = True
            txtkodejenis.SetFocus
        Else
            MsgBox "Data maksimal 5 digit, silahkan di isi kembali !", vbInformation, "Informasi"
            txtkodebuku.SetFocus
        End If
    End If
End Sub

Private Sub txtkodejenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtjudul.Enabled = True
        txtjudul.SetFocus
    End If
    
End Sub

Private Sub txtjudulbuku_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtkodepengarang.Enabled = True
        txtkodepengarang.SetFocus
    End If
End Sub

Private Sub txtkodepengarang_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtstok.Enabled = True
        txtstok.SetFocus
    End If
End Sub

Private Sub txtstok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        save.Enabled = True
        save.SetFocus
    End If
End Sub
