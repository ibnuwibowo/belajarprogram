VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form5"
   ScaleHeight     =   6885
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton menu 
      Caption         =   "Menu"
      Height          =   495
      Left            =   4680
      TabIndex        =   25
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton reset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   2640
      TabIndex        =   24
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox txttanggal 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton pinjam 
      Caption         =   "Pinjam"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
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
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   7095
      Begin VB.TextBox txtstok 
         Height          =   285
         Left            =   4560
         TabIndex        =   19
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtpengarang 
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtjudul 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtkode 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Stok :"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Pengarang :"
         Height          =   255
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Judul Buku :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Kode Buku :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtkodepinjam 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   23
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtcari 
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtnama 
         Height          =   285
         Left            =   4440
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox cmbfilter 
         Height          =   315
         ItemData        =   "peminjaman.frx":0000
         Left            =   1320
         List            =   "peminjaman.frx":0010
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cmbkode 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Kode Pinjam"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cari Buku"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nama Anggota"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Filter Pencarian"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Anggota"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Tanggal Pinjam :"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   5760
      Width           =   1335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cbanggota As New ADODB.Recordset
Dim rsbuku As New ADODB.Recordset
Dim xxfilter As New ADODB.Recordset
Dim xx As New ADODB.Recordset
Dim koneksi As ADODB.Connection

Private Sub Command1_Click()
    Hide Me
    Load menu
    menu.Show
End Sub

Private Sub Form_Load()
kosong
pinjam.Enabled = False
Set koneksi = New ADODB.Connection
koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"
        
        rsbuku.CursorLocation = adUseClient
        rsbuku.Open "select * from tblBuku", koneksi, adOpenStatic, adLockOptimistic
        Set DataGrid1.DataSource = rsbuku
        DataGrid1.Refresh

        cmbkode.Clear
        cbanggota.CursorLocation = adUseClient
        Set cbanggota = New ADODB.Recordset
        cbanggota.Open "select * from tblAnggota", koneksi, adOpenStatic, adLockOptimistic
        Do Until cbanggota.EOF
            cmbkode.AddItem cbanggota!kode_anggota
            cbanggota.MoveNext
        Loop
        
End Sub

Private Sub menu_Click()
    Me.Hide
    Load menu
    menu.Show
End Sub

Private Sub reset_Click()
kosong
End Sub

Private Sub txtkodepinjam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(txtkodepinjam.Text) < 5 Then
            MsgBox "Kode pinjam harus di isi 5 digit, silahkan di isi kembali !", vbInformation, "Informasi"
            txtkodepinjam.SetFocus
        Else
            cmbkode.Enabled = True
            cmbkode.SetFocus
        End If
        End If
End Sub
Private Sub cmbkode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cbanggota.State = 1 Then cbanggota.Close
    cbanggota.Open "select * from tblAnggota where [kode_anggota]='" & cmbkode.Text & "'", koneksi, adOpenStatic, adLockOptimistic
    If cbanggota.RecordCount <> 0 Then
        txtnama.Text = cbanggota.Fields("nama_anggota").Value
        cmbfilter.SetFocus
    Else
        MsgBox "Data anggota tidak ada.", vbInformation, "Informasi"
        cmbkode.SetFocus
    End If
End If
End Sub

Private Sub cmbfilter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtcari.Text = ""
    txtcari.SetFocus
End If
End Sub

Private Sub txtcari_Change()
If rsbuku.State = 1 Then rsbuku.Close
rsbuku.Open "select * from tblBuku where [" & cmbfilter.Text & "] like '" & txtcari.Text & "%'", koneksi, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = rsbuku
DataGrid1.Refresh
DataGrid1.SetFocus
End Sub

Private Sub DataGrid1_Click()
txtkode.Text = rsbuku.Fields("kode_buku").Value
txtjudul.Text = rsbuku.Fields("judul").Value
txtpengarang.Text = rsbuku.Fields("pengarang").Value
txtstok.Text = rsbuku.Fields("stok").Value
txttanggal.Text = ""
End Sub

Private Sub pinjam_Click()
Dim psn
Dim a As String
a = "Belum Kembali"
    psn = MsgBox("Apakah anda yakin data akan di simpan ?", vbYesNo + vbInformation, "Informasi")
    
    xx.CursorLocation = adUseClient
    xx.Open "select * from tblPeminjaman", koneksi, adOpenStatic, adLockOptimistic

    If psn = vbYes Then
        With xx
            .AddNew
            .Fields("kode_peminjaman").Value = txtkodepinjam.Text
            .Fields("kode_buku").Value = txtkode.Text
            .Fields("kode_anggota").Value = cmbkode.Text
            .Fields("tgl_peminjaman").Value = txttanggal.Text
            .Fields("history").Value = a
            .Update
        End With
        xx.Close
        rsbuku.Fields("stok").Value = rsbuku.Fields("stok").Value - 1
        rsbuku.Update
        Set DataGrid1.DataSource = rsbuku
        DataGrid1.Refresh
        kosong
    End If
End Sub
Private Sub kosong()
    txtkodepinjam.Text = ""
    txtnama.Text = ""
    txtcari.Text = ""
    txtkode.Text = ""
    txtjudul.Text = ""
    txtpengarang.Text = ""
    txtstok.Text = ""
    txttanggal.Text = ""
End Sub

Private Sub kunci(a As Boolean)
    cmbkode.Enabled = a
    cmbfilter.Enabled = a
    txtnama.Enabled = a
    txtcari.Enabled = a
    txtkode.Enabled = a
    txtjudul.Enabled = a
    txtpengarang.Enabled = a
    txtstok.Enabled = a
    txttanggal.Enabled = a
End Sub

Private Sub txttanggal_Change()
If txttanggal = "" Then
    pinjam.Enabled = False
    Else
    pinjam.Enabled = True
End If
End Sub
