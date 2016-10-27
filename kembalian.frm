VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "form kembalian"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415
   LinkTopic       =   "Form6"
   ScaleHeight     =   5835
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton menu 
      Caption         =   "Menu"
      Height          =   495
      Left            =   4440
      TabIndex        =   19
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   4440
      TabIndex        =   18
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "buku kembali"
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtdenda 
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text7"
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txttanggalkembali 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txttanggalpinjam 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtkodebuku 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtkodeanggota 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtkodepinjam 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3000
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3625
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
   Begin VB.TextBox txtcaridata 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.ComboBox cmbsaring 
      Height          =   315
      ItemData        =   "kembalian.frx":0000
      Left            =   1920
      List            =   "kembalian.frx":0010
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Denda Rp."
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Tanggal Kembali"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Tanggal Pinjam"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Kode Buku"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Kode Anggota"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Kode Pinjam"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Cari data peminjaman"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Saring Pencarian"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsbuku As New ADODB.Recordset
Dim rbbuku As New ADODB.Recordset
Dim rskembali As New ADODB.Recordset
Dim xx As New ADODB.Recordset
Dim yy As New ADODB.Recordset
Dim tambah As New ADODB.Recordset
Dim koneksi As ADODB.Connection

Private Sub Command2_Click()
Set DataGrid1.DataSource = rsbuku
    DataGrid1.Refresh
    kosong
End Sub

Private Sub Form_Load()
Set koneksi = New ADODB.Connection
koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"
      
        rsbuku.CursorLocation = adUseClient
        rsbuku.Open "select * from tblPeminjaman where history like 'belum kembali'", koneksi, adOpenStatic, adLockOptimistic
        Set DataGrid1.DataSource = rsbuku
        DataGrid1.Refresh
        
kosong
kunci
End Sub

Private Sub menu_Click()
    Me.Hide
    Load menu
    menu.Show
End Sub

Private Sub txtcaridata_Change()
If rsbuku.State = 1 Then rsbuku.Close
rsbuku.Open "select * from tblPeminjaman where [" & cmbsaring.Text & "] like '" & txtcaridata.Text & "%'", koneksi, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = rsbuku
DataGrid1.Refresh
End Sub

Private Sub DataGrid1_Click()
txtkodepinjam.Text = rsbuku.Fields("kode_peminjaman").Value
txtkodeanggota.Text = rsbuku.Fields("kode_anggota").Value
txtkodebuku.Text = rsbuku.Fields("kode_buku").Value
txttanggalpinjam.Text = rsbuku.Fields("tgl_peminjaman").Value
End Sub


Private Sub txttanggalkembali_KeyPress(KeyAscii As Integer)
Dim a
If KeyAscii = 13 Then
     a = Val(txttanggalkembali.Text) - Val(txttanggalpinjam.Text)
        If a > 3 Then
            txtdenda.Text = (a - 3) * 2000
        Else
            txtdenda.Text = 0
    End If
End If
End Sub

Private Sub Command1_Click()
    xx.CursorLocation = adUseClient
    xx.Open "select * from tblPengembalian", koneksi, adOpenStatic, adLockOptimistic
        With xx
            .AddNew
            .Fields("kode_pinjam").Value = txtkodepinjam.Text
            .Fields("tanggal_pinjam").Value = txttanggalpinjam.Text
            .Fields("tanggal_kembali").Value = txttanggalkembali.Text
            .Fields("denda").Value = txtdenda.Text
            .Update
        End With
        xx.Close
    
    Dim b As String
    b = "done"
    yy.CursorLocation = adUseClient
    yy.Open "select * from tblPeminjaman where kode_peminjaman ='" & txtkodepinjam.Text & "'", koneksi, adOpenStatic, adLockOptimistic
        With yy
            .Update
            .Fields("history").Value = b
            .Update
        End With
        yy.Close
          
        tambah.CursorLocation = adUseClient
        tambah.Open "select * from tblBuku where kode_buku = '" & txtkodebuku.Text & "'", koneksi, adOpenStatic, adLockOptimistic
        
        'tambah.Fields("stok").Value = rsbuku.Fields("stok").Value + 1
        'tambah.Update
  
        
End Sub
Private Sub kosong()
txtkodepinjam.Text = ""
txtkodeanggota.Text = ""
txtkodebuku.Text = ""
txttanggalpinjam.Text = ""
txttanggalkembali.Text = ""
txtdenda.Text = ""
End Sub

Private Sub kunci()
txtkodepinjam.Enabled = False
txtkodeanggota.Enabled = False
txtkodebuku.Enabled = False
txttanggalpinjam.Enabled = False
txtdenda.Enabled = False
End Sub
