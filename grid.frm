VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form Input Data"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8145
   LinkTopic       =   "Form3"
   ScaleHeight     =   4620
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Menu 
      Caption         =   "Menu"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2990
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
   Begin VB.TextBox txtalamat 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton add 
      Caption         =   "Add"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtkode 
      Height          =   375
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Anggota"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Anggota"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim koneksi As ADODB.Connection

Private Sub add_Click()
    If add.Caption = "&Add" Then
        add.Caption = "&Cancel"
        txtkode.Enabled = True
        txtkode.SetFocus
        save.Enabled = False
        DataGrid1.Enabled = False
    Else
        add.Caption = "&Add"
        save.Enabled = False
        DataGrid1.Enabled = True
        kosong
        kunci (False)
    End If
            
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub menu_Click()
    Me.Hide
    Load menu
    menu.Show
End Sub

Private Sub save_Click()
    Dim psn
    psn = MsgBox("Apakah anda yakin data akan di simpan ?", vbYesNo + vbInformation, "Informasi")
    If psn = vbYes Then
        With rs
            .AddNew
            .Fields("kode_anggota").Value = txtkode.Text
            .Fields("nama_anggota").Value = txtnama.Text
            .Fields("alamat").Value = txtalamat.Text
            .Update
        End With
        kosong
        kunci (False)
        add.Caption = "&Add"
        save.Enabled = False
        add.SetFocus
    Else
        add_Click
    End If
End Sub
Private Sub Form_Load()
Set koneksi = New ADODB.Connection
    If rs.State = 1 Then rs.Close
        koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
        koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"
        rs.CursorLocation = adUseClient
        rs.Open "select * from tblAnggota order by [kode_anggota]", koneksi, adOpenStatic, adLockOptimistic
        Set DataGrid1.DataSource = rs
    save.Enabled = False
    kosong
    kunci (False)
End Sub

Private Sub kosong()
    txtkode.Text = ""
    txtnama.Text = ""
    txtalamat.Text = ""
End Sub

Private Sub kunci(a As Boolean)
    txtkode.Enabled = a
    txtnama.Enabled = a
    txtalamat.Enabled = a
End Sub

Private Sub txtalamat_Change()
    If txtalamat.Text = "" Then
        save.Enabled = False
    Else
        save.Enabled = True
    End If
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtkode.Text) = 5 Then
            txtnama.Enabled = True
            txtnama.SetFocus
        Else
            MsgBox "Data kode harus 5 digit, silahkan di isi kembali !", vbInformation, "Informasi"
            txtkode.SetFocus
        End If
    End If
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            txtalamat.Enabled = True
            txtalamat.SetFocus
           
            
        End If
        
End Sub
Private Sub txtalamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        save.Enabled = True
        save.SetFocus
    End If
End Sub
