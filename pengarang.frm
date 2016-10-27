VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pengarang 
   Caption         =   "Input Pengarang"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   LinkTopic       =   "Form4"
   ScaleHeight     =   2685
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton menu 
      Caption         =   "Menu"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2143
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
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton add 
      Caption         =   "Add"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtket 
      Height          =   405
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox txtkode 
      Height          =   375
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Pengarang"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "pengarang"
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
        kosong
    Else
        add.Caption = "&Add"
        save.Enabled = False
        DataGrid1.Enabled = True
        kosong
        kunci (False)
    End If
            
End Sub
Private Sub Form_Load()
Set koneksi = New ADODB.Connection
kosong
kunci (False)
    If rs.State = 1 Then rs.Close
        koneksi.Provider = "Microsoft.Jet.OLEDB.4.0;" & "jet oledb:database password="
        koneksi.Open "D:\DATA\KULIAH\SEM5\Pem Tingkat Lanjut\LATIHAN\VB\sipusta.mdb"
        rs.CursorLocation = adUseClient
        rs.Open "select * from tblPengarang order by [kode_pengarang]", koneksi, adOpenStatic, adLockOptimistic
        Set DataGrid1.DataSource = rs
    save.Enabled = False
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
            .Fields("kode_pengarang").Value = txtkode.Text
            .Fields("keterangan").Value = txtket.Text
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
Private Sub kosong()
    txtkode.Text = ""
    txtket.Text = ""
End Sub

Private Sub kunci(a As Boolean)
    txtkode.Enabled = a
    txtket.Enabled = a
End Sub

Private Sub txtket_Change()
    If txtket.Text = "" Then
        save.Enabled = False
    Else
        save.Enabled = True
    End If
End Sub
Private Sub txtkode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(txtkode.Text) = 5 Then
            txtket.Enabled = True
            txtket.SetFocus
        Else
            MsgBox "Data harus di isi 5 digit, silahkan di isi kembali !", vbInformation, "Informasi"
            txtkode.SetFocus
        End If
    End If
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        save.Enabled = True
        save.SetFocus
    End If
End Sub


