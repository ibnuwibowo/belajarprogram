VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Operator Test"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Op Logika : "
      Height          =   975
      Left            =   360
      TabIndex        =   15
      Top             =   4680
      Width           =   3495
      Begin VB.OptionButton Option14 
         Caption         =   "And"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Or"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Not"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.OptionButton Option10 
      Caption         =   "<="
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Op Perbandingan :"
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   3495
      Begin VB.OptionButton Option11 
         Caption         =   ">="
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "<>"
         Height          =   195
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option8 
         Caption         =   ">"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option7 
         Caption         =   "="
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         Caption         =   "<"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.OptionButton Option5 
      Caption         =   "&&"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "*"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "+"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Op Aritmatika : "
      Height          =   1335
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
      Begin VB.OptionButton Option4 
         Caption         =   "/"
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "-"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Var 2 :"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Var 1 :"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim var1 As Single, var2 As Single
Dim hasil As Single
Private Sub Option1_Click()
    var1 = Text1.Text
    var2 = Text2.Text
    hasil = var1 + var2
    Label3.Caption = hasil
End Sub

Private Sub Option10_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = (var1 <= var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option11_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = (var1 >= var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option12_Click()
    var1 = IIf(Text1.Text = "true", -1, 0)
    hasil = Not (var1)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option13_Click()
    var1 = IIf(Text1.Text = "true", -1, 0)
    var2 = IIf(Text2.Text = "true", -1, 0)
    hasil = (var1 Or var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option14_Click()
var1 = IIf(Text1.Text = "true", -1, 0)
    var2 = IIf(Text2.Text = "true", -1, 0)
    hasil = (var1 And var2)
    Label3.Caption = Format(hasil, "true/false")

End Sub

Private Sub Option2_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = var1 * var2
    Label3.Caption = hasil
End Sub

Private Sub Option3_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = var1 - var2
    Label3.Caption = hasil
End Sub

Private Sub Option4_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = var1 / var2
    Label3.Caption = hasil
End Sub

Private Sub Option5_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = var1 & var2
    Label3.Caption = hasil
End Sub

Private Sub Option6_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = (var1 < var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option7_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = (var1 = var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option8_Click()
    var1 = Text1.Text
    var2 = Text2.Text
    hasil = (var1 > var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub

Private Sub Option9_Click()
var1 = Text1.Text
    var2 = Text2.Text
    hasil = (var1 <> var2)
    Label3.Caption = Format(hasil, "true/false")
End Sub
