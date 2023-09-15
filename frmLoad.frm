VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   13680
      Top             =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HOSPITAL"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   975
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MARIA MATER MASSERICODEA "
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   13215
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6720
      TabIndex        =   1
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   -960
      Picture         =   "frmLoad.frx":0000
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   16455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim counts As Integer

Private Sub Timer1_Timer()
counts = counts + 2
If counts = 10 Then
lbl2.Caption = "10%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(253 * Rnd, 1 * Rnd, 2 * Rnd)
ElseIf counts = 20 Then
lbl2.Caption = "20%"
lbl1.ForeColor = RGB(300 * Rnd, 253 * Rnd, 367 * Rnd)
lbl2.ForeColor = RGB(200 * Rnd, 60 * Rnd, 5 * Rnd)
ElseIf counts = 30 Then
lbl2.Caption = "30%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(190 * Rnd, 80 * Rnd, 10 * Rnd)
ElseIf counts = 40 Then
lbl2.Caption = "40%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(50 * Rnd, 100 * Rnd, 15 * Rnd)
ElseIf counts = 50 Then
lbl2.Caption = "50%"
lbl1.ForeColor = RGB(0 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(2 * Rnd, 130 * Rnd, 50 * Rnd)
ElseIf counts = 60 Then
lbl2.Caption = "60%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(0 * Rnd, 150 * Rnd, 80 * Rnd)
ElseIf counts = 70 Then
lbl2.Caption = "70%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 257 * Rnd)
lbl2.ForeColor = RGB(0 * Rnd, 100 * Rnd, 150 * Rnd)
ElseIf counts = 80 Then
lbl2.Caption = "80%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(0 * Rnd, 50 * Rnd, 200 * Rnd)
ElseIf counts = 90 Then
lbl2.Caption = "90%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(0 * Rnd, 10 * Rnd, 240 * Rnd)
ElseIf counts = 100 Then
lbl2.Caption = "100%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(0 * Rnd, 0 * Rnd, 267 * Rnd)
ElseIf counts = 110 Then
lbl2.Caption = "110%"
lbl1.ForeColor = RGB(253 * Rnd, 253 * Rnd, 267 * Rnd)
lbl2.ForeColor = RGB(0 * Rnd, 0 * Rnd, 267 * Rnd)
Unload Me
MsgBox "loading complete", vbInformation, "loading message"
Form1.Show
End If
End Sub
