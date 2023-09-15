VERSION 5.00
Begin VB.Form frmloading 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   10170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18030
   FillColor       =   &H00800000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form5.frx":0CCA
   Picture         =   "Form5.frx":1994
   ScaleHeight     =   10170
   ScaleWidth      =   18030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   4560
      TabIndex        =   0
      Top             =   2400
      Width           =   11655
      Begin VB.Timer timer 
         Interval        =   300
         Left            =   10440
         Top             =   5520
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         TabIndex        =   12
         Top             =   6720
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   855
         Left            =   5040
         Shape           =   3  'Circle
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label lblload 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   11
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "2017"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6120
         TabIndex        =   10
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "WINDOWS BASED         PROGRAM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   4800
         TabIndex        =   9
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Piracy is highly prohibited"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         TabIndex        =   8
         Top             =   5760
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "This Product is Protected by HTECH && NBC  Software Production Companies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   120
         MouseIcon       =   "Form5.frx":12162
         TabIndex        =   7
         Top             =   4080
         Width           =   4215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "This Program is Designed       mainly for Hospitals"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8640
         TabIndex        =   6
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Health is wealth..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   7560
         TabIndex        =   5
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enhancing New Techniques in the World of Health"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         TabIndex        =   4
         Top             =   1320
         Width           =   6975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Catholic Hospital"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   3120
         TabIndex        =   3
         Top             =   2280
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "he"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   1320
         TabIndex        =   2
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00400000&
         BorderWidth     =   15
         X1              =   480
         X2              =   2040
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderWidth     =   15
         X1              =   1200
         X2              =   1200
         Y1              =   2280
         Y2              =   3480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "HOSPITAL  MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   9375
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   240
         Picture         =   "Form5.frx":12E2C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim today  As Variant
Dim counts As Integer
Private Sub Timer1_Timer()
Timer1 = Time$
End Sub

Private Sub timer_Timer()
counts = counts + 1
timer = counts
If counts = 10 Then
lblload.ForeColor = &HFF0000
lblload.Caption = "Initializing...."
ElseIf counts = 20 Then
lblload.ForeColor = &HFF00&
lblload.Caption = "Loading Images..."
ElseIf counts = 40 Then
lblload.ForeColor = &HFF00FF
lblload.Caption = "Loading Font and Styles..."
ElseIf counts = 60 Then
lblload.Caption = "Loading Database..."
ElseIf counts = 70 Then
lblload.Caption = "Loading Interface..."
ElseIf counts = 80 Then
lblload.Caption = "Please wait..."
ElseIf counts = 90 Then
lblload.Caption = "Completing..."
ElseIf counts = 100 Then
Unload Me
frmLogin.Show
End If
    'Unload Me
    'frmLogin.Show
End Sub
