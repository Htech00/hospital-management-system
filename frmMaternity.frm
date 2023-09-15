VERSION 5.00
Begin VB.Form frmMaternity 
   BackColor       =   &H00808000&
   Caption         =   "Form4"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14490
   FillColor       =   &H00404000&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   14490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text21 
      Height          =   615
      Left            =   7560
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label22 
      Caption         =   "L.M.P :"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   21
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label21 
      Caption         =   "E.M.P :"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label20 
      Caption         =   "No of children alive:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   19
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label19 
      Caption         =   "No of children Dead:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      TabIndex        =   18
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label18 
      Caption         =   "Condition on Admission:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label17 
      Caption         =   "Admitted:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   16
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label16 
      Caption         =   "Discharge:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   15
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label15 
      Caption         =   "Readmitted:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label14 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   13
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label10 
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Occupation:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   7
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Parent Name's:"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Husband's Name"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Doctor"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MATERNITY REGISTRATION"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmMaternity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Text19_Change()

End Sub

