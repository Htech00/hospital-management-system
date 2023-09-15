VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLabour 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9870
   ClientLeft      =   3660
   ClientTop       =   1470
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLabour.frx":0000
   LinkTopic       =   "Form4"
   MouseIcon       =   "frmLabour.frx":0CCA
   Picture         =   "frmLabour.frx":1994
   ScaleHeight     =   9870
   ScaleWidth      =   16335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   20
      Top             =   5400
      Width           =   18615
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   12600
         Top             =   2040
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2640
         TabIndex        =   44
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   16711935
         Format          =   38731777
         CurrentDate     =   42906
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00404000&
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00404000&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00404000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00404000&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtBP 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2640
         TabIndex        =   39
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtPulse 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2640
         TabIndex        =   38
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox txtUrine 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   8640
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtVE 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   8640
         TabIndex        =   36
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtContractions 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   8640
         TabIndex        =   35
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtPDR 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   13800
         TabIndex        =   34
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtMembranes 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   8640
         TabIndex        =   33
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtFH 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   13800
         TabIndex        =   32
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblTime 
         Caption         =   "hh:mm:ss"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2640
         TabIndex        =   46
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Urine:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   5280
         TabIndex        =   45
         Top             =   600
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1680
         Top             =   3480
         Width           =   12375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   120
         X2              =   18600
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Urine:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   30
         Top             =   480
         Width           =   15
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "F.H :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11160
         TabIndex        =   29
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Membranes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   5280
         TabIndex        =   28
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Vaginal Examination:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5280
         TabIndex        =   27
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Contractions:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5280
         TabIndex        =   26
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Piticon Drip Rate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11160
         TabIndex        =   25
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Pressure:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   360
         TabIndex        =   24
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Pulse:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "PERSONAL INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         TabIndex        =   19
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtChartno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   4200
         TabIndex        =   18
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txtNCD 
         Height          =   495
         Left            =   7560
         TabIndex        =   16
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2880
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtHusband 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2880
         TabIndex        =   14
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtNCA 
         Height          =   495
         Left            =   2880
         TabIndex        =   13
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtDoctor 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7560
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtPatientname 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7560
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtOccupation 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7560
         TabIndex        =   1
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ChartNo:"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1800
         TabIndex        =   17
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1440
         Top             =   3840
         Width           =   7935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5280
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Children Dead:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5280
         TabIndex        =   9
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5280
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Children Alive:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Husband's Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5280
         TabIndex        =   4
         Top             =   2040
         Width           =   2055
      End
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT LABOUR RECORDS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   0
      TabIndex        =   31
      Top             =   4800
      Width           =   18375
   End
End
Attribute VB_Name = "frmLabour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim today As Variant
Dim str As String
Private Sub Text5_Change()

End Sub

Private Sub cmdearch_Click()

End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
If txtUrine.Text = "" Or txtFH.Text = "" Or txtVE.Text = "" Or txtPDR.Text = "" Or txtBP.Text = "" Or txtContractions.Text = "" Or txtPulse.Text = "" Or txtMembranes.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from LabourRecords", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!VE = txtVE.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!PDR = txtPDR.Text
rs.Fields!Time = lblTime.Caption
rs.Fields!Age = txtAge.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!BP = txtBP.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Contractions = txtContractions.Text
rs.Fields!Pulse = txtPulse.Text
rs.Fields!Membranes = txtMembranes.Text
rs.Fields!Chartno = txtChartno.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient Labour Records Modified Successfully", vbInformation, "Delete"
Else
frmMain.Show
frmMain.Enabled = True
cmdSearch.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
c.Close
End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

Private Sub cmdSave_Click()
If txtPatientname.Text = "" Or txtNCA.Text = "" Or txtDoctor.Text = "" Or txtNCD.Text = "" Or txtPDR.Text = "" Or txtAge.Text = "" Or txtHusband.Text = "" Or txtOccupation.Text = "" Or txtFH.Text = "" Or txtVE.Text = "" Or txtaddress.Text = "" Or txtUrine.Text = "" Or txtContractions.Text = "" Or txtChartno.Text = "" Or txtMembranes.Text = "" Or txtPulse.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from LabourRecords", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!VE = txtVE.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Urine = txtUrine.Text
rs.Fields!FH = txtFH.Text
rs.Fields!PDR = txtPDR.Text
rs.Fields!Time = lblTime.Caption
rs.Fields!Age = txtAge.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!BP = txtBP.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Contractions = txtContractions.Text
rs.Fields!Pulse = txtPulse.Text
rs.Fields!Membranes = txtMembranes.Text
rs.Fields!Chartno = txtChartno.Text
rs.Update
MsgBox "Patient Labour Records Saved Successfully", vbInformation, "Record Saved"
Unload Me
frmMain.Show
End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If txtChartno.Text = "" Then
MsgBox "Required field(s) empty,Please check your REGISTRATION NUMBER", vbCritical, "Error!!!"
frmLabour.Show
txtChartno.SetFocus
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from MaternityRecords", c, adOpenDynamic, adLockOptimistic
'If txtRegno.Text <> rs.Fields!reg_no Then
rs.Find "Chartno ='" & txtChartno.Text & "'"
If rs.EOF Then
MsgBox "Record Not Find", vbExclamation, "error"
rs.MoveFirst
txtPatientname.Text = ""
txtDoctor.Text = ""
txtAge.Text = ""
txtaddress.Text = ""
txtHusband.Text = ""
txtOccupation.Text = ""
txtNCA.Text = ""
txtNCD.Text = ""
cmdSave.Enabled = False
Else
txtPatientname.Text = rs.Fields!Patientname
txtDoctor.Text = rs.Fields!Doctor
txtAge.Text = rs.Fields!Age
txtaddress.Text = rs.Fields!Address
txtHusband.Text = rs.Fields!Husband
txtOccupation.Text = rs.Fields!Occupation
txtNCA.Text = rs.Fields!NCA
txtNCD.Text = rs.Fields!NCD
txtPatientname.Enabled = False
txtDoctor.Enabled = False
txtAge.Enabled = False
txtaddress.Enabled = False
txtHusband.Enabled = False
txtOccupation.Enabled = False
txtNCA.Enabled = False
txtNCD.Enabled = False
cmdSave.Enabled = True
MsgBox "Chart Number found,Please Enter Your Labour Records", vbInformation, "Chart Number found!!!"
cmdSave.Enabled = True
DTPicker1.SetFocus
c.Close
End If
End If
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
If txtUrine.Text = "" Or txtFH.Text = "" Or txtVE.Text = "" Or txtPDR.Text = "" Or txtBP.Text = "" Or txtContractions.Text = "" Or txtPulse.Text = "" Or txtMembranes.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from LabourRecords", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!VE = txtVE.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!PDR = txtPDR.Text
lblTime.Caption = Format(Time$, "hh:mm:ss ampm")
rs.Fields!Time = lblTime.Caption
rs.Fields!Age = txtAge.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!BP = txtBP.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Contractions = txtContractions.Text
rs.Fields!Pulse = txtPulse.Text
rs.Fields!Membranes = txtMembranes.Text
rs.Fields!Chartno = txtChartno.Text
rs.Update
Unload Me
MsgBox "Patient Labour Records Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
cmdSearch.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
c.Close
End If
End Sub

Private Sub Timer1_Timer()
today = Now
lblTime.Caption = Format(today, "hh:mm:ss ampm")
End Sub

