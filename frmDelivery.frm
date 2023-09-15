VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDelivery 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   10245
   ClientLeft      =   3975
   ClientTop       =   1365
   ClientWidth     =   16110
   Icon            =   "frmDelivery.frx":0000
   MouseIcon       =   "frmDelivery.frx":0CCA
   Picture         =   "frmDelivery.frx":1994
   ScaleHeight     =   10245
   ScaleWidth      =   16110
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Height          =   6375
      Left            =   0
      TabIndex        =   21
      Top             =   3840
      Width           =   16095
      Begin VB.ComboBox cmbSex 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmDelivery.frx":B1E2
         Left            =   13920
         List            =   "frmDelivery.frx":B1EC
         TabIndex        =   61
         Top             =   2160
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2520
         TabIndex        =   60
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   99352577
         CurrentDate     =   42917
      End
      Begin VB.CommandButton cmdUpdate 
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
         Left            =   5400
         TabIndex        =   59
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton cmdDelete 
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
         Left            =   7680
         TabIndex        =   58
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   10440
         TabIndex        =   57
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   2640
         TabIndex        =   56
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox txtDN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox txtHOL 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   53
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtMR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   52
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtTIB 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   51
         Top             =   3960
         Width           =   2175
      End
      Begin VB.TextBox txtTD 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   50
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtPD 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   49
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtPME 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   48
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtPL 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   47
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtBLA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   46
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtSutures 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13920
         TabIndex        =   45
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtBPAD 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13920
         TabIndex        =   44
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtMidwife 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13920
         TabIndex        =   43
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtCAB 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13920
         TabIndex        =   42
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox txtNTC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   41
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox txtWeight 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   40
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox txtTLS 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   2040
         Top             =   5400
         Width           =   11295
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctors Notes:"
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
         Left            =   11040
         TabIndex        =   54
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Hours Of Labour:"
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
         Height          =   975
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "TimeAt Which Labour Starts:"
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
         Height          =   975
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Note Of The Case:"
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
         Left            =   5160
         TabIndex        =   37
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Midwife in Charge:"
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
         Height          =   495
         Left            =   11040
         TabIndex        =   36
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Infant Born:"
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
         TabIndex        =   35
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Infant Born:"
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
         Left            =   240
         TabIndex        =   34
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Delivery:"
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
         Height          =   495
         Left            =   5160
         TabIndex        =   33
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Placenta Delivery:"
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
         Height          =   495
         Left            =   5160
         TabIndex        =   32
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Placenta and Membrane Examined:"
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
         Left            =   5160
         TabIndex        =   31
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Perineal Laceration:"
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
         Height          =   495
         Left            =   5160
         TabIndex        =   30
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Sutures:"
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
         Height          =   495
         Left            =   11040
         TabIndex        =   29
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "BP After Delivery:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   11040
         TabIndex        =   28
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Loss Approx:"
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
         Height          =   495
         Left            =   5160
         TabIndex        =   27
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
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
         Height          =   495
         Left            =   11040
         TabIndex        =   26
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight:"
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
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Condition At Birth:"
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
         Height          =   495
         Left            =   11040
         TabIndex        =   24
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Membrane Rapture:"
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
         TabIndex        =   22
         Top             =   2160
         Width           =   1935
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
      Height          =   3375
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      Begin VB.TextBox txtOccupation 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtPatientname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtDoctor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtNCA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtHusband 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtNCD 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtChartno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   2520
         Width           =   2175
      End
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
         Left            =   6240
         TabIndex        =   1
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         Left            =   1680
         TabIndex        =   19
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
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
         Left            =   4800
         TabIndex        =   18
         Top             =   1680
         Width           =   2055
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
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
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Children Alive"
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
         Left            =   9000
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor"
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
         Left            =   4920
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Children Dead"
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
         Left            =   9000
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
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
         Left            =   4920
         TabIndex        =   12
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's Name"
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
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         FillColor       =   &H00404000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1320
         Top             =   2400
         Width           =   7095
      End
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   0
      Picture         =   "frmDelivery.frx":B1FE
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT DELIVERY INFORMATION"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   -120
      TabIndex        =   20
      Top             =   3360
      Width           =   16215
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Private Sub cmdDelete_Click()
On Error Resume Next
If txtTLS.Text = "" Or txtTD.Text = "" Or txtSutures = "" Or txtHOL.Text = "" Or txtPD.Text = "" Or txtBPAD.Text = "" Or txtMR.Text = "" Or txtPME.Text = "" Or cmbsex.Text = "" Or txtPL.Text = "" Or txtMidwife.Text = "" Or txtTIB.Text = "" Or txtCAB.Text = "" Or txtWeight.Text = "" Or txtNTC.Text = "" Or txtDN.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Delivery", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Age = txtAge.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Chartno = txtChartno.Text
rs.Fields!TLS = txtTLS.Text
rs.Fields!TD = txtTD.Text
rs.Fields!Sutures = txtSutures.Text
rs.Fields!HOL = txtHOL.Text
rs.Fields!PD = txtPD.Text
rs.Fields!BPAD = txtBPAD.Text
rs.Fields!MR = txtMR.Text
rs.Fields!PME = txtPME.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!DIB = DTPicker1.Value
rs.Fields!PL = txtPL.Text
rs.Fields!Midwife = txtMidwife.Text
rs.Fields!TIB = txtTIB.Text
rs.Fields!BLA = txtBLA.Text
rs.Fields!CAB = txtCAB.Text
rs.Fields!Weight = txtWeight.Text
rs.Fields!NTC = txtNTC.Text
rs.Fields!DN = txtDN.Text
str = MsgBox("Are you sure you want to delete these records", vbCritical + vbYesNo, "Caution!!!")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient Labour Records Deleted Successful", vbInformation, "Update"
frmMain.Show
frmMain.Enabled = True
Else
frmMain.Show
frmMain.Enabled = True
End If
End If
c.Close
End Sub

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If txtTLS.Text = "" Or txtTD.Text = "" Or txtSutures.Text = "" Or txtHOL.Text = "" Or txtPD.Text = "" Or txtBPAD.Text = "" Or txtMR.Text = "" Or txtPME.Text = "" Or cmbsex.Text = "" Or txtPL.Text = "" Or txtMidwife.Text = "" Or txtTIB.Text = "" Or txtBLA.Text = "" Or txtCAB.Text = "" Or txtWeight.Text = "" Or txtNTC.Text = "" Then
MsgBox "Required Field(s) Empty,Please check your Records", vbCritical, "error"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Delivery ", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Age = txtAge.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Chartno = txtChartno.Text
rs.Fields!TLS = txtTLS.Text
rs.Fields!TD = txtTD.Text
rs.Fields!Sutures = txtSutures.Text
rs.Fields!HOL = txtHOL.Text
rs.Fields!PD = txtPD.Text
rs.Fields!BPAD = txtBPAD.Text
rs.Fields!MR = txtMR.Text
rs.Fields!PME = txtPME.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!DIB = DTPicker1.Value
rs.Fields!PL = txtPL.Text
rs.Fields!Midwife = txtMidwife.Text
rs.Fields!TIB = txtTIB.Text
rs.Fields!BLA = txtBLA.Text
rs.Fields!CAB = txtCAB.Text
rs.Fields!Weight = txtWeight.Text
rs.Fields!NTC = txtNTC.Text
rs.Fields!DN = txtDN.Text
rs.Update
Unload Me
MsgBox "Patient Labour Records Saved Successful", vbInformation, "Save"
frmMain.Show
frmMain.Enabled = True
End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If txtChartno.Text = "" Then
MsgBox "Required field(s) empty,Please check your REGISTRATION NUMBER", vbCritical, "Error!!!"
frmDelivery.Show
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
txtTLS.SetFocus
c.Close
End If
End If
End Sub

Private Sub cmdUpdate_Click()
If txtTLS.Text = "" Or txtTD.Text = "" Or txtSutures = "" Or txtHOL.Text = "" Or txtPD.Text = "" Or txtBPAD.Text = "" Or txtMR.Text = "" Or txtPME.Text = "" Or cmbsex.Text = "" Or txtPL.Text = "" Or txtMidwife.Text = "" Or txtTIB.Text = "" Or txtCAB.Text = "" Or txtWeight.Text = "" Or txtNTC.Text = "" Or txtDN.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Delivery", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Age = txtAge.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Chartno = txtChartno.Text
rs.Fields!TLS = txtTLS.Text
rs.Fields!TD = txtTD.Text
rs.Fields!Sutures = txtSutures.Text
rs.Fields!HOL = txtHOL.Text
rs.Fields!PD = txtPD.Text
rs.Fields!BPAD = txtBPAD.Text
rs.Fields!MR = txtMR.Text
rs.Fields!PME = txtPME.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!DIB = DTPicker1.Value
rs.Fields!PL = txtPL.Text
rs.Fields!Midwife = txtMidwife.Text
rs.Fields!TIB = txtTIB.Text
rs.Fields!BLA = txtBLA.Text
rs.Fields!CAB = txtCAB.Text
rs.Fields!Weight = txtWeight.Text
rs.Fields!NTC = txtNTC.Text
rs.Fields!DN = txtDN.Text
rs.Update
Unload Me
MsgBox "Patient Labour Records Updated Successful", vbInformation, "Update"
frmMain.Show
frmMain.Enabled = True
End If
End Sub
