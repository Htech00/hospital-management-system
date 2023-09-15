VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPressure 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9630
   ClientLeft      =   3915
   ClientTop       =   1380
   ClientWidth     =   12030
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "TEMPERATURE INFO"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   4455
      Left            =   0
      TabIndex        =   13
      Top             =   5760
      Width           =   12015
      Begin VB.TextBox txtPulse 
         Height          =   495
         Left            =   9000
         TabIndex        =   20
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox cmbHour 
         Height          =   315
         ItemData        =   "frmPressure.frx":0000
         Left            =   9000
         List            =   "frmPressure.frx":000A
         TabIndex        =   17
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   15
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   14
         Top             =   2640
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Format          =   48824321
         CurrentDate     =   42905
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2880
         TabIndex        =   19
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         Format          =   48824321
         CurrentDate     =   42905
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
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
         TabIndex        =   24
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Day Of Diseases:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
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
         TabIndex        =   23
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Pulse:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6600
         TabIndex        =   22
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Hour:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6720
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   0
         X2              =   12000
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   12000
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "PERSONAL INFORMATION"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txtPatientname 
         Height          =   495
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtHusband 
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtAddress 
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtChartno 
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         TabIndex        =   2
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtOccupation 
         Height          =   495
         Left            =   3360
         TabIndex        =   1
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Husband's Name:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3720
         Width           =   6135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Chart No:"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF80&
         BorderWidth     =   3
         X1              =   0
         X2              =   6360
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   11520
      Picture         =   "frmPressure.frx":0021
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PULSE RECORDS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   6480
      Picture         =   "frmPressure.frx":25C7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5535
   End
End
Attribute VB_Name = "frmPressure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmdSave_Click()
'On Error Resume Next
If cmbHour.Text = "" Or txtPressure.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/HMS.mdb;Persist Security Info=false")
rs.Open "select * from PatientPressure", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Address = txtAddress.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!ChartNo = txtChartno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!DOD = DTPicker2.Value
rs.Fields!Hour = cmbHour.Text
rs.Fields!Pressure = txtPressure.Text
rs.Update
MsgBox "Patient Pressure Record Save  Successful", vbInformation, "Registered"
Unload Me
Form3.Show
End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If txtChartno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from MaternityRecords", c, adOpenDynamic, adLockOptimistic
'If txtRegno.Text <> rs.Fields!reg_no Then
rs.Find "ChartNo ='" & txtChartno.Text & "'"
If rs.EOF Then
MsgBox "Record Not Find", vbExclamation, "error"
rs.MoveFirst
txtPatientname.Text = ""
txtHusband.Text = ""
txtAddress.Text = ""
txtOccupation.Text = ""
cmdSave.Enabled = False
Else
txtPatientname.Text = rs.Fields!Patientname
txtHusband.Text = rs.Fields!Husband
txtAddress.Text = rs.Fields!Address
txtOccupation.Text = rs.Fields!Occupation
cmdSave.Enabled = True
c.Close
End If
End If
End Sub

Private Sub Image2_Click()
Unload Me
Form3.Show
Form3.Enabled = True
End Sub
