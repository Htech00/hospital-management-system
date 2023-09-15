VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtemp 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9105
   ClientLeft      =   4530
   ClientTop       =   1530
   ClientWidth     =   10755
   Icon            =   "frmtemp.frx":0000
   LinkTopic       =   "Form4"
   MouseIcon       =   "frmtemp.frx":0CCA
   ScaleHeight     =   9105
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "TEMPERATURE INFO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3495
      Left            =   -240
      TabIndex        =   12
      Top             =   5640
      Width           =   11055
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   23
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   22
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   21
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox cmbHour 
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
         Height          =   405
         ItemData        =   "frmtemp.frx":1994
         Left            =   7680
         List            =   "frmtemp.frx":199E
         TabIndex        =   20
         Text            =   "Morning"
         Top             =   720
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   2880
         TabIndex        =   19
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   40239105
         CurrentDate     =   42905
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   40239105
         CurrentDate     =   42905
      End
      Begin VB.TextBox txtTemp 
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
         Left            =   7680
         TabIndex        =   17
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   240
         X2              =   11040
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         X1              =   240
         X2              =   11040
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Hour:"
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
         Left            =   5640
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature:"
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
         Left            =   5520
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Day Of Diseases:"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "PERSONAL INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
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
      Top             =   600
      Width           =   6375
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
         Left            =   3360
         TabIndex        =   24
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         TabIndex        =   11
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtChartno 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   3960
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
         Left            =   3360
         TabIndex        =   8
         Top             =   2040
         Width           =   2175
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
         Left            =   3360
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
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
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF80&
         BorderWidth     =   3
         X1              =   0
         X2              =   6360
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Chart No:"
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
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   1935
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
      Begin VB.Label Label5 
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
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   10320
      Picture         =   "frmtemp.frx":19B4
      Top             =   0
      Width           =   450
   End
   Begin VB.Label frmOperationsfrmOperationsfrmOperationsfrmOperationsfrmOperations 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TEMPERATURE RECORDS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   -360
      TabIndex        =   1
      Top             =   0
      Width           =   11175
   End
   Begin VB.Image frmOperationsfrmOperationsfrmOperationsfrmOperations 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   5055
      Left            =   6360
      Picture         =   "frmtemp.frx":3F5A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "frmtemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Private Sub Text6_Change()

End Sub

Private Sub txtSearch_Click()

End Sub
Private Sub cmdDelete_Click()
On Error Resume Next
If txtPatientname.Text = "" Or txtHusband.Text = "" Or txtaddress.Text = "" Or txtOccupation.Text = "" Or txtChartno.Text = "" Or cmbHour.Text = "" Or txtTemp.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/HMS.mdb;Persist Security Info=false")
rs.Open "select * from PatientTemperature", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Chartno = txtChartno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!DOD = DTPicker2.Value
rs.Fields!Hour = cmbHour.Text
rs.Fields!Temperature = txtTemp.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
MsgBox "Patient Temperature Record Deleted  Successfully", vbInformation, "Delete"
frmMain.Show
frmMain.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
Else
frmMain.Show
frmMain.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
End If
End If
End Sub

Private Sub cmdSave_Click()
'On Error Resume Next
If cmbHour.Text = "" Or txtTemp.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/HMS.mdb;Persist Security Info=false")
rs.Open "select * from PatientTemperature", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Chartno = txtChartno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!DOD = DTPicker2.Value
rs.Fields!Hour = cmbHour.Text
rs.Fields!Temperature = txtTemp.Text
rs.Update
Unload Me
MsgBox "Patient Temperature Record Save  Successful", vbInformation, "Registered"
frmMain.Show
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
txtaddress.Text = ""
txtOccupation.Text = ""
cmdSave.Enabled = False
Else
txtPatientname.Text = rs.Fields!Patientname
txtHusband.Text = rs.Fields!Husband
txtaddress.Text = rs.Fields!Address
txtOccupation.Text = rs.Fields!Occupation
frmtemp.txtPatientname.Enabled = False
frmtemp.txtHusband.Enabled = False
frmtemp.txtaddress.Enabled = False
frmtemp.txtOccupation.Enabled = False
frmtemp.txtChartno.Enabled = False
cmdSave.Enabled = True
c.Close
End If
End If
End Sub

Private Sub y_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
If txtPatientname.Text = "" Or txtHusband.Text = "" Or txtaddress.Text = "" Or txtOccupation.Text = "" Or txtChartno.Text = "" Or cmbHour.Text = "" Or txtTemp.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/HMS.mdb;Persist Security Info=false")
rs.Open "select * from PatientTemperature", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Chartno = txtChartno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!DOD = DTPicker2.Value
rs.Fields!Hour = cmbHour.Text
rs.Fields!Temperature = txtTemp.Text
rs.Update
MsgBox "Patient Temperature Record Updated  Successfully", vbInformation, "Updated"
Unload Me
frmMain.Show
frmMain.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = True
End If
End Sub

Private Sub Image1_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub
