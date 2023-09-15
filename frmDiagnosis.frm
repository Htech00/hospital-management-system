VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiagnosis 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "frmDiagnosis.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDiagnosis.frx":0CCA
   ScaleHeight     =   7980
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   1440
      TabIndex        =   21
      Top             =   4320
      Width           =   2295
      _ExtentX        =   4048
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
      Format          =   99352577
      CurrentDate     =   42902
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Picture         =   "frmDiagnosis.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   18
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   17
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtCodeno 
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
      Height          =   615
      Left            =   4560
      TabIndex        =   16
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txtDiagnosis 
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
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Caption         =   "PATIENT INFORMATION"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.ComboBox cmbsex 
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
         ItemData        =   "frmDiagnosis.frx":4028
         Left            =   1800
         List            =   "frmDiagnosis.frx":4032
         TabIndex        =   20
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtSurname 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   615
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtOthernames 
         Enabled         =   0   'False
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
         Height          =   615
         Left            =   6840
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
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
         Height          =   615
         Left            =   6840
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtRegno 
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
         Height          =   615
         Left            =   3240
         TabIndex        =   2
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   1
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   9240
         Picture         =   "frmDiagnosis.frx":4044
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label11 
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Othernames:"
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
         Left            =   4440
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label13 
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
         Left            =   4560
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "RegNo"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1200
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   7695
      End
   End
   Begin VB.Label Label4 
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
      Left            =   0
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   0
      X2              =   9600
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   0
      X2              =   9600
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   1215
      Left            =   1320
      Top             =   5280
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Code No:"
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
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosis:"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      BorderWidth     =   3
      FillColor       =   &H00FFFF00&
      Height          =   4215
      Left            =   -120
      Top             =   4080
      Width           =   9855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT DIAGNOSIS INFORMATION"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   9615
   End
End
Attribute VB_Name = "frmDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str  As String
Private Sub Command2_Click()

End Sub
Private Sub cmdAdd_Click()
'On Error Resume Next
If txtDiagnosis.Text = "" Or txtCodeno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/HMS.mdb;Persist Security Info=false")
rs.Open "select * from PatienTDiagnosisHistory", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Diagnosis = txtDiagnosis.Text
rs.Fields!Codeno = txtCodeno.Text
rs.Update
MsgBox "Patient Diagnosis Added  Successful", vbInformation, "Registered"
Unload Me
frmMain.Show
frmMain.Enabled = True
End If

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cmdDelete_Click()
'On Error Resume Next
If txtDiagnosis.Text = "" Or txtCodeno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientDiagnosisHistory", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Diagnosis = txtDiagnosis.Text
rs.Fields!Codeno = txtCodeno.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient Diagnosis Record  Deleted ", vbInformation, "Deleted"
frmMain.Show
frmMain.Enabled = True
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
Else
frmMain.Show
frmMain.Enabled = True
cmdAddPatient.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
c.Close
End If
End If
End Sub


Private Sub cmdSearch_Click()
On Error Resume Next
If txtRegno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
'If txtRegno.Text <> rs.Fields!reg_no Then
rs.Find "reg_no ='" & txtRegno.Text & "'"
If rs.EOF Then
MsgBox "Record Not Find", vbExclamation, "error"
rs.MoveFirst
txtSurname.Text = ""
txtOthernames.Text = ""
cmbsex.Text = ""
txtaddress.Text = ""
cmdAdd.Enabled = False
Else
txtSurname.Text = rs.Fields!Surname
txtOthernames.Text = rs.Fields!Othernames
cmbsex = rs.Fields!Sex
txtaddress.Text = rs.Fields!Address
cmdAdd.Enabled = True
c.Close
End If
End If
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
If txtDiagnosis.Text = "" Or txtCodeno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientDiagnosisHistory", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Diagnosis = txtDiagnosis.Text
rs.Fields!Codeno = txtCodeno.Text
rs.Update
Unload Me
MsgBox "Patient Record Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
cmdAddPatient.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = True
c.Close
End If
End Sub

Private Sub Image3_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

