VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHospitalHistory 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9600
   ClientLeft      =   5565
   ClientTop       =   1560
   ClientWidth     =   9210
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmHospitalHistory.frx":0000
   LinkTopic       =   "Form4"
   MouseIcon       =   "frmHospitalHistory.frx":0CCA
   ScaleHeight     =   9600
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "PATIENT HISTORY"
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
      Height          =   5415
      Left            =   -120
      TabIndex        =   12
      Top             =   4200
      Width           =   9855
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   6960
         TabIndex        =   32
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
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
         CurrentDate     =   42919
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
         Height          =   615
         Left            =   4800
         TabIndex        =   30
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "UPDATE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   28
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         TabIndex        =   27
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&ADD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtOthers 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   7200
         TabIndex        =   25
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtDischarge 
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
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   7200
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtWard 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   3120
         TabIndex        =   23
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtReferred 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   3120
         TabIndex        =   14
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtPhysician 
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   3120
         TabIndex        =   13
         Top             =   2280
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3120
         TabIndex        =   15
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
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
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   16711935
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   8421504
         Format          =   99352577
         CurrentDate     =   42902
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   9840
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Attended or Admitted"
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
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Referred by:"
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
         TabIndex        =   21
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ward/Clinic"
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
         TabIndex        =   20
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Discharge:"
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
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Physician/Surgeon:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Others"
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
         TabIndex        =   17
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Discharge to"
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
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
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
      TabIndex        =   2
      Top             =   840
      Width           =   9855
      Begin VB.ComboBox cmbSex 
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
         Left            =   1920
         TabIndex        =   31
         Top             =   1680
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
         TabIndex        =   29
         Top             =   2520
         Width           =   1695
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
         TabIndex        =   11
         Top             =   2520
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
         TabIndex        =   10
         Top             =   1560
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
         TabIndex        =   9
         Top             =   600
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
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   7695
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
         TabIndex        =   7
         Top             =   2520
         Width           =   1815
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   600
         Width           =   1935
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
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
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
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.TextBox Text9 
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT HOSPITAL HISTORY"
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
      TabIndex        =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmHospitalHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub cmdAdd_Click()
'On Error Resume Next
If txtReferred.Text = "" Or txtDischarge.Text = "" Or txtPhysician.Text = "" Or txtWard.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/HMS.mdb;Persist Security Info=false")
rs.Open "select * from PatientHistory", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Regno = txtRegno.Text
rs.Fields!Dateattended = DTPicker1.Value
rs.Fields!Datedischarge = DTPicker2.Value
rs.Fields!Referred = txtReferred.Text
rs.Fields!Dischargeto = txtDischarge.Text
rs.Fields!Physician = txtPhysician.Text
rs.Fields!Others = txtOthers.Text
rs.Fields!Ward = txtWard.Text
rs.Update
MsgBox "Patient history added  Successful", vbInformation, "Registered"
Unload Me
frmMain.Show
frmMain.Enabled = True
End If
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
'If txtSurname.Text = ""
'MsgBox "Required field(s) empty", vbCritical, "Error!!!"
'Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientHistory", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!Dateattended = DTPicker1.Value
rs.Fields!Datedischarge = DTPicker2.Value
rs.Fields!Dischargeto = txtDischarge.Text
rs.Fields!Referred = txtReferred.Text
rs.Fields!Physician = txtPhysician.Text
rs.Fields!Others = txtOthers.Text
rs.Fields!Ward = txtWard.Text
str = MsgBox("Are You Sure you want Delete", vbInformation + vbYesNo, "Delete")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient  Hospital History Records Deleted Successfully", vbInformation, "Deleted"
frmMain3.Show
frmMain.Enabled = True
cmdAdd.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
cmdExit.Enabled = True
Else
frmMain.Show
frmMain.Enabled = True
cmdAdd.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
cmdExit.Enabled = True
c.Close
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
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
txtaddress.Text = ""
cmdAdd.Enabled = False
Else
txtSurname.Text = rs.Fields!Surname
txtOthernames.Text = rs.Fields!Othernames
cmbsex.Text = rs.Fields!Sex
txtaddress.Text = rs.Fields!Address
cmdAdd.Enabled = True
c.Close
End If
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdUpdate_Click()
'On Error Resume Next
'If txtSurname.Text = ""
'MsgBox "Required field(s) empty", vbCritical, "Error!!!"
'Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientHistory", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!Dateattended = DTPicker1.Value
rs.Fields!Datedischarge = DTPicker2.Value
rs.Fields!Dischargeto = txtDischarge.Text
rs.Fields!Referred = txtReferred.Text
rs.Fields!Physician = txtPhysician.Text
rs.Fields!Others = txtOthers.Text
rs.Fields!Ward = txtWard.Text
rs.Update
Unload Me
MsgBox "Patient  Hospital History Records Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = True
cmdExit.Enabled = True
c.Close
'End If
End Sub

