VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInput 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form4"
   MouseIcon       =   "frmInput.frx":0CCA
   ScaleHeight     =   8175
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C000&
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
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
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
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
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
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "FLUID INPUT RECORD"
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
      Height          =   3015
      Left            =   0
      TabIndex        =   13
      Top             =   3960
      Width           =   8895
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   6840
         Top             =   2280
      End
      Begin VB.TextBox txtVolume 
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
         Left            =   6600
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtDN 
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
         Left            =   6600
         TabIndex        =   22
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFluid 
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
         TabIndex        =   21
         Top             =   2040
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         CurrentDate     =   42906
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
         Left            =   1800
         TabIndex        =   19
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fluid Type:"
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
         TabIndex        =   18
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label5 
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
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor/Nurse:"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Volume:"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   600
         Width           =   1815
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   3135
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
      Left            =   -120
      TabIndex        =   0
      Top             =   600
      Width           =   9015
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         Left            =   6720
         TabIndex        =   4
         Top             =   1440
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
         Left            =   6720
         TabIndex        =   3
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
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
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
         ItemData        =   "frmInput.frx":1994
         Left            =   1800
         List            =   "frmInput.frx":199E
         TabIndex        =   1
         Top             =   1560
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404000&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   0
      Top             =   7080
      Width           =   8895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   8400
      Picture         =   "frmInput.frx":19B0
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT FLUID INPUT RECORDS"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   -240
      TabIndex        =   12
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim today As Variant
Dim str As String
Private Sub cmdDelete_Click()
If txtVolume.Text = "" Or txtDN = "" Or txtFluid.Text = "" Then
MsgBox "Empty Field(s),Please check your records!!!", vbCritical, "Error"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientInput", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Volume = txtVolume.Text
rs.Fields!Time = lblTime.Caption
rs.Fields!Fluid = txtFluid.Text
rs.Fields!DN = txtDN.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
MsgBox "Patient Input Fluid Deleted Successfully", vbInformation, "Deleted"
frmMain.Show
frmMain.Enabled = True
Else
frmMain.Show
frmMain.Enabled = True
c.Close
End If
End If
End Sub

Private Sub cmdSave_Click()
If txtVolume.Text = "" Or txtDN = "" Or txtFluid.Text = "" Then
MsgBox "Empty Field(s),Please check your records!!!", vbCritical, "Error"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientInput", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Volume = txtVolume.Text
lblTime.Caption = Format(Time$, "hh:mm:ss ampm")
rs.Fields!Time = lblTime.Caption
rs.Fields!Fluid = txtFluid.Text
rs.Fields!DN = txtDN.Text
rs.Update
Unload Me
MsgBox "Patient Input Fluid Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
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
MsgBox "Record Not Find", vbCritical, "error"
rs.MoveFirst
txtSurname.Text = ""
txtOthernames.Text = ""
cmbsex.Text = ""
txtaddress.Text = ""
cmdSave.Enabled = False
Else
txtSurname.Text = rs.Fields!Surname
txtOthernames.Text = rs.Fields!Othernames
cmbsex = rs.Fields!Sex
txtaddress.Text = rs.Fields!Address
cmdSave.Enabled = True
txtSurname.Enabled = False
txtOthernames.Enabled = False
cmbsex.Enabled = False
txtaddress.Enabled = False
txtRegno.Enabled = False
c.Close
End If
End If
End Sub

Private Sub cmdUpdate_Click()
If txtVolume.Text = "" Or txtDN = "" Or txtFluid.Text = "" Then
MsgBox "Empty Field(s),Please check your records!!!", vbCritical, "Error"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientInput", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Volume = txtVolume.Text
lblTime.Caption = Format(Time$, "hh:mm:ss ampm")
rs.Fields!Time = lblTime.Caption
rs.Fields!Fluid = txtFluid.Text
rs.Fields!DN = txtDN.Text
rs.Update
Unload Me
MsgBox "Patient Input Fluid Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
End If
End Sub

Private Sub Image1_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

Private Sub Timer1_Timer()
today = Now
lblTime.Caption = Format(today, "hh:mm:ss ampm")
End Sub

