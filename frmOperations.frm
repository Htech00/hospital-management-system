VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOperations 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8205
   ClientLeft      =   5070
   ClientTop       =   1530
   ClientWidth     =   9870
   Icon            =   "frmOperations.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmOperations.frx":0CCA
   ScaleHeight     =   8205
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00808000&
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00808000&
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
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00808000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      Caption         =   "OPERATION INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   0
      TabIndex        =   13
      Top             =   3960
      Width           =   9975
      Begin VB.TextBox txtOperation 
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
         Height          =   525
         Left            =   8040
         TabIndex        =   24
         Top             =   840
         Width           =   1815
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
         Height          =   645
         Left            =   4680
         TabIndex        =   17
         Top             =   2040
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         Top             =   840
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
         CalendarTitleBackColor=   16711935
         Format          =   40828929
         CurrentDate     =   42902
      End
      Begin VB.TextBox txtSurgeon 
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
         Height          =   525
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Operation:"
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
         Left            =   6480
         TabIndex        =   23
         Top             =   840
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   4
         Height          =   1095
         Left            =   1440
         Top             =   1800
         Width           =   6975
      End
      Begin VB.Label Label4 
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
         Height          =   495
         Left            =   2280
         TabIndex        =   19
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Surgeon:"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   840
         Width           =   1455
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
         Top             =   840
         Width           =   1215
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
      TabIndex        =   1
      Top             =   600
      Width           =   9975
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtOthernames 
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
         Left            =   6720
         TabIndex        =   4
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
         TabIndex        =   3
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
         ItemData        =   "frmOperations.frx":1994
         Left            =   1800
         List            =   "frmOperations.frx":199E
         TabIndex        =   2
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   9480
      Picture         =   "frmOperations.frx":19B0
      Top             =   0
      Width           =   450
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404000&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   7200
      Width           =   9975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT OPERATION INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frmOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub cmdDelete_Click()
On Error Resume Next
If txtSurname.Text = "" Or txtOthernames.Text = "" Or txtaddress.Text = "" Or txtRegno.Text = "" Or txtOperation.Text = "" Or txtSurgeon = "" Or txtCodeno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientOperations", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Surgeon = txtSurgeon.Text
rs.Fields!Operation = txtOperation.Text
rs.Fields!Codeno = txtCodeno.Text
rs.Fields!reg_no = txtRegno.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient  Operation Records Deleted Successfully", vbInformation, "Delete"
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
c.Close
End If
End If
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If txtSurname.Text = "" Or txtOthernames.Text = "" Or txtaddress.Text = "" Or txtRegno.Text = "" Or txtOperation.Text = "" Or txtSurgeon = "" Or txtCodeno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientOperations", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Surgeon = txtSurgeon.Text
rs.Fields!Operation = txtOperation.Text
rs.Fields!Codeno = txtCodeno.Text
rs.Fields!reg_no = txtRegno.Text
rs.Update
Unload Me
rs.Update
MsgBox "Patient Operation Saved Successfully", vbInformation, "Registered"
Unload Me
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
cmdAdd.Enabled = True
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
On Error Resume Next
If txtSurname.Text = "" Or txtOthernames.Text = "" Or txtaddress.Text = "" Or txtRegno.Text = "" Or txtOperation.Text = "" Or txtSurgeon = "" Or txtCodeno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientOperations", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthernames.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Surgeon = txtSurgeon.Text
rs.Fields!Operation = txtOperation.Text
rs.Fields!Codeno = txtCodeno.Text
rs.Fields!reg_no = txtRegno.Text
rs.Update
Unload Me
MsgBox "Patient  Operation Records Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
cmdSave.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = True
c.Close
End If
End Sub

Private Sub Image1_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

