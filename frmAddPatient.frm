VERSION 5.00
Begin VB.Form frmAddPatient 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   6825
   ClientTop       =   1530
   ClientWidth     =   9015
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFF80&
   Icon            =   "frmAddPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAddPatient.frx":0CCA
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9750
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReligion 
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
      Left            =   6480
      TabIndex        =   43
      Top             =   4440
      Width           =   2535
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
      Height          =   615
      Left            =   1800
      TabIndex        =   41
      Top             =   4440
      Width           =   2535
   End
   Begin VB.ComboBox cmbkinsex 
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
      ItemData        =   "frmAddPatient.frx":1994
      Left            =   1920
      List            =   "frmAddPatient.frx":199E
      TabIndex        =   39
      Top             =   6840
      Width           =   2415
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
      ItemData        =   "frmAddPatient.frx":19B0
      Left            =   1800
      List            =   "frmAddPatient.frx":19BA
      TabIndex        =   38
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox cmbkinmaritalstatus 
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
      ItemData        =   "frmAddPatient.frx":19CC
      Left            =   6480
      List            =   "frmAddPatient.frx":19D6
      TabIndex        =   37
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      Picture         =   "frmAddPatient.frx":19EB
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8880
      Width           =   2295
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      Picture         =   "frmAddPatient.frx":434F
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Picture         =   "frmAddPatient.frx":6B9E
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddPatient 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "frmAddPatient.frx":90E6
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8880
      Width           =   1815
   End
   Begin VB.ComboBox cmbmaritalstatus 
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
      ItemData        =   "frmAddPatient.frx":B77A
      Left            =   6480
      List            =   "frmAddPatient.frx":B784
      TabIndex        =   32
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox txtkin 
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
      ItemData        =   "frmAddPatient.frx":B799
      Left            =   1920
      List            =   "frmAddPatient.frx":B7AF
      TabIndex        =   31
      Top             =   8040
      Width           =   2415
   End
   Begin VB.TextBox txtkinaddress 
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
      TabIndex        =   27
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox txtkintelephone 
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
      Left            =   6480
      TabIndex        =   26
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox txtkinsurname 
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
      TabIndex        =   21
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txtkinothername 
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
      Left            =   6480
      TabIndex        =   20
      Top             =   6120
      Width           =   2535
   End
   Begin VB.TextBox txttribe 
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
      Left            =   6480
      TabIndex        =   18
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txtnationality 
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
      TabIndex        =   16
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txttelephone 
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
      Left            =   6480
      TabIndex        =   14
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtaddress 
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
      TabIndex        =   12
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtlga 
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
      Left            =   6480
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtstate 
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
      TabIndex        =   8
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtOthername 
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
      Left            =   6480
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtSurname 
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
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Religion:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   42
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   40
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Next of Kin:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   28
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   24
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Names:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   22
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PARENT /GUARDIAN DETAILS "
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
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   5280
      Width           =   9135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tribe:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   17
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   13
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "L.G.A"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "State of Origin"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Names:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname:"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATIENT ENTRY FORM"
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
      Height          =   735
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmAddPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rand As String
Function Random(Lowerbound As Long, Upperbound As Long)
Randomize
Random = Int(Rnd * Upperbound) + Lowerbound
End Function
Private Sub cmdAddPatient_Click()
On Error Resume Next
rand = "HMS/" & Year(Date$) & "/" & Random(1000, 1111)
If txtSurname.Text = "" Or txtOthername.Text = "" Or txtstate.Text = "" Or txtlga.Text = "" Or txtaddress.Text = "" Or txttelephone.Text = "" Or txtnationality.Text = "" Or txttribe.Text = "" Or txtOccupation.Text = "" Or txtReligion.Text = "" Or txtkinsurname.Text = "" Or txtkinothername.Text = "" Or txtkinaddress.Text = "" Or txtkintelephone.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthername.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!MaritalStatus = cmbmaritalstatus.Text
rs.Fields!State = txtstate.Text
rs.Fields!LGA = txtlga.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!PhoneNo = txttelephone.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Religion = txtReligion.Text
rs.Fields!Nationality = txtnationality.Text
rs.Fields!Tribe = txttribe.Text
rs.Fields!Psurname = txtkinsurname.Text
rs.Fields!Pothernames = txtkinothername.Text
rs.Fields!Psex = cmbkinsex.Text
rs.Fields!Pmaritalstatus = cmbkinmaritalstatus.Text
rs.Fields!Paddress = txtkinaddress.Text
rs.Fields!Pphoneno = txtkintelephone.Text
rs.Fields!Pnok = txtkin.Text
rs.Fields!reg_no = rand
rs.Update
MsgBox "Patient Added Successfully" & " " & rand, vbInformation, "Registered"
MsgBox "The Patient Registration Number is" & " " & rand & " " & " " & "                                  Do well to copy down the REGISTRATION NUMBER", vbInformation, "Registered"
Unload Me
frmAddPatient.Show
End If
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
If txtSurname.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthername.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!MaritalStatus = cmbmaritalstatus.Text
rs.Fields!State = txtstate.Text
rs.Fields!LGA = txtlga.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!PhoneNo = txttelephone.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Religion = txtReligion.Text
rs.Fields!Nationality = txtnationality.Text
rs.Fields!Tribe = txttribe.Text
rs.Fields!Psurname = txtkinsurname.Text
rs.Fields!Pothernames = txtkinothername.Text
rs.Fields!Psex = cmbkinsex.Text
rs.Fields!Pmaritalstatus = cmbkinmaritalstatus.Text
rs.Fields!Paddress = txtkinaddress.Text
rs.Fields!Pphoneno = txtkintelephone.Text
rs.Fields!Pnok = txtkin.Text
rs.Fields!reg_no = rand
rs.Delete
MsgBox "Patient Record deleted Succesfully", vbInformation, "Deleted"
Unload Me
frmMain.Show
frmMain.Enabled = True
cmdAddPatient.Enabled = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = True
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
If txtSurname.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.Fields!Surname = txtSurname.Text
rs.Fields!Othernames = txtOthername.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!MaritalStatus = cmbmaritalstatus.Text
rs.Fields!State = txtstate.Text
rs.Fields!LGA = txtlga.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!PhoneNo = txttelephone.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Religion = txtReligion.Text
rs.Fields!Nationality = txtnationality.Text
rs.Fields!Tribe = txttribe.Text
rs.Fields!Psurname = txtkinsurname.Text
rs.Fields!Pothernames = txtkinothername.Text
rs.Fields!Psex = cmbkinsex.Text
rs.Fields!Pmaritalstatus = cmbkinmaritalstatus.Text
rs.Fields!Paddress = txtkinaddress.Text
rs.Fields!Pphoneno = txtkintelephone.Text
rs.Fields!Pnok = txtkin.Text
rs.Update
Unload Me
frmAddPatient.Show
cmdAddPatient.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = True
End If
End Sub
