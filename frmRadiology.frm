VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRadiology 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   Icon            =   "frmRadiology.frx":0000
   LinkTopic       =   "Form4"
   MouseIcon       =   "frmRadiology.frx":0CCA
   ScaleHeight     =   7275
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmRadiology.frx":1994
      Left            =   2280
      List            =   "frmRadiology.frx":199E
      TabIndex        =   24
      Top             =   2640
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
      Left            =   1800
      TabIndex        =   23
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C0C000&
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
      Left            =   3120
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C000&
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C000&
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C000&
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtXrayno 
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
      TabIndex        =   17
      Top             =   4800
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.TextBox txtName 
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
      Left            =   2280
      TabIndex        =   15
      Top             =   840
      Width           =   2175
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
      Height          =   615
      Left            =   2280
      TabIndex        =   14
      Top             =   1680
      Width           =   2175
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
      Height          =   615
      Left            =   7320
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtLMP 
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
      Left            =   7320
      TabIndex        =   12
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtCI 
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
      Left            =   2280
      TabIndex        =   11
      Top             =   3480
      Width           =   2175
   End
   Begin VB.ComboBox cmbExamination 
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
      ItemData        =   "frmRadiology.frx":19B0
      Left            =   7320
      List            =   "frmRadiology.frx":19FF
      TabIndex        =   10
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Regno :"
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
      Left            =   480
      TabIndex        =   22
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   9600
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   9600
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   1095
      Left            =   120
      Top             =   4560
      Width           =   9015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Clinical Information : "
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
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Examination Requested"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "L.M.P :"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
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
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex :"
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
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "X-Ray No :"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RADIOLOGY REQUEST FORM"
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
      Height          =   615
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmRadiology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim rand As String
Function Random(Lowerbound As Long, Upperbound As Long)
Randomize
Random = Int(Rnd * Upperbound) + Lowerbound
End Function
Private Sub Command1_Click()

End Sub

Private Sub cmdDelete_Click()
If txtName.Text = "" Or txtAge.Text = "" Or txtLMP.Text = "" Or cmbsex.Text = "" Or txtaddress.Text = "" Or txtCI.Text = "" Or cmbExamination.Text = "" Or txtRegno.Text = "" Or txtXrayno.Text = "" Then
MsgBox "Required field(s) Empty,Please check your records", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Radiology ", c, adOpenDynamic, adLockOptimistic
rs.Fields!Name = txtName.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Age = txtAge.Text
rs.Fields!LMP = txtLMP.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!CI = txtCI.Text
rs.Fields!Examination = cmbExamination.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Xrayno = txtXrayno.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient radiological Request  Deleted successfully", vbInformation, "Delete"
frmMain.Show
frmMain.Enabled = True
Else
frmMain.Show
frmMain.Enabled = True
End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If txtName.Text = "" Or txtAge.Text = "" Or txtLMP.Text = "" Or cmbsex.Text = "" Or txtaddress.Text = "" Or txtCI.Text = "" Or cmbExamination.Text = "" Or txtRegno.Text = "" Or txtXrayno.Text = "" Then
MsgBox "Required field(s) Empty,Please check your records", vbCritical, "Error!!!"
ElseIf Not IsNumeric(txtXrayno.Text) Then
MsgBox "Xray Field Must Be a Number", vbExclamation, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Radiology ", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Name = txtName.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Age = txtAge.Text
rs.Fields!LMP = txtLMP.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!CI = txtCI.Text
rs.Fields!Examination = cmbExamination.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Xrayno = txtXrayno.Text
rs.Update
MsgBox "Patient radiological Request Saved successfully", vbInformation, "Saved"
frmRadiology.Show
txtName.Text = ""
txtAge.Text = ""
txtLMP.Text = ""
cmbsex.Text = ""
txtaddress.Text = ""
txtCI.Text = ""
cmbExamination.Text = ""
txtRegno.Text = ""
txtXrayno.Text = ""
Form3.Enabled = True
End If
c.Close
End Sub

Private Sub cmdUpdate_Click()
If txtName.Text = "" Or txtAge.Text = "" Or txtLMP.Text = "" Or cmbsex.Text = "" Or txtaddress.Text = "" Or txtCI.Text = "" Or cmbExamination.Text = "" Or txtRegno.Text = "" Or txtXrayno.Text = "" Then
MsgBox "Required field(s) Empty,Please check your records", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Radiology ", c, adOpenDynamic, adLockOptimistic
rs.Fields!Name = txtName.Text
rs.Fields!Date = DTPicker1.Value
rs.Fields!Age = txtAge.Text
rs.Fields!LMP = txtLMP.Text
rs.Fields!Sex = cmbsex.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!CI = txtCI.Text
rs.Fields!Examination = cmbExamination.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Xrayno = txtXrayno.Text
rs.Update
Unload Me
MsgBox "Patient radiological Request Updated successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
End If
End Sub

