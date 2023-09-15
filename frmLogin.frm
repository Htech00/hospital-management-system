VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   585
   ClientTop       =   1725
   ClientWidth     =   15030
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   29.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":0CCA
   Picture         =   "frmLogin.frx":1994
   ScaleHeight     =   8880
   ScaleMode       =   0  'User
   ScaleWidth      =   14673.93
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   7170
      TabIndex        =   0
      Top             =   4000
      Width           =   6615
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00404000&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3600
         Width           =   6615
      End
      Begin VB.TextBox txtpassword 
         DataField       =   "password"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   735
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtusername 
         DataField       =   "username"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   735
         Left            =   3000
         TabIndex        =   4
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   6120
         Picture         =   "frmLogin.frx":74A64
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "LOGIN FORM"
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
         TabIndex        =   1
         Top             =   600
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset




Private Sub Command1_Click()

End Sub

Private Sub cmdLogin_Click()
On Error Resume Next
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from login", c, adOpenDynamic
If txtusername.Text = "" Or txtpassword.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
ElseIf rs.Fields(0) = txtusername.Text And rs.Fields(1) = txtpassword.Text Then
Unload Me
MsgBox "Login Successful", vbInformation, "WELCOME"
frmMain.Show
Else
MsgBox "Invalid username or password", vbCritical, "Error!!!"
End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
Frame1.Move (Screen.Width - Frame1.Width) / 2, (Screen.Height - Frame1.Height) / 2
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Label4.Left + Label1.Width < 0 Then
Label4.Left = Me.Width
End If
Label4.Left = Label4.Left - 70
End Sub

