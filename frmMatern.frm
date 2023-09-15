VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMatern 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8400
   ClientLeft      =   2505
   ClientTop       =   1845
   ClientWidth     =   16830
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMatern.frx":0000
   LinkTopic       =   "Form4"
   MouseIcon       =   "frmMatern.frx":0CCA
   Picture         =   "frmMatern.frx":1994
   ScaleHeight     =   8400
   ScaleWidth      =   16830
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6120
      Top             =   4560
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
      Left            =   2160
      TabIndex        =   37
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1455
      Left            =   360
      TabIndex        =   34
      Top             =   6600
      Width           =   6015
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   615
         Left            =   4200
         TabIndex        =   44
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   615
         Left            =   360
         TabIndex        =   43
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update"
         Enabled         =   0   'False
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      ForeColor       =   &H8000000E&
      Height          =   8415
      Left            =   6600
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   495
         Left            =   2400
         TabIndex        =   41
         Top             =   4680
         Width           =   2535
         _ExtentX        =   4471
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
         CalendarTitleBackColor=   4210688
         Format          =   40239105
         CurrentDate     =   42905
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   7680
         TabIndex        =   40
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
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
         Left            =   7680
         TabIndex        =   39
         Top             =   3120
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.TextBox txtDoctor 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   33
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtPatientname 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   32
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtEMP 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   31
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtCOA 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   30
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox txtAge 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   29
         Top             =   5400
         Width           =   2535
      End
      Begin VB.TextBox txtHusband 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   28
         Top             =   6120
         Width           =   2535
      End
      Begin VB.TextBox txtNCA 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7680
         TabIndex        =   27
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtNCD 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7680
         TabIndex        =   26
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtOccupation 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7680
         TabIndex        =   25
         Top             =   5400
         Width           =   2415
      End
      Begin VB.TextBox txtLMP 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7680
         TabIndex        =   24
         Top             =   6120
         Width           =   2415
      End
      Begin VB.TextBox txtReligion 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7680
         TabIndex        =   23
         Top             =   6840
         Width           =   2415
      End
      Begin VB.TextBox txtParameters 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   7680
         TabIndex        =   22
         Top             =   7560
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   21
         Top             =   6840
         Width           =   2535
      End
      Begin VB.TextBox txtEDD 
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
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   2400
         TabIndex        =   20
         Top             =   7560
         Width           =   2535
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   9600
         Picture         =   "frmMatern.frx":D51C
         Top             =   120
         Width           =   450
      End
      Begin VB.Label lblTime 
         Caption         =   "hh:mm:ss "
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
         Height          =   375
         Left            =   7680
         TabIndex        =   42
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Children Dead"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   19
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Admitted"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   18
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Discharge"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   17
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Registered:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   16
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "L.M.P :"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   15
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   14
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Religion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   13
         Top             =   6960
         Width           =   2055
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters:"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   5520
         TabIndex        =   12
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "E.D.D :"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MATERNITY RECORDS"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   10215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Children Alive"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5520
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Husband's Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Readmitted:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Conditions of Admission"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "E.M.P :"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   1695
      End
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No:"
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
      Left            =   600
      TabIndex        =   36
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   360
      Top             =   5040
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404000&
      BorderWidth     =   4
      Height          =   3735
      Left            =   0
      Top             =   4680
      Width           =   6855
   End
End
Attribute VB_Name = "frmMatern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim chart As String
Dim today As Variant
Dim str As String
Private Sub Text1_Change()
End Sub

Private Sub Text4_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Label12_Click()

End Sub
Function Random(Lowerbound As Long, Upperbound As Long)
Randomize
Random = Int(Rnd * Upperbound) + Lowerbound
End Function

Private Sub cmdDelete_Click()
On Error Resume Next
If txtPatientname.Text = "" Or txtNCA.Text = "" Or txtDoctor.Text = "" Or txtNCD.Text = "" Or txtEMP.Text = "" Or txtCOA.Text = "" Or txtAge.Text = "" Or txtHusband.Text = "" Or txtOccupation.Text = "" Or txtReligion.Text = "" Or txtLMP.Text = "" Or txtaddress.Text = "" Or txtEDD.Text = "" Or txtParameters.Text = "" Or txtRegno.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from MaternityRecords", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!EMP = txtEMP.Text
rs.Fields!Admitted = DTPicker1.Value
rs.Fields!COA = txtCOA.Text
rs.Fields!Discharge = DTPicker2.Value
rs.Fields!Readmitted = DTPicker3.Value
rs.Fields!Time = Time$
rs.Fields!Age = txtAge.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!LMP = txtLMP.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Religion = txtReligion.Text
rs.Fields!EDD = txtEDD.Text
rs.Fields!Parameters = txtParameters.Text
rs.Fields!reg_no = txtRegno.Text
str = MsgBox("Are you sure you want to delete patient records", vbInformation + vbYesNo, "Caution")
If str = vbYes Then
rs.Delete
Unload Me
MsgBox "Patient  Maternity Records Deleted Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
cmdSearch.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
Else
frmMain.Show
frmMain.Enabled = True
cmdSearch.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = True
cmdUpdate.Enabled = False
c.Close
c.Close
End If
End If
End Sub

Private Sub cmdSave_Click()
chart = "MC/" & Year(Date$) & "/" & Random(1000, 1800)
If txtNCA.Text = "" Or txtDoctor.Text = "" Or txtNCD.Text = "" Or txtEMP.Text = "" Or txtCOA.Text = "" Or txtAge.Text = "" Or txtOccupation.Text = "" Or txtHusband.Text = "" Or txtLMP.Text = "" Or txtReligion.Text = "" Or txtaddress.Text = "" Or txtEDD.Text = "" Or txtParameters.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  MaternityRecords", c, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!EMP = txtEMP.Text
rs.Fields!Admitted = DTPicker1.Value
rs.Fields!COA = txtCOA.Text
rs.Fields!Discharge = DTPicker2.Value
rs.Fields!Readmitted = DTPicker3.Value
rs.Fields!Time = lblTime.Caption
rs.Fields!Age = txtAge.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!LMP = txtLMP.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Religion = txtReligion.Text
rs.Fields!EDD = txtEDD.Text
rs.Fields!Parameters = txtParameters.Text
rs.Fields!reg_no = txtRegno.Text
rs.Fields!Chartno = chart
rs.Update
MsgBox "Patient Records Saved Successfully" & " " & chart, vbInformation, "Registered"
MsgBox "The Patient Chart Number is" & " " & chart & " " & " " & "                                                   Do well to copy down the CHART NUMBER", vbInformation, "Registered"
Unload Me
frmMatern.Show
cmdSave.Enabled = True
End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If txtRegno.Text = "" Then
MsgBox "Required field(s) empty,Please check your REGISTRATION NUMBER", vbCritical, "Error!!!"
frmMatern.Show
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
'If txtRegno.Text <> rs.Fields!reg_no Then
rs.Find "reg_no ='" & txtRegno.Text & "'"
If rs.EOF Then
MsgBox "Record Not Find", vbExclamation, "error"
rs.MoveFirst
txtPatientname.Text = ""
cmdSave.Enabled = False
Else
txtPatientname.Text = rs.Fields!Surname
MsgBox "Registration Number found,Please Enter Your Records", vbInformation, "Registration Number found!!!"
cmdSave.Enabled = True
txtNCA.SetFocus
c.Close
End If
End If
End Sub

Private Sub cmdUpdate_Click()
'On Error Resume Next
If txtPatientname.Text = "" Or txtNCA.Text = "" Or txtDoctor.Text = "" Or txtNCD.Text = "" Or txtEMP.Text = "" Or txtCOA.Text = "" Or txtAge.Text = "" Or txtHusband.Text = "" Or txtOccupation.Text = "" Or txtReligion.Text = "" Or txtLMP.Text = "" Or txtaddress.Text = "" Or txtEDD.Text = "" Or txtParameters.Text = "" Then
MsgBox "Required field(s) empty", vbCritical, "Error!!!"
Else
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from MaternityRecords", c, adOpenDynamic, adLockOptimistic
rs.Fields!Patientname = txtPatientname.Text
rs.Fields!NCA = txtNCA.Text
rs.Fields!Doctor = txtDoctor.Text
rs.Fields!NCD = txtNCD.Text
rs.Fields!EMP = txtEMP.Text
rs.Fields!Admitted = DTPicker1.Value
rs.Fields!COA = txtCOA.Text
rs.Fields!Discharge = DTPicker2.Value
rs.Fields!Readmitted = DTPicker3.Value
lblTime.Caption = Format(Time$, "hh:mm:ss ampm")
rs.Fields!Time = lblTime.Caption
rs.Fields!Age = txtAge.Text
rs.Fields!Occupation = txtOccupation.Text
rs.Fields!Husband = txtHusband.Text
rs.Fields!LMP = txtLMP.Text
rs.Fields!Address = txtaddress.Text
rs.Fields!Religion = txtReligion.Text
rs.Fields!EDD = txtEDD.Text
rs.Fields!Parameters = txtParameters.Text
rs.Update
Unload Me
MsgBox "Patient  Maternity Records Modified Successfully", vbInformation, "Updated"
frmMain.Show
frmMain.Enabled = True
cmdSearch.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = True
c.Close
End If
End Sub

Private Sub cndSave_Click()

End Sub

Private Sub Image3_Click()
Unload Me
frmMain.Show
frmMain.Enabled = True
End Sub

Private Sub Timer1_Timer()
today = Now
lblTime.Caption = Format(today, "hh:mm:ss ampm")
End Sub

