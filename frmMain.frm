VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HOSPITAL MANAGEMENT SYSTEM"
   ClientHeight    =   7860
   ClientLeft      =   2580
   ClientTop       =   1230
   ClientWidth     =   14745
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   MouseIcon       =   "frmMain.frx":0CCA
   Picture         =   "frmMain.frx":1994
   ScaleHeight     =   7860
   ScaleWidth      =   14745
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   7680
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3360
      Top             =   3240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   -360
      TabIndex        =   1
      Top             =   4560
      Width           =   3975
      Begin VB.Label lbltime 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
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
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14415
   End
   Begin VB.Menu mnufile 
      Caption         =   "FILE"
      Begin VB.Menu mnuModityPatientRecord 
         Caption         =   "Modify Patient Records"
      End
      Begin VB.Menu mnuDeletePatientRecord 
         Caption         =   "Delete Patient Records"
      End
      Begin VB.Menu mnuPatientdiagnosishospital 
         Caption         =   " Modify Patient Diagnosis "
         Begin VB.Menu mnuupdatepatientdiagnosis 
            Caption         =   "Update Patient Diagnosis"
         End
         Begin VB.Menu mnuDeletepatientdiagnosis 
            Caption         =   "Delete Patient Diagnosis"
         End
      End
      Begin VB.Menu mnuModifypatienthospitalhistory 
         Caption         =   "Modify Patient Hospital History"
         Begin VB.Menu mnuUpdatepatienthospitalhistory 
            Caption         =   "Update Patient Hospital History"
         End
         Begin VB.Menu mnuDeletepatienthospitalhistory 
            Caption         =   "Delete Patient Hospital History"
         End
      End
      Begin VB.Menu mnumodifymaternitycharts 
         Caption         =   "Modify Maternity Charts"
         Begin VB.Menu mnuModifymaternityrecords 
            Caption         =   "Modify Maternity Records"
            Begin VB.Menu mnuUpdatematernityrecords 
               Caption         =   "Update Maternity Records"
            End
            Begin VB.Menu mnuDeletematernityrecords 
               Caption         =   "Delete Maternity Records"
            End
         End
         Begin VB.Menu mnuModifyTemperaturerecords 
            Caption         =   "Modify Temperature Records"
            Begin VB.Menu mnuUpdatetemperaturerecords 
               Caption         =   "Update Temperature Records"
            End
            Begin VB.Menu mnuDeletetemperaturerecords 
               Caption         =   "Delete Temperature Records"
            End
         End
         Begin VB.Menu mnuModifypulserecords 
            Caption         =   "Modify Pulse Records"
            Begin VB.Menu mnuUpdatepulserecords 
               Caption         =   "Update Pulse Records"
            End
            Begin VB.Menu mnuDeletepulserecords 
               Caption         =   "Delete Pulse Records"
            End
         End
         Begin VB.Menu mnuModifylabourrecords 
            Caption         =   "Modify Labour Records"
            Begin VB.Menu mnuUpdatelabourrecords 
               Caption         =   "Update Labour Records"
            End
            Begin VB.Menu mnuDeletelabourrecords 
               Caption         =   "Delete Labour Records"
            End
         End
         Begin VB.Menu mnuMDeliveryrecords 
            Caption         =   "Modify Delivery Record"
            Begin VB.Menu mnuMUDeliveryrecords 
               Caption         =   "Update Delivery Records"
            End
            Begin VB.Menu mnuMDDeliveryrecords 
               Caption         =   "Delete Delivery Records"
            End
         End
      End
      Begin VB.Menu mnuModifypatientoperationracord 
         Caption         =   "Modify Patient Operation Records"
         Begin VB.Menu mnuUpdatepatientoperationrecords 
            Caption         =   "Update Patient Operation Records"
         End
         Begin VB.Menu mnuDeletepatientoperationrecords 
            Caption         =   "Delete Patient Operation Records"
         End
      End
      Begin VB.Menu mnuModifyfluidchartrecords 
         Caption         =   "Modify Fluid Chart Records"
         Begin VB.Menu mnuInputfluidchart 
            Caption         =   "Modify Input Fluid Chart"
            Begin VB.Menu mnuUpdateinputfluidrecords 
               Caption         =   "Update Input Fluid Records"
            End
            Begin VB.Menu mnuDeleteinputFluidrecords 
               Caption         =   "Delete Input Fluid Records"
            End
         End
         Begin VB.Menu mnuOutputfluidrecords 
            Caption         =   "Modify Output Records"
            Begin VB.Menu mnuUpdateoutputfluidrecords 
               Caption         =   "Update Output Fluid Records"
            End
            Begin VB.Menu mnuDeleteoutputfluidrecords 
               Caption         =   "Delete Output Fluid Records"
            End
         End
      End
      Begin VB.Menu mnuRadiologicalRequest 
         Caption         =   "Modify Radiological Request"
         Begin VB.Menu mnuRUpdateradiologicalrequest 
            Caption         =   "Update Patient Radiological Request"
         End
         Begin VB.Menu mnuRDeleteradiologicalrequest 
            Caption         =   "Delete Patient Radiological Request"
         End
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "ADMINISTRATOR"
      Begin VB.Menu mnuadd 
         Caption         =   "Add New Patient"
      End
      Begin VB.Menu mnuHospitalhistory 
         Caption         =   "Patient Hospital History"
      End
      Begin VB.Menu mnuDiagnosis 
         Caption         =   "Patient Diagnosis"
         Index           =   4
      End
      Begin VB.Menu mnuOpreations 
         Caption         =   "Patient Operations"
      End
      Begin VB.Menu mnumaternityC 
         Caption         =   "Maternity Chart"
         Begin VB.Menu mnuMaternity 
            Caption         =   "Maternity Records"
         End
         Begin VB.Menu mnuMTemp 
            Caption         =   "Temperature"
         End
         Begin VB.Menu mnuMpulse 
            Caption         =   "Pulse"
         End
         Begin VB.Menu mnuLabour 
            Caption         =   "Labour Record"
         End
         Begin VB.Menu mnuDelivery 
            Caption         =   "Delivery"
         End
      End
      Begin VB.Menu mnuFluid 
         Caption         =   "Fluid  Chart"
         Begin VB.Menu mnuInput 
            Caption         =   "Input Chart"
         End
         Begin VB.Menu mnuOutput 
            Caption         =   "Output Chart"
         End
      End
      Begin VB.Menu mnuRadiology 
         Caption         =   "Radiology Request"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "VIEW"
      Begin VB.Menu mnuViewallpatient 
         Caption         =   "View All Patient"
      End
      Begin VB.Menu mnuPatienthospitalhistory 
         Caption         =   "Patient Hospital History"
      End
      Begin VB.Menu mnuVPDiagnosis 
         Caption         =   "Patient Diagnosis History"
      End
      Begin VB.Menu mnuOperations 
         Caption         =   "Patient Operations"
      End
      Begin VB.Menu mnuVMaternityChart 
         Caption         =   "Maternity Charts"
         Begin VB.Menu mnuMMaternityRecords 
            Caption         =   "Maternity Records"
         End
         Begin VB.Menu mnuMTempRecords 
            Caption         =   "Temperature Records"
         End
         Begin VB.Menu mnuVPulseRecords 
            Caption         =   "Pulse Records"
         End
         Begin VB.Menu mnuLabourRecords 
            Caption         =   "Labour Records"
         End
         Begin VB.Menu mnuVDeliveryrecords 
            Caption         =   "Delivery Records"
         End
      End
      Begin VB.Menu mnuVFluidChart 
         Caption         =   "Fluid Chart"
         Begin VB.Menu mnuVInputfluid 
            Caption         =   "Input Fluid"
         End
         Begin VB.Menu mnuVOutputfluid 
            Caption         =   "Output Fluid"
         End
      End
      Begin VB.Menu mnuVRadiology 
         Caption         =   "Radiology Request"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "SEARCH"
      Begin VB.Menu mnPatientRecords 
         Caption         =   "Patient Records"
      End
      Begin VB.Menu mnuSpatienthospitalhistoryrecords 
         Caption         =   "Patient Hospital History Records"
      End
      Begin VB.Menu mnuSpatientdiagnosisrecords 
         Caption         =   "Patient Diagnosis Records"
      End
      Begin VB.Menu mnuSpatientoperationsrecords 
         Caption         =   "Patient Operation Records"
      End
      Begin VB.Menu mnuSmaternitychartrecords 
         Caption         =   "Maternity Chart Records"
         Begin VB.Menu mnuSMMaternityrecords 
            Caption         =   "Maternity Records"
         End
         Begin VB.Menu mnuSMLabourrecords 
            Caption         =   "Labour Records"
         End
         Begin VB.Menu mnuSMTemperaturerecords 
            Caption         =   "Temperature Records"
         End
         Begin VB.Menu mnuSMPulserecords 
            Caption         =   "Pulse Records"
         End
         Begin VB.Menu mnuSDelivery 
            Caption         =   "Delivery Records"
         End
      End
      Begin VB.Menu mnuSFluidchartrecords 
         Caption         =   "Fluid Chart Records"
         Begin VB.Menu mnuSFInputfluidrecords 
            Caption         =   "Input Fluid Records"
         End
         Begin VB.Menu mnuSFOutputfluidrecords 
            Caption         =   "Output Fluid Records"
         End
      End
      Begin VB.Menu mnuSPRadiologicalrecords 
         Caption         =   "Patient Radiological Request Records"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentLength As Byte
Const msg  As String = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
Dim today As Variant
Dim str As String
Dim c As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
lbltitle.Move (Screen.Width - lbltitle.Width) / 2
'lblTime.Caption = Time$
End Sub
Private Sub mnPatientRecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Search", "Search Patient Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmAddPatient.Show
frmAddPatient.txtSurname = rs.Fields!Surname
frmAddPatient.txtOthername = rs.Fields!Othernames
frmAddPatient.cmbsex = rs.Fields!Sex
frmAddPatient.cmbmaritalstatus = rs.Fields!MaritalStatus
frmAddPatient.txtstate = rs.Fields!State
frmAddPatient.txtlga = rs.Fields!LGA
frmAddPatient.txtaddress.Text = rs.Fields!Address
frmAddPatient.txtnationality = rs.Fields!Nationality
frmAddPatient.txttelephone = rs.Fields!PhoneNo
frmAddPatient.txttribe = rs.Fields!Tribe
frmAddPatient.txtOccupation = rs.Fields!Occupation
frmAddPatient.txtReligion = rs.Fields!Religion
frmAddPatient.txtkinsurname = rs.Fields!Psurname
frmAddPatient.txtkinothername = rs.Fields!Pothernames
frmAddPatient.cmbkinsex = rs.Fields!Psex
frmAddPatient.cmbkinmaritalstatus = rs.Fields!Pmaritalstatus
frmAddPatient.txtkinaddress.Text = rs.Fields!Paddress
frmAddPatient.txtkintelephone.Text = rs.Fields!Pphoneno
frmAddPatient.txtkin.Text = rs.Fields!Pnok
frmAddPatient.cmdAddPatient.Enabled = False
frmAddPatient.cmdDelete.Enabled = False
frmAddPatient.cmdUpdate.Enabled = False
End If
End Sub

Private Sub mnuadd_Click()
frmMain.Enabled = False
frmAddPatient.Show
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmAddPatient.cmdDelete.Enabled = False
frmAddPatient.cmdUpdate.Enabled = False
End Sub

Private Sub mnuDeleteinputFluidrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  PatientOutput", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmInput.Show
frmInput.txtSurname = rs.Fields!Surname
frmInput.txtOthernames = rs.Fields!Othernames
frmInput.cmbsex = rs.Fields!Sex
frmInput.txtaddress = rs.Fields!Address
frmInput.txtRegno = rs.Fields!reg_no
frmInput.DTPicker1 = rs.Fields!Date
frmInput.txtVolume = rs.Fields!Volume
frmInput.txtDN = rs.Fields!DN
frmInput.txtFluid = rs.Fields!Fluid
frmInput.lblTime = rs.Fields!Time
frmInput.cmdSearch.Enabled = False
frmInput.cmdSave.Enabled = False
frmInput.cmdUpdate.Enabled = False
frmInput.cmdDelete.Enabled = True
End If
End Sub

Private Sub mnuDeletelabourrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  LabourRecords", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
MsgBox "patient info found,please modify", vbInformation, "info"
frmLabour.Show
frmLabour.txtPatientname = rs.Fields!Patientname
frmLabour.txtDoctor = rs.Fields!Doctor
frmLabour.txtAge = rs.Fields!Age
frmLabour.txtaddress = rs.Fields!Address
frmLabour.txtHusband = rs.Fields!Husband
frmLabour.txtOccupation = rs.Fields!Occupation
frmLabour.txtNCA = rs.Fields!NCA
frmLabour.txtNCD = rs.Fields!NCD
frmLabour.txtChartno = rs.Fields!Chartno
frmLabour.DTPicker1 = rs.Fields!Date
frmLabour.txtUrine = rs.Fields!Urine
frmLabour.txtFH = rs.Fields!FH
frmLabour.lblTime = rs.Fields!Time
frmLabour.txtVE = rs.Fields!VE
frmLabour.txtPDR = rs.Fields!PDR
frmLabour.txtBP = rs.Fields!BP
frmLabour.txtContractions = rs.Fields!Contractions
frmLabour.txtPulse = rs.Fields!Pulse
frmLabour.txtMembranes = rs.Fields!Membranes
frmLabour.cmdSave.Visible = False
frmLabour.cmdDelete.Enabled = True
frmLabour.cmdSearch.Visible = False
frmLabour.cmdUpdate.Visible = False
End If
End Sub

Private Sub mnuDeletematernityrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  MaternityRecords", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmMatern.Show
frmMatern.txtPatientname = rs.Fields!Patientname
frmMatern.txtDoctor = rs.Fields!Doctor
frmMatern.txtNCA = rs.Fields!NCA
frmMatern.txtNCD = rs.Fields!NCD
frmMatern.txtEMP = rs.Fields!EMP
frmMatern.DTPicker1 = rs.Fields!Admitted
frmMatern.txtCOA = rs.Fields!COA
frmMatern.DTPicker2 = rs.Fields!Discharge
frmMatern.DTPicker3 = rs.Fields!Readmitted
frmMatern.lblTime = rs.Fields!Time
frmMatern.txtAge = rs.Fields!Age
frmMatern.txtHusband = rs.Fields!Husband
frmMatern.txtOccupation = rs.Fields!Occupation
frmMatern.txtLMP = rs.Fields!LMP
frmMatern.txtaddress = rs.Fields!Address
frmMatern.txtReligion = rs.Fields!Religion
frmMatern.txtEDD = rs.Fields!EDD
frmMatern.txtParameters = rs.Fields!Parameters
frmMatern.txtRegno = rs.Fields!reg_no
frmMatern.cmdSave.Visible = False
frmMatern.cmdDelete.Enabled = True
frmMatern.cmdSearch.Visible = False
frmMatern.cmdUpdate.Visible = False
End If
End Sub

Private Sub mnuDeleteoutputfluidrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  PatientOutput", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmOutput.Show
frmOutput.txtSurname = rs.Fields!Surname
frmOutput.txtOthernames = rs.Fields!Othernames
frmOutput.cmbsex = rs.Fields!Sex
frmOutput.txtaddress = rs.Fields!Address
frmOutput.txtRegno = rs.Fields!reg_no
frmOutput.DTPicker1 = rs.Fields!Date
frmOutput.txtVolume = rs.Fields!Volume
frmOutput.txtDN = rs.Fields!DN
frmOutput.txtFluid = rs.Fields!Fluid
frmOutput.lblTime = rs.Fields!Time
frmOutput.cmdSearch.Enabled = False
frmOutput.cmdSave.Enabled = False
frmOutput.cmdUpdate.Enabled = False
frmOutput.cmdDelete.Enabled = True
End If
End Sub
Private Sub mnuDeletepatientdiagnosis_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientDiagnosisHistory", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmDiagnosis.Show
frmDiagnosis.txtSurname.Enabled = True
frmDiagnosis.txtOthernames.Enabled = True
frmDiagnosis.cmbsex.Enabled = True
frmDiagnosis.txtaddress.Enabled = True
frmDiagnosis.txtSurname = rs.Fields!Surname
frmDiagnosis.txtOthernames = rs.Fields!Othernames
frmDiagnosis.cmbsex = rs.Fields!Sex
frmDiagnosis.txtaddress = rs.Fields!Address
frmDiagnosis.txtRegno = rs.Fields!reg_no
frmDiagnosis.DTPicker1 = rs.Fields!Date
frmDiagnosis.txtDiagnosis = rs.Fields!Diagnosis
frmDiagnosis.txtCodeno = rs.Fields!Codeno
frmDiagnosis.cmdUpdate.Enabled = False
frmDiagnosis.cmdAdd.Enabled = False
frmDiagnosis.cmdSearch.Enabled = False
frmDiagnosis.cmdDelete.Enabled = True
c.Close
End If
End Sub

Private Sub mnuDeletepatienthospitalhistory_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientHistory", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmHospitalHistory.Show
frmHospitalHistory.txtSurname.Enabled = False
frmHospitalHistory.txtOthernames.Enabled = False
frmHospitalHistory.cmbsex.Enabled = False
frmHospitalHistory.txtaddress.Enabled = False
frmHospitalHistory.txtRegno.Enabled = False
frmHospitalHistory.txtSurname = rs.Fields!Surname
frmHospitalHistory.txtOthernames = rs.Fields!Othernames
frmHospitalHistory.cmbsex = rs.Fields!Sex
frmHospitalHistory.txtaddress = rs.Fields!Address
frmHospitalHistory.txtRegno = rs.Fields!Regno
frmHospitalHistory.DTPicker1 = rs.Fields!DTPicker1
frmHospitalHistory.DTPicker2 = rs.Fields!DTPicker2
frmHospitalHistory.txtReferred = rs.Fields!Referred
frmHospitalHistory.txtDischarge = rs.Fields!Dischargeto
frmHospitalHistory.txtPhysician = rs.Fields!Physician
frmHospitalHistory.txtOthers = rs.Fields!Others
frmHospitalHistory.txtWard = rs.Fields!Ward
frmHospitalHistory.cmdUpdate.Enabled = False
frmHospitalHistory.cmdAdd.Enabled = False
frmHospitalHistory.cmdSearch.Visible = False
frmHospitalHistory.cmdDelete.Enabled = True
c.Close
End If
End Sub

Private Sub mnuDeletepatientoperationrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientOperations", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmOperations.Show
frmOperations.txtSurname.Enabled = True
frmOperations.txtOthernames.Enabled = True
frmOperations.cmbsex.Enabled = True
frmOperations.txtaddress.Enabled = True
frmOperations.txtSurname = rs.Fields!Surname
frmOperations.txtOthernames = rs.Fields!Othernames
frmOperations.cmbsex = rs.Fields!Sex
frmOperations.txtaddress = rs.Fields!Address
frmOperations.txtRegno = rs.Fields!reg_no
frmOperations.DTPicker1 = rs.Fields!Date
frmOperations.txtSurgeon = rs.Fields!Surgeon
frmOperations.txtOperation = rs.Fields!Operation
frmOperations.txtCodeno = rs.Fields!Codeno
frmOperations.cmdUpdate.Enabled = False
frmOperations.cmdSave.Enabled = False
frmOperations.cmdSearch.Enabled = False
frmOperations.cmdDelete.Enabled = True
c.Close
End If
End Sub

Private Sub mnuDeletePatientRecord_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmAddPatient.Show
frmAddPatient.txtSurname = rs.Fields!Surname
frmAddPatient.txtOthername = rs.Fields!Othernames
frmAddPatient.cmbsex = rs.Fields!Sex
frmAddPatient.cmbmaritalstatus = rs.Fields!MaritalStatus
frmAddPatient.txtstate = rs.Fields!State
frmAddPatient.txtlga = rs.Fields!LGA
frmAddPatient.txtaddress.Text = rs.Fields!Address
frmAddPatient.txtnationality = rs.Fields!Nationality
frmAddPatient.txttelephone = rs.Fields!PhoneNo
frmAddPatient.txttribe = rs.Fields!Tribe
frmAddPatient.txtOccupation = rs.Fields!Occupation
frmAddPatient.txtReligion = rs.Fields!Religion
frmAddPatient.txtkinsurname = rs.Fields!Psurname
frmAddPatient.txtkinothername = rs.Fields!Pothernames
frmAddPatient.cmbkinsex = rs.Fields!Psex
frmAddPatient.cmbkinmaritalstatus = rs.Fields!Pmaritalstatus
frmAddPatient.txtkinaddress.Text = rs.Fields!Paddress
frmAddPatient.txtkintelephone.Text = rs.Fields!Pphoneno
frmAddPatient.txtkin.Text = rs.Fields!Pnok
frmAddPatient.txtSurname.Enabled = False
frmAddPatient.txtOthername.Enabled = False
frmAddPatient.cmbsex.Enabled = False
frmAddPatient.cmbmaritalstatus.Enabled = False
frmAddPatient.txtstate.Enabled = False
frmAddPatient.txtlga.Enabled = False
frmAddPatient.txtaddress.Enabled = False
frmAddPatient.txtnationality.Enabled = False
frmAddPatient.txttelephone.Enabled = False
frmAddPatient.txttribe.Enabled = False
frmAddPatient.txtOccupation.Enabled = False
frmAddPatient.txtReligion.Enabled = False
frmAddPatient.txtkinsurname.Enabled = False
frmAddPatient.txtkinothername.Enabled = False
frmAddPatient.cmbkinsex.Enabled = False
frmAddPatient.cmbkinmaritalstatus.Enabled = False
frmAddPatient.txtkinaddress.Enabled = False
frmAddPatient.txtkintelephone.Enabled = False
frmAddPatient.cmdAddPatient.Enabled = False
frmAddPatient.cmdUpdate.Enabled = False
frmAddPatient.cmdDelete.Enabled = True
End If
End Sub

Private Sub mnuDeletetemperaturerecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Chart Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientTemperature", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmtemp.Show
frmtemp.DTPicker1.SetFocus
frmtemp.txtPatientname.Enabled = False
frmtemp.txtOccupation.Enabled = False
frmtemp.txtaddress.Enabled = False
frmtemp.txtChartno.Enabled = False
frmtemp.txtHusband.Enabled = False
frmtemp.txtPatientname = rs.Fields!Patientname
frmtemp.txtOccupation = rs.Fields!Occupation
frmtemp.txtHusband = rs.Fields!Husband
frmtemp.DTPicker1 = rs.Fields!Date
frmtemp.cmbHour = rs.Fields!Hour
frmtemp.DTPicker2 = rs.Fields!DOD
frmtemp.txtTemp = rs.Fields!Temperature
frmtemp.txtChartno = rs.Fields!Chartno
frmtemp.cmdUpdate.Enabled = False
frmtemp.cmdSave.Visible = False
frmtemp.cmdSearch.Visible = False
frmtemp.cmdDelete.Visible = True
c.Close
End If
End Sub
Private Sub mnuDelivery_Click()
frmMain.Enabled = False
frmDelivery.Show
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmMain.Enabled = True
frmDelivery.cmdDelete.Enabled = False
frmDelivery.cmdUpdate.Enabled = False
frmDelivery.cmdSave.Enabled = True
frmDelivery.txtChartno.SetFocus
frmDelivery.txtPatientname.Enabled = False
frmDelivery.txtDoctor.Enabled = False
frmDelivery.txtAge.Enabled = False
frmDelivery.txtaddress.Enabled = False
frmDelivery.txtHusband.Enabled = False
frmDelivery.txtOccupation.Enabled = False
frmDelivery.txtNCA.Enabled = False
frmDelivery.txtNCD.Enabled = False
End Sub

Private Sub mnuDiagnosis_Click(Index As Integer)
frmMain.Enabled = False
frmDiagnosis.Show
Timer1.Enabled = False
frmDiagnosis.txtRegno.SetFocus
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmDiagnosis.cmdDelete.Enabled = False
frmDiagnosis.cmdUpdate.Enabled = False
End Sub

Private Sub mnuHospitalhistory_Click()
frmMain.Enabled = False
frmHospitalHistory.Show
Timer1.Enabled = False
frmHospitalHistory.txtRegno.SetFocus
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmHospitalHistory.cmdDelete.Enabled = False
frmHospitalHistory.cmdUpdate.Enabled = False
End Sub

Private Sub mnuInput_Click()
frmMain.Enabled = False
frmInput.Show
frmInput.txtRegno.SetFocus
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmInput.cmdDelete.Visible = False
frmInput.cmdUpdate.Visible = False
End Sub

Private Sub mnuLabour_Click()
frmMain.Enabled = False
frmLabour.Show
frmLabour.txtChartno.SetFocus
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmMain.Enabled = True
frmLabour.cmdDelete.Enabled = False
frmLabour.cmdUpdate.Enabled = False
frmLabour.cmdSave.Enabled = False
frmLabour.txtChartno.SetFocus
frmLabour.txtPatientname.Enabled = False
frmLabour.txtDoctor.Enabled = False
frmLabour.txtAge.Enabled = False
frmLabour.txtaddress.Enabled = False
frmLabour.txtHusband.Enabled = False
frmLabour.txtOccupation.Enabled = False
frmLabour.txtNCA.Enabled = False
frmLabour.txtNCD.Enabled = False
End Sub

Private Sub mnuLabourRecords_Click()
frmViewPatientLabour.Show
Me.Enabled = False
End Sub

Private Sub mnuMaternity_Click()
frmMain.Enabled = False
frmMatern.Show
Timer1.Enabled = False
frmMatern.txtRegno.SetFocus
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmMatern.cmdDelete.Enabled = False
frmMatern.cmdUpdate.Enabled = False
frmMatern.cmdDelete.Enabled = False
frmMatern.cmdSave.Enabled = False
End Sub

Private Sub mnuMDDeliveryrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  Delivery", c, adOpenDynamic, adLockOptimistic
rs.Find "Chartno='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
MsgBox "patient info found,please modify", vbInformation, "found"
frmDelivery.Show
frmDelivery.txtPatientname.Enabled = False
frmDelivery.txtDoctor.Enabled = False
frmDelivery.txtAge.Enabled = False
frmDelivery.txtaddress.Enabled = False
frmDelivery.txtHusband.Enabled = False
frmDelivery.txtOccupation.Enabled = False
frmDelivery.txtNCA.Enabled = False
frmDelivery.txtNCD.Enabled = False
frmDelivery.txtChartno.Enabled = False
frmDelivery.cmdSave.Visible = False
frmDelivery.cmdDelete.Enabled = True
frmDelivery.cmdSearch.Visible = False
frmDelivery.cmdUpdate.Visible = False
frmDelivery.txtPatientname = rs.Fields!Patientname
frmDelivery.txtDoctor = rs.Fields!Doctor
frmDelivery.txtAge = rs.Fields!Age
frmDelivery.txtaddress = rs.Fields!Address
frmDelivery.txtHusband = rs.Fields!Husband
frmDelivery.txtOccupation = rs.Fields!Occupation
frmDelivery.txtNCA = rs.Fields!NCA
frmDelivery.txtNCD = rs.Fields!NCD
frmDelivery.txtChartno = rs.Fields!Chartno
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtTD = rs.Fields!TD
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtSutures = rs.Fields!Sutures
frmDelivery.txtHOL = rs.Fields!HOL
frmDelivery.txtPD = rs.Fields!PD
frmDelivery.txtBPAD = rs.Fields!BPAD
frmDelivery.txtMR = rs.Fields!MR
frmDelivery.txtPME = rs.Fields!PME
frmDelivery.cmbsex = rs.Fields!Sex
frmDelivery.txtMidwife = rs.Fields!Midwife
frmDelivery.DTPicker1 = rs.Fields!DIB
frmDelivery.txtPL = rs.Fields!Midwife
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtTIB = rs.Fields!TIB
frmDelivery.txtBLA = rs.Fields!BLA
frmDelivery.txtCAB = rs.Fields!CAB
frmDelivery.txtWeight = rs.Fields!Weight
frmDelivery.txtNTC = rs.Fields!NTC
frmDelivery.txtDN = rs.Fields!DN
End If
End Sub

Private Sub mnuMMaternityRecords_Click()
frmViewPatientMaternity.Show
Me.Enabled = False
End Sub

Private Sub mnuModityPatientRecord_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmAddPatient.Show
frmAddPatient.txtSurname = rs.Fields!Surname
frmAddPatient.txtOthername = rs.Fields!Othernames
frmAddPatient.cmbsex = rs.Fields!Sex
frmAddPatient.cmbmaritalstatus = rs.Fields!MaritalStatus
frmAddPatient.txtstate = rs.Fields!State
frmAddPatient.txtlga = rs.Fields!LGA
frmAddPatient.txtaddress.Text = rs.Fields!Address
frmAddPatient.txtnationality = rs.Fields!Nationality
frmAddPatient.txttelephone = rs.Fields!PhoneNo
frmAddPatient.txttribe = rs.Fields!Tribe
frmAddPatient.txtOccupation = rs.Fields!Occupation
frmAddPatient.txtReligion = rs.Fields!Religion
frmAddPatient.txtkinsurname = rs.Fields!Psurname
frmAddPatient.txtkinothername = rs.Fields!Pothernames
frmAddPatient.cmbkinsex = rs.Fields!Psex
frmAddPatient.cmbkinmaritalstatus = rs.Fields!Pmaritalstatus
frmAddPatient.txtkinaddress.Text = rs.Fields!Paddress
frmAddPatient.txtkintelephone.Text = rs.Fields!Pphoneno
frmAddPatient.txtkin.Text = rs.Fields!Pnok
frmAddPatient.cmdAddPatient.Enabled = False
frmAddPatient.cmdDelete.Enabled = False
End If
End Sub

Private Sub mnuMpressure_Click()
frmMain.Enabled = False
frmPressure.Show
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
End Sub

Private Sub mnuMpulse_Click()
frmMain.Enabled = False
frmPulse.Show
Timer1.Enabled = False
frmPulse.txtChartno.SetFocus
frmMain.Enabled = True
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmPulse.cmdDelete.Visible = False
frmPulse.cmdUpdate.Visible = False
frmPulse.txtPatientname.Enabled = False
frmPulse.txtHusband.Enabled = False
frmPulse.txtaddress.Enabled = False
frmPulse.txtOccupation.Enabled = False
frmPulse.txtChartno.SetFocus
End Sub

Private Sub mnuMtemp_Click()
frmMain.Enabled = False
frmtemp.Show
Timer1.Enabled = False
frmtemp.txtChartno.SetFocus
frmMain.Enabled = True
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmtemp.cmdDelete.Enabled = False
frmtemp.cmdUpdate.Enabled = False
frmtemp.txtPatientname.Enabled = False
frmtemp.txtHusband.Enabled = False
frmtemp.txtaddress.Enabled = False
frmtemp.txtOccupation.Enabled = False
End Sub


Private Sub mnuMTempRecords_Click()
frmViewPatientTemperature.Show
Me.Enabled = False
End Sub

Private Sub mnuMUDeliveryrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  Delivery", c, adOpenDynamic, adLockOptimistic
rs.Find "Chartno='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
MsgBox "patient info found,please modify", vbInformation, "found"
frmDelivery.Show
frmDelivery.txtPatientname.Enabled = False
frmDelivery.txtDoctor.Enabled = False
frmDelivery.txtAge.Enabled = False
frmDelivery.txtaddress.Enabled = False
frmDelivery.txtHusband.Enabled = False
frmDelivery.txtOccupation.Enabled = False
frmDelivery.txtNCA.Enabled = False
frmDelivery.txtNCD.Enabled = False
frmDelivery.txtChartno.Enabled = False
frmDelivery.cmdSave.Visible = False
frmDelivery.cmdDelete.Enabled = False
frmDelivery.cmdSearch.Visible = False
frmDelivery.cmdUpdate.Visible = True
frmDelivery.txtTLS.SetFocus
frmDelivery.txtPatientname = rs.Fields!Patientname
frmDelivery.txtDoctor = rs.Fields!Doctor
frmDelivery.txtAge = rs.Fields!Age
frmDelivery.txtaddress = rs.Fields!Address
frmDelivery.txtHusband = rs.Fields!Husband
frmDelivery.txtOccupation = rs.Fields!Occupation
frmDelivery.txtNCA = rs.Fields!NCA
frmDelivery.txtNCD = rs.Fields!NCD
frmDelivery.txtChartno = rs.Fields!Chartno
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtTD = rs.Fields!TD
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtSutures = rs.Fields!Sutures
frmDelivery.txtHOL = rs.Fields!HOL
frmDelivery.txtPD = rs.Fields!PD
frmDelivery.txtBPAD = rs.Fields!BPAD
frmDelivery.txtMR = rs.Fields!MR
frmDelivery.txtPME = rs.Fields!PME
frmDelivery.txtMidwife = rs.Fields!Midwife
frmDelivery.cmbsex = rs.Fields!Sex
frmDelivery.DTPicker1 = rs.Fields!DIB
frmDelivery.txtPL = rs.Fields!Midwife
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtTIB = rs.Fields!TIB
frmDelivery.txtBLA = rs.Fields!BLA
frmDelivery.txtCAB = rs.Fields!CAB
frmDelivery.txtWeight = rs.Fields!Weight
frmDelivery.txtNTC = rs.Fields!NTC
frmDelivery.txtDN = rs.Fields!DN
End If
End Sub

Private Sub mnuOperations_Click()
frmViewPatientOperation.Show
Me.Enabled = False
End Sub

Private Sub mnuOpreations_Click()
frmMain.Enabled = False
frmOperations.Show
frmOperations.txtRegno.SetFocus
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmOperations.cmdDelete.Enabled = False
frmOperations.cmdUpdate.Enabled = False
End Sub

Private Sub mnuOutput_Click()
frmMain.Enabled = False
frmOutput.Show
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmOutput.cmdDelete.Enabled = False
frmOutput.cmdUpdate.Enabled = False
End Sub

Private Sub mnuPatienthospitalhistory_Click()
frmViewPatientHistory.Show
Me.Enabled = False
End Sub

Private Sub mnuRadiology_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient", "Search Patient Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from AddPatient", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmRadiology.Show
frmRadiology.txtRegno = rs.Fields!reg_no
MsgBox "Registration Number valid!!!", vbInformation, "info"
frmMain.Enabled = False
frmRadiology.Show
Timer1.Enabled = False
frmRadiology.txtName.SetFocus
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
frmRadiology.cmdDelete.Enabled = False
frmRadiology.cmdUpdate.Enabled = False
End If
End Sub

Private Sub mnuupadatepatientdiagnosis_Click()

End Sub

Private Sub mnuRDeleteradiologicalrequest_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL,ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:/HMS/crack/hms.mdb;Persist Security Info= false")
rs.Open " select * from radiology", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmRadiology.Show
frmRadiology.txtName = rs.Fields!Name
frmRadiology.DTPicker1 = rs.Fields!Date
frmRadiology.txtAge = rs.Fields!Age
frmRadiology.txtLMP = rs.Fields!LMP
frmRadiology.cmbsex = rs.Fields!Sex
frmRadiology.txtaddress = rs.Fields!Address
frmRadiology.txtCI = rs.Fields!CI
frmRadiology.cmbExamination = rs.Fields!Examination
frmRadiology.txtRegno = rs.Fields!reg_no
frmRadiology.txtXrayno = rs.Fields!Xrayno
frmRadiology.cmdUpdate.Enabled = False
frmRadiology.cmdSave.Enabled = False
frmMain.Enabled = True
c.Close
End If
End Sub

Private Sub mnuRUpdateradiologicalrequest_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL,ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:/HMS/crack/hms.mdb;Persist Security Info= false")
rs.Open " select * from Radiology", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmRadiology.Show
frmRadiology.txtName = rs.Fields!Name
frmRadiology.DTPicker1 = rs.Fields!Date
frmRadiology.txtAge = rs.Fields!Age
frmRadiology.txtLMP = rs.Fields!LMP
frmRadiology.cmbsex = rs.Fields!Sex
frmRadiology.txtaddress = rs.Fields!Address
frmRadiology.txtCI = rs.Fields!CI
frmRadiology.cmbExamination = rs.Fields!Examination
frmRadiology.txtRegno = rs.Fields!reg_no
frmRadiology.txtXrayno = rs.Fields!Xrayno
frmRadiology.cmdDelete.Enabled = False
frmRadiology.cmdSave.Enabled = False
frmMain.Enabled = True
End If
End Sub

Private Sub mnuSDelivery_Click()
'On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Delivery", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmDelivery.txtPatientname = rs.Fields!Patientname
frmDelivery.txtDoctor = rs.Fields!Doctor
frmDelivery.txtAge = rs.Fields!Age
frmDelivery.txtaddress = rs.Fields!Address
frmDelivery.txtHusband = rs.Fields!Husband
frmDelivery.txtOccupation = rs.Fields!Occupation
frmDelivery.txtNCA = rs.Fields!NCA
frmDelivery.txtNCD = rs.Fields!NCD
frmDelivery.txtTLS = rs.Fields!TLS
frmDelivery.txtTD = rs.Fields!TD
frmDelivery.txtSutures = rs.Fields!Sutures
frmDelivery.txtHOL = rs.Fields!HOL
frmDelivery.txtPD = rs.Fields!PD
frmDelivery.txtBPAD = rs.Fields!BPAD
frmDelivery.txtMR = rs.Fields!MR
frmDelivery.txtPME = rs.Fields!PME
frmDelivery.cmbsex = rs.Fields!Sex
frmDelivery.txtPatientname = rs.Fields!Patientname
frmDelivery.DTPicker1 = rs.Fields!DIB
frmDelivery.txtPL = rs.Fields!PL
frmDelivery.txtMidwife = rs.Fields!Midwife
frmDelivery.txtTIB = rs.Fields!TIB
frmDelivery.txtBLA = rs.Fields!BLA
frmDelivery.txtCAB = rs.Fields!CAB
frmDelivery.txtWeight = rs.Fields!Weight
frmDelivery.txtNTC = rs.Fields!NTC
frmDelivery.txtDN = rs.Fields!Patientname
End If
End Sub

Private Sub mnuSFInputfluidrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  PatientInput", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmInput.txtSurname = rs.Fields!Surname
frmInput.txtOthernames = rs.Fields!Othernames
frmInput.cmbsex = rs.Fields!Sex
frmInput.txtaddress = rs.Fields!Address
frmInput.txtRegno = rs.Fields!reg_no
frmInput.DTPicker1 = rs.Fields!Date
frmInput.txtVolume = rs.Fields!Volume
frmInput.txtDN = rs.Fields!DN
frmInput.txtFluid = rs.Fields!Fluid
frmInput.lblTime = rs.Fields!Time
frmInput.cmdSearch.Enabled = False
frmInput.cmdSave.Enabled = False
frmInput.cmdDelete.Enabled = False
frmInput.cmdUpdate.Enabled = False
MsgBox "patient info found,please check", vbInformation, "Search"
frmInput.Show
End If
End Sub

Private Sub mnuSFOutputfluidrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  PatientOutput", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmOutput.txtSurname = rs.Fields!Surname
frmOutput.txtOthernames = rs.Fields!Othernames
frmOutput.cmbsex = rs.Fields!Sex
frmOutput.txtaddress = rs.Fields!Address
frmOutput.txtRegno = rs.Fields!reg_no
frmOutput.DTPicker1 = rs.Fields!Date
frmOutput.txtVolume = rs.Fields!Volume
frmOutput.txtDN = rs.Fields!DN
frmOutput.txtFluid = rs.Fields!Fluid
frmOutput.lblTime = rs.Fields!Time
frmOutput.cmdSearch.Enabled = False
frmOutput.cmdSave.Enabled = False
frmOutput.cmdDelete.Enabled = False
frmOutput.cmdUpdate.Enabled = False
MsgBox "Patient info found,please check!!!", vbInformation, "searched"
frmOutput.Show
End If
End Sub

Private Sub mnuSMLabourrecords_Click()
'On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  LabourRecords", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmLabour.txtPatientname = rs.Fields!Patientname
frmLabour.txtDoctor = rs.Fields!Doctor
frmLabour.txtAge = rs.Fields!Age
frmLabour.txtaddress = rs.Fields!Address
frmLabour.txtHusband = rs.Fields!Husband
frmLabour.txtOccupation = rs.Fields!Occupation
frmLabour.txtNCA = rs.Fields!NCA
frmLabour.txtNCD = rs.Fields!NCD
frmLabour.txtChartno = rs.Fields!Chartno
frmLabour.DTPicker1 = rs.Fields!Date
frmLabour.txtUrine = rs.Fields!Urine
frmLabour.txtFH = rs.Fields!FH
frmLabour.lblTime = rs.Fields!Time
frmLabour.txtVE = rs.Fields!VE
frmLabour.txtPDR = rs.Fields!PDR
frmLabour.txtBP = rs.Fields!BP
frmLabour.txtContractions = rs.Fields!Contractions
frmLabour.txtPulse = rs.Fields!Pulse
frmLabour.txtMembranes = rs.Fields!Membranes
frmLabour.cmdSave.Enabled = False
frmLabour.cmdDelete.Enabled = False
frmLabour.cmdSearch.Enabled = False
frmLabour.cmdUpdate.Enabled = False
MsgBox "patient info found,please Check", vbInformation, "found"
frmLabour.Show
End If
End Sub

Private Sub mnuSMMaternityrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  MaternityRecords", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmMatern.Show
frmMatern.txtPatientname = rs.Fields!Patientname
frmMatern.txtDoctor = rs.Fields!Doctor
frmMatern.txtNCA = rs.Fields!NCA
frmMatern.txtNCD = rs.Fields!NCD
frmMatern.txtEMP = rs.Fields!EMP
frmMatern.DTPicker1 = rs.Fields!Admitted
frmMatern.txtCOA = rs.Fields!COA
frmMatern.DTPicker2 = rs.Fields!Discharge
frmMatern.DTPicker3 = rs.Fields!Readmitted
frmMatern.lblTime = rs.Fields!Time
frmMatern.txtAge = rs.Fields!Age
frmMatern.txtHusband = rs.Fields!Husband
frmMatern.txtOccupation = rs.Fields!Occupation
frmMatern.txtLMP = rs.Fields!LMP
frmMatern.txtaddress = rs.Fields!Address
frmMatern.txtReligion = rs.Fields!Religion
frmMatern.txtEDD = rs.Fields!EDD
frmMatern.txtParameters = rs.Fields!Parameters
frmMatern.txtRegno = rs.Fields!reg_no
frmMatern.cmdSave.Enabled = False
frmMatern.cmdDelete.Enabled = False
frmMatern.cmdSearch.Enabled = False
frmMatern.cmdUpdate.Enabled = False
End If
End Sub

Private Sub mnuSMPulserecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Chart Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientPulse", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmPulse.DTPicker1.SetFocus
frmPulse.txtPatientname.Enabled = True
frmPulse.txtOccupation.Enabled = True
frmPulse.txtaddress.Enabled = True
frmPulse.txtChartno.Enabled = True
frmPulse.txtHusband.Enabled = True
frmPulse.txtPatientname = rs.Fields!Patientname
frmPulse.txtOccupation = rs.Fields!Occupation
frmPulse.txtaddress = rs.Fields!Address
frmPulse.txtHusband = rs.Fields!Husband
frmPulse.DTPicker1 = rs.Fields!Date
frmPulse.cmbHour = rs.Fields!Hour
frmPulse.DTPicker2 = rs.Fields!DOD
frmPulse.txtPulse = rs.Fields!Pulse
frmPulse.txtChartno = rs.Fields!Chartno
frmPulse.cmdUpdate.Enabled = False
frmPulse.cmdSave.Enabled = False
frmPulse.cmdSearch.Enabled = False
frmPulse.cmdDelete.Enabled = False
MsgBox "Patient info found,please check", vbInformation, "info"
frmPulse.Show
c.Close
End If
End Sub

Private Sub mnuSMTemperaturerecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Chart Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientTemperature", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmtemp.Show
frmtemp.DTPicker1.SetFocus
frmtemp.txtPatientname.Enabled = True
frmtemp.txtOccupation.Enabled = True
frmtemp.txtaddress.Enabled = True
frmtemp.txtChartno.Enabled = True
frmtemp.txtHusband.Enabled = True
frmtemp.txtPatientname = rs.Fields!Patientname
frmtemp.txtOccupation = rs.Fields!Occupation
frmtemp.txtHusband = rs.Fields!Husband
frmtemp.DTPicker1 = rs.Fields!Date
frmtemp.cmbHour = rs.Fields!Hour
frmtemp.DTPicker2 = rs.Fields!DOD
frmtemp.txtTemp = rs.Fields!Temperature
frmtemp.txtChartno = rs.Fields!Chartno
frmtemp.cmdUpdate.Enabled = False
frmtemp.cmdSave.Enabled = False
frmtemp.cmdSearch.Enabled = False
frmtemp.cmdDelete.Enabled = False
c.Close
End If
End Sub


Private Sub mnuSpatientdiagnosisrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientDiagnosisHistory", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmDiagnosis.Show
frmDiagnosis.txtSurname.Enabled = True
frmDiagnosis.txtOthernames.Enabled = True
frmDiagnosis.cmbsex.Enabled = True
frmDiagnosis.txtaddress.Enabled = True
frmDiagnosis.txtSurname = rs.Fields!Surname
frmDiagnosis.txtOthernames = rs.Fields!Othernames
frmDiagnosis.cmbsex = rs.Fields!Sex
frmDiagnosis.txtaddress = rs.Fields!Address
frmDiagnosis.txtRegno = rs.Fields!reg_no
frmDiagnosis.DTPicker1 = rs.Fields!DTPicker1
frmDiagnosis.txtDiagnosis = rs.Fields!Diagnosis
frmDiagnosis.txtCodeno = rs.Fields!Codeno
frmDiagnosis.cmdUpdate.Enabled = False
frmDiagnosis.cmdAdd.Enabled = False
frmDiagnosis.cmdSearch.Enabled = False
frmDiagnosis.cmdDelete.Enabled = False
c.Close
End If
End Sub

Private Sub mnuSpatienthospitalhistoryrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientHistory", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmHospitalHistory.Show
frmHospitalHistory.txtSurname.Enabled = True
frmHospitalHistory.txtOthernames.Enabled = True
frmHospitalHistory.cmbsex.Enabled = True
frmHospitalHistory.txtaddress.Enabled = True
frmHospitalHistory.txtRegno.Enabled = True
frmHospitalHistory.txtSurname = rs.Fields!Surname
frmHospitalHistory.txtOthernames = rs.Fields!Othernames
frmHospitalHistory.cmbsex = rs.Fields!Sex
frmHospitalHistory.txtaddress = rs.Fields!Address
frmHospitalHistory.txtRegno = rs.Fields!Regno
frmHospitalHistory.DTPicker1 = rs.Fields!DTPicker1
frmHospitalHistory.DTPicker2 = rs.Fields!DTPicker2
frmHospitalHistory.txtReferred = rs.Fields!Referred
frmHospitalHistory.txtDischarge = rs.Fields!Dischargeto
frmHospitalHistory.txtPhysician = rs.Fields!Physician
frmHospitalHistory.txtOthers = rs.Fields!Others
frmHospitalHistory.txtWard = rs.Fields!Ward
frmHospitalHistory.cmdUpdate.Enabled = False
frmHospitalHistory.cmdAdd.Enabled = False
frmHospitalHistory.cmdSearch.Enabled = False
frmHospitalHistory.cmdDelete.Enabled = False
c.Close
End If
End Sub

Private Sub mnuSpatientoperationsrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientOperations", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmOperations.Show
frmOperations.txtSurname.Enabled = True
frmOperations.txtOthernames.Enabled = True
frmOperations.cmbsex.Enabled = True
frmOperations.txtaddress.Enabled = True
frmOperations.txtSurname = rs.Fields!Surname
frmOperations.txtOthernames = rs.Fields!Othernames
frmOperations.cmbsex = rs.Fields!Sex
frmOperations.txtaddress = rs.Fields!Address
frmOperations.txtRegno = rs.Fields!reg_no
frmOperations.DTPicker1 = rs.Fields!Date
frmOperations.txtSurgeon = rs.Fields!Surgeon
frmOperations.txtOperation = rs.Fields!Operation
frmOperations.txtCodeno = rs.Fields!Codeno
frmOperations.cmdUpdate.Enabled = False
frmOperations.cmdSave.Enabled = False
frmOperations.cmdSearch.Enabled = False
frmOperations.cmdDelete.Enabled = False
c.Close
End If
End Sub

Private Sub mnuSPRadiologicalrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from Radiology ", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmRadiology.txtName = rs.Fields!Name
frmRadiology.DTPicker1 = rs.Fields!Date
frmRadiology.txtAge = rs.Fields!Age
frmRadiology.txtLMP = rs.Fields!LMP
frmRadiology.cmbsex = rs.Fields!Sex
frmRadiology.txtaddress = rs.Fields!Address
frmRadiology.txtCI = rs.Fields!CI
frmRadiology.cmbExamination = rs.Fields!Examination
frmRadiology.txtRegno = rs.Fields!reg_no
frmRadiology.txtXrayno = rs.Fields!Xrayno
MsgBox "Patient radiological Request form found,Please check your records!!!", vbInformation, "Updated"
frmRadiology.Show
frmRadiology.cmdDelete.Enabled = False
frmRadiology.cmdUpdate.Enabled = False
frmRadiology.cmdSave.Enabled = False
frmMain.Enabled = True
End If
End Sub

Private Sub mnuUpdateinputfluidrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  PatientInput", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmInput.Show
frmInput.Timer1 = False
MsgBox "patient info found,please modify", vbInformation, "found"
frmInput.txtSurname.Enabled = False
frmInput.txtOthernames.Enabled = False
frmInput.cmbsex.Enabled = False
frmInput.txtaddress.Enabled = False
frmInput.txtRegno.Enabled = False
frmInput.txtSurname = rs.Fields!Surname
frmInput.txtOthernames = rs.Fields!Othernames
frmInput.cmbsex = rs.Fields!Sex
frmInput.txtaddress = rs.Fields!Address
frmInput.txtRegno = rs.Fields!reg_no
frmInput.DTPicker1 = rs.Fields!Date
frmInput.txtVolume = rs.Fields!Volume
frmInput.txtDN = rs.Fields!DN
frmInput.txtFluid = rs.Fields!Fluid
frmInput.lblTime = rs.Fields!Time
frmInput.cmdSearch.Enabled = False
frmInput.cmdSave.Enabled = False
frmInput.cmdDelete.Enabled = False
End If
End Sub
Private Sub mnuUpdatelabourrecords_Click()
On Error Resume Next
Form3.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  LabourRecords", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmLabour.Timer1 = False
frmLabour.txtPatientname = rs.Fields!Patientname
frmLabour.txtDoctor = rs.Fields!Doctor
frmLabour.txtAge = rs.Fields!Age
frmLabour.txtaddress = rs.Fields!Address
frmLabour.txtHusband = rs.Fields!Husband
frmLabour.txtOccupation = rs.Fields!Occupation
frmLabour.txtNCA = rs.Fields!NCA
frmLabour.txtNCD = rs.Fields!NCD
frmLabour.txtChartno = rs.Fields!Chartno
frmLabour.DTPicker1 = rs.Fields!Date
frmLabour.txtUrine = rs.Fields!Urine
frmLabour.txtFH = rs.Fields!FH
frmLabour.lblTime = rs.Fields!Time
frmLabour.txtVE = rs.Fields!VE
frmLabour.txtPDR = rs.Fields!PDR
frmLabour.txtBP = rs.Fields!BP
frmLabour.txtContractions = rs.Fields!Contractions
frmLabour.txtPulse = rs.Fields!Pulse
frmLabour.txtMembranes = rs.Fields!Membranes
frmLabour.cmdSave.Visible = False
frmLabour.cmdDelete.Enabled = False
frmLabour.cmdSearch.Visible = False
frmLabour.cmdUpdate.Visible = True
MsgBox "patient info found,please modify", vbInformation, "found"
frmLabour.Show
frmLabour.txtPatientname.Enabled = False
frmLabour.txtDoctor.Enabled = False
frmLabour.txtAge.Enabled = False
frmLabour.txtaddress.Enabled = False
frmLabour.txtHusband.Enabled = False
frmLabour.txtOccupation.Enabled = False
frmLabour.txtNCA.Enabled = False
frmLabour.txtNCD.Enabled = False
frmLabour.txtChartno.Enabled = False
frmLabour.DTPicker1.SetFocus
End If
End Sub

Private Sub mnuUpdatematernityrecords_Click()
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Chart Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  MaternityRecords", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmMatern.Timer1 = False
frmMatern.txtPatientname = rs.Fields!Patientname
frmMatern.txtDoctor = rs.Fields!Doctor
frmMatern.txtNCA = rs.Fields!NCA
frmMatern.txtNCD = rs.Fields!NCD
frmMatern.txtEMP = rs.Fields!EMP
frmMatern.DTPicker1 = rs.Fields!Admitted
frmMatern.txtCOA = rs.Fields!COA
frmMatern.DTPicker2 = rs.Fields!Discharge
frmMatern.DTPicker3 = rs.Fields!Readmitted
frmMatern.lblTime = rs.Fields!Time
frmMatern.txtAge = rs.Fields!Age
frmMatern.txtHusband = rs.Fields!Husband
frmMatern.txtOccupation = rs.Fields!Occupation
frmMatern.txtLMP = rs.Fields!LMP
frmMatern.txtaddress = rs.Fields!Address
frmMatern.txtReligion = rs.Fields!Religion
frmMatern.txtEDD = rs.Fields!EDD
frmMatern.txtParameters = rs.Fields!Parameters
frmMatern.txtRegno = rs.Fields!reg_no
frmMatern.cmdSave.Visible = False
frmMatern.cmdDelete.Visible = False
frmMatern.cmdSearch.Visible = False
frmMatern.cmdUpdate.Enabled = True
frmMatern.txtRegno.Enabled = False
MsgBox "patient info found,please modify", vbInformation, "found"
frmMatern.Show
frmMatern.txtPatientname.Enabled = False
frmMatern.txtNCA.SetFocus
End If
End Sub

Private Sub mnuUpdateoutputfluidrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Registration Number of the Patient to Modify", "Modify Record")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from  PatientOutput", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmOutput.Show
frmOutput.Timer1 = False
frmOutput.txtSurname.Enabled = False
frmOutput.txtOthernames.Enabled = False
frmOutput.cmbsex.Enabled = False
frmOutput.txtaddress.Enabled = False
frmOutput.txtRegno.Enabled = False
frmOutput.txtSurname = rs.Fields!Surname
frmOutput.txtOthernames = rs.Fields!Othernames
frmOutput.cmbsex = rs.Fields!Sex
frmOutput.txtaddress = rs.Fields!Address
frmOutput.txtRegno = rs.Fields!reg_no
frmOutput.DTPicker1 = rs.Fields!Date
frmOutput.txtVolume = rs.Fields!Volume
frmOutput.txtDN = rs.Fields!DN
frmOutput.txtFluid = rs.Fields!Fluid
frmOutput.lblTime = rs.Fields!Time
frmOutput.cmdSearch.Enabled = False
frmOutput.cmdSave.Enabled = False
frmOutput.cmdDelete.Enabled = False
End If
End Sub
Private Sub mnuupdatepatientdiagnosis_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientDiagnosisHistory", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmDiagnosis.Show
frmDiagnosis.txtSurname.Enabled = True
frmDiagnosis.txtOthernames.Enabled = True
frmDiagnosis.cmbsex.Enabled = True
frmDiagnosis.txtaddress.Enabled = True
frmDiagnosis.txtSurname = rs.Fields!Surname
frmDiagnosis.txtOthernames = rs.Fields!Othernames
frmDiagnosis.cmbsex = rs.Fields!Sex
frmDiagnosis.txtaddress = rs.Fields!Address
frmDiagnosis.txtRegno = rs.Fields!reg_no
frmDiagnosis.DTPicker1 = rs.Fields!DTPicker1
frmDiagnosis.txtDiagnosis = rs.Fields!Diagnosis
frmDiagnosis.txtCodeno = rs.Fields!Codeno
frmDiagnosis.cmdUpdate.Enabled = True
frmDiagnosis.cmdAdd.Enabled = False
frmDiagnosis.cmdSearch.Enabled = False
frmDiagnosis.cmdDelete.Enabled = False
c.Close
End If
End Sub

Private Sub mnuUpdatepatienthospitalhistory_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientHistory", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmHospitalHistory.Show
frmHospitalHistory.txtSurname.Enabled = False
frmHospitalHistory.txtOthernames.Enabled = False
frmHospitalHistory.cmbsex.Enabled = False
frmHospitalHistory.txtaddress.Enabled = False
frmHospitalHistory.txtRegno.Enabled = False
frmHospitalHistory.txtSurname = rs.Fields!Surname
frmHospitalHistory.txtOthernames = rs.Fields!Othernames
frmHospitalHistory.cmbsex = rs.Fields!Sex
frmHospitalHistory.txtaddress = rs.Fields!Address
frmHospitalHistory.txtRegno = rs.Fields!Regno
frmHospitalHistory.DTPicker1 = rs.Fields!DTPicker1
frmHospitalHistory.DTPicker2 = rs.Fields!DTPicker2
frmHospitalHistory.txtReferred = rs.Fields!Referred
frmHospitalHistory.txtDischarge = rs.Fields!Dischargeto
frmHospitalHistory.txtPhysician = rs.Fields!Physician
frmHospitalHistory.txtOthers = rs.Fields!Others
frmHospitalHistory.txtWard = rs.Fields!Ward
frmHospitalHistory.cmdUpdate.Enabled = True
frmHospitalHistory.cmdAdd.Enabled = False
frmHospitalHistory.cmdSearch.Visible = False
frmHospitalHistory.cmdDelete.Enabled = False
c.Close
End If
End Sub
Private Sub mnuUpdatepatientoperationrecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Registration Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientOperations", c, adOpenDynamic, adLockOptimistic
rs.Find "reg_no='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmOperations.Show
frmOperations.txtSurname.Enabled = True
frmOperations.txtOthernames.Enabled = True
frmOperations.cmbsex.Enabled = True
frmOperations.txtaddress.Enabled = True
frmOperations.txtSurname = rs.Fields!Surname
frmOperations.txtOthernames = rs.Fields!Othernames
frmOperations.cmbsex = rs.Fields!Sex
frmOperations.txtaddress = rs.Fields!Address
frmOperations.txtRegno = rs.Fields!reg_no
frmOperations.DTPicker1 = rs.Fields!Date
frmOperations.txtSurgeon = rs.Fields!Surgeon
frmOperations.txtOperation = rs.Fields!Operation
frmOperations.txtCodeno = rs.Fields!Codeno
frmOperations.cmdUpdate.Enabled = True
frmOperations.cmdSave.Enabled = False
frmOperations.cmdSearch.Enabled = False
frmOperations.cmdDelete.Enabled = False
c.Close
End If
End Sub

Private Sub mnuUpdatepulserecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Chart Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientPulse", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Registration Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmPulse.Show
frmPulse.DTPicker1.SetFocus
frmPulse.txtPatientname.Enabled = False
frmPulse.txtOccupation.Enabled = False
frmPulse.txtaddress.Enabled = False
frmPulse.txtChartno.Enabled = False
frmPulse.txtHusband.Enabled = False
frmPulse.txtPatientname = rs.Fields!Patientname
frmPulse.txtOccupation = rs.Fields!Occupation
frmPulse.txtaddress = rs.Fields!Address
frmPulse.txtHusband = rs.Fields!Husband
frmPulse.DTPicker1 = rs.Fields!Date
frmPulse.cmbHour = rs.Fields!Hour
frmPulse.DTPicker2 = rs.Fields!DOD
frmPulse.txtPulse = rs.Fields!Pulse
frmPulse.txtChartno = rs.Fields!Chartno
frmPulse.cmdUpdate.Enabled = True
frmPulse.cmdSave.Visible = False
frmPulse.cmdSearch.Visible = False
frmPulse.cmdDelete.Visible = False
c.Close
End If
End Sub

Private Sub mnuUpdatetemperaturerecords_Click()
On Error Resume Next
frmMain.Enabled = False
Timer1.Enabled = False
lbltitle.Caption = "WELCOME TO CATHOLIC HOSPITAL, ONDO."
str = InputBox("Enter the Patient Chart Number", "Update info")
c.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/HMS/crack/hms.mdb;Persist Security Info=false")
rs.Open "select * from PatientTemperature", c, adOpenDynamic, adLockOptimistic
rs.Find "ChartNo='" & str & "'"
If rs.EOF Then
MsgBox "Chart Number Not Found", vbCritical, "Error Message"
frmMain.Show
frmMain.Enabled = True
rs.MoveFirst
Else
Me.Enabled = False
frmtemp.Show
frmtemp.DTPicker1.SetFocus
frmtemp.txtPatientname.Enabled = False
frmtemp.txtOccupation.Enabled = False
frmtemp.txtaddress.Enabled = False
frmtemp.txtChartno.Enabled = False
frmtemp.txtHusband.Enabled = False
frmtemp.txtPatientname = rs.Fields!Patientname
frmtemp.txtOccupation = rs.Fields!Occupation
frmtemp.txtaddress = rs.Fields!Address
frmtemp.txtHusband = rs.Fields!Husband
frmtemp.DTPicker1 = rs.Fields!Date
frmtemp.cmbHour = rs.Fields!Hour
frmtemp.DTPicker2 = rs.Fields!DOD
frmtemp.txtTemp = rs.Fields!Temperature
frmtemp.txtChartno = rs.Fields!Chartno
frmtemp.cmdUpdate.Enabled = True
frmtemp.cmdSave.Visible = False
frmtemp.cmdSearch.Visible = False
frmtemp.cmdDelete.Visible = False
c.Close
End If
End Sub

Private Sub mnuVDiagnosis_Click()

End Sub

Private Sub mnuVDeliveryrecords_Click()
frmViewPatientDelivery.Show
Me.Enabled = False
End Sub

Private Sub mnuViewallpatient_Click()
frmViewAllPatient.Show
Me.Enabled = False
End Sub

Private Sub mnuVInputfluid_Click()
frmViewPatientInputFluid.Show
Me.Enabled = False
End Sub

Private Sub mnuVOutputfluid_Click()
frmViewPatientOutput.Show
Me.Enabled = False
End Sub

Private Sub mnuVPDiagnosis_Click()
frmViewPatientDiagnosis.Show
Me.Enabled = False
End Sub

Private Sub mnuVPulseRecords_Click()
frmViewPatientPulse.Show
Me.Enabled = False
End Sub

Private Sub mnuVRadiology_Click()
frmViewPatientRadiology.Show
Me.Enabled = False
End Sub

Private Sub Timer1_Timer()
lbltitle.Caption = Right(msg, currentLength)
currentLength = (currentLength + 1) Mod (Len(msg) + 1)
End Sub


Private Sub Timer2_Timer()
today = Now
lblTime.Caption = Format(today, "hh:mm:ss ampm")
End Sub
