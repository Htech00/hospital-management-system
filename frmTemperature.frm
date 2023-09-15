VERSION 5.00
Begin VB.Form frmTemperature 
   Caption         =   "Form4"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   LinkTopic       =   "Form4"
   Picture         =   "frmTemperature.frx":0000
   ScaleHeight     =   7575
   ScaleLeft       =   50
   ScaleMode       =   0  'User
   ScaleWidth      =   13545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   4080
      TabIndex        =   0
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TEMPERATURE RECORDS"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmTemperature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
