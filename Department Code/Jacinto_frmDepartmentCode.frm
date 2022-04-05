VERSION 5.00
Begin VB.Form frmDepartmentCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Code"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmDepartmentCode.frx":0000
   ScaleHeight     =   4095
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdenter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Enter Department Code"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lbldepartment 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   4935
   End
End
Attribute VB_Name = "frmDepartmentCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdenter_Click()
Dim code As Integer
code = InputBox("Enter Code Here:", "System")

Select Case code

Case "101"
     lbldepartment.Caption = "Marketing"
Case "102"
     lbldepartment.Caption = "Advertising"
Case "103"
     lbldepartment.Caption = "Sales"
Case "104"
     lbldepartment.Caption = "Production"
Case "105"
     lbldepartment.Caption = "Budget and Finance"
Case "106"
     lbldepartment.Caption = "Accounting"
Case "107"
     lbldepartment.Caption = "Personnel"
Case Else
     MsgBox "No record found", vbExclamation + vbOKOnly, "Notice"
     
End Select
End Sub
