VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLoadingForm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Form"
   ClientHeight    =   4335
   ClientLeft      =   6435
   ClientTop       =   3450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7440
   Begin VB.Timer tmr 
      Interval        =   100
      Left            =   360
      Top             =   1680
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   4095
   End
End
Attribute VB_Name = "frmLoadingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmr_Timer()
ProgressBar.Value = ProgressBar.Value + 1
  If ProgressBar.Value = 10 Then
     lbldisplay.Caption = "Loading."
  ElseIf ProgressBar.Value = 20 Then
     lbldisplay.Caption = "Loading.."
  ElseIf ProgressBar.Value = 30 Then
     lbldisplay.Caption = "Loading..."
  ElseIf ProgressBar.Value = 40 Then
     lbldisplay.Caption = "Initializing."
  ElseIf ProgressBar.Value = 50 Then
     lbldisplay.Caption = "Initializing.."
  ElseIf ProgressBar.Value = 60 Then
     lbldisplay.Caption = "Initializing..."
  ElseIf ProgressBar.Value = 70 Then
     lbldisplay.Caption = "Please Wait."
  ElseIf ProgressBar.Value = 80 Then
     lbldisplay.Caption = "Please Wait.."
  ElseIf ProgressBar.Value = 90 Then
     lbldisplay.Caption = "Please Wait..."
  ElseIf ProgressBar.Value = 100 Then
     lbldisplay.Caption = "Loading Successful"
     tmr.Enabled = False
     frmLoadingForm.Hide
     frmTestwithTimer.Show
  End If
End Sub
