VERSION 5.00
Begin VB.Form frmconvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Conversion"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmconvert.frx":0000
   ScaleHeight     =   4050
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddisplaymin 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox txthours 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lbldisplaymin 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Converted Hours to Minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblhours 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Hours:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "frmconvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddisplaymin_Click()
Dim hours As Integer
Dim minutes As Integer

hours = Val(txthours.Text)
minutes = hours * 60
lbldisplaymin.Caption = minutes
End Sub

