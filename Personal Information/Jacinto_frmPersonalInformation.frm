VERSION 5.00
Begin VB.Form frmPersonalInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Information "
   ClientHeight    =   3510
   ClientLeft      =   6480
   ClientTop       =   3945
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmPersonalInformation.frx":0000
   ScaleHeight     =   3510
   ScaleWidth      =   6600
   Begin VB.TextBox txtnumber 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label lblphonenumber 
      BackColor       =   &H00FFFF80&
      Caption         =   "Phone Number:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblname 
      BackColor       =   &H00FFFF80&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmPersonalInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 64 And KeyAscii <> 46 Then
KeyAscii = 0
End If
End Sub

Private Sub txtnumber_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 44 Or KeyAscii >= 46 _
And KeyAscii <= 47 Or KeyAscii >= 65 And KeyAscii <= 122 Then
KeyAscii = 0
End If
End Sub
