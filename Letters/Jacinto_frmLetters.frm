VERSION 5.00
Begin VB.Form frmLetters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letters"
   ClientHeight    =   4200
   ClientLeft      =   7080
   ClientTop       =   3750
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmLetters.frx":0000
   ScaleHeight     =   4200
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdevaluate 
      BackColor       =   &H0080FFFF&
      Caption         =   "EVALUATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtletter 
      Alignment       =   2  'Center
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter Letter:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vowel or Consonats"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdevaluate_Click()
txtletter.Text = UCase(txtletter)
Select Case txtletter
     Case "A"
         lbldisplay.Caption = "Letter is a Vowel"
     Case "E"
         lbldisplay.Caption = "Letter is a Vowel"
     Case "I"
         lbldisplay.Caption = "Letter is a Vowel"
     Case "O"
         lbldisplay.Caption = "Letter is a Vowel"
     Case "U"
         lbldisplay.Caption = "Letter is a Vowel"
     Case Else
         lbldisplay.Caption = "Letter is a Consonant"
End Select
End Sub

Private Sub txtletter_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 64 Then
KeyAscii = 0
End If
End Sub
