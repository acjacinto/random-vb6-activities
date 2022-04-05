VERSION 5.00
Begin VB.Form frmRomanNumerals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roman Numerals"
   ClientHeight    =   4320
   ClientLeft      =   7275
   ClientTop       =   3555
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmRomanNumerals.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdevaluate 
      BackColor       =   &H000080FF&
      Caption         =   "EVALUATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtnum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label lblsubtitle 
      BackColor       =   &H0080C0FF&
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Roman Numerals"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmRomanNumerals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdevaluate_Click()
Select Case txtnum
     Case "1"
         lbldisplay.Caption = "I"
     Case "5"
         lbldisplay.Caption = "V"
     Case "10"
         lbldisplay.Caption = "X"
     Case "50"
         lbldisplay.Caption = "L"
     Case "100"
         lbldisplay.Caption = "C"
     Case "500"
         lbldisplay.Caption = "D"
     Case "1000"
         lbldisplay.Caption = "M"
     Case Else
         lbldisplay.Caption = "Unrecognized Roman Numeral"
End Select
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 44 Or KeyAscii >= 46 And KeyAscii <= 47 Or KeyAscii >= 65 And KeyAscii <= 122 Then
KeyAscii = 0
End If
End Sub
