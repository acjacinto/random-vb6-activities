VERSION 5.00
Begin VB.Form frmConversions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversions"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmConversions.frx":0000
   ScaleHeight     =   5430
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Inches to Feet"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton opt10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Year to Decade"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.OptionButton opt8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Minutes to Seconds"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.OptionButton opt6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Minutes to Hours"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.OptionButton opt4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Celcius to Fahrenheit"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Feet to Yard"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton opt9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Seconds to Minutes"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.OptionButton opt7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Hours to Minutes"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton opt5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Area of square"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton opt3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fahrenheit to Celcius"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton cmdconvert 
      BackColor       =   &H0080FFFF&
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtnum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label lblnote 
      BackStyle       =   0  'Transparent
      Caption         =   $"Jacinto_frmConversions.frx":AB64
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   3975
   End
End
Attribute VB_Name = "frmConversions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer

Private Sub cmdconvert_Click()
num = txtnum.Text
If opt1.Value = True Then
   lbldisplay.Caption = (num / 12) & " ft"
ElseIf opt2.Value = True Then
   lbldisplay.Caption = (num * 0.3333333) & " yd"
ElseIf opt3.Value = True Then
   lbldisplay.Caption = (num - 32) * 5 / 9 & " C"
ElseIf opt4.Value = True Then
   lbldisplay.Caption = (num * 9) / 5 + 32 & " F"
ElseIf opt5.Value = True Then
   lbldisplay.Caption = (num * num)
ElseIf opt6.Value = True Then
   lbldisplay.Caption = (num / 60) & " hr"
ElseIf opt7.Value = True Then
   lbldisplay.Caption = (num * 60) & " min"
ElseIf opt8.Value = True Then
   lbldisplay.Caption = (num * 60) & " sec"
ElseIf opt9.Value = True Then
   lbldisplay.Caption = (num / 60) & " min"
ElseIf opt10.Value = True Then
   lbldisplay.Caption = (num * 10) & " Decade"
   
End If
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
If KeyAscii >= 33 And KeyAscii <= 44 Or KeyAscii >= 46 And KeyAscii <= 47 Or KeyAscii >= 65 And KeyAscii <= 122 Then
KeyAscii = 0
End If
End Sub
