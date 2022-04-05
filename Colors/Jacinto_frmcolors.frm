VERSION 5.00
Begin VB.Form frmcolors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colors"
   ClientHeight    =   4260
   ClientLeft      =   7680
   ClientTop       =   3555
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmcolors.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   5895
   Begin VB.TextBox fcolor 
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
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox bcolor 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton cmdchange 
      BackColor       =   &H008080FF&
      Caption         =   "Change Color"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Coloring Program"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblfcolor 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Font Color:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblbcolor 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back Color:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
End
Attribute VB_Name = "frmcolors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdchange_Click()
fcolor = UCase(fcolor.Text)
bcolor = UCase(bcolor.Text)

Select Case fcolor
     Case "RED"
       lblfcolor.ForeColor = vbRed
     Case "BLUE"
       lblfcolor.ForeColor = vbBlue
     Case "GREEN"
       lblfcolor.ForeColor = vbGreen
     Case "WHITE"
       lblfcolor.ForeColor = vbWhite
     Case "BLACK"
       lblfcolor.ForeColor = vbBlack
     Case "YELLOW"
       lblfcolor.ForeColor = vbYellow
     Case Else
       lblfcolor.Caption = "COLOR IS NOT AVAILABLE"
End Select

Select Case bcolor
     Case "RED"
       lblbcolor.BackColor = vbRed
     Case "BLUE"
       lblbcolor.BackColor = vbBlue
     Case "GREEN"
       lblbcolor.BackColor = vbGreen
     Case "WHITE"
       lblbcolor.BackColor = vbWhite
     Case "BLACK"
       lblbcolor.BackColor = vbBlack
     Case "YELLOW"
       lblbcolor.BackColor = vbYellow
     Case Else
       lblbcolor.Caption = "COLOR IS NOT AVAILABLE"
End Select
End Sub


