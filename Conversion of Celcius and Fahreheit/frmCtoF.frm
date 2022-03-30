VERSION 5.00
Begin VB.Form frmCtoF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Celcius to Fahrenheit"
   ClientHeight    =   3390
   ClientLeft      =   7485
   ClientTop       =   3750
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCtoF.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   5070
   Begin VB.CommandButton cmdfahrenheit 
      BackColor       =   &H00FF0000&
      Caption         =   "Convert to Fahrenheit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdcelcius 
      BackColor       =   &H000000FF&
      Caption         =   "Convert to Celcius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txttemp 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblconvert 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lbltemp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmCtoF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcelcius_Click()
Dim temp As Integer
temp = txttemp.Text
lblconvert.Caption = temp - 32 * 5 / 9
cmdfahrenheit.Visible = True
cmdcelcius.Visible = False
End Sub

Private Sub cmdfahrenheit_Click()
Dim temp As Integer
temp = txttemp.Text
lblconvert.Caption = temp * 9 / 5 + 32
cmdfahrenheit.Visible = False
cmdcelcius.Visible = True
End Sub


