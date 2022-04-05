VERSION 5.00
Begin VB.Form frmYearLevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High School Level"
   ClientHeight    =   4305
   ClientLeft      =   7275
   ClientTop       =   3750
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmYearLevel.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   6225
   Begin VB.TextBox txtnum 
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
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton cmdevaluate 
      BackColor       =   &H008080FF&
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label lblsubtitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Number:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   5175
   End
End
Attribute VB_Name = "frmYearLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdevaluate_Click()
Select Case txtnum
     Case "1"
         lbldisplay.Caption = "First Year"
     Case "2"
         lbldisplay.Caption = "Second Year"
     Case "3"
         lbldisplay.Caption = "Third Year"
     Case "4"
         lbldisplay.Caption = "Fourth Year"
     Case Else
         lbldisplay.Caption = "Unrecognized Year Level"
End Select
End Sub
