VERSION 5.00
Begin VB.Form frmFootToInches 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foot to Inches"
   ClientHeight    =   3345
   ClientLeft      =   7275
   ClientTop       =   3945
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFootToInches.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Compute"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtfeet 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblinch 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inches"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblfeet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Number of Feet Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblinches 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
End
Attribute VB_Name = "frmFootToInches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcompute_Click()
Dim feet As Integer
feet = txtfeet
lblinches.Caption = feet * 12
End Sub
