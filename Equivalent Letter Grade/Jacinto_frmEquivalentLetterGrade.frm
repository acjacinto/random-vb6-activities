VERSION 5.00
Begin VB.Form frmEquivalentLetterGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equivalent Letter Grade"
   ClientHeight    =   4500
   ClientLeft      =   7275
   ClientTop       =   3555
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmEquivalentLetterGrade.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   5895
   Begin VB.TextBox txtgrade 
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
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton cmdevaluate 
      BackColor       =   &H000080FF&
      Caption         =   "Evaluate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Grade's Remark"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblsubtitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Grade Here:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   5175
   End
End
Attribute VB_Name = "frmEquivalentLetterGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdevaluate_Click()
If Val(txtgrade) >= 90 Then
          lbldisplay.Caption = "A"
   ElseIf Val(txtgrade) >= 85 Then
          lbldisplay.Caption = "B"
   ElseIf Val(txtgrade) >= 80 Then
          lbldisplay.Caption = "C"
   ElseIf Val(txtgrade) >= 75 Then
          lbldisplay.Caption = "D"
   Else
          lbldisplay.Caption = "F"
End If
End Sub
