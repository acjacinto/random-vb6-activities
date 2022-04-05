VERSION 5.00
Begin VB.Form frmAverageGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Average Grade"
   ClientHeight    =   5565
   ClientLeft      =   7485
   ClientTop       =   3150
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmAverageGrade.frx":0000
   ScaleHeight     =   5565
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H000080FF&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox txtg3 
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtg2 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtg1 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Grade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lblg2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblg3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblg1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grade 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Average Grade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmAverageGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcompute_Click()
Dim g1 As Integer
Dim g2 As Integer
Dim g3 As Integer
Dim average As Single
g1 = Val(txtg1.Text)
g2 = Val(txtg2.Text)
g3 = Val(txtg3.Text)
average = (g1 + g2 + g3) / 3
If average >= 90 Then
   lbldisplay.Caption = "A"
ElseIf average >= 80 Then
   lbldisplay.Caption = "B"
ElseIf average >= 70 Then
   lbldisplay.Caption = "C"
ElseIf average >= 60 Then
   lbldisplay.Caption = "D"
Else
   lbldisplay.Caption = "F"
End If
End Sub


