VERSION 5.00
Begin VB.Form frmOperator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Bar  "
   ClientHeight    =   3615
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmOperator.frx":0000
   ScaleHeight     =   3615
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtnum2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3960
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtnum1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7560
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblequals 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbloperator 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2400
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblname 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   9495
   End
   Begin VB.Menu mnuoperators 
      Caption         =   "Operators"
      Begin VB.Menu mnuaddition 
         Caption         =   "Addition"
      End
      Begin VB.Menu mnusubtraction 
         Caption         =   "Subtraction"
      End
      Begin VB.Menu mnumultiplication 
         Caption         =   "Multiplication"
      End
      Begin VB.Menu mnudivision 
         Caption         =   "Division"
      End
      Begin VB.Menu mnuseparatorbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
      Begin VB.Menu mnuprogram 
         Caption         =   "Program"
      End
      Begin VB.Menu mnuseparatorbar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuauthor 
         Caption         =   "Author"
      End
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuaddition_Click()
lblname.Caption = "Addition"
lbloperator.Caption = "+"
total = Val(txtnum1) + Val(txtnum2)
lbltotal.Caption = total
End Sub

Private Sub mnuauthor_Click()
MsgBox "Created By: Alyda Caye Jacinto, BSCS-II", vbInformation + vbOKOnly, "Author"
End Sub

Private Sub mnudivision_Click()
lblname.Caption = "Division"
lbloperator.Caption = "÷"
If Val(txtnum2.Text) = 0 Then
   MsgBox "You cannot use zero as your divisor", vbOKOnly + vbInformation, "Menu Editor"
Else
total = Val(txtnum1) / Val(txtnum2)
lbltotal.Caption = total
End If
End Sub

Private Sub mnumultiplication_Click()
lblname.Caption = "Multiplication"
lbloperator.Caption = "X"
total = Val(txtnum1) * Val(txtnum2)
lbltotal.Caption = total
End Sub

Private Sub mnuprogram_Click()
MsgBox "This program tends to compute the two inputed numbers with different mathematical operators", vbInformation + vbOKOnly, "About"
End Sub

Private Sub mnusubtraction_Click()
lblname.Caption = "Subtraction"
lbloperator.Caption = "-"
total = Val(txtnum1) - Val(txtnum2)
lbltotal.Caption = total
End Sub
