VERSION 5.00
Begin VB.Form frmQuiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pop Quiz"
   ClientHeight    =   7920
   ClientLeft      =   5670
   ClientTop       =   1770
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmPopQuiz.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   9270
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H000080FF&
      Caption         =   "Submit Answer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   1
      Left            =   4680
      TabIndex        =   20
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox txtans7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   33
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PBA star who is the husband of Kris Aquino"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtans1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "In Los Angeles Lakers, who wears jersey number 24?"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.TextBox txtans2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   8
      Left            =   4680
      TabIndex        =   8
      Top             =   5400
      Width           =   4455
      Begin VB.TextBox txtans10 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Mother of Philippines Democracy"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   7
      Left            =   4680
      TabIndex        =   7
      Top             =   4080
      Width           =   4455
      Begin VB.TextBox txtans9 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "House where the American President resides"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   6
      Left            =   4680
      TabIndex        =   6
      Top             =   2760
      Width           =   4455
      Begin VB.TextBox txtans8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Movie from India that won Oscar best movie of 2009"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   4455
      Begin VB.TextBox txtans5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "2009 NBA Eastern Conference Champion"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4455
      Begin VB.TextBox txtans4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Team where Lebron James played"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4455
      Begin VB.TextBox txtans3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "She sang love story and you belong with me"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Filipino American that stars in the show High School Musical"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.TextBox txtans6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Question #6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "FIlipino World Boxing Champion"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label lblremark 
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
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label lblcorrect 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label lblincorrect 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   24
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrect Answer:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answer:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   6840
      Width           =   1455
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SCORE As Integer
Dim INCORRECT As Integer
Private Sub cmdsubmit_Click()

If txtans1 = "KOBE BRYANT" Or txtans1 = "KOBE" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans2 = "VANESSA HUDGENS" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans3 = "TAYLOR SWIFT" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans4 = "CLEVELAND" Or txtans4 = "CAVALIERS" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans5 = "MAGIC" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans6 = "MANNY PAQUIAO" Or txtans6 = "PAQUIAO" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans7 = "JAMES YAP" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans8 = "SLUMDOG MILLIONAIRE" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans9 = "WHITE HOUSE" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


If txtans10 = "CORY" Or txtans10 = "CORAZON AQUINO" Then
   SCORE = SCORE + 1
Else
   INCORRECT = INCORRECT + 1
End If


cmdsubmit.Enabled = False
lblcorrect.Caption = SCORE
lblincorrect.Caption = INCORRECT


If SCORE >= 9 Then
   lblremark.Caption = "EXELLENT"
ElseIf SCORE >= 7 Then
   lblremark.Caption = "VERYGOOD"
ElseIf SCORE >= 5 Then
   lblremark.Caption = "GOOD"
ElseIf SCORE >= 3 Then
   lblremark.Caption = "NEEDS IMPROVEMENT"
Else
   lblremark.Caption = "FAILED"
End If
End Sub


Private Sub txtans1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans1.Text = UCase(txtans1)
txtans2.SetFocus
End If
End Sub

Private Sub txtans10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans10.Text = UCase(txtans10)
End If
End Sub

Private Sub txtans2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans2.Text = UCase(txtans2)
txtans3.SetFocus
End If
End Sub

Private Sub txtans3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans3.Text = UCase(txtans3)
txtans4.SetFocus
End If
End Sub

Private Sub txtans4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans4.Text = UCase(txtans4)
txtans5.SetFocus
End If
End Sub

Private Sub txtans5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans5.Text = UCase(txtans5)
txtans6.SetFocus
End If
End Sub

Private Sub txtans6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans6.Text = UCase(txtans6)
txtans7.SetFocus
End If
End Sub

Private Sub txtans7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans7.Text = UCase(txtans7)
txtans8.SetFocus
End If
End Sub

Private Sub txtans8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans8.Text = UCase(txtans8)
txtans9.SetFocus
End If
End Sub

Private Sub txtans9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtans9.Text = UCase(txtans9)
txtans10.SetFocus
End If
End Sub
