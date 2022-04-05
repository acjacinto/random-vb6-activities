VERSION 5.00
Begin VB.Form frmStopWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Improvised Stop Watch"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmStopWatch.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrwatch 
      Interval        =   1000
      Left            =   1320
      Top             =   4920
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdreset 
      BackColor       =   &H0080FFFF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdstart 
      BackColor       =   &H0080FF80&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdstop 
      BackColor       =   &H008080FF&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lbltitle3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "SECONDS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lbltitle2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "MINUTES"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lbltitle1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "HOURS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblhyper2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblhyper1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblsecond 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblminute 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblhour 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdquit_Click()
If MsgBox("Are you sure to terminate this program?", vbExclamation + vbYesNo) = vbYes Then
End
End If
End Sub

Private Sub cmdreset_Click()
lblsecond.Caption = "00"
lblminute.Caption = "00"
lblhour.Caption = "00"
End Sub

Private Sub cmdstart_Click()
tmrwatch.Enabled = True
End Sub

Private Sub cmdstop_Click()
tmrwatch.Enabled = False
End Sub

Private Sub Form_Load()
tmrwatch.Enabled = False
End Sub


Private Sub tmrwatch_Timer()
   lblsecond.Caption = Val(lblsecond.Caption) + 1
   lblsecond.Caption = Format(lblsecond.Caption, "00")
If Val(lblsecond.Caption) = 60 Then
   lblminute.Caption = Val(lblminute.Caption) + 1
   lblminute.Caption = Format(lblminute.Caption, "00")
   lblsecond.Caption = 0
If Val(lblminute.Caption) = 60 Then
   lblhour.Caption = Val(lblhour.Caption) + 1
   lblhour.Caption = Format(lblhour.Caption, "00")
   lblminute.Caption = 0
End If
End If
End Sub
