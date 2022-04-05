VERSION 5.00
Begin VB.Form frmCase 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Case Grade"
   ClientHeight    =   6120
   ClientLeft      =   7275
   ClientTop       =   2760
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmCase.frx":0000
   ScaleHeight     =   6120
   ScaleWidth      =   6030
   Begin VB.TextBox pgrade 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox mgrade 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox fgrade 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label lblgrade 
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
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label lbleg 
      BackStyle       =   0  'Transparent
      Caption         =   "Equivalent Grade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Average Grade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblp1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prelim Grades:"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblf3 
      BackStyle       =   0  'Transparent
      Caption         =   "Finals Grade:"
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
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblm2 
      BackStyle       =   0  'Transparent
      Caption         =   "Midterm Grade:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblfg 
      BackStyle       =   0  'Transparent
      Caption         =   "FGrade:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblfgrade 
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
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   3480
      Width           =   2895
   End
End
Attribute VB_Name = "frmCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcompute_Click()
Dim tgrade As Integer
Const prelim = 0.3
Const midterm = 0.3
Const finals = 0.4

tgrade = (Val(pgrade) * prelim) + (Val(mgrade) * midterm) + (Val(fgrade) * finals)
lblfgrade.Caption = tgrade

Select Case lblfgrade
     Case Is >= 98
        lblgrade.Caption = "1.00"
     Case Is >= 95
        lblgrade.Caption = "1.25"
     Case Is >= 91
        lblgrade.Caption = "1.50"
     Case Is >= 88
        lblgrade.Caption = "1.75"
     Case Is >= 85
        lblgrade.Caption = "2.00"
     Case Is >= 82
        lblgrade.Caption = "2.25"
     Case Is >= 80
        lblgrade.Caption = "2.50"
     Case Is >= 77
        lblgrade.Caption = "2.75"
     Case Is >= 75
        lblgrade.Caption = "3.00"
     Case Is >= 70
        lblgrade.Caption = "4.00"
     Case Is < 70
        lblgrade.Caption = "5.00"
End Select
End Sub
