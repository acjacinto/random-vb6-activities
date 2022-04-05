VERSION 5.00
Begin VB.Form frmPoundLost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pound Lost"
   ClientHeight    =   7080
   ClientLeft      =   7680
   ClientTop       =   2160
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmPoundLost.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   4950
   Begin VB.CommandButton lblcompute 
      BackColor       =   &H000080FF&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "Century Gothic"
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
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtrunning 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox txtbiking 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtbasketball 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label lblrunning 
      BackStyle       =   0  'Transparent
      Caption         =   "Running"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblbiking 
      BackStyle       =   0  'Transparent
      Caption         =   "Biking"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblbasketball 
      BackStyle       =   0  'Transparent
      Caption         =   "Basketball"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblhours 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblactivities 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Activities"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Burning Calories"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmPoundLost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblcompute_Click()
Dim actbasketball As Integer
Dim actbiking As Integer
Dim actrunning As Integer
Dim pounds As Integer
Const calories = 3500
Const basketball = 300
Const biking = 400
Const running = 500
actbasketball = Val(txtbasketball.Text) * basketball
actbiking = Val(txtbiking.Text) * biking
actrunning = Val(txtrunning.Text) * running
pounds = (actbasketball + actbiking + actrunning) / calories
lbldisplay.Caption = pounds & " Pounds"
End Sub

