VERSION 5.00
Begin VB.Form frmSlotMachine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slot Machine"
   ClientHeight    =   5085
   ClientLeft      =   7080
   ClientTop       =   3150
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmSlotMachine.frx":0000
   ScaleHeight     =   5085
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdroll 
      BackColor       =   &H0000C0C0&
      Caption         =   "Roll"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Third 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Second 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label First 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Slot Machine"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmSlotMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdroll_Click()
   First.Caption = Int(4 * Rnd) + 1
   Second.Caption = Int(4 * Rnd) + 1
   Third.Caption = Int(4 * Rnd) + 1
If First.Caption = Second.Caption And First.Caption = Third.Caption Then
   lbldisplay.Caption = "You Won"
Else
   lbldisplay.Caption = "Try Again"
End If
End Sub

