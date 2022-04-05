VERSION 5.00
Begin VB.Form frmTestwithTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test with Timer"
   ClientHeight    =   3960
   ClientLeft      =   6630
   ClientTop       =   3450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmLoadingForm2.frx":0000
   ScaleHeight     =   3960
   ScaleWidth      =   7335
   Begin VB.TextBox txtuser 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox txtpass 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H000080FF&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H000000FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lbluser 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "frmTestwithTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

