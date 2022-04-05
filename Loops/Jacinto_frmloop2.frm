VERSION 5.00
Begin VB.Form frmloop2 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loops"
   ClientHeight    =   4440
   ClientLeft      =   7680
   ClientTop       =   3555
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmloop2.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   5160
   Begin VB.CommandButton cmddisplay 
      BackColor       =   &H008080FF&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   360
      ScaleHeight     =   2940
      ScaleWidth      =   4380
      TabIndex        =   0
      Top             =   240
      Width           =   4440
   End
End
Attribute VB_Name = "frmloop2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddisplay_Click()
Dim num As Integer
Dim sum As Integer
num = 0
Do While num < 10
   Picture1.Print Tab(5); sum & "+" & num & "=" & sum + num
   sum = sum + num
   num = num + 1
Loop
End Sub


