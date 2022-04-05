VERSION 5.00
Begin VB.Form frmloop1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loops"
   ClientHeight    =   3030
   ClientLeft      =   7875
   ClientTop       =   4155
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmloop1.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4830
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
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
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmloop1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmddisplay_Click()
Dim num As Integer
num = 1
Do While num <= 10
   Picture1.Print num;
   num = num + 1
Loop
End Sub
