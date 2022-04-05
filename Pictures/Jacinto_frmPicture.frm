VERSION 5.00
Begin VB.Form frmPicture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmPicture.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfind 
      BackColor       =   &H0080C0FF&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   6960
      Width           =   4095
   End
   Begin VB.CommandButton cmdclear 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.ComboBox cmbpic 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2400
      TabIndex        =   2
      Text            =   "Choose a Picture"
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton cmdshow 
      BackColor       =   &H0080FFFF&
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.PictureBox picdisplay 
      Height          =   4095
      Left            =   960
      ScaleHeight     =   4035
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   2160
      Width           =   7215
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdclear_Click()
picdisplay.Picture = LoadPicture("")
Set picdisplay = Nothing
End Sub

Private Sub cmdfind_Click()
picdisplay.Picture = LoadPicture("C:\Users\AC\Pictures\Pics\" & txtname.Text & ".jpg")
End Sub

Private Sub cmdshow_Click()
picdisplay.Picture = LoadPicture("C:\Users\AC\Pictures\Pics\" & cmbpic.Text & ".jpg")
End Sub

Private Sub Form_Load()
cmbpic.AddItem "Baby Rose"
cmbpic.AddItem "Bird"
cmbpic.AddItem "Cat"
cmbpic.AddItem "Dog"
cmbpic.AddItem "Fox"
cmbpic.AddItem "Panda"
cmbpic.AddItem "Penguin"
cmbpic.AddItem "Pig"
cmbpic.AddItem "Rabbit"
cmbpic.AddItem "Seal"
cmbpic.AddItem "Tiger"
End Sub

