VERSION 5.00
Begin VB.Form frmMenuEditor 
   Caption         =   "Menu Editor"
   ClientHeight    =   3735
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALIGNMENT"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   8655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuseparatorbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuAlignment 
      Caption         =   "&Alignment"
      Begin VB.Menu mnuCenter 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "Right Align"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Left Align"
      End
   End
   Begin VB.Menu mnuBackcolor 
      Caption         =   "&Backcolor"
      Begin VB.Menu mnuForm 
         Caption         =   "Form"
         Begin VB.Menu mnufRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnufBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnufGreen 
            Caption         =   "Green"
         End
      End
      Begin VB.Menu mnuLabel 
         Caption         =   "Label"
         Begin VB.Menu mnuRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "Green"
         End
      End
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuBlue_Click()
lbldisplay.BackColor = vbBlue
End Sub

Private Sub mnuCenter_Click()
lbldisplay.Alignment = 2
End Sub

Private Sub mnufBlue_Click()
frmMenuEditor.BackColor = vbBlue
End Sub

Private Sub mnufGreen_Click()
frmMenuEditor.BackColor = vbGreen
End Sub

Private Sub mnufRed_Click()
frmMenuEditor.BackColor = vbRed
End Sub

Private Sub mnuGreen_Click()
lbldisplay.BackColor = vbGreen
End Sub

Private Sub mnuLeft_Click()
lbldisplay.Alignment = 0
End Sub

Private Sub mnuRed_Click()
lbldisplay.BackColor = vbRed
End Sub

Private Sub mnuRight_Click()
lbldisplay.Alignment = 1
End Sub
