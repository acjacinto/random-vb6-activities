VERSION 5.00
Begin VB.Form frmtax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Value Added Tax "
   ClientHeight    =   4665
   ClientLeft      =   7875
   ClientTop       =   3150
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmtax.frx":0000
   ScaleHeight     =   4665
   ScaleWidth      =   5145
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H0080FF80&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtprice 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtitem 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblamountwithvat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount w/ VAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblprice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblamount 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   2760
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblitem 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Value Added Tax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmtax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcompute_Click()
Dim price As Integer
Dim taxedamount As Integer
Const vat = 0.12
price = Val(txtprice.Text)
taxedamount = price * vat
lblamount.Caption = taxedamount + price
End Sub
