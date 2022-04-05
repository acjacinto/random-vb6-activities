VERSION 5.00
Begin VB.Form frmburgerhouse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCJB Burger House"
   ClientHeight    =   8115
   ClientLeft      =   6870
   ClientTop       =   1965
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmburgerhouse.frx":0000
   ScaleHeight     =   8115
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdcompchange 
      BackColor       =   &H000040C0&
      Caption         =   "Compute for Change"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H000080FF&
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtpayment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   29
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Frame fmesoftdrinks 
      BackColor       =   &H0080C0FF&
      Caption         =   "Softdrinks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   6255
      Begin VB.TextBox txtlarges 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   24
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtmeds 
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
         Height          =   495
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtregs 
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
         Height          =   495
         Left            =   1680
         TabIndex        =   22
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbllarges 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblmeds 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblregs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblnumorders3 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame fmefries 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   6255
      Begin VB.TextBox txtxlargef 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtlargef 
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
         Height          =   495
         Left            =   3360
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtregf 
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
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.Label txtxlarge 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exrta Large"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label txtlarge1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label txtreg1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblnumorders2 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame fmeburgers 
      BackColor       =   &H0080C0FF&
      Caption         =   "Burgers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtspecialb 
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
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtcheeseb 
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
         Height          =   495
         Left            =   3360
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtregb 
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
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblspecialb 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special Burger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblcheeseb 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cheese Burger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblregb 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Regular Burger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblnumorders1 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Orders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Label lblchange 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   30
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label lblamount 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   2040
      TabIndex        =   28
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lblchge 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
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
      Left            =   240
      TabIndex        =   27
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label lblpayment 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment:"
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
      Left            =   240
      TabIndex        =   26
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lbltotalA 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
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
      Left            =   240
      TabIndex        =   25
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmburgerhouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim regularb As Single
Dim cheeseb As Integer
Dim specialb As Integer
Dim regfries As Integer
Dim largefries As Integer
Dim exfries As Integer
Dim regdrinks As Integer
Dim meddrinks As Integer
Dim largedrinks As Integer
Dim total As Single
Dim burgers As Single
Dim fries As Integer
Dim softdrinks As Integer

Const regburger = 27.5
Const cheeseburger = 45
Const specialburger = 70
Const regularfries = 25
Const lfries = 37
Const extrafries = 45
Const regulardrinks = 20
Const mediumdrinks = 30
Const ldrinks = 35

Private Sub cmdcompchange_Click()
Dim change As Single
Dim payment As Integer

regularb = Val(txtregb.Text) * regburger
cheeseb = Val(txtcheeseb.Text) * cheeseburger
specialb = Val(txtspecialb.Text) * specialburger
regfries = Val(txtregf.Text) * regularfries
largefries = Val(txtlargef.Text) * lfries
exfries = Val(txtxlargef.Text) * extrafries
regdrinks = Val(txtregs.Text) * regulardrinks
meddrinks = Val(txtmeds.Text) * mediumdrinks
largedrinks = Val(txtlarges.Text) * ldrinks

burgers = regularb + cheeseb + specialb
fries = regfries + largefries + exfries
softdrinks = regdrinks + meddrinks + largedrinks
total = burgers + fries + softdrinks

lblamount.Caption = total
payment = txtpayment.Text
change = payment - total

lblchange.Caption = change
cmdcompute.Enabled = True
cmdcompchange.Enabled = False
End Sub
Private Sub cmdcompute_Click()
regularb = Val(txtregb.Text) * regburger
cheeseb = Val(txtcheeseb.Text) * cheeseburger
specialb = Val(txtspecialb.Text) * specialburger
regfries = Val(txtregf.Text) * regularfries
largefries = Val(txtlargef.Text) * lfries
exfries = Val(txtxlargef.Text) * extrafries
regdrinks = Val(txtregs.Text) * regulardrinks
meddrinks = Val(txtmeds.Text) * mediumdrinks
largedrinks = Val(txtlarges.Text) * ldrinks

burgers = regularb + cheeseb + specialb
fries = regfries + largefries + exfries
softdrinks = regdrinks + meddrinks + largedrinks
total = burgers + fries + softdrinks

lblamount.Caption = total
cmdcompute.Enabled = False
cmdcompchange.Enabled = True
End Sub


