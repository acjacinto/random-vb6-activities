VERSION 5.00
Begin VB.Form frmOrderList 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order List"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcompute 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Compute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Frame fmelist 
      BackColor       =   &H00C0E0FF&
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox chkchicken 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Chicken w/ Spaghetti"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   3360
         Width           =   3495
      End
      Begin VB.CheckBox chkspaghetti 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Spaghetti"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   5
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox chkdouble 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Double Cheese Burger"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   3375
      End
      Begin VB.CheckBox chkdrinks 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Soft Drinks"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkfries 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fries"
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
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CheckBox chkburger 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Burger"
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
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Label lbldisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   5655
   End
End
Attribute VB_Name = "frmOrderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcompute_Click()
Dim dburger As Integer
Dim dfries As Integer
Dim ddrinks As Integer
Dim ddouble As Integer
Dim dspaghetti As Integer
Dim dchickenspag As Integer

Const burger = 30
Const fries = 25
Const drinks = 20
Const dblcheese = 90
Const Spaghetti = 59
Const chickenspag = 100


If chkburger.Value = 1 Then
   dburger = burger
End If

If chkfries.Value = 1 Then
   dfries = fries
End If

If chkdrinks.Value = 1 Then
   ddrinks = drinks
End If

If chkdouble.Value = 1 Then
   ddouble = dblcheese
End If

If chkspaghetti.Value = 1 Then
   dspaghetti = Spaghetti
End If

If chkchicken.Value = 1 Then
   dchickenspag = chickenspag
End If

lbldisplay.Caption = FormatCurrency(dburger + dfries + ddrinks + ddouble + dspaghetti + dchickenspag, 2)
   
End Sub
