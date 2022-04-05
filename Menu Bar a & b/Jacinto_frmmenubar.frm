VERSION 5.00
Begin VB.Form frmmenubar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Bar"
   ClientHeight    =   7095
   ClientLeft      =   5865
   ClientTop       =   2400
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Jacinto_frmmenubar.frx":0000
   ScaleHeight     =   7095
   ScaleWidth      =   9105
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunewproject 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopenproject 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuseparatorbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaddproject 
         Caption         =   "A&dd Project..."
      End
      Begin VB.Menu mnuremoveproject 
         Caption         =   "&Remove Project"
      End
      Begin VB.Menu mnuseparatorbar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusaveproject 
         Caption         =   "Sa&ve Project"
      End
      Begin VB.Menu mnusaveprojectas 
         Caption         =   "Sav&e Project &As..."
      End
      Begin VB.Menu mnuseparatorbar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnusaveform 
         Caption         =   "&Save Form1"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnusaveformas 
         Caption         =   "Save Form &As..."
      End
      Begin VB.Menu mnusaveselection 
         Caption         =   "Save Se&lection"
      End
      Begin VB.Menu mnusavechangescript 
         Caption         =   "Save C&hange Script"
      End
      Begin VB.Menu mnuseparatorbar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuprintsetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuseparatorbar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnumakeprojecexe 
         Caption         =   "Ma&ke Project1.exe..."
      End
      Begin VB.Menu mnumakeprojectgroup 
         Caption         =   "Make Project &Group..."
      End
      Begin VB.Menu mnuseparatorbar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu1 
         Caption         =   "&1 ...\book\TAB FUNCTIONS\Project1.vbp"
      End
      Begin VB.Menu mnud1 
         Caption         =   "&2 D:\vb ni dom\Project1.vbp"
      End
      Begin VB.Menu mnuD2 
         Caption         =   "&3 D:\ian visual basic\Project1.vbp"
      End
      Begin VB.Menu mnud3 
         Caption         =   "&4 D:\LUIS\Project1.vbp"
      End
      Begin VB.Menu mnuseparatorbar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnutoolbox 
         Caption         =   "Tool Box"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnucolorbox 
         Caption         =   "Color Box"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnustatusbar 
         Caption         =   "Status Bar"
      End
      Begin VB.Menu mnutexttoolbar 
         Caption         =   "Text Toolbar"
      End
      Begin VB.Menu mnusebaratorbar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuzoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnunormalsize 
            Caption         =   "Normal Size"
         End
         Begin VB.Menu mnulargesize 
            Caption         =   "Large Size"
         End
         Begin VB.Menu mnucustom 
            Caption         =   "Custom..."
         End
         Begin VB.Menu mnuseparatorbar9 
            Caption         =   "-"
         End
         Begin VB.Menu mnushowgrid 
            Caption         =   "Show Grid"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnushowthumbnail 
            Caption         =   "Show Thumbnail"
         End
      End
      Begin VB.Menu mnuviewbitmap 
         Caption         =   "View Bitmap "
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuproject 
      Caption         =   "&Project"
   End
   Begin VB.Menu mnuformat 
      Caption         =   "&Format"
   End
   Begin VB.Menu mnudebug 
      Caption         =   "&Debug"
   End
   Begin VB.Menu mnurun 
      Caption         =   "&Run"
   End
   Begin VB.Menu mnuquery 
      Caption         =   "Q&uery"
   End
   Begin VB.Menu mnudiagram 
      Caption         =   "D&iagram"
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuaddins 
      Caption         =   "&Add-Ins"
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmmenubar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

