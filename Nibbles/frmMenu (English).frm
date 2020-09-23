VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Menu's"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuPauze 
      Caption         =   "&Pause"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSpel1 
      Caption         =   "Game"
      Begin VB.Menu mnuSpel 
         Caption         =   "New game"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "Close game"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "Highscore"
         Index           =   3
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "Leveleditor"
         Index           =   5
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSpel 
         Caption         =   "Exit"
         Index           =   7
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuOpties1 
      Caption         =   "Options"
      Begin VB.Menu mnuOptie 
         Caption         =   "Speed"
         Index           =   0
         Begin VB.Menu mnuSnelheid 
            Caption         =   "Slow"
            Index           =   0
         End
         Begin VB.Menu mnuSnelheid 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuSnelheid 
            Caption         =   "Fast"
            Index           =   2
         End
      End
      Begin VB.Menu mnuOptie 
         Caption         =   "Don't place food along the wall"
         Index           =   1
      End
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "File"
      Begin VB.Menu mnuEditor1 
         Caption         =   "New map"
         Index           =   0
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "Open map..."
         Index           =   1
         Begin VB.Menu mnuOpenen 
            Caption         =   "[level]"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "Save map"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "Remove map..."
         Index           =   5
         Begin VB.Menu mnuVerwijderen 
            Caption         =   "[level]"
            Index           =   0
         End
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuEditor1 
         Caption         =   "Exit"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuEditor1_Click(Index As Integer)
    frmLevelEditor.mnuEditor1_Click Index
End Sub

Private Sub mnuOpenen_Click(Index As Integer)
    frmLevelEditor.mnuOpenen_Click Index
End Sub

Private Sub mnuOptie_Click(Index As Integer)
    frmNibbles.mnuOptie_Click Index
End Sub

Private Sub mnuPauze_Click()
    frmNibbles.mnuPauze_Click
End Sub

Private Sub mnuSnelheid_Click(Index As Integer)
    frmNibbles.mnuSnelheid_Click Index
End Sub

Private Sub mnuSpel_Click(Index As Integer)
    frmNibbles.mnuSpel_Click Index
End Sub

Private Sub mnuVerwijderen_Click(Index As Integer)
    frmLevelEditor.mnuVerwijderen_Click Index
End Sub
