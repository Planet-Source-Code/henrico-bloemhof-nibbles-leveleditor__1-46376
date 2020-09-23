VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLevelEditor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Leveleditor - [Lege map - level 0.dat]"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   300
   ClientWidth     =   7950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLevelEditor (English).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00F3DED6&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   390
      ScaleHeight     =   2295
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   2580
      Width           =   495
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   240
         Picture         =   "frmLevelEditor (English).frx":0CCA
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   0
         Picture         =   "frmLevelEditor (English).frx":0F74
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   0
         Picture         =   "frmLevelEditor (English).frx":121E
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X: 0"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y: 0"
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.PictureBox picMenu 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5280
      Picture         =   "frmLevelEditor (English).frx":14C8
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   6
      Top             =   270
      Width           =   855
      Begin VB.Label lblMenu 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox picVeld 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F19470&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4065
      Left            =   885
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   451
      TabIndex        =   0
      Top             =   855
      Width           =   6765
   End
   Begin MSComctlLib.Toolbar TLB 
      Height          =   2610
      Left            =   390
      TabIndex        =   1
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   4604
      ButtonWidth     =   979
      ButtonHeight    =   900
      ImageList       =   "img"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Empty"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Wall"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            ImageIndex      =   2
            Style           =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList img 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLevelEditor (English).frx":945A
               Key             =   "blok"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLevelEditor (English).frx":9714
               Key             =   "begin"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLevelEditor (English).frx":99CE
               Key             =   "leeg"
            EndProperty
         EndProperty
      End
   End
   Begin Nibbles.Theme Theme 
      Height          =   5220
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9208
      Caption         =   "Leveleditor - [empty map - level 0.dat]"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxButton       =   0   'False
      MinButton       =   0   'False
   End
End
Attribute VB_Name = "frmLevelEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim nieuw As Boolean, level As Long
Dim startpositie(1) As Long
Dim sLevel(36) As String
Dim sNieuwLevel(17) As String
Dim X1 As Long, Y1 As Long
Dim keuze As Long, klik As Boolean

Private Sub OnStart(Index As Integer)
    Select Case Index
        Case 0
            With Theme
                .Top = 0
                .Left = 0
                .Height = Me.ScaleHeight
                .Width = Me.ScaleWidth
                .SethWndRgn Me
            End With
    End Select
End Sub

Private Sub Form_Load()
    Call OnStart(0)
    picVeld.Cls: Set picVeld.Picture = Nothing
    
    For Y = 0 To picVeld.Height Step 15
        picVeld.Line (0, Y)-(picVeld.Width, Y), &HFFFFFF
    Next Y
    For X = 0 To picVeld.Width Step 15
        picVeld.Line (X, 0)-(X, picVeld.Height), &HFFFFFF
    Next X
    For X = 1 To picVeld.Width Step 15
        BitBlt picVeld.hDC, X, 1, 14, 14, pic(1).hDC, 0, 0, vbSrcCopy
        BitBlt picVeld.hDC, X, picVeld.Height - 15, 14, 14, pic(1).hDC, 0, 0, vbSrcCopy
    Next X
    For Y = 1 To picVeld.Height Step 15
        BitBlt picVeld.hDC, 1, Y, 14, 14, pic(1).hDC, 0, 0, vbSrcCopy
        BitBlt picVeld.hDC, picVeld.Width - 15, Y, 14, 14, pic(1).hDC, 0, 0, vbSrcCopy
    Next Y
    Set picVeld.Picture = picVeld.Image: Cls
    
    MenuLevelLaden
    
    NieuweMap
    keuze = 1
End Sub

Private Sub MenuLevelLaden()
    For i = frmMenu.mnuOpenen.Count - 1 To 1 Step -1: Unload frmMenu.mnuOpenen(i): Unload frmMenu.mnuVerwijderen(i): Next i
    For i = 1 To 99
        If Dir$(App.Path & "\Levels\Level " & i & ".dat") <> vbNullString Then
            If i > 1 Then Load frmMenu.mnuOpenen(i - 1): Load frmMenu.mnuVerwijderen(i - 1)
            frmMenu.mnuOpenen(i - 1).Caption = "Level " & i: frmMenu.mnuVerwijderen(i - 1).Caption = "Level " & i
            frmMenu.mnuOpenen(i - 1).Tag = i: frmMenu.mnuVerwijderen(i - 1).Tag = i
            frmMenu.mnuOpenen(i - 1).Enabled = True: frmMenu.mnuVerwijderen(i - 1).Enabled = True
        Else: Exit For
        End If
    Next i
End Sub
Public Sub mnuEditor1_Click(Index As Integer)
    Select Case Index
        Case 0 'nieuwe map
            NieuweMap
        Case 1 'map openen
        Case 2 '-
        Case 3 'map opslaan
            If LevelOpslaan = True Then
                Caption = "Leveleditor - [empty map - level " & level & ".dat]"
                Theme.Caption = "Leveleditor - [existing map - level " & level & ".dat]"
                MenuLevelLaden
                nieuw = False
                MsgBox "Do you think this is a nice level? Send it to h.bloemhof@planet.nl and I will put it in the next update!", vbInformation, "Leveleditor"
            End If
        Case 4 '-
        Case 5 'map verwijderen
        Case 6 '-
        Case 7 'einde
            Unload Me
    End Select
End Sub

Private Sub NieuweMap()
    sLevel(0) = "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
    sLevel(1) = "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
    For i = 2 To 33
        sLevel(i) = "MM........................................................MM"
    Next i
    sLevel(34) = "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
    sLevel(35) = "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
    
'    sNieuwLevel(0) = "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
'    For i = 1 To 16
'        sNieuwLevel(i) = "M.............................M"
'    Next i
'    sNieuwLevel(17) = "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
    
    Caption = "Leveleditor - [new map - level " & frmMenu.mnuOpenen.Count + 1 & ".dat]"
    Theme.Caption = "Leveleditor - [new map - level " & frmMenu.mnuOpenen.Count + 1 & ".dat]"
    level = frmMenu.mnuOpenen.Count + 1
    nieuw = True
    frmMenu.mnuEditor1(3).Enabled = False: TLB.Buttons(5).Enabled = True
    picVeld.Cls
End Sub
Public Sub mnuOpenen_Click(Index As Integer)
    LevelLaden Index + 1
    Caption = "Leveleditor - [existing map - level " & Index + 1 & ".dat]"
    Theme.Caption = "Leveleditor - [existing map - level " & Index + 1 & ".dat]"
    nieuw = False
End Sub

Public Sub mnuVerwijderen_Click(Index As Integer)
    LevelVerwijderen Index + 1
End Sub

Private Sub Form_Resize()
    Call OnStart(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmNibbles.Show
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            PopupMenu frmMenu.mnuEditor, , picMenu.Left + lblMenu(Index).Left, picMenu.Top + lblMenu(Index).Top + lblMenu(Index).Height + 2
    End Select
End Sub

Private Sub picVeld_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'PopupMenu frmMenu.mnuEditor, , picMenu.Left + lblMenu(0).Left, picMenu.Top + lblMenu(0).Top + lblMenu(0).Height + 2
    klik = True
    WijzigVeld
End Sub

Private Sub WijzigVeld()
    'Mid$(sLevel(Y1 * 2), (X1 + 1) * 2, 1)
    
    If Not X1 <= 0 And Not X1 >= 29 And Not Y1 <= 0 And Not Y1 >= 17 Then
        BitBlt picVeld.hDC, X1 * 15 + 1, Y1 * 15 + 1, 14, 14, pic(keuze).hDC, 0, 0, vbSrcCopy
    
        'MsgBox sLevel(Y1 * 2) & " > " & Len(sLevel(Y1 * 2)) & vbCrLf & sLevel(Y1 * 2 + 1) & " > " & Len(sLevel(Y1 * 2 + 1))
        Select Case keuze
            Case 0 'leeg
                If startpositie(0) = X1 And startpositie(1) = Y1 Then
                    frmMenu.mnuEditor1(3).Enabled = False: TLB.Buttons(5).Enabled = True
                End If
                
                If Not Mid$(sLevel(Y1 * 2), (X1 + 1) * 2, 1) = "." Then
                    sLevel(Y1 * 2) = Left$(sLevel(Y1 * 2), (X1 + 1) * 2 - 2) & ".." & Mid$(sLevel(Y1 * 2), (X1 + 1) * 2 + 1)
                    sLevel(Y1 * 2 + 1) = Left$(sLevel(Y1 * 2 + 1), (X1 + 1) * 2 - 2) & ".." & Mid$(sLevel(Y1 * 2 + 1), (X1 + 1) * 2 + 1)
                End If
            Case 1 'muur
                If startpositie(0) = X1 And startpositie(1) = Y1 Then
                    frmMenu.mnuEditor1(3).Enabled = False: TLB.Buttons(5).Enabled = True
                End If
                
                If Not Mid$(sLevel(Y1 * 2), (X1 + 1) * 2, 1) = "M" Then
                    sLevel(Y1 * 2) = Left$(sLevel(Y1 * 2), (X1 + 1) * 2 - 2) & "MM" & Mid$(sLevel(Y1 * 2), (X1 + 1) * 2 + 1)
                    sLevel(Y1 * 2 + 1) = Left$(sLevel(Y1 * 2 + 1), (X1 + 1) * 2 - 2) & "MM" & Mid$(sLevel(Y1 * 2 + 1), (X1 + 1) * 2 + 1)
                End If
            Case 2 'beginpunt
                If Mid$(sLevel(Y1 * 2), (X1 + 1) * 2, 1) = "M" Then
                    sLevel(Y1 * 2) = Left$(sLevel(Y1 * 2), (X1 + 1) * 2 - 2) & ".." & Mid$(sLevel(Y1 * 2), (X1 + 1) * 2 + 1)
                    sLevel(Y1 * 2 + 1) = Left$(sLevel(Y1 * 2 + 1), (X1 + 1) * 2 - 2) & "MM" & Mid$(sLevel(Y1 * 2 + 1), (X1 + 1) * 2 + 1)
                End If
                
                TLB_ButtonClick TLB.Buttons(1)
                startpositie(0) = X1: startpositie(1) = Y1: TLB.Buttons(5).Enabled = False
                frmMenu.mnuEditor1(3).Enabled = True
        End Select
        picVeld.Refresh
        'MsgBox sLevel(Y1 * 2) & " > " & Len(sLevel(Y1 * 2)) & vbCrLf & sLevel(Y1 * 2 + 1) & " > " & Len(sLevel(Y1 * 2 + 1))
    End If
End Sub
Private Sub picVeld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InStr(1, (X + 1) / 15, ",") <> 0 Then
        X1 = Left$((X + 1) / 15, InStr(1, (X + 1) / 15, ",") - 1)
    Else: X1 = (X + 1) / 15
    End If
    If InStr(1, (Y + 1) / 15, ",") <> 0 Then
        Y1 = Left$((Y + 1) / 15, InStr(1, (Y + 1) / 15, ",") - 1)
    Else: Y1 = (Y + 1) / 15
    End If
    If Not lblX.Caption = "X: " & X1 Then lblX.Caption = "X: " & X1
    If Not lblY.Caption = "Y: " & Y1 Then lblY.Caption = "Y: " & Y1
    
    If klik = True Then WijzigVeld
End Sub

Private Sub LevelVerwijderen(verwijderen_level As Long)
    If MsgBox("Do you want remove " & verwijderen_level & " ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Kill App.Path & "\Levels\Level " & verwijderen_level & ".dat"
    For i = verwijderen_level + 1 To frmMenu.mnuOpenen.Count
        FileCopy App.Path & "\Levels\Level " & i & ".dat", App.Path & "\Levels\Level " & i - 1 & ".dat"
        Kill App.Path & "\Levels\Level " & i & ".dat"
    Next i
    MenuLevelLaden
End Sub
Private Function LevelOpslaan() As Boolean
    If nieuw = False Then
        If MsgBox("Do you want replace the existing map?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    If Dir$(App.Path & "\Levels", vbDirectory) = vbNullString Then MkDir (App.Path & "\Levels")
    Open App.Path & "\Levels\Level " & level & ".dat" For Output As #1
        For i = 0 To 35
            Print #1, sLevel(i)
        Next i
        Print #1, startpositie(0) & "," & startpositie(1)
    Close #1
    
    LevelOpslaan = True
End Function
Private Sub LevelLaden(nieuw_level)
    Erase sLevel
    If Dir$(App.Path & "\Levels\Level " & nieuw_level & ".dat") <> vbNullString Then
        picVeld.Cls
        
        Open App.Path & "\Levels\Level " & nieuw_level & ".dat" For Input As #1
            For i = 0 To 35
                Line Input #1, tekst
                sLevel(i) = Trim$(tekst)
            Next i
            Line Input #1, tekst
            startpositie(0) = CLng(Mid$(tekst, 1, InStr(1, tekst, ",") - 1))
            startpositie(1) = CLng(Mid$(tekst, InStr(1, tekst, ",") + 1))
        Close #1
    
        For Y = 0 To 35 Step 2
            For X = 1 To Len(sLevel(Y)) Step 2
                If Mid$(sLevel(Y), X, 1) = "M" Then
                    BitBlt picVeld.hDC, (X - 1) * 7.5 + 1, Y * 7.5 + 1, 14, 14, pic(1).hDC, 0, 0, vbSrcCopy
                End If
            Next X
        Next Y
        BitBlt picVeld.hDC, startpositie(0) * 15 + 1, startpositie(1) * 15 + 1, 14, 14, pic(2).hDC, 0, 0, vbSrcCopy
        
        level = nieuw_level
        TLB.Buttons(5).Enabled = False: frmMenu.mnuEditor1(3).Enabled = True
        TLB_ButtonClick TLB.Buttons(1)
        picVeld.Refresh
        
    Else: MsgBox "Can't load the level!", vbInformation
    End If
End Sub

Private Sub picVeld_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    klik = False
End Sub

Private Sub TLB_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Value = tbrUnpressed Then Button.Value = tbrPressed
    
    If Button.Index = 1 Then
        keuze = 0
    ElseIf Button.Index = 3 Then
        keuze = 1
    ElseIf Button.Index = 5 Then
        keuze = 2
    End If
    For i = 1 To TLB.Buttons.Count
        If Not Button.Index = i Then TLB.Buttons(i).Value = tbrUnpressed
    Next i
End Sub
