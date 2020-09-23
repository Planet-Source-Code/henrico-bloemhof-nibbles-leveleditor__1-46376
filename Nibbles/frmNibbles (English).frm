VERSION 5.00
Begin VB.Form frmNibbles 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Nibbles"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNibbles (English).frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   120
      Picture         =   "frmNibbles (English).frx":0CCA
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   360
      Picture         =   "frmNibbles (English).frx":0F74
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   600
      Picture         =   "frmNibbles (English).frx":121E
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   840
      Picture         =   "frmNibbles (English).frx":14C8
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   1560
      Picture         =   "frmNibbles (English).frx":1772
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   1320
      Picture         =   "frmNibbles (English).frx":1A1C
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picBlok 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   1080
      Picture         =   "frmNibbles (English).frx":1CC6
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   120
      Picture         =   "frmNibbles (English).frx":1F70
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   360
      Picture         =   "frmNibbles (English).frx":221A
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picSegment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Index           =   2
      Left            =   600
      Picture         =   "frmNibbles (English).frx":24C4
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picEten 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   600
      Picture         =   "frmNibbles (English).frx":25AE
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picMenu 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2325
      Picture         =   "frmNibbles (English).frx":2698
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   8
      Top             =   270
      Width           =   2535
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1875
         TabIndex        =   11
         Top             =   60
         Width           =   555
      End
      Begin VB.Label lblMenu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   10
         Top             =   60
         Width           =   405
      End
      Begin VB.Label lblMenu 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pause"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3DED6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   225
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   2
      Top             =   4530
      Visible         =   0   'False
      Width           =   6300
      Begin VB.Timer tmrSpel 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2280
         Top             =   0
      End
      Begin VB.Label lblEtenstijd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dinnertime: 100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level: 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5535
         TabIndex        =   6
         Top             =   0
         Width           =   645
      End
      Begin VB.Label lblPunten 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3015
         TabIndex        =   5
         Top             =   30
         Width           =   525
      End
      Begin VB.Label lblLevens 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lifes: 3/3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5385
         TabIndex        =   4
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblEten 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Food: 0/10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   900
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
      Height          =   3780
      Left            =   225
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   420
      TabIndex        =   1
      Top             =   690
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.PictureBox picNaam 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3DED6&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   795
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtNaam 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   27
         Top             =   525
         Width           =   3015
      End
      Begin Nibbles.Knop cmdOK 
         Height          =   285
         Left            =   4020
         TabIndex        =   26
         Top             =   525
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Enabled         =   0   'False
         BeginProperty Lettertype {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tekst           =   "OK"
         Voorgrondkleur  =   14737632
      End
      Begin VB.Label lblNaam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Je staat in de topscore. Geef je naam op en klik vervolgens op OK."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   4800
      End
      Begin VB.Label lblNaam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Naam:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   555
         Width           =   465
      End
   End
   Begin VB.PictureBox picLaden 
      Appearance      =   0  'Flat
      BackColor       =   &H00F3DED6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   6135
      TabIndex        =   23
      Top             =   2640
      Width           =   6135
      Begin VB.Label lblLaden 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bezig met laden van level 1..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   2475
      End
   End
   Begin Nibbles.Theme Theme 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9128
      Caption         =   "Nibbles"
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
Attribute VB_Name = "frmNibbles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim blok As Long
Dim segmenten(30, 30) As Double
Dim X1(1000) As Long, Y1(1000) As Long
Dim richting As String, oud_richting As String
Dim stappen As New Collection, erbijgroeien As Long
Dim aantal As Long, groei As Long
Dim levens As Long, Score As Long, level As Long
Dim posEten(1) As Long, etenstijd As Long
Dim pauze As Boolean, Topscore1(4) As Long, Topscore2(4) As String
Dim sLevel(36) As String, VolgendeLevel As Boolean
Dim startpositie(1) As Long, ok As Boolean
Dim CanMove As Boolean 'thank you Jason
Dim Vote As Boolean

Private Sub cmdOK_Click()
    
End Sub

Private Sub cmdOK_klik()
    ok = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Nieuw_spel
    ElseIf KeyCode = vbKeyF3 Then
        Beëindig_spel
    ElseIf KeyCode = vbKeyF3 Then
        Unload Me
    End If
    If picVeld.Visible Then picVeld.SetFocus
End Sub

Private Sub Form_Load()
    Call OnStart(0)
    blok = 0
    TopscoreLaden
    
    If Dir$(App.Path & "\Settings.dat", vbNormal) <> vbNullString Then Vote = True Else Vote = False
End Sub

Private Sub LevelLaden(nieuw_level As Long)
    Dim tijd As Single
    
    picVeld.Visible = False: picVeld.Cls: Set picVeld.Picture = Nothing
    picInfo.Visible = False
    
    Erase sLevel
    If Dir$(App.Path & "\Levels\Level " & nieuw_level & ".dat") <> vbNullString Then
        tijd = Timer + 1
        lblLaden.Caption = "Loading level " & nieuw_level & "..."
        lblLaden.Visible = True
        Do Until Timer > tijd: DoEvents: Loop
        lblLaden.Visible = False

        Open App.Path & "\Levels\Level " & nieuw_level & ".dat" For Input As #1
            For i = 0 To 35
                Line Input #1, tekst
                sLevel(i) = Trim$(tekst)
            Next i
            Line Input #1, tekst
            startpositie(0) = CLng(Mid$(tekst, 1, InStr(1, tekst, ",") - 1))
            startpositie(1) = CLng(Mid$(tekst, InStr(1, tekst, ",") + 1))
            
            For i = 0 To aantal
               X1(i) = 14 * startpositie(0) - (0.5 * i): Y1(i) = 14 * startpositie(1)
            Next i
            
        Close #1
        
        For Y = 0 To picVeld.Height Step 14
            For X = 0 To picVeld.Width Step 14
                If Mid$(sLevel(Y / 7), X / 7 + 1, 1) = "M" Then
                    BitBlt picVeld.hDC, X, Y, 14, 14, picBlok(blok).hDC, 0, 0, vbSrcCopy
                End If
            Next X
        Next Y
        
        picVeld.Picture = picVeld.Image: picVeld.Cls: picVeld.Visible = True
        picInfo.Visible = True
    ElseIf nieuw_level > 1 Then
        MsgBox "You have played all the levels!. You can start playing level 1.", vbInformation, "Nibbles"
        level = 1
        lblLevel.Caption = "Level: 1"
        LevelLaden 1
    Else
        tijd = Timer + 1
        lblLaden.Caption = "Loading level 1 " & nieuw_level & "..."
        lblLaden.Visible = True
        Do Until Timer > tijd: DoEvents: Loop
        lblLaden.Visible = False
        
        MsgBox "Can't load level " & nieuw_level & "." & vbCrLf & vbCrLf & "Your score is " & Score & ". Game will be closed.", vbInformation, "Nibbles"
        Beëindig_spel
    End If
End Sub

Private Sub Nieuw_spel()
    tmrSpel.Enabled = False
    frmMenu.mnuSpel(0).Enabled = False: frmMenu.mnuSpel(1).Enabled = True: frmMenu.mnuSpel(5).Enabled = False
    
    LevelLaden 1
    For i = stappen.Count To 1 Step -1: stappen.Remove i: Next i
    
    groei = 5: erbijgroeien = 0
    Score = 0: levens = 3: aantal = 2: level = 1
    lblPunten.Caption = "Score" & vbCrLf & Score
    lblLevens.Caption = "Lifes: " & levens & "/3"
    lblLevel.Caption = "Level: " & level
    lblEten.Caption = "Food: 0/10"
    For i = 0 To aantal
        X1(i) = 14 * startpositie(0) - (0.5 * i): Y1(i) = 14 * startpositie(1)
    Next i
    
    Randomize
    Nieuw_eten
    Eten
    Teken
    
    richting = "RECHTS"
    
    'frmMenu.mnuPauze.Caption = "&Pauze": frmMenu.mnuPauze.Visible = True:
    lblMenu(0).Enabled = True: lblMenu(0).Caption = "&Pause"
    picVeld.Visible = True: picInfo.Visible = True
    picVeld.SetFocus
    tmrSpel.Enabled = True
End Sub

Private Sub Beëindig_spel()
    tmrSpel.Enabled = False
    erbijgroeien = 0: pauze = False
    picVeld.Visible = False
    picInfo.Visible = False
    
    frmMenu.mnuSpel(0).Enabled = True: frmMenu.mnuSpel(1).Enabled = False
    frmMenu.mnuPauze.Visible = False: frmMenu.mnuSpel(5).Enabled = True
    lblMenu(0).Enabled = False: lblMenu(0).Caption = "Pause"
    Theme.SetFocus
End Sub

Private Sub Nieuw_eten()
    Dim geraakt As Boolean
    
    Do
        posEten(0) = (Int(Rnd * (picVeld.Width / 14))) * 14
        posEten(1) = (Int(Rnd * (picVeld.Height / 14))) * 14
        geraakt = False
        For i = 1 To aantal
            If posEten(0) = X1(i) And posEten(1) = Y1(i) Then geraakt = True: Exit For
        Next i
        If geraakt = False Then
            If Mid$(sLevel(posEten(1) / 7), posEten(0) / 7 + 1, 1) = "M" Then geraakt = True
        End If
        If frmMenu.mnuOptie(1).Checked = True And geraakt = False Then
            If Mid$(sLevel(posEten(1) / 7 - 1), posEten(0) / 7 + 1, 1) = "M" Then
                geraakt = True
            ElseIf Mid$(sLevel(posEten(1) / 7 + 1), posEten(0) / 7 + 1, 1) = "M" Then geraakt = True
            ElseIf Mid$(sLevel(posEten(1) / 7), posEten(0) / 7 + 2, 1) = "M" Then geraakt = True
            ElseIf Mid$(sLevel(posEten(1) / 7), posEten(0) / 7 - 0, 1) = "M" Then geraakt = True
            End If
        End If
    Loop Until geraakt = False

    etenstijd = 100
End Sub
Private Sub Eten()
    BitBlt picVeld.hDC, posEten(0), posEten(1), 14, 14, picEten.hDC, 0, 0, vbSrcCopy
End Sub
Private Sub Teken()
    Dim info() As String
    picVeld.Cls
    
    Eten
    
    If erbijgroeien > 0 Then aantal = aantal + 1: erbijgroeien = erbijgroeien - 1
    For i = 1 To aantal
        If stappen.Count > i Then
            info = Split(stappen(stappen.Count - i), "|")
            X1(i) = CLng(info(0))
            Y1(i) = CLng(info(1))
        ElseIf stappen.Count > 0 Then
            info = Split(stappen(stappen.Count), "|")
            X1(i) = CLng(info(0))
            Y1(i) = CLng(info(1))
        End If
    Next i
    
    For i = 0 To aantal
        BitBlt picVeld.hDC, X1(i), Y1(i), 14, 14, picSegment(2).hDC, 0, 0, vbSrcCopy
    Next i
    'BitBlt picVeld.hDC, X1, Y1, 14, 14, picSegment(1).hDC, 0, 0, vbSrcAnd
    'BitBlt picVeld.hDC, X1, Y1, 14, 14, picSegment(2).hDC, 0, 0, vbSrcPaint
End Sub

Public Sub mnuOptie_Click(Index As Integer)
    Select Case Index
        Case 1 'eten niet langs muren plaatsen
            frmMenu.mnuOptie(Index).Checked = Not (frmMenu.mnuOptie(Index).Checked)
    End Select
End Sub

Public Sub mnuPauze_Click()
    picVeld_KeyDown vbKeyP, 0&
End Sub

Public Sub mnuSnelheid_Click(Index As Integer)
    Select Case Index
        Case 0: tmrSpel.Interval = 90
        Case 1: tmrSpel.Interval = 60
        Case 2: tmrSpel.Interval = 30
    End Select
    For i = 0 To frmMenu.mnuSnelheid.Count - 1
        If Not Index = i Then
            If Not frmMenu.mnuSnelheid(i).Checked = False Then frmMenu.mnuSnelheid(i).Checked = False
        Else: If Not frmMenu.mnuSnelheid(i).Checked = True Then frmMenu.mnuSnelheid(i).Checked = True
        End If
    Next i
End Sub

Public Sub mnuSpel_Click(Index As Integer)
    Select Case Index
        Case 0: Nieuw_spel
        Case 1: Beëindig_spel
        Case 2: '-
        Case 3
            MsgBox "Highscore:" _
                    & vbCrLf & vbCrLf & "1. " & Topscore1(0) & " points from " & Topscore2(0) _
            & vbCrLf & "2. " & Topscore1(1) & " points from " & Topscore2(1) & vbCrLf _
            & "3. " & Topscore1(2) & " points from " & Topscore2(2) & vbCrLf & _
            "4. " & Topscore1(3) & " points from " & Topscore2(3) & vbCrLf & _
            "5. " & Topscore1(4) & " points from " & Topscore2(4), vbInformation, "Nibbles"
            
        Case 4: '-
        Case 5: 'leveleditor
            frmNibbles.Hide
            frmLevelEditor.Show
            'frmNibbles.Show
        Case 6: '-
        Case 7: Unload Me
    End Select
        
End Sub

Private Sub Form_Resize()
    Call OnStart(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As Variant
    Dim sDirectory As Variant

    Beëindig_spel
    
    If Vote = True Then End
    
    If MsgBox("If you like this game, please vote for me! If you want to vote click on yes.", vbQuestion + vbYesNo, "Nibbles") = vbYes Then
        Open App.Path & "\Settings.dat" For Output As #1: Close #1
        
        sTopic = "Open"
        sFile = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=46376&lngWId=1"
        sParams = 0&
        sDirectory = 0&
        
        Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
    End If
    
    End
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0 'pauze
            mnuPauze_Click
        Case 1 'spel
            PopupMenu frmMenu.mnuSpel1, , picMenu.Left + lblMenu(Index).Left, picMenu.Top + lblMenu(Index).Top + lblMenu(Index).Height + 2
        Case 2 'opties
            PopupMenu frmMenu.mnuOpties1, , picMenu.Left + lblMenu(Index).Left, picMenu.Top + lblMenu(Index).Top + lblMenu(Index).Height + 2
    End Select
End Sub

Private Sub picVeld_KeyDown(KeyCode As Integer, Shift As Integer)
    If pauze = True Then KeyCode = vbKeyP
    
    Select Case KeyCode
        Case vbKeyF2
            Nieuw_spel
        Case vbKeyF3
            Beëindig_spel
        Case vbKeyF3
            Unload Me
        Case vbKeyUp
            If Not oud_richting = "BENEDEN" And CanMove = True Then richting = "BOVEN": CanMove = False
        Case vbKeyDown
            If Not oud_richting = "BOVEN" And CanMove = True Then richting = "BENEDEN": CanMove = False
        Case vbKeyLeft
            If Not oud_richting = "RECHTS" And CanMove = True Then richting = "LINKS": CanMove = False
        Case vbKeyRight
            If Not oud_richting = "LINKS" And CanMove = True Then richting = "RECHTS": CanMove = False
        Case vbKeyEscape
            Beëindig_spel
        Case vbKeyReturn
            If tmrSpel.Enabled = False Then
                Nieuw_spel
            End If
        Case vbKeyN: VolgendeLevel = True
        Case vbKeyP
            'picVeld.Cls
            pauze = Not (pauze)
            If pauze = True Then
                'mnuPauze.Caption = "Verder gaan"
                lblMenu(0).Caption = "Continue"
                Printtekst "PAUSE"
                Printtekst "press a key to continue", 0, picVeld.Height / 2 + 30, 9
            Else: lblMenu(0).Caption = "&Pause"
            End If
            picVeld.SetFocus
        Case Else: Exit Sub
    End Select
End Sub

Private Sub Theme_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub tmrSpel_Timer()
    CanMove = True
    Dim geraakt As Boolean
    
    Do Until pauze = False: DoEvents: Loop
    
    oud_richting = richting
    Select Case richting
        Case "BOVEN"
            Y1(0) = Y1(0) - 7
        Case "BENEDEN"
            Y1(0) = Y1(0) + 7
        Case "LINKS"
            X1(0) = X1(0) - 7
        Case "RECHTS"
            X1(0) = X1(0) + 7
    End Select
    
    stappen.Add X1(0) & "|" & Y1(0)
    If etenstijd > 10 Then etenstijd = etenstijd - 1
    lblEtenstijd.Caption = "Dinnertime: " & etenstijd
    
    If X1(0) = posEten(0) And Y1(0) = posEten(1) Or VolgendeLevel = True Then
        Score = Score + etenstijd: erbijgroeien = groei
        lblEten.Caption = "Food: " & (aantal - 2) / groei + 1 & "/10"
        If (aantal - 2) / groei >= 9 Or VolgendeLevel = True Then
            VolgendeLevel = False
            level = level + 1
            lblLevel.Caption = "Level: " & level
            lblEten.Caption = "Eten: 0/10"
            richting = "RECHTS"
            erbijgroeien = 0
            aantal = 2
            For i = stappen.Count To 1 Step -1: stappen.Remove i: Next i
            LevelLaden level
        End If
        lblPunten.Caption = "Score" & vbCrLf & Score
        Nieuw_eten
        Teken
        picVeld.SetFocus
        Exit Sub
    End If
    
    geraakt = False
    For i = 1 To aantal
        If X1(0) = X1(i) And Y1(0) = Y1(i) Then
            geraakt = True: Exit For
        End If
    Next i
    
    If Mid$(sLevel(Y1(0) / 7), X1(0) / 7 + 1, 1) = "M" Then geraakt = True
    
    If geraakt = True Then
        levens = levens - 1: richting = "RECHTS"
        erbijgroeien = 0
        aantal = 2
        For i = stappen.Count To 1 Step -1: stappen.Remove i: Next i
        If levens = -1 Then
            picVeld.Visible = False: picInfo.Visible = False: frmMenu.mnuSpel1.Enabled = False
            frmMenu.mnuPauze.Visible = False: frmMenu.mnuOpties1.Enabled = False
            lblMenu(0).Enabled = False
            CheckTopscore
            picVeld.Visible = True: picInfo.Visible = True: frmMenu.mnuSpel1.Enabled = True: frmMenu.mnuPauze.Visible = True: frmMenu.mnuOpties1.Enabled = True
            lblMenu(0).Enabled = True
            picVeld.SetFocus
            If MsgBox("You have no lifes more. Do you want start again?", vbQuestion + vbYesNo, "Nibbles") = vbYes Then
                Nieuw_spel
            Else: Beëindig_spel
            End If
        Else:
            lblEten.Caption = "Food: 0/10"
            lblLevens.Caption = "Lifes: " & levens & "/3"
            For i = 0 To aantal
                X1(i) = 14 * startpositie(0) - (0.5 * i): Y1(i) = 14 * startpositie(1)
            Next i
            Nieuw_eten
            Teken
        End If
    Else: Teken
    End If
End Sub

Private Sub Printtekst(tekst As String, Optional X As Long = 0, Optional Y As Long = 0, Optional grootte As Single = 0)
    Dim oude_grootte As Single
    
    If grootte > 0 Then
        oude_grootte = picVeld.FontSize
        picVeld.FontSize = grootte
    End If
    
    If X > 0 Then
        If Y > 0 Then
            picVeld.CurrentY = Y
        Else: picVeld.CurrentY = picVeld.Height / 2 - picVeld.TextHeight(tekst) / 2
        End If
        picVeld.CurrentX = X
    ElseIf Y > 0 Then
        picVeld.CurrentY = Y
        picVeld.CurrentX = picVeld.Width / 2 - picVeld.TextWidth(tekst) / 2
    Else
        picVeld.CurrentY = picVeld.Height / 2 - picVeld.TextHeight(tekst) / 2
        picVeld.CurrentX = picVeld.Width / 2 - picVeld.TextWidth(tekst) / 2
    End If
    picVeld.Print tekst
    
    If oude_grootte > 0 Then picVeld.FontSize = oude_grootte
End Sub

Private Sub GameOver()
    picVeld.Cls
    Printtekst "GAME OVER"
    Printtekst "press on 'escape' to stop", 0, picVeld.Height / 2 + 30, 9
    Printtekst "press a other key to start a new game", 0, picVeld.Height / 2 + 43, 9
    tmrSpel.Enabled = False
End Sub

Private Sub TopscoreLaden()
    Dim tekst As String, info() As String
    If Dir$(App.Path & "\Topscore.dat") <> vbNullString Then
        Open App.Path & "\Topscore.dat" For Input As #1
            For i = 0 To 4
                Line Input #1, tekst
                info = Split(Trim$(tekst), "|")
                Topscore1(i) = CLng(info(0))
                Topscore2(i) = info(1)
            Next i
        Close #1
    Else
        Topscore2(0) = "HB": Topscore1(0) = "1000"
        Topscore2(1) = "HB": Topscore1(1) = "800"
        Topscore2(2) = "HB": Topscore1(2) = "600"
        Topscore2(3) = "HB": Topscore1(3) = "400"
        Topscore2(4) = "HB": Topscore1(4) = "200"
    End If
End Sub

Private Sub TopscoreOpslaan()
    Open App.Path & "\Topscore.dat" For Output As #1
        For i = 0 To 4: Print #1, Topscore1(i) & "|" & Topscore2(i): Next i
    Close #1
End Sub

Private Sub CheckTopscore()
    If Score > Topscore1(4) Then
        lblMenu(1).Enabled = False: lblMenu(2).Enabled = False
        picNaam.Visible = True: picNaam.ZOrder vbBringToFront
        txtNaam.SetFocus: ok = False
        Do Until ok = True: DoEvents: Loop
        picNaam.Visible = False
        lblMenu(1).Enabled = True: lblMenu(2).Enabled = True
        Dim naam As String: naam = txtNaam.text
        
        
        If Score > Topscore1(0) Then
            Topscore1(4) = Topscore1(3): Topscore2(4) = Topscore2(3)
            Topscore1(3) = Topscore1(2): Topscore2(3) = Topscore2(2)
            Topscore1(2) = Topscore1(1): Topscore2(2) = Topscore2(1)
            Topscore1(1) = Topscore1(0): Topscore2(1) = Topscore2(0)
            Topscore1(0) = Score
            Topscore2(0) = naam
        ElseIf Score > Topscore1(1) Then
            Topscore1(4) = Topscore1(3): Topscore2(4) = Topscore2(3)
            Topscore1(3) = Topscore1(2): Topscore2(3) = Topscore2(2)
            Topscore1(2) = Topscore1(1): Topscore2(2) = Topscore2(1)
            Topscore1(1) = Score
            Topscore2(1) = naam
        ElseIf Score > Topscore1(2) Then
            Topscore1(4) = Topscore1(3): Topscore2(4) = Topscore2(3)
            Topscore1(3) = Topscore1(2): Topscore2(3) = Topscore2(2)
            Topscore1(2) = Score
            Topscore2(2) = naam
        ElseIf Score > Topscore1(3) Then
            Topscore1(4) = Topscore1(3): Topscore2(4) = Topscore2(3)
            Topscore1(3) = Score
            Topscore2(3) = naam
        ElseIf Score > Topscore1(4) Then
            Topscore1(4) = Score
            Topscore2(4) = naam
        End If
        TopscoreOpslaan
        
        MsgBox "Your score is " & Score & ".  Highscore:" _
            & vbCrLf & vbCrLf & "1. " & Topscore1(0) & " points from " & Topscore2(0) _
            & vbCrLf & "2. " & Topscore1(1) & " points from " & Topscore2(1) & vbCrLf _
            & "3. " & Topscore1(2) & " points from " & Topscore2(2) & vbCrLf & _
            "4. " & Topscore1(3) & " points from " & Topscore2(3) & vbCrLf & _
            "5. " & Topscore1(4) & " points from " & Topscore2(4), vbInformation, "Nibbles"
    
    End If
End Sub

Private Sub txtNaam_Change()
    If Trim$(txtNaam.text) <> vbNullString Then cmdOK.Enabled = True Else cmdOK.Enabled = False
End Sub

Private Sub txtNaam_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: cmdOK_klik
End Sub

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


