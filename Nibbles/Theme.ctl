VERSION 5.00
Begin VB.UserControl Theme 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   FillStyle       =   0  'Solid
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   7
      Left            =   3120
      Picture         =   "Theme.ctx":0000
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   6
      Left            =   1560
      Picture         =   "Theme.ctx":1FF2
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Timer TimerTheme 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   4080
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   5
      Left            =   0
      Picture         =   "Theme.ctx":3FE4
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   0
      Picture         =   "Theme.ctx":5FD6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   3
      Left            =   3660
      Picture         =   "Theme.ctx":7FF8
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   3
      Top             =   150
      Width           =   390
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   4470
      Picture         =   "Theme.ctx":885A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   4035
      Picture         =   "Theme.ctx":9364
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      Top             =   150
      Width           =   390
   End
   Begin VB.PictureBox PictureBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   4035
      Picture         =   "Theme.ctx":9BC6
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image IconBox 
      Height          =   240
      Index           =   1
      Left            =   1920
      Picture         =   "Theme.ctx":A428
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IconBox 
      Height          =   240
      Index           =   0
      Left            =   150
      Picture         =   "Theme.ctx":A683
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MicroX Controls"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   450
      TabIndex        =   8
      Top             =   300
      Width           =   1350
   End
   Begin VB.Image ImageBox 
      Height          =   180
      Index           =   6
      Left            =   4905
      Picture         =   "Theme.ctx":A8DE
      Top             =   3300
      Width           =   180
   End
   Begin VB.Image ImageBox 
      Height          =   135
      Index           =   5
      Left            =   285
      Picture         =   "Theme.ctx":AB00
      Stretch         =   -1  'True
      Top             =   3345
      Width           =   4620
   End
   Begin VB.Image ImageBox 
      Height          =   2610
      Index           =   7
      Left            =   4965
      Picture         =   "Theme.ctx":ABBA
      Stretch         =   -1  'True
      Top             =   690
      Width           =   120
   End
   Begin VB.Image ImageBox 
      Height          =   270
      Index           =   4
      Left            =   0
      Picture         =   "Theme.ctx":AC70
      Top             =   3210
      Width           =   285
   End
   Begin VB.Image ImageBox 
      Height          =   2520
      Index           =   3
      Left            =   0
      Picture         =   "Theme.ctx":ADF2
      Stretch         =   -1  'True
      Top             =   690
      Width           =   225
   End
   Begin VB.Image ImageBox 
      Height          =   690
      Index           =   1
      Left            =   225
      Picture         =   "Theme.ctx":B1DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2985
   End
   Begin VB.Image ImageBox 
      Height          =   690
      Index           =   0
      Left            =   0
      Picture         =   "Theme.ctx":B965
      Top             =   0
      Width           =   225
   End
   Begin VB.Image ImageBox 
      Height          =   690
      Index           =   2
      Left            =   3120
      Picture         =   "Theme.ctx":BC8C
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "Theme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum ColorSchemeEnm
    [Blue (Default)] = 0
    [OliveGreen] = 1
    [Silver] = 2
End Enum

Private m_AlwaysOnTop As Boolean
Private m_BackColor As OLE_COLOR
Private m_Caption As String
Private m_CaptionOffsetX, m_CaptionOffsetY, m_IconOffsetX, m_IconOffsetY As Long
Private m_ColorScheme As ColorSchemeEnm
Private m_Font As Font
Private m_ForeColor As OLE_COLOR
Private m_Icon As Picture
Private m_MaxButton As Boolean
Private m_MinButton As Boolean

Private hWnd As Form

Private hTop, hLeft, hWidth, hHeight As Long
Private SetMaxRes As Boolean

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
    m_Caption = NewValue
    Label(0).Caption = m_Caption
    PropertyChanged "Caption"
End Property

Public Property Get CaptionOffsetX() As Long
    CaptionOffsetX = m_CaptionOffsetX
End Property

Public Property Let CaptionOffsetX(ByVal NewValue As Long)
    m_CaptionOffsetX = NewValue
    Label(0).Left = m_CaptionOffsetX
    PropertyChanged "CaptionOffsetX"
End Property

Public Property Get CaptionOffsetY() As Long
    CaptionOffsetY = m_CaptionOffsetY
End Property

Public Property Let CaptionOffsetY(ByVal NewValue As Long)
    m_CaptionOffsetY = NewValue
    Label(0).Top = m_CaptionOffsetY
    PropertyChanged "CaptionOffsetY"
End Property

Public Property Get ColorScheme() As ColorSchemeEnm
    ColorScheme = m_ColorScheme
End Property

Public Property Let ColorScheme(ByVal NewValue As ColorSchemeEnm)
    m_ColorScheme = NewValue
    If m_ColorScheme = 0 Then
        'Call SetColorShm(0)
    ElseIf m_ColorScheme = 1 Then
        Call SetColorShm(1)
    ElseIf m_ColorScheme = 2 Then
        Call SetColorShm(2)
    End If
    PropertyChanged "ColorScheme"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    Label(0).ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal NewValue As Font)
    Set m_Font = NewValue
    Set Label(0).Font = m_Font
    PropertyChanged "Font"
End Property

'Public Property Get Icon() As Picture
'    Set Icon = m_Icon
'End Property
'
'Public Property Set Icon(ByVal NewValue As Picture)
'    Set m_Icon = NewValue
'    Set IconBox(0).Picture = m_Icon
'    PropertyChanged "Icon"
'End Property
'
'Public Property Get IconOffsetX() As Long
'    IconOffsetX = m_IconOffsetX
'End Property
'
'Public Property Let IconOffsetX(ByVal NewValue As Long)
'    m_IconOffsetX = NewValue
'    IconBox(0).Left = m_IconOffsetX
'    PropertyChanged "IconOffsetX"
'End Property
'
'Public Property Get IconOffsetY() As Long
'    IconOffsetY = m_IconOffsetY
'End Property
'
'Public Property Let IconOffsetY(ByVal NewValue As Long)
'    m_IconOffsetY = NewValue
'    IconBox(0).Top = m_IconOffsetY
'    PropertyChanged "IconOffsetY"
'End Property

Public Property Get MaxButton() As Boolean
    MaxButton = m_MaxButton
End Property

Public Property Let MaxButton(ByVal NewValue As Boolean)
    m_MaxButton = NewValue
    If m_MaxButton = False Then
        Call PictureBox(1).PaintPicture(PictureBox(5).Picture, -78, 0, PictureBox(5).Width, PictureBox(5).Height)
    Else
        Call PictureBox(1).PaintPicture(PictureBox(5).Picture, 0, 0, PictureBox(5).Width, PictureBox(5).Height)
    End If
    PropertyChanged "MaxButton"
End Property

Public Property Get MinButton() As Boolean
    MinButton = m_MinButton
End Property

Public Property Let MinButton(ByVal NewValue As Boolean)
    m_MinButton = NewValue
    If m_MinButton = False Then
        Call PictureBox(3).PaintPicture(PictureBox(7).Picture, -78, 0, PictureBox(7).Width, PictureBox(7).Height)
    Else
        Call PictureBox(3).PaintPicture(PictureBox(7).Picture, 0, 0, PictureBox(7).Width, PictureBox(7).Height)
    End If
    PropertyChanged "MinButton"
End Property

Private Sub ImageBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Param As Long

    Select Case Index
        Case 1
            If SetMaxRes = False Then
                Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, &HA1, 2, 0)
            End If
    End Select
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            If SetMaxRes = False Then
                Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, &HA1, 2, 0)
            End If
    End Select
End Sub

Private Sub PictureBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            TimerTheme.Enabled = False
            Call PictureBox(0).PaintPicture(PictureBox(4).Picture, -30, 0, PictureBox(4).Width, PictureBox(4).Height)
        
        Case 1
            If m_MaxButton = True Then
                TimerTheme.Enabled = False
                Call PictureBox(1).PaintPicture(PictureBox(5).Picture, -26, 0, PictureBox(5).Width, PictureBox(5).Height)
            End If
        
        Case 2
            TimerTheme.Enabled = False
            Call PictureBox(2).PaintPicture(PictureBox(6).Picture, -26, 0, PictureBox(6).Width, PictureBox(6).Height)
        Case 3
            If m_MinButton = True Then
                TimerTheme.Enabled = False
                Call PictureBox(3).PaintPicture(PictureBox(7).Picture, -26, 0, PictureBox(7).Width, PictureBox(7).Height)
            End If
    End Select
End Sub

Private Sub PictureBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
            TimerTheme.Enabled = True
            Call Timeout(0.1): Unload UserControl.Parent
            
        Case 1
            If m_MaxButton = True Then
                TimerTheme.Enabled = True
                Call Timeout(0.1)
                With UserControl.Parent
                    If SetMaxRes = False Then
                        SetMaxRes = True
                        hTop = .Top: hLeft = .Left: hWidth = .Width: hHeight = .Height
                        PictureBox(1).Visible = False
                        PictureBox(2).Visible = True
                        .Move .SysInfo.WorkAreaLeft, .SysInfo.WorkAreaTop, .SysInfo.WorkAreaWidth, .SysInfo.WorkAreaHeight
                    End If
                End With
            End If
        Case 2
            TimerTheme.Enabled = True
            Call Timeout(0.1)
            With UserControl.Parent
                If SetMaxRes = True Then
                    SetMaxRes = False
                    PictureBox(1).Visible = True
                    PictureBox(2).Visible = False
                    .Move hLeft, hTop, hWidth, hHeight
                End If
            End With
        Case 3
            If m_MinButton = True Then
                TimerTheme.Enabled = True
                Call Timeout(0.1): UserControl.Parent.WindowState = 1
            End If
    End Select
End Sub

Private Sub TimerTheme_Timer()
Dim Pt As POINTAPI
    GetCursorPos Pt
    
    If WindowFromPoint(Pt.X, Pt.Y) = PictureBox(0).hWnd Then
        Call PictureBox(0).PaintPicture(PictureBox(4).Picture, -60, 0, PictureBox(4).Width, PictureBox(4).Height)
    Else
        Call PictureBox(0).PaintPicture(PictureBox(4).Picture, 0, 0, PictureBox(4).Width, PictureBox(4).Height)
    End If
    
    If m_MaxButton = True Then
        If WindowFromPoint(Pt.X, Pt.Y) = PictureBox(1).hWnd Then
            Call PictureBox(1).PaintPicture(PictureBox(5).Picture, -52, 0, PictureBox(5).Width, PictureBox(5).Height)
        Else
            Call PictureBox(1).PaintPicture(PictureBox(5).Picture, 0, 0, PictureBox(5).Width, PictureBox(5).Height)
        End If
    End If
    
    If WindowFromPoint(Pt.X, Pt.Y) = PictureBox(2).hWnd Then
        Call PictureBox(2).PaintPicture(PictureBox(6).Picture, -52, 0, PictureBox(6).Width, PictureBox(5).Height)
    Else
        Call PictureBox(2).PaintPicture(PictureBox(6).Picture, 0, 0, PictureBox(6).Width, PictureBox(5).Height)
    End If
    
    If m_MinButton = True Then
        If WindowFromPoint(Pt.X, Pt.Y) = PictureBox(3).hWnd Then
            Call PictureBox(3).PaintPicture(PictureBox(7).Picture, -52, 0, PictureBox(7).Width, PictureBox(7).Height)
        Else
            Call PictureBox(3).PaintPicture(PictureBox(7).Picture, 0, 0, PictureBox(7).Width, PictureBox(7).Height)
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    Call SetFramePos
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    Call SetFramePos
End Sub

Private Sub SetFramePos()
Dim X, Y As Single
    
    X = UserControl.ScaleWidth
    Y = UserControl.ScaleHeight

    With ImageBox
        .Item(0).Left = 0
        .Item(0).Top = 0
        
        .Item(1).Left = .Item(0).Width
        .Item(1).Top = 0
        .Item(1).Width = X - .Item(2).Width
        
        .Item(2).Left = .Item(1).Width
        .Item(2).Top = 0
        
        .Item(3).Left = 0
        .Item(3).Top = .Item(0).Height
        .Item(3).Height = Y - 64
        
        .Item(4).Left = 0
        .Item(4).Top = Y - .Item(4).Height
        
        .Item(5).Left = .Item(4).Width
        .Item(5).Top = Y - .Item(5).Height
        .Item(5).Width = X - 31
        
        .Item(6).Left = X - .Item(6).Width
        .Item(6).Top = Y - .Item(6).Height
        
        .Item(7).Left = X - .Item(7).Width
        .Item(7).Top = .Item(2).Height
        .Item(7).Height = Y - 58
    End With
    
    With PictureBox
        .Item(0).Left = X - 41
        .Item(0).Top = 8
        .Item(1).Left = X - 70
        .Item(1).Top = 10
        .Item(2).Left = X - 70
        .Item(2).Top = 10
        .Item(3).Left = X - 95
        .Item(3).Top = 10
    End With
End Sub

Public Sub SethWndRgn(hWnd)
Dim CrtRctRgn, X, Y As Single, i As Integer
    
    X = hWnd.Width / Screen.TwipsPerPixelX
    Y = hWnd.Height / Screen.TwipsPerPixelY
    
    CrtRctRgn = CreateRectRgn(X - 100, 0, X - 9, 1)
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 104, 1, X - 6, 2), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 106, 2, X - 4, 3), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 108, 3, X - 3, 4), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 109, 4, X - 2, 5), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 110, 5, X - 2, 6), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 111, 6, X - 1, 7), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 112, 7, X - 1, 8), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 114, 8, X, 9), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 116, 9, X, 10), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 118, 10, X, 11), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 121, 11, X, 12), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(X - 125, 12, X, 13), 2
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, 13, X, 14), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(5, 14, X, 15), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(3, 15, X, 16), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(2, 16, X, 17), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(1, 17, X, 18), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(1, 18, X, 19), 2
    
    For i = 19 To Y - 18
        CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(0, i, X, i + 1), 2
    Next i
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(1, Y - 17, X, Y - 16), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(2, Y - 16, X, Y - 15), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(3, Y - 15, X, Y - 14), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(4, Y - 14, X, Y - 13), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(5, Y - 13, X, Y - 12), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(6, Y - 12, X, Y - 11), 2
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, Y - 11, X, Y - 10), 2
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, Y - 10, X, Y - 9), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, Y - 9, X, Y - 8), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, Y - 8, X - 1, Y - 7), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, Y - 7, X - 1, Y - 6), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(7, Y - 6, X - 2, Y - 5), 2
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(8, Y - 5, X - 2, Y - 4), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(9, Y - 4, X - 3, Y - 3), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(10, Y - 3, X - 4, Y - 2), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(11, Y - 2, X - 6, Y - 1), 2
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(12, Y - 1, X - 9, Y - 0), 2
    
    CombineRgn CrtRctRgn, CrtRctRgn, CreateRectRgn(0, Y - 0, X, Y), 2
    SetWindowRgn UserControl.ContainerHwnd, CrtRctRgn, True
    
    TimerTheme.Enabled = True
End Sub

Private Sub SetColorShm(Index As Integer)
Dim SetDirPath(0 To 2) As String

    Select Case Index
        Case 0
            SetDirPath(0) = App.Path & "\Graphics\Blue (Default)"
            
            With ImageBox
                .Item(0).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(0)" & ".gif")
                .Item(1).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(1)" & ".gif")
                .Item(2).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(2)" & ".gif")
                .Item(3).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(3)" & ".gif")
                .Item(4).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(4)" & ".gif")
                .Item(5).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(5)" & ".gif")
                .Item(6).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(6)" & ".gif")
                .Item(7).Picture = LoadPicture(SetDirPath(0) & "\" & "ImageBox(7)" & ".gif")
            End With
            
            With PictureBox
                .Item(0).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(0)" & ".gif")
                .Item(1).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(1)" & ".gif")
                .Item(2).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(2)" & ".gif")
                .Item(3).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(3)" & ".gif")
                .Item(4).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(4)" & ".gif")
                .Item(5).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(5)" & ".gif")
                .Item(6).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(6)" & ".gif")
                .Item(7).Picture = LoadPicture(SetDirPath(0) & "\" & "PictureBox(7)" & ".gif")
            End With
        
        Case 1
            SetDirPath(1) = App.Path & "\Graphics\OliveGreen"
            
            With ImageBox
                .Item(0).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(0)" & ".gif")
                .Item(1).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(1)" & ".gif")
                .Item(2).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(2)" & ".gif")
                .Item(3).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(3)" & ".gif")
                .Item(4).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(4)" & ".gif")
                .Item(5).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(5)" & ".gif")
                .Item(6).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(6)" & ".gif")
                .Item(7).Picture = LoadPicture(SetDirPath(1) & "\" & "ImageBox(7)" & ".gif")
            End With
            
            With PictureBox
                .Item(0).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(0)" & ".gif")
                .Item(1).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(1)" & ".gif")
                .Item(2).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(2)" & ".gif")
                .Item(3).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(3)" & ".gif")
                .Item(4).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(4)" & ".gif")
                .Item(5).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(5)" & ".gif")
                .Item(6).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(6)" & ".gif")
                .Item(7).Picture = LoadPicture(SetDirPath(1) & "\" & "PictureBox(7)" & ".gif")
            End With
            
        Case 2
            SetDirPath(2) = App.Path & "\Graphics\Silver"
            
            With ImageBox
                .Item(0).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(0)" & ".gif")
                .Item(1).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(1)" & ".gif")
                .Item(2).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(2)" & ".gif")
                .Item(3).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(3)" & ".gif")
                .Item(4).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(4)" & ".gif")
                .Item(5).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(5)" & ".gif")
                .Item(6).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(6)" & ".gif")
                .Item(7).Picture = LoadPicture(SetDirPath(2) & "\" & "ImageBox(7)" & ".gif")
            End With
            
            With PictureBox
                .Item(0).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(0)" & ".gif")
                .Item(1).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(1)" & ".gif")
                .Item(2).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(2)" & ".gif")
                .Item(3).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(3)" & ".gif")
                .Item(4).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(4)" & ".gif")
                .Item(5).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(5)" & ".gif")
                .Item(6).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(6)" & ".gif")
                .Item(7).Picture = LoadPicture(SetDirPath(2) & "\" & "PictureBox(7)" & ".gif")
            End With
    End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AlwaysOnTop = PropBag.ReadProperty("AlwaysOnTop", False)
    If m_AlwaysOnTop = True Then
        SetWindowPos UserControl.ContainerHwnd, -1, 0, 0, 0, 0, &H2 Or &H1
    Else
        SetWindowPos UserControl.ContainerHwnd, -2, 0, 0, 0, 0, &H2 Or &H1
    End If
    m_BackColor = PropBag.ReadProperty("BackColor", &HF3DED6)
    UserControl.BackColor = m_BackColor
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Label(0).Caption = m_Caption
    m_CaptionOffsetX = PropBag.ReadProperty("CaptionOffsetX", 30)
    m_CaptionOffsetY = PropBag.ReadProperty("CaptionOffsetY", 20)
    Label(0).Left = m_CaptionOffsetX
    Label(0).Top = m_CaptionOffsetY
    m_ColorScheme = PropBag.ReadProperty("ColorScheme", 0)
    If m_ColorScheme = 0 Then
        'Call SetColorShm(0)
    ElseIf m_ColorScheme = 1 Then
        Call SetColorShm(1)
    ElseIf m_ColorScheme = 2 Then
        Call SetColorShm(2)
    End If
    m_ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    Label(0).ForeColor = m_ForeColor
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Label(0).Font = m_Font
    'Set Icon = PropBag.ReadProperty("Icon", IconBox(1).Picture)
    'Set Icon = IconBox(0).Picture
'    m_IconOffsetX = PropBag.ReadProperty("IconOffsetX", 10)
'    m_IconOffsetY = PropBag.ReadProperty("IconOffsetY", 20)
'    IconBox(0).Left = m_IconOffsetX
'    IconBox(0).Top = m_IconOffsetY
    m_MaxButton = PropBag.ReadProperty("MaxButton", True)
    If m_MaxButton = False Then
        Call PictureBox(1).PaintPicture(PictureBox(5).Picture, -78, 0, PictureBox(5).Width, PictureBox(5).Height)
    Else
        Call PictureBox(1).PaintPicture(PictureBox(5).Picture, 0, 0, PictureBox(5).Width, PictureBox(5).Height)
    End If
    m_MinButton = PropBag.ReadProperty("MinButton", True)
    If m_MinButton = False Then
        Call PictureBox(3).PaintPicture(PictureBox(7).Picture, -78, 0, PictureBox(7).Width, PictureBox(7).Height)
    Else
        Call PictureBox(3).PaintPicture(PictureBox(7).Picture, 0, 0, PictureBox(7).Width, PictureBox(7).Height)
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AlwaysOnTop", m_AlwaysOnTop, False)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HF3DED6)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("CaptionOffsetX", m_CaptionOffsetX, 30)
    Call PropBag.WriteProperty("CaptionOffsetY", m_CaptionOffsetY, 20)
    Call PropBag.WriteProperty("ColorScheme", m_ColorScheme, 0)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    'Call PropBag.WriteProperty("Icon", m_Icon, IconBox(1).Picture)
'    Call PropBag.WriteProperty("IconOffsetX", m_IconOffsetX, 10)
'    Call PropBag.WriteProperty("IconOffsetY", m_IconOffsetY, 20)
    Call PropBag.WriteProperty("MaxButton", m_MaxButton, True)
    Call PropBag.WriteProperty("MinButton", m_MinButton, True)
End Sub

Public Function Timeout(Interval)
Dim i As String
    i = Timer

    Do While Timer - i < Val(Interval)
        DoEvents
    Loop
End Function
