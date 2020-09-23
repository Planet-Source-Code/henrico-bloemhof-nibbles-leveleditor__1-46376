VERSION 5.00
Begin VB.UserControl Knop 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00BAD6D4&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Knop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

'Default Property Values:
Const m_def_Klikkleur = &H80C0FF
Const m_def_Beweegkleur = &H70A093
Const m_def_Voorgrondkleur = &HFFFFFF
Const m_def_Voorgrondkleur1 = &HFFFFFF
Const m_def_Achtergrondkleur = &HF19470
Const m_def_Tekst = "[tekst]"
Const m_def_Randkleur = &HFFFFFF
'Property Variables:
Dim m_Klikkleur As OLE_COLOR
Dim m_Beweegkleur As OLE_COLOR
Dim m_Voorgrondkleur As OLE_COLOR
Dim m_Voorgrondkleur1 As OLE_COLOR
Dim m_Achtergrondkleur As OLE_COLOR
Dim m_Tekst As String
Dim m_Randkleur As OLE_COLOR
'Event Declarations:
Event klik() 'Mappinginfo=UserControl,UserControl,-1,Click
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim muisklik As Boolean
Dim beweging As Boolean

Private Sub Timer1_Timer()
    
    Dim pos As POINTAPI
    GetCursorPos pos
    
    ScreenToClient UserControl.hwnd, pos

    If pos.X < UserControl.ScaleLeft Or _
        pos.Y < UserControl.ScaleTop Or _
        pos.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
        pos.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        
        beweging = False
        If Not BackColor = m_Achtergrondkleur Then UserControl_Resize
    Else
        If UserControl.Parent.Enabled = True Then
            beweging = True
            If Not BackColor = m_Beweegkleur Then UserControl_Resize
        Else
            beweging = False
            If Not BackColor = m_Achtergrondkleur Then UserControl_Resize
        End If
    End If

End Sub

Private Sub UserControl_Initialize()
    If App.LogMode <> 1 Then
        If Not Timer1.Enabled = False Then Timer1.Enabled = False
    Else: If Not Timer1.Enabled = True Then Timer1.Enabled = True
    End If
    'If Not Timer1.Enabled = True Then Timer1.Enabled = True
    
    Cls
    BackColor = m_Achtergrondkleur
    ForeColor = m_Voorgrondkleur
    Line (DrawWidth - 1, DrawWidth - 1)-(ScaleWidth - 1, ScaleHeight - 1), m_Randkleur, B
    CurrentX = ScaleWidth / 2 - TextWidth(m_Tekst) / 2
    CurrentY = ScaleHeight / 2 - TextHeight(m_Tekst) / 2
    Print m_Tekst
End Sub

'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Memberinfo=14,0,0,&HFFFFFF
Public Property Get Randkleur() As OLE_COLOR
    Randkleur = m_Randkleur
End Property

Public Property Let Randkleur(ByVal New_Randkleur As OLE_COLOR)
    m_Randkleur = New_Randkleur
    PropertyChanged "Randkleur"
    UserControl_Resize
End Property

'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Mappinginfo=UserControl,UserControl,-1,Font
Public Property Get Lettertype() As Font
    Set Lettertype = UserControl.Font
End Property

Public Property Set Lettertype(ByVal New_Lettertype As Font)
    Set UserControl.Font = New_Lettertype
    PropertyChanged "Lettertype"
    UserControl_Resize
End Property

'initialize Properties for User Control
Private Sub UserControl_initProperties()
    Enabled = True
    m_Randkleur = m_def_Randkleur
    Set UserControl.Font = Ambient.Font
    m_Tekst = m_def_Tekst
    m_Beweegkleur = m_def_Beweegkleur
    m_Voorgrondkleur = m_def_Voorgrondkleur
    m_Achtergrondkleur = m_def_Achtergrondkleur
    m_Klikkleur = m_def_Klikkleur
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then muisklik = True
    If Not BackColor = m_Klikkleur Then UserControl_Resize
End Sub
'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'    beweging = True
'    If Not BackColor = m_Beweegkleur Then UserControl_Resize
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then
        muisklik = False
        If beweging = True Or Timer1.Enabled = False Then
            LostFocus
            RaiseEvent klik
        End If
    End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    m_Randkleur = PropBag.ReadProperty("Randkleur", m_def_Randkleur)
    Set UserControl.Font = PropBag.ReadProperty("Lettertype", Ambient.Font)
    m_Tekst = PropBag.ReadProperty("Tekst", m_def_Tekst)
    m_Beweegkleur = PropBag.ReadProperty("Beweegkleur", m_def_Beweegkleur)
    m_Voorgrondkleur = PropBag.ReadProperty("Voorgrondkleur", m_def_Voorgrondkleur)
    m_Achtergrondkleur = PropBag.ReadProperty("Achtergrondkleur", m_def_Achtergrondkleur)
    m_Klikkleur = PropBag.ReadProperty("Klikkleur", m_def_Klikkleur)
End Sub

Private Sub UserControl_Resize()
    Cls
    
    If beweging = False Or Enabled = False Then
        BackColor = m_Achtergrondkleur
    ElseIf muisklik = True Then
        BackColor = m_Klikkleur
    ElseIf beweging = True Then
        BackColor = m_Beweegkleur
    Else: BackColor = m_Achtergrondkleur
    End If
    
    ForeColor = m_Voorgrondkleur
    Line (DrawWidth - 1, DrawWidth - 1)-(ScaleWidth - 1, ScaleHeight - 1), m_Randkleur, B
    CurrentX = ScaleWidth / 2 - TextWidth(m_Tekst) / 2
    CurrentY = ScaleHeight / 2 - TextHeight(m_Tekst) / 2
    Print m_Tekst
End Sub

Private Sub UserControl_Show()
    Cls
    BackColor = m_Achtergrondkleur
    ForeColor = m_Voorgrondkleur
    Line (DrawWidth - 1, DrawWidth - 1)-(ScaleWidth - 1, ScaleHeight - 1), m_Randkleur, B
    CurrentX = ScaleWidth / 2 - TextWidth(m_Tekst) / 2
    CurrentY = ScaleHeight / 2 - TextHeight(m_Tekst) / 2
    Print m_Tekst
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Randkleur", m_Randkleur, m_def_Randkleur)
    Call PropBag.WriteProperty("Lettertype", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Tekst", m_Tekst, m_def_Tekst)
    Call PropBag.WriteProperty("Beweegkleur", m_Beweegkleur, m_def_Beweegkleur)
    Call PropBag.WriteProperty("Voorgrondkleur", m_Voorgrondkleur, m_def_Voorgrondkleur)
    Call PropBag.WriteProperty("Achtergrondkleur", m_Achtergrondkleur, m_def_Achtergrondkleur)
    Call PropBag.WriteProperty("Klikkleur", m_Klikkleur, m_def_Klikkleur)
End Sub

'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Memberinfo=13,0,0,[tekst]
Public Property Get tekst() As String
    tekst = m_Tekst
End Property

Public Property Let tekst(ByVal New_Tekst As String)
    m_Tekst = New_Tekst
    PropertyChanged "Tekst"
    UserControl_Resize
End Property

'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Memberinfo=10,0,0,&HFFFFFF
Public Property Get Beweegkleur() As OLE_COLOR
    Beweegkleur = m_Beweegkleur
End Property

Public Property Let Beweegkleur(ByVal New_Beweegkleur As OLE_COLOR)
    m_Beweegkleur = New_Beweegkleur
    PropertyChanged "Beweegkleur"
End Property

'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Memberinfo=10,0,0,&HFFFFFF
Public Property Get Voorgrondkleur() As OLE_COLOR
    Voorgrondkleur = m_Voorgrondkleur
End Property

Public Property Let Voorgrondkleur(ByVal New_Voorgrondkleur As OLE_COLOR)
    m_Voorgrondkleur = New_Voorgrondkleur
    m_Voorgrondkleur1 = New_Voorgrondkleur
    PropertyChanged "Voorgrondkleur"
    UserControl_Resize
End Property

'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Memberinfo=10,0,0,&HF19470
Public Property Get Achtergrondkleur() As OLE_COLOR
    Achtergrondkleur = m_Achtergrondkleur
End Property

Public Property Let Achtergrondkleur(ByVal New_Achtergrondkleur As OLE_COLOR)
    m_Achtergrondkleur = New_Achtergrondkleur
    PropertyChanged "Achtergrondkleur"
    UserControl_Resize
End Property

Public Sub LostFocus()
    If Not BackColor = m_Achtergrondkleur Then UserControl_Resize
End Sub
'WARNiNG! DO NOT REMOVE OR MODiFY THE FOLLOWiNG COMMENTED LiNES!
'Memberinfo=10,0,0,&HFFFFFF
Public Property Get Klikkleur() As OLE_COLOR
    Klikkleur = m_Klikkleur
End Property

Public Property Let Klikkleur(ByVal New_Klikkleur As OLE_COLOR)
    m_Klikkleur = New_Klikkleur
    PropertyChanged "Klikkleur"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    If Enabled = False Then
        m_Voorgrondkleur = &HE0E0E0
    Else
        m_Voorgrondkleur = m_Voorgrondkleur1
    End If
    UserControl_Resize
End Property


