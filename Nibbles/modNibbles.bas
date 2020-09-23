Attribute VB_Name = "modNibbles"
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
   hWndDesk = GetDesktopWindow()

   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)
   
  If success = SE_ERR_NOASSOC Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
   
End Sub


