Attribute VB_Name = "APIStuff"
Option Explicit
Public OldWindowProc As Long
Public TheForm As Form
Public TheMenu As Menu
Public Howmany As Integer
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const WM_USER = &H400
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private TheData As NOTIFYICONDATA

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_NULL = &H0
Private Const APP_SYSTRAY_ID = 999 'unique identifier

Private Const NOTIFYICON_VERSION = &H3



Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10


Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)
Public Sub ShellTrayModifyTip(frm As Form, nIconIndex As Long)


   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With TheData
      .cbSize = NOTIFYICONDATA_SIZE
      .hWnd = frm.hWnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      
      'InfoTitle is the balloon tip title,
      'and szInfo is the message displayed.
      'Terminating both with vbNullChar prevents
      'the display of the unused padding in the
      'strings defined as fixed-length in NOTIFYICONDATA.
      .szInfoTitle = "HeadLines Today!" & vbNullChar
      .szInfo = "New headline is available." & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, TheData)

End Sub


Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub


Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

' The replacement window proc.
Public Function NewWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Const WM_NCDESTROY = &H82

    ' If we're being destroyed, remove the tray icon
    ' and restore the original WindowProc.
    If Msg = WM_NCDESTROY Then
        RemoveFromTray
    ElseIf Msg = TRAY_CALLBACK Then
        ' The user clicked on the tray icon.
        ' Look for click events.
        If lParam = WM_RBUTTONUP Then
            ' On right click, show the menu.
            SetForegroundWindow TheForm.hWnd
            TheForm.PopupMenu TheMenu
            If Not (TheForm Is Nothing) Then
                PostMessage TheForm.hWnd, WM_NULL, ByVal 0&, ByVal 0&
            End If
            Exit Function
        ElseIf lParam = WM_LBUTTONUP Then
            SetForegroundWindow TheForm.hWnd
            TheForm.mnuTrayDoStuff_Click
            If Not (TheForm Is Nothing) Then
                PostMessage TheForm.hWnd, WM_NULL, ByVal 0&, ByVal 0&
            End If
            Exit Function
        End If
    End If

    ' Send other messages to the original
    ' window proc.
    NewWindowProc = CallWindowProc( _
        OldWindowProc, hWnd, Msg, _
        wParam, lParam)
End Function
' Add the form's icon to the tray.
Public Sub AddToTray(frm As Form, mnu As Menu)
On Error Resume Next
    ' ShowInTaskbar must be set to False at
    ' design time because it is read-only at
    ' run time.

    ' Save the form and menu for later use.
    Set TheForm = frm
    Set TheMenu = mnu
    
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   ' Install the new WindowProc.
    OldWindowProc = SetWindowLong(frm.hWnd, _
        GWL_WNDPROC, AddressOf NewWindowProc)
   
  'set up the type members
   With TheData
   
      .cbSize = NOTIFYICONDATA_SIZE
      .hWnd = frm.hWnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = frm.Icon
      
      .uFlags = NIF_ICON
      .uCallbackMessage = TRAY_CALLBACK
      .uFlags = .uFlags Or NIF_MESSAGE
      'szTip is the tooltip shown when the
      'mouse hovers over the systray icon.
      'Terminate it since the strings are
      'fixed-length in NOTIFYICONDATA
      .szTip = "Headlines Today! Click to View." & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      
   End With
   
  'add the icon ...
   Call Shell_NotifyIcon(NIM_ADD, TheData)
   
  '... and inform the system of the
  'NOTIFYICON version in use
   Call Shell_NotifyIcon(NIM_SETVERSION, TheData)
End Sub
' Remove the icon from the system tray.
Public Sub RemoveFromTray()
On Error Resume Next
    ' Remove the icon from the tray.
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData

    ' Restore the original window proc.
    SetWindowLong TheForm.hWnd, GWL_WNDPROC, _
        OldWindowProc

    ' Clean up.
    Set TheForm = Nothing
End Sub
' Set a new tray tip.
Public Sub SetTrayTip(tip As String)
On Error Resume Next
    With TheData
        .szTip = tip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
' Set a new tray icon.
Public Sub SetTrayIcon(Pic As Picture)
On Error Resume Next
    ' Do nothing if the picture is not an icon.
    If Pic.Type <> vbPicTypeIcon Then Exit Sub

    ' Update the tray icon.
    With TheData
        .hIcon = Pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
