VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Headlines 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   270
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7785
   ControlBox      =   0   'False
   Icon            =   "GetHTTPSource.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   15
      Picture         =   "GetHTTPSource.frx":0E42
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   15
      Width           =   240
   End
   Begin VB.PictureBox PicHide 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   7245
      Picture         =   "GetHTTPSource.frx":1184
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   5
      ToolTipText     =   "Hide"
      Top             =   30
      Width           =   240
   End
   Begin VB.Timer TmrGetNews 
      Interval        =   65000
      Left            =   2205
      Top             =   -90
   End
   Begin VB.PictureBox PicClose 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7500
      Picture         =   "GetHTTPSource.frx":1466
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   240
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get"
      Default         =   -1  'True
      Height          =   315
      Left            =   4500
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   45
      TabIndex        =   1
      Text            =   "http://ww1.mid-day.com/includes/tick.htm"
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1800
      Top             =   1320
   End
   Begin VB.TextBox txtSource 
      Height          =   3015
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   4050
      Width           =   5415
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label LblNews 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   375
      TabIndex        =   4
      Top             =   30
      Width           =   990
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayDoStuff 
         Caption         =   "&Show Headlines"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Headlines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Dim sURL As String
Dim sHost As String
Dim sPage As String
Dim lPort As Long
Dim FileArray() As String
Dim ShowFrom As Integer
Dim GotNew As Boolean
Dim Shown As Boolean
Private Sub GetHTMLSource(ByVal sURL As String)
On Error GoTo Errhandler
  txtSource = ""
  GotNew = False
  sHost = Mid(sURL, InStr(sURL, "://") + 3)
  If InStr(sHost, "/") > 0 Then
    sPage = Mid(sHost, InStr(sHost, "/"))
    sHost = Left(sHost, InStr(sHost, "/") - 1)
  Else
    sPage = "/"
  End If
  If InStr(sHost, ":") > 0 Then
    lPort = Mid(sHost, InStr(sHost, ":") + 1)
    sHost = Left(sHost, InStr(sHost, ":") - 1)
  Else
    lPort = 80
  End If
  With Winsock1
    If .State <> sckClosed Then .Close
    .RemoteHost = sHost
    .RemotePort = lPort
    .Connect
  End With
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub
Private Sub Form_Load()
On Error GoTo Errhandler
    If App.PrevInstance = True Then
        Unload Me
        End
    End If
    ShowFrom = 1
    Me.Top = 5
    Me.Left = 5
    Me.Width = Screen.Width - 5
    PicClose.Left = Me.Width - 290
    PicHide.Left = Me.Width - 550
    DrawGradient Me.hDC, Me.Width, Me.Height + 200, RGB(10, 36, 106), RGB(165, 201, 239), 0
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    If Me.Visible = False And Timer1.Enabled = False Then Call GetHTMLSource(txtURL)
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Errhandler
    Timer1.Enabled = False
    TmrGetNews.Enabled = False
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub



Private Sub PicClose_Click()
On Error GoTo Errhandler
    DoEvents
    If MsgBox("Are you sure you want to close Headlines Today?" & vbNewLine & "You will not receive any news update if you close the program.", vbQuestion + vbYesNo, "HeadLines Today") = vbYes Then
        Unload Me
    End If
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub PicHide_Click()
On Error GoTo Errhandler
    Timer1.Enabled = False
    TmrGetNews.Enabled = True
    Me.Hide
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub Timer1_Timer()
On Error GoTo Errhandler
  Winsock1.Close
  txtSource = "Connection timeout"
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub


Private Sub TmrGetNews_Timer()
On Error GoTo Errhandler
    If Me.Visible = False And Timer1.Enabled = False Then Call GetHTMLSource(txtURL)
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub Winsock1_Close()
On Error GoTo Errhandler
    Call AppendToFile
    Call ReadFile
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub Winsock1_Connect()
On Error GoTo Errhandler
  Timer1.Enabled = True
  Winsock1.SendData "GET " & sPage & " HTTP/1.0" & Chr(10) & Chr(10)
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Errhandler
  Timer1.Enabled = False
  Dim sBuffer As String
  Winsock1.GetData sBuffer
  txtSource = txtSource & sBuffer
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo Errhandler
  Winsock1.Close
  'MsgBox "Error " & Number & ": " & Description, vbCritical, "WinSock Error"
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Private Sub AppendToFile()
On Error GoTo Errhandler
    tmp = txtSource.Text
    If InStr(1, Trim(tmp), "theSummaries[0] = '") Then
        tmp = Mid(Trim(tmp), InStr(1, tmp, "theSummaries[0] = '"))
        If InStr(1, Trim(tmp), "theSummaries[0] = '") Then
            tmparray = Split(Trim(tmp), "';")
        End If
    End If
    
'    If Dir(App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt") <> "" Then
'        Open App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt" For Input As #1
'            Do While Not EOF(1)
'                Line Input #1, getline
'                If GotNew = True Then
'                    ShowFrom = ShowFrom + 1
'                End If
'            Loop
'        Close #1
'    End If
    
    
    For X = 0 To UBound(tmparray)
        If InStr(1, tmparray(X), "theSummaries[") Then
            For Y = 0 To UBound(tmparray)
                tmpfinal = Replace(Trim(tmparray(X)), "theSummaries[" & Y & "] = ", "")
                If InStr(1, Trim(tmparray(X)), "theSummaries[" & Y & "] = ") Then Exit For
            Next
            tmpfinal = Right(Trim(tmpfinal), Len(tmpfinal) - 1)
            'tmpfinal = Left(Trim(tmpfinal), Len(tmpfinal) - 1)
            If Trim(tmpfinal) <> "" Then SaveNews (tmpfinal)
        End If
    Next
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub
Public Sub SaveNews(News As String)
On Error GoTo Errhandler
    Dim found As Boolean
    If Dir(App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt") <> "" Then
        Open App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt" For Input As #1
            lctr = 1
            Do While Not EOF(1)
                lctr = lctr + 1
                Line Input #1, tmpline
                If Trim(tmpline) = Trim("--> " & Replace(News, "\'", "'")) Then
                    found = True
                    'Exit Do
                End If
            Loop
        Close #1
    End If
    
    If found <> True Then
        If Shown = True Then ShowFrom = lctr
        Shown = False
        Open App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt" For Append As #1
            Print #1, "--> " & Replace(News, "\'", "'")
            GotNew = True
        Close #1
    End If
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Public Sub Display(Showheadline As String)
On Error GoTo Errhandler
    Dim tmpcaption As String
    For X = 0 To Len(Showheadline)
        DoEvents
        Sleep 50
        LblNews.Caption = "- " & Left(Showheadline, X)
        LblNews.Refresh
    Next
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub
Private Sub mnuTrayClose_Click()
    If MsgBox("Are you sure you want to close Headlines Today?" & vbNewLine & "You will not receive any news update if you close the program.", vbQuestion + vbYesNo, "HeadLines Today") = vbYes Then
        Unload Me
    End If
End Sub

Public Sub mnuTrayDoStuff_Click()
On Error GoTo Errhandler
    TmrGetNews.Enabled = False
    RemoveFromTray
    Me.Visible = True
    Shown = True
    Howmany = 0
continue:
    If Howmany = 5 Then
        If Me.Visible = True Then
            DoEvents
            Call Display("HeadLines Today : Brought to you by Nitin Shinde, Sr. Web Developer, Email:nitinps@hotmail.com")
            DoEvents
            Sleep 2750
        End If
        Call PicHide_Click
    Else
        If Me.Visible = True Then
            Dim NewsLoop As Boolean
            If ShowFrom > UBound(FileArray) Then ShowFrom = 1
            For X = ShowFrom To UBound(FileArray)
                If Me.Visible = True Then
                    NewsLoop = True
                    DoEvents
                    Call Display(Replace(FileArray(X), "-->", ""))
                    DoEvents
                    Sleep 2750
                End If
            Next
            If Me.Visible = True And NewsLoop = True Then
                Howmany = Howmany + 1
            End If
            GoTo continue
        End If
    End If
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub

Public Sub ReadFile()
On Error GoTo Errhandler
    DoEvents
    ReDim FileArray(0) As String
    If Dir(App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt") <> "" Then
        Open App.Path & "\" & Format(Date, "dd_mmm_yy") & ".txt" For Input As #1
            While Not EOF(1)
                ReDim Preserve FileArray(UBound(FileArray) + 1) As String
                Line Input #1, tmpline
                FileArray(UBound(FileArray)) = tmpline
            Wend
        Close #1
        
        If FileArray(UBound(FileArray)) <> "" And GotNew = True Then
            RemoveFromTray
            AddToTray Me, mnuTray
            ShellTrayModifyTip Me, 1
            SetTrayTip "Headlines Today! Click to View."
        End If
    End If
Exit Sub
Errhandler:
    MsgBox Err.Description, vbCritical, "HeadLines Today"
End Sub
