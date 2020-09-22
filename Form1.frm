VERSION 5.00
Object = "{6AC6B294-E9DB-4D2D-8B31-1032A1756D70}#1.0#0"; "jcOffice2003.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'Kein
   Caption         =   "NSB ultra"
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   1  'Minimiert
   Begin jcOffice2003.JCToolbar JCToolbar1 
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   688
      BackColor       =   -2147483633
      Begin VB.TextBox countit 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
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
         Height          =   200
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   90
         Width           =   375
      End
      Begin jcOffice2003.JCF_Button JCF_Button1 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         ToolTipText     =   "Schließt NSB"
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         Caption         =   "x"
      End
      Begin jcOffice2003.ucVertical3DLine ucVertical3DLine2 
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   0
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   661
      End
      Begin jcOffice2003.JCF_Button closes 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         ToolTipText     =   "Msgs schließen?"
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Close Msgs?"
         IsCheckButton   =   -1  'True
      End
      Begin jcOffice2003.ucVertical3DLine ucVertical3DLine1 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   0
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   661
      End
      Begin jcOffice2003.JCF_Button Command2 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Msgs verstecken?"
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         State           =   4
         Enabled         =   0   'False
         Caption         =   "Hide Msg"
      End
      Begin jcOffice2003.JCF_Button Command1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Msgs zeigen?"
         Top             =   0
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   661
         Caption         =   "Show Msg"
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   1560
      Top             =   600
   End
   Begin NSB_by_bp.TextBalloon TextBalloon1 
      Left            =   120
      Top             =   2160
      _extentx        =   873
      _extenty        =   873
   End
   Begin NSB_by_bp.TrayIcon TrayIcon1 
      Left            =   120
      Top             =   1680
      _extentx        =   873
      _extenty        =   873
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1080
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' NOTE: This Project contains several codes/snippets which are not made by me.
' Some of them are written by authers whose names included, others are unknown to me.
' You need to register an OCX-file. I do not have the source of this, you can
' download the file at this location:

' --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' http://rapidshare.de/files/9201501/jcOffice2003.ocx.html
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' If you know where to get the sourcecode, contact me. thanx

' DO NOT FORGET TO SET MSGWINDTITLE TO YOUR NAME OF THE NETSEND-WINDOW

Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Const WM_CLOSE = &H10
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const tBarClass = "Shell_TrayWnd"

Const msgwindowtitle = "Nachrichtendienst "
' Note: Set this to the name of your netsend windowtitle! Maybe you need a
' following empty space after the name, try it out. good luck.

Dim tBarHwnd As Long
Dim tBarRect As RECT
Dim netsend ' Dim your way

Private Sub Command1_Click()
    Command2.Enabled = True
    Command1.Enabled = False
    Command2.State = 0
    Command1.State = 4
    Timer1.Enabled = False
    Timer2.Enabled = False
    WindowHandle netsend, 1 ' Shows the !FIRST! Net Send Window.
End Sub

Private Sub Command2_Click()
    Command1.Enabled = True
    Command2.Enabled = False
    Command2.State = 4
    Command1.State = 0
    Timer2.Enabled = True
End Sub

Private Sub Form_Load()
    TrayIcon1.TrayIcon = Me.Icon
    ' Set the systray Icon
    TrayIcon1.InTray = True
    ' Show the systray Icon
End Sub

Private Sub JCF_Button1_Click()
    TrayIcon1.InTray = False
    ' Removes the Icon from systray
    End
End Sub

Private Sub Timer1_Timer()
    netsend = FindWindow(vbNullString, msgwindowtitle)
    ' Re-Find Handle
    If netsend <> 0 Then
        X = WindowHandle(netsend, 2) ' Hide it! No Net Send will appear! But it will beep!
        If X <> 0 And closes.Value = 0 Then TrayIcon1.PopupBalloon "Recently hid a message!", "HIDDEN MESSAGE!", bsIconExclamation
        DoEvents
        If closes.Value = 1 Then X = WindowHandle(netsend, 0) ' Close the Net Send Window
        DoEvents
        If X = 0 And closes.Value = 1 Then
            TrayIcon1.PopupBalloon "Blocked " & countit.Text + 1 & " messages during runtime", "NetSendBlocker by BP", bsIconInformation
            countit.Text = countit.Text + 1 ' Show a baloonTip and count closed msgs!
        End If
    Else
        Timer1.Enabled = False
        Timer2.Enabled = True
    End If
End Sub

Private Sub timer2_timer()
    ' Searching for a Window called like on of Windows Net Send Service
    netsend = FindWindow(vbNullString, msgwindowtitle)
    If netsend <> 0 Then Timer2.Enabled = False: Timer1.Enabled = True
    ' Found one, then Hide/Close it!
End Sub

Private Function WindowHandle(win, cas As Long)

    'by storm  *** Big thanks ;)
    'Case 0 = CloseWindow
    'Case 1 = Show Win
    'Case 2 = Hide Win
    'Case 3 = Max Win
    'Case 4 = Min Win

    Select Case cas
        Case 0:
        WindowHandle = SendMessage(win, WM_CLOSE, 0, 0)
        Case 1:
        WindowHandle = ShowWindow(win, SW_SHOW)
        Case 2:
        WindowHandle = ShowWindow(win, SW_HIDE)
        Case 3:
        WindowHandle = ShowWindow(win, SW_MAXIMIZE)
        Case 4:
        WindowHandle = ShowWindow(win, SW_MINIMIZE)
    End Select

'any questions e-mail at storm@n2.com
'
End Function

Private Sub Timer3_Timer()
    ' Checks if Window is minimized
    If Me.WindowState = 1 Then Me.Visible = False Else Me.WindowState = 0
    DoEvents
End Sub

Private Sub TrayIcon1_MouseDown(Button As Integer, X As Single, Y As Single)
    ' Window is invisible, then show it!
    If Me.Visible = False Then Me.Visible = True: Me.WindowState = 0: TaskBarLocation: Me.SetFocus Else Me.Visible = False
    DoEvents
End Sub

Private Sub TaskBarLocation()
    ' I found this somewhere, do not know excactly where
    ' Attempt to get the handle to the taskbar
    tBarHwnd = FindWindow(tBarClass, "")
    If tBarHwnd = 0 Then
        ' Failed to find window
        MsgBox "Window Not Found !", 0, "Unable to retrieve hWnd"
    Else
        ' Window was found
        ' Get the location of the taskbar
        GetWindowRect tBarHwnd, tBarRect
        If tBarRect.Left = -2 And tBarRect.Top > -2 Then
            ' Taskbar is located at the bottom of the screen
            'MsgBox "Taskbar is at the bottom of the screen.."
            TaskBarLocationHeight = (Screen.Height / 15) - tBarRect.Top
            Me.Top = Screen.Height - ((TaskBarLocationHeight * 15) + Me.Height)
            Me.Left = Screen.Width - Me.Width
        End If
        
        If tBarRect.Left > -2 And tBarRect.Bottom = Screen.Height / Screen.TwipsPerPixelY + 2 Then
            ' Taskbar is at the right hand edge of the screen
            TaskBarLocationLeft = (Screen.Width / 15) - tBarRect.Left
            Me.Top = Screen.Height - Me.Height
            Me.Left = Screen.Width - (TaskBarLocationLeft * 15) - Me.Width
        End If
        
        If tBarRect.Bottom <> Screen.Height / Screen.TwipsPerPixelY + 2 And tBarRect.Right = Screen.Width / Screen.TwipsPerPixelX + 2 Then
            ' Taskbar is at the top of the screen
            TaskBarLocationHeight = tBarRect.Bottom
            Me.Top = TaskBarLocationHeight * 15
            Me.Left = Screen.Width - Me.Width
        End If
        
        If tBarRect.Right <> Screen.Width / Screen.TwipsPerPixelX + 2 And tBarRect.Bottom = Screen.Height / Screen.TwipsPerPixelY + 2 Then
            ' Taskbar is located at the left of the screen
            TaskBarLocationRight = tBarRect.Right
            Me.Top = Screen.Height - Me.Height
            Me.Left = TaskBarLocationRight * 15
        End If
    End If
End Sub

