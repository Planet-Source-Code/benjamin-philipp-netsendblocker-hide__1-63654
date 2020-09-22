Attribute VB_Name = "modCtls"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32" () As Long                    ' Used to enable xp style
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long      ' Used to display the
                                                                                        ' Balloon for textboxes

Public Type EDITBALLOONTIP          ' Type used to store the balloon data for textboxes
   cbStruct As Long
   pszTitle As String
   pszText As String
   ttiIcon As Long
End Type

Public Const ECM_FIRST As Long = &H1500                     ' Stores the start settings for the two below
Public Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)    ' Shows a balloon on a textbox
Public Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)    ' Hides a balloon from a textbox
