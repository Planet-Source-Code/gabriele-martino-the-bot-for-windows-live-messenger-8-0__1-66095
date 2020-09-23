Attribute VB_Name = "module"
Option Explicit
 
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public mArrc() As Messengerapi.IMessengerContact
Public currentlog As String
Public countform As Byte
Public request As Boolean
Public frm(16) As Form2


Public Sub SendMessage(thetext As String)
Dim imwindowclass As Long, richedita As Long
imwindowclass = FindWindow("imwindowclass", vbNullString)
richedita = FindWindowEx(imwindowclass, 0&, "richedit20a", vbNullString)
richedita = FindWindowEx(imwindowclass, richedita, "richedit20a", vbNullString)
Call SendMessageByString(richedita, WM_SETTEXT, 0&, thetext)
Call clicksend
End Sub

Public Sub clicksend()
Dim imwindowclass As Long, Button As Long
imwindowclass = FindWindow("imwindowclass", vbNullString)
Button = FindWindowEx(imwindowclass, 0&, "button", vbNullString)
Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
End Sub

'function to get a string from the status code
Public Function GetStatus(ByVal st As Messengerapi.MISTATUS) As String
Select Case st
Case Is = MISTATUS_AWAY
GetStatus = "Away"
Case Is = MISTATUS_BE_RIGHT_BACK
GetStatus = "Be right back"
Case Is = MISTATUS_BUSY
GetStatus = "Busy"
Case Is = MISTATUS_IDLE
GetStatus = "Idle"
Case Is = MISTATUS_INVISIBLE
GetStatus = "Invisible"
Case Is = MISTATUS_OFFLINE
GetStatus = "Offline"
Case Is = MISTATUS_ON_THE_PHONE
GetStatus = "On the phone"
Case Is = MISTATUS_ONLINE
GetStatus = "Online"
Case Is = MISTATUS_OUT_TO_LUNCH
GetStatus = "Out to lunch"
Case Is = MISTATUS_UNKNOWN
GetStatus = "Unknown"
End Select
End Function

Public Sub ShowEvent(pUp As String)
On Error GoTo errh
Dim x, k As Byte
k = 1
Set frm(0) = New Form2
frm(0).Visible = False
For x = 1 To 16
If frm(x).Visible = True Then k = k + 1
Next
If frm(countform).Visible = False Then
frm(countform + 1).Left = Screen.Width - frm(0).Width
Else
frm(countform + 1).Left = Screen.Width - frm(0).Width * k
End If
request = True
frm(countform + 1).eventtext.Caption = pUp
countform = countform + 1
Set frm(0) = Nothing
request = False
Exit Sub
errh:
countform = 0
For x = 1 To 16
frm(x).Visible = False
Next
Resume
End Sub

