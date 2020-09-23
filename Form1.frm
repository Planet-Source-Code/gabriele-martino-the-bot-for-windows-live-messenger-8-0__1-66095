VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Bot! for Windows Live Messenger"
   ClientHeight    =   5880
   ClientLeft      =   3720
   ClientTop       =   2535
   ClientWidth     =   8265
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8265
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Refresh"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Auto Login"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3600
      Top             =   2280
   End
   Begin VB.Timer agg 
      Interval        =   60000
      Left            =   3000
      Top             =   2280
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "About..."
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Go to Tray"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Events of this session"
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   8055
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   7815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08CA
            Key             =   "Online"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11A4
            Key             =   "tray"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A7E
            Key             =   "cOff"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2358
            Key             =   "cBusy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3DE6
            Key             =   "cAway"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":46C0
            Key             =   "offline"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5874
            Key             =   "cOn"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":614E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7302
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":84B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":966A
            Key             =   "trayicon"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   3960
      Top             =   3480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Contacts"
      Height          =   4455
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sort By Group (experimental)"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   525
         Left            =   4200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "Form1.frx":9F44
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5280
         Top             =   2640
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5640
         TabIndex        =   16
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5280
         TabIndex        =   15
         Top             =   3480
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Send not now, but at"
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Send instant message"
         Enabled         =   0   'False
         Height          =   975
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox blocked 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Blocked"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3855
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6800
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "message:"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label status 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label phone 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone number:"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label mail 
         BackStyle       =   0  'Transparent
         Caption         =   "mail"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail address:"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Check mail"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log in to check email"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   960
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Menu pop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu rest 
         Caption         =   "Restore"
      End
      Begin VB.Menu fdjhfgtkh 
         Caption         =   "-"
      End
      Begin VB.Menu xit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'set declarations for windows live! messenger

'this line creates events of messenger
Dim WithEvents winmesseng As Messengerapi.Messenger
Attribute winmesseng.VB_VarHelpID = -1

'this is useful to work with contacts
Dim oCons As Messengerapi.IMessengerContacts
Dim oCon As Messengerapi.IMessengerContact
Public theNode As MSComctlLib.Node

Private Sub agg_Timer()
If Not agg.Enabled = False Then Form_Load
End Sub

Private Sub Check1_Click()
Timer2.Enabled = True
End Sub

Private Sub Check2_Click()
Form_Load
End Sub

Private Sub Combo1_Click()
'this code change status in messenger when you check a status in the combo box
'if you have not logged in, this sub will be ignored
If winmesseng.MyStatus = MISTATUS_OFFLINE Then Exit Sub
Select Case Combo1.Text
Case Is = "Away"
winmesseng.MyStatus = MISTATUS_AWAY
Case Is = "Be right back"
winmesseng.MyStatus = MISTATUS_BE_RIGHT_BACK
Case Is = "Busy"
winmesseng.MyStatus = MISTATUS_BUSY
Case Is = "Idle"
winmesseng.MyStatus = MISTATUS_IDLE
Case Is = "Invisible"
winmesseng.MyStatus = MISTATUS_INVISIBLE
Case Is = "Offline"
agg.Enabled = False
Timer4.Enabled = True
winmesseng.Signout
Case Is = "On the phone"
winmesseng.MyStatus = MISTATUS_ON_THE_PHONE
Case Is = "Online"
winmesseng.MyStatus = MISTATUS_ONLINE
Case Is = "Out to lunch"
winmesseng.MyStatus = MISTATUS_OUT_TO_LUNCH
Case Is = "Unknown"
winmesseng.MyStatus = MISTATUS_UNKNOWN
End Select
End Sub

Private Sub Command1_Click()
'this open the inbox in internet explorer
'if you have not logged in, this sub will be ignored
If winmesseng.MyStatus = MISTATUS_OFFLINE Then Exit Sub
winmesseng.OpenInbox
End Sub

Private Sub Command2_Click()
If winmesseng.MyStatus = MISTATUS_OFFLINE Then Exit Sub
'this open a new chat window
winmesseng.InstantMessage mail.Caption
End Sub

Private Sub Command3_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Command4_Click()
frmAbout.Show
End Sub

Private Sub Command5_Click()
On Error Resume Next
winmesseng.AutoSignin
End Sub

Private Sub Command6_Click()
agg_Timer
End Sub

Private Sub Form_Initialize()
'create instances for events windows (i gave this name to form2 because it reports each event of messenger)
For x = 1 To 16
Set frm(x) = New Form2
Next
Set winmesseng = New Messengerapi.Messenger
If winmesseng.MyStatus = MISTATUS_OFFLINE Then winmesseng.AutoSignin
'create log file
Set fso = CreateObject("scripting.filesystemobject")
st = Str(Day(Date)) + "-" + Str(Month(Date)) + "-" + Str(Year(Date)) + "§" + Format(Time)
currentlog = winmesseng.ReceiveFileDirectory + "\" + st + ".log"
Set t = fso.createtextfile(winmesseng.ReceiveFileDirectory + "\" + st + ".log", 1)
t.Close
Me.Icon = ImageList1.ListImages("trayicon").Picture
'load setting
Slider1.Value = GetSetting(App.EXEName, "Options", "trasparent", 254)
'something about tray icon
Me.Show
Me.Refresh
       With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = ImageList1.ListImages("trayicon").Picture
        .szTip = "The Bot! for Messenger" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
'ok, let's go with main program
Command2.Picture = ImageList1.ListImages.Item(11).Picture
'add status names to the combo box
Combo1.AddItem ("Away")
Combo1.AddItem ("Be right back")
Combo1.AddItem ("Busy")
Combo1.AddItem ("Idle")
Combo1.AddItem ("Invisible")
Combo1.AddItem ("Offline")
Combo1.AddItem ("On the phone")
Combo1.AddItem ("Online")
Combo1.AddItem ("Out to lunch")
Combo1.AddItem ("Unknown")

End Sub

Private Sub Form_Load()
On Error GoTo aaa
aaa:
TreeView1.Nodes.Clear
If Check2.Value = 1 Then
Dim grp As Messengerapi.IMessengerGroup
For Each grp In winmesseng.MyGroups
Call TreeView1.Nodes.Add(, 2, grp.Name, grp.Name, "Online") 'this creates a node in treeview for each group
Set oCons = grp.Contacts
For Each oCon In oCons
If (oCon.status <> MISTATUS_OFFLINE) And (oCon.status <> MISTATUS_UNKNOWN) Then
TreeView1.Nodes.Add grp.Name, 4, , oCon.FriendlyName, "cOn"
Else
Call TreeView1.Nodes.Add(grp.Name, 4, , oCon.FriendlyName, "cOff")
End If
Next
Set oCons = Nothing
Set oCon = Nothing
Next
Else
For Each oCon In winmesseng.MyContacts
If (oCon.status <> MISTATUS_OFFLINE) And (oCon.status <> MISTATUS_UNKNOWN) Then
TreeView1.Nodes.Add , 2, , oCon.FriendlyName, "cOn"
Else
Call TreeView1.Nodes.Add(, , , oCon.FriendlyName, "cOff")
End If
Next oCon
Set oCon = Nothing
End If
'show an event window
If winmesseng.MyStatus = MISTATUS_OFFLINE Then ShowEvent ("You aren't online. Please log in")
End Sub

 Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
      'this procedure receives the callbacks from the System Tray icon.
      Dim Result As Long
      Dim Msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        Msg = x
       Else
        Msg = x / Screen.TwipsPerPixelX
       End If
       Select Case Msg
        Case WM_LBUTTONUP        '514 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu pop
       End Select
      End Sub
   
      Private Sub Form_Resize()
       'this is necessary to assure that the minimized window is hidden
       If Me.WindowState = vbMinimized Then Me.Hide
      End Sub
   
Private Sub Form_Unload(Cancel As Integer)
'this removes the icon from the system tray
Shell_NotifyIcon NIM_DELETE, nid
'close the log file
Open currentlog For Append As 1
Print #1, "Session terminated"
Close
End
      End Sub

Private Sub rest_Click()
'restore main window from the system tray
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.Show
End Sub

Private Sub Slider1_change()
inizio:
'sets transparency of main window
Call MakeTransparent(Me.hwnd, Slider1.Value)
SaveSetting App.EXEName, "Options", "trasparent", Slider1.Value
If Slider1.Value < 100 Then
a = MsgBox("The window may be too much transparent to be visible. Do you want to keep this setting anyway?", vbExclamation + vbYesNo, "Be Careful!!")
If a = vbYes Then
Else
Slider1.Value = 255
SaveSetting App.EXEName, "Options", "trasparent", Slider1.Value
GoTo inizio
End If
End If
End Sub

Private Sub Timer1_Timer()
Image1.Picture = ImageList1.ListImages("Online").Picture
Combo1.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Select Case winmesseng.MyStatus
Case Is = MISTATUS_AWAY
Combo1.Text = "Away"
Image1.Picture = ImageList1.ListImages("cAway").Picture
Case Is = MISTATUS_BE_RIGHT_BACK
Combo1.Text = "Be right back"
Image1.Picture = ImageList1.ListImages("cAway").Picture
Case Is = MISTATUS_BUSY
Combo1.Text = "Busy"
Image1.Picture = ImageList1.ListImages("cBusy").Picture
Case Is = MISTATUS_IDLE
Combo1.Text = "Idle"
Case Is = MISTATUS_INVISIBLE
Combo1.Text = "Invisible"
Case Is = MISTATUS_OFFLINE
Combo1.Text = "Offline"
Image1.Picture = ImageList1.ListImages("offline").Picture
Combo1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Case Is = MISTATUS_ON_THE_PHONE
Combo1.Text = "On the phone"
Case Is = MISTATUS_ONLINE
Combo1.Text = "Online"
Case Is = MISTATUS_OUT_TO_LUNCH
Combo1.Text = "Out to lunch"
Case Is = MISTATUS_UNKNOWN
Combo1.Text = "Unknown"
Image1.Picture = ImageList1.ListImages("offline").Picture
End Select
End Sub

Private Sub Timer2_Timer()
Check1.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
If Hour(Time) >= Val(Text1.Text) And Minute(Time) >= Val(Text2.Text) Then
Command2_Click
SendKeys (Text3.Text + "{Enter}")
Check1.Enabled = True
Check1.Value = 0
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
Form_Load
End Sub

Private Sub Timer4_Timer()
agg.Enabled = True
Timer4.Enabled = False
agg_Timer
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Children = 0 Then
Command2.Enabled = True
Set theNode = Node
theNode = Node
For Each grp In winmesseng.MyGroups
Set oCons = grp.Contacts
For Each oCon In oCons
If oCon.FriendlyName = theNode Then
status.Caption = GetStatus(oCon.status)
mail.Caption = oCon.SigninName
aaa = oCon.PhoneNumber(MPHONE_TYPE_HOME)
phone.Caption = aaa
blocked.Value = oCon.blocked
End If
Next oCon
Set oCons = Nothing
Set oCon = Nothing
Next
Else
Exit Sub
End If
End Sub

Private Sub winmesseng_OnContactFriendlyNameChange(ByVal hr As Long, ByVal pMcontact As Object, ByVal bstrPrevFriendlyName As String)
'this event is generated when a contact change its name
If agg.Enabled = False Then Exit Sub
ShowEvent pMcontact.FriendlyName + " has changed the name: the new is " + bstrPrevFriendlyName
Form_Load
End Sub

Private Sub winmesseng_OnContactStatusChange(ByVal pMcontact As Object, ByVal mStatus As Messengerapi.MISTATUS)
'this event is generated when a contact change its status
If agg.Enabled = False Then Exit Sub
ShowEvent pMcontact.FriendlyName + " has changed status: new status " + GetStatus(mStatus)
Form_Load
End Sub

Private Sub winmesseng_OnSignin(ByVal hr As Long)
'this event is generated when you sign in
agg.Enabled = False
Timer4.Enabled = True
Label2.Caption = "There aren't new emails"
ShowEvent "You have signed in succesfully"
End Sub

Private Sub winmesseng_OnSignout()
'this event is generated when you log off
agg.Enabled = False
Timer4.Enabled = True
ShowEvent "You have logged out"
End Sub

Private Sub winmesseng_OnUnreadEmailChange(ByVal mFolder As Messengerapi.MUAFOLDER, ByVal cUnreadEmail As Long, pBoolfEnableDefault As Boolean)
'this event is generated when you receive an e-mail
ShowEvent "You have " + Str(cUnreadEmail) + " new email(s)"
Label2.Caption = "You have " + Str(cUnreadEmail) + " new email(s)"
End Sub

Private Sub xit_Click()
End
End Sub
