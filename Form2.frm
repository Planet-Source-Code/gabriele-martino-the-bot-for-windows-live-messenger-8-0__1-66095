VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1665
   ClientLeft      =   11715
   ClientTop       =   6780
   ClientWidth     =   2310
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Left            =   1560
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "New Event!!"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   600
         Top             =   840
      End
      Begin VB.Label eventtext 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub eventtext_Change()
Form1.List1.AddItem (Format(Time) + " " + eventtext.Caption)
Open currentlog For Append As 1
Print #1, eventtext.Caption
Close
Form_Load
End Sub

Private Sub Form_Load()
If request = False Then Exit Sub
Me.Show
Me.Top = 0
Timer2.Interval = 2300
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 30
If Me.Top > Screen.Height Then
Timer1.Enabled = False
Me.Hide
'Set frm(countform) = Nothing
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer1.Enabled = True
End Sub
