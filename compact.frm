VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Diary"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "compact.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   840
      TabIndex        =   11
      Top             =   1800
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12640511
      Appearance      =   1
      StartOfWeek     =   24576001
      CurrentDate     =   36826
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   3600
      Picture         =   "compact.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "SET SYSTEM TIME"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0FF&
      Height          =   735
      Left            =   2880
      Picture         =   "compact.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "SWITCH TO COMPACT MODE"
      Top             =   840
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Height          =   615
      Left            =   1800
      Picture         =   "compact.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "IMAGE VIEWER"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   3600
      Picture         =   "compact.frx":1458
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "SWITCH TO FULL MODE"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   2400
      Picture         =   "compact.frx":189A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "STORE INFORMATION"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   3000
      Picture         =   "compact.frx":1CDC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "TIME / DATE / CALENDAR"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   1200
      Picture         =   "compact.frx":241E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "PLAY MOVIES OR SONGS"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   600
      Picture         =   "compact.frx":2860
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "SURF THE WEB"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   0
      Picture         =   "compact.frx":5002
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "CALCULATOR / PHONE DIALER"
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Show
End Sub

Private Sub Command2_Click()
Form8.Show
End Sub

Private Sub Command3_Click()
form6.Show
End Sub

Private Sub Command4_Click()
Form4.Height = Form4.Height + 3700
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
memo = 2
Form3.Show
End Sub

Private Sub Command6_Click()
Unload Me
form1.Show
Unload Form5
Unload Form3
Unload Form7
End Sub

Private Sub Command7_Click()
Form10.Show
End Sub

Private Sub Command8_Click()
Form4.Height = Form4.Height - 3700
Command4.Enabled = True
End Sub

Private Sub Command9_Click()
Dim mytime
mytime = InputBox("ENTER NEW TIME")
Time = mytime
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload form1
Unload Form2
Unload Form3
Unload Form5
Unload Form7
Unload Form8
Unload Form9
Unload Form10
Unload Me
End Sub

Private Sub Timer1_Timer()
Text1.Text = Format(Now, "dddd,mmm dd,yyyy")
Text2.Text = Format(Now, "hh:mm:ss AM/PM")
End Sub
