VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Browser"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Text            =   "http://www.gauravcreations.com"
      Top             =   960
      Width           =   8415
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Picture         =   "Form8.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Favs"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Picture         =   "Form8.frx":2C94
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Picture         =   "Form8.frx":30D2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7335
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   12938
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command7 
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Picture         =   "Form8.frx":354C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      Picture         =   "Form8.frx":3A02
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      Picture         =   "Form8.frx":3EF4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      Picture         =   "Form8.frx":41A6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Picture         =   "Form8.frx":465C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Picture         =   "Form8.frx":4A7A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "Form8.frx":4F6C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9840
      Picture         =   "Form8.frx":521E
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Combo1_Click()
Combo1_KeyPress 13
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    WebBrowser1.Navigate Combo1.Text
    Combo1.AddItem Combo1.Text
End If
End Sub

Private Sub Command1_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command10_Click()
WebBrowser1.GoSearch
End Sub

Private Sub Command2_Click()
Combo1.Text = ("http://www.www.gauravcreations.com")
    WebBrowser1.Navigate Combo1.Text
    Combo1.AddItem Combo1.Text
End Sub

Private Sub Command3_Click()
WebBrowser1.Refresh
End Sub

Private Sub Command4_Click()
WebBrowser1.Stop
End Sub

Private Sub Command5_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command6_Click()
ShellExecute hwnd, "open", "mailto:" & Text6, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub Command7_Click()
Call ShellExecute(hwnd, "Open", "C:\Windows\History\", "", App.Path, 1)
End Sub

Private Sub Command8_Click()
Printer.Print WebBrowser1.Document
End Sub

Private Sub Command9_Click()
Call ShellExecute(hwnd, "Open", "C:\Windows\Favorites\", "", App.Path, 1)
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Status.Max = ProgressMax
Status.Value = Progress
End Sub


