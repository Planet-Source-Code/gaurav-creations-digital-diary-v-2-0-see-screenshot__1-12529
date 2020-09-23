VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MEMORY"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "memo.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   27
      Top             =   3480
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PHONE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   26
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REMOVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   240
      Width           =   2295
   End
   Begin VB.Data Data1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MEMORY"
      Connect         =   "Access"
      DatabaseName    =   "C:\Diary\DIARY.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MEMO"
      Top             =   4800
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      DataField       =   "ICQ"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      DataField       =   "HOMEPAGE"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      DataField       =   "E2"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      DataField       =   "E1"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      DataField       =   "ANNIVERSARY"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      DataField       =   "BIRTHDAY"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      DataField       =   "ADDRESS"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      DataField       =   "PHONE"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      DataField       =   "NAME"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line8 
      X1              =   7200
      X2              =   7200
      Y1              =   120
      Y2              =   4440
   End
   Begin VB.Line Line7 
      X1              =   4680
      X2              =   7200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line6 
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   4440
   End
   Begin VB.Line Line5 
      X1              =   4680
      X2              =   7200
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4560
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   120
      Y2              =   4440
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ICQ"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ANNIVERSARY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "HOMEPAGE"
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
      Left            =   360
      MouseIcon       =   "memo.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "CLICK TO VISIT THE WEBSITE"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E-mail 2"
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
      Left            =   360
      MouseIcon       =   "memo.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   14
      ToolTipText     =   "CLICK TO E-MAIL"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E-mail 1"
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
      Left            =   360
      MouseIcon       =   "memo.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   13
      ToolTipText     =   "CLICK TO E-MAIL"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PHONE"
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
      Left            =   360
      MouseIcon       =   "memo.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "CLICK TO CALL"
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "BIRTHDAY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Dim rs As Recordset
Private Sub Command1_Click()
newName = Text1
newPhone = Text2
newAddress = Text3
newBirthday = Text4
newAnniversary = Text5
newE1 = Text6
newE2 = Text7
newHomepage = Text8
newIcq = Text9
With rs
.AddNew
!Name = LCase(newName)
!Phone = LCase(newPhone)
!Address = LCase(newAddress)
!Birthday = LCase(newBirthday)
!Anniversary = LCase(newAnniversary)
!E1 = LCase(newE1)
!E2 = LCase(newE2)
!Homepage = LCase(newHomepage)
 Icq = newIcq
.Update
End With
End Sub

Private Sub Command2_Click()
If Data1.Recordset.BOF Or Data1.Recordset.EOF = True Then
MsgBox "  END OF FILE", vbOKOnly, "Diary"
Else
Data1.Recordset.Delete
End If
End Sub

Private Sub Command3_Click()

If Data1.Recordset.EOF = True Then
MsgBox "  END OF FILE", vbOKOnly, "Diary"
Else
Data1.Recordset.MoveNext
End If
End Sub

Private Sub Command4_Click()

If Data1.Recordset.BOF = True Then
MsgBox "  END OF FILE", vbOKOnly, "Diary"
Else
Data1.Recordset.MovePrevious
End If
End Sub

Private Sub Command5_Click()
If Data1.Recordset.BOF = True Then
MsgBox " NO RECORD", vbOKOnly, "Diary"
Else
Data1.Recordset.MoveFirst
End If
End Sub

Private Sub Command6_Click()
If Data1.Recordset.EOF = True Then
MsgBox " NO RECORD", vbOKOnly, "Diary"
Else
Data1.Recordset.MoveLast
End If
End Sub

Private Sub Command7_Click()
Unload Me
If memo = 1 Then
form1.Show
Else
If memo = 2 Then
Form4.Show
End If
End If
End Sub

Private Sub Command8_Click()
If Option1.Value = True Then
Set db = OpenDatabase(App.Path + "/DIARY.mdb")
Set rs = db.OpenRecordset("SELECT * FROM MEMO")
NameQuery = InputBox("Enter A Name To Search For", "Name Query")
rs.MoveFirst
Do Until rs.EOF
If rs.Fields("name") Like "*" & LCase(NameQuery) & "*" Then
Text1 = rs.Fields("Name")
Text2 = rs.Fields("PHONE")
Text3 = rs.Fields("ADDRESS")
Text4 = rs.Fields("BIRTHDAY")
Text5 = rs.Fields("ANNIVERSARY")
Text6 = rs.Fields("E1")
Text7 = rs.Fields("E2")
Text8 = rs.Fields("HOMEPAGE")
Text9 = rs.Fields("ICQ")
Exit Sub
Else
rs.MoveNext
End If
Loop
End If

If Option2.Value = True Then
Set db = OpenDatabase(App.Path + "/DIARY.mdb")
Set rs = db.OpenRecordset("SELECT * FROM MEMO")
NameQuery = InputBox("Enter A Phone To Search For", "Phone Query")
rs.MoveFirst
Do Until rs.EOF
If rs.Fields("Phone") Like "*" & LCase(phoneQuery) & "*" Then
Text1 = rs.Fields("Name")
Text2 = rs.Fields("PHONE")
Text3 = rs.Fields("ADDRESS")
Text4 = rs.Fields("BIRTHDAY")
Text5 = rs.Fields("ANNIVERSARY")
Text6 = rs.Fields("E1")
Text7 = rs.Fields("E2")
Text8 = rs.Fields("HOMEPAGE")
Text9 = rs.Fields("ICQ")
Exit Sub
Else
rs.MoveNext
End If
Loop
End If
End Sub

Private Sub Form_Load()
Set db = OpenDatabase(App.Path + "/Diary.mdb")
Set rs = db.OpenRecordset("memo")
End Sub

Private Sub Label3_Click()
Form5.Show
End Sub

Private Sub Label5_Click()
ShellExecute hwnd, "open", "mailto:" & Text6, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub label6_Click()
ShellExecute hwnd, "open", "mailto:" & Text7, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub Label7_Click()
Form8.Show
End Sub
