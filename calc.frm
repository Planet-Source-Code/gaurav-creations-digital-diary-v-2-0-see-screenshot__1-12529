VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator/Phone "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2730
   Icon            =   "calc.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   2730
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dial Number"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0FFFF&
      Caption         =   "DISCONNECT"
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
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command21 
      Caption         =   "T  a  n"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Tangent of an angle "
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      Caption         =   "C   o   s"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "Cosine of an angle "
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      Caption         =   "S  i  n"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Sine of an angle "
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   18
      ToolTipText     =   "Squareroot of a number"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      ToolTipText     =   "Display result"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      ToolTipText     =   "Clear Display"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "Division"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      ToolTipText     =   "Multiplication"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   13
      ToolTipText     =   "Subtraction"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      ToolTipText     =   "Addition"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "Decimal"
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command22 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command24 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command25 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "0."
      ToolTipText     =   "Display Output & Accept Input"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   0
      Top             =   -120
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Double
Dim a(100) As String
Dim add As Double
Dim d As Integer
Dim e As Double
Dim su As Double
Dim s As Integer
Dim g As Double
Dim mul As Double
Dim f As Integer
Dim m As Integer
Dim n As Double
Dim div As Double
Dim o As Integer
Dim p As Double
Dim sq As Double

Dim si As Double
Dim co As Double
Dim ta As Double
Dim CancelFlag
Private Sub Command26_Click()
Text1.Text = "1"
a(i) = (Text1.Text)

i = i + 1
For b = 0 To i - 1
    c = c + a(b)
    Text1.Text = c
Next b

End Sub

Private Sub Command10_Click()
Text1.Text = "0"
a(i) = (Text1.Text)

i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command11_Click()
If f = 0 Then
Text1.Text = "."
a(i) = (Text1.Text)

i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
f = 1
Else
MsgBox "DON't TRY TO USE TWO DECIMAL POINTS", vbCritical, "ERROR"
End If
End Sub

Private Sub Command12_Click()
f = 0
add = Text1.Text
Text1.Text = "0."

d = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command13_Click()
f = 0
su = Text1.Text
Text1.Text = "0."

s = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command14_Click()
f = 0
mul = Text1.Text
Text1.Text = "0."

m = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command15_Click()
f = 0
div = Text1.Text
Text1.Text = "0."

o = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command16_Click()
For b = 0 To i - 1
a(b) = 0
Next b
i = 0
Text1.Text = "0."
End Sub

Private Sub Command17_Click()
If d = 1 Then
e = Text1.Text + add
Text1.Text = e
d = 0
ElseIf s = 1 Then
g = su - Text1.Text
Text1.Text = g
s = 0
ElseIf m = 1 Then
n = mul * Text1.Text
Text1.Text = n
m = 0
ElseIf o = 1 Then
If (Text1.Text > 0) Then
p = div / Text1.Text
Text1.Text = p
o = 0
ElseIf (Text1.Text) = 0 Then
MsgBox ("INVALID CAN'T DIVIDE BY 0"), vbCritical, "ERROR"
For b = 0 To i - 1
a(b) = 0
Next b
Text1.Text = "0."
End If
End If

For b = 0 To i - 1
a(b) = 0
Next b
End Sub

Private Sub Command18_Click()
If (Text1.Text) > 0 Then
sq = Sqr(Text1.Text)
Text1.Text = sq

ElseIf Text1.Text = 0 Then
MsgBox ("INVALID: NOT DEFINED"), vbCritical, "ERROR"
For b = 0 To i - 1
a(b) = 0
Next b
Text1.Text = "0."
End If
End Sub

Private Sub Command19_Click()
si = Sin(Text1.Text)
Text1.Text = si
End Sub

Private Sub Command25_Click()
Text1.Text = "2"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command20_Click()
co = Cos(Text1.Text)
Text1.Text = co
End Sub

Private Sub Command21_Click()
ta = Tan(Text1.Text)
Text1.Text = ta
End Sub

Private Sub Command24_Click()
Text1.Text = "3"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command23_Click()
Text1.Text = "4"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command22_Click()
Text1.Text = "5"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub



Private Sub Command32_Click()
    Dim Number$, Temp$
    
    Command32.Enabled = False
    Command33.Enabled = True
    
    ' Get the number to dial.
    Number$ = Text1.Text
        If Number$ = "" Then Exit Sub
    Temp$ = Status
    Status = "Dialing - " + Number$
    
    ' Dial the selected phone number.
    Dial Number$

   Command32.Enabled = True
   Command33.Enabled = False

    Status = Temp$
End Sub

Private Sub Command33_Click()
    ' CancelFlag tells the Dial procedure to exit.
   CancelFlag = True
   Command33.Enabled = False
End Sub

Private Sub Command6_Click()
Text1.Text = "6"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command7_Click()
Text1.Text = "7"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command8_Click()
Text1.Text = "8"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command9_Click()
Text1.Text = "9"
a(i) = (Text1.Text)
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub
Private Sub Dial(Number$)
    Dim DialString$, FromModem$, dummy

    ' AT is the Hayes compatible ATTENTION command and is required to send commands to the modem.
    ' DT means "Dial Tone." The Dial command uses touch tones, as opposed to pulse (DP = Dial Pulse).
    ' Numbers$ is the phone number being dialed.
    ' A semicolon tells the modem to return to command mode after dialing (important).
    ' A carriage return, vbCr, is required when sending commands to the modem.
    DialString$ = "ATDT" + Number$ + ";" + vbCr

    ' Communications port settings.
    ' Assuming that a mouse is attached to COM1, CommPort is set to 2
    MSComm1.CommPort = 2
    MSComm1.Settings = "9600,N,8,1"
    
    ' Open the communications port.
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    ' Flush the input buffer.
    MSComm1.InBufferCount = 0
    
    ' Dial the number.
    MSComm1.Output = DialString$
    
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, "OK") Then
             ' Notify the user to pick up the phone.
             Beep
             MsgBox "Please pick up the phone and either press Enter or click OK", , "Digital Diary"
             Exit Do
          End If
       End If
        
       ' Did the user choose Cancel?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
    Loop
    
    ' Disconnect the modem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Close the port.
    MSComm1.PortOpen = False
End Sub

Private Sub Form_Load()
MSComm1.InputLen = 0
End Sub
