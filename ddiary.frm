VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Diary"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   FillColor       =   &H00FFFFC0&
   FillStyle       =   0  'Solid
   Icon            =   "ddiary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C000&
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
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0C000&
      Caption         =   "Dial This Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CALENDAR"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00FFFFFF&
      Caption         =   "f"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "IMAGE VIEWER WITH MORE OPTIONS"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   9720
      Picture         =   "ddiary.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "SWITCH TO COMPACT MODE"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H000080FF&
      Caption         =   "  S  U  R  F          T  H  E        W E  B"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   3855
      Left            =   120
      Picture         =   "ddiary.frx":074C
      ScaleHeight     =   3795
      ScaleWidth      =   4395
      TabIndex        =   36
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SET SYSTEM TIME"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text3 
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
      Height          =   405
      Left            =   7680
      TabIndex        =   29
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
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
      Height          =   435
      Left            =   7680
      TabIndex        =   28
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9360
      Top             =   4320
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "PLAY MOVIES    AND    LISTEN SONGS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8280
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   4800
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   25
      Text            =   "0."
      ToolTipText     =   "Display Output"
      Top             =   4800
      Width           =   2535
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5760
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
      Left            =   6120
      TabIndex        =   23
      Top             =   5760
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
      Left            =   6840
      TabIndex        =   22
      Top             =   5760
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
      Left            =   5400
      TabIndex        =   21
      Top             =   6240
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
      Left            =   6120
      TabIndex        =   20
      Top             =   6240
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
      Left            =   6840
      TabIndex        =   19
      Top             =   6240
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
      Left            =   5400
      TabIndex        =   18
      Top             =   6720
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
      Left            =   6120
      TabIndex        =   17
      Top             =   6720
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
      Left            =   6840
      TabIndex        =   16
      Top             =   6720
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
      Left            =   5400
      TabIndex        =   15
      Top             =   7200
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
      Left            =   6120
      TabIndex        =   14
      ToolTipText     =   "Decimal"
      Top             =   7200
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
      Left            =   6840
      TabIndex        =   13
      ToolTipText     =   "Addition"
      Top             =   7200
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
      Left            =   5400
      TabIndex        =   12
      ToolTipText     =   "Subtraction"
      Top             =   7680
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
      Left            =   6120
      TabIndex        =   11
      ToolTipText     =   "Multiplication"
      Top             =   7680
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
      Left            =   6840
      TabIndex        =   10
      ToolTipText     =   "Division"
      Top             =   7680
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
      Left            =   6840
      TabIndex        =   9
      ToolTipText     =   "Clear Display"
      Top             =   8160
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
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Display result"
      Top             =   8160
      Width           =   495
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
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Squareroot of a number"
      Top             =   8160
      Width           =   495
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
      Left            =   4800
      TabIndex        =   6
      ToolTipText     =   "Sine of an angle "
      Top             =   5760
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
      Left            =   4800
      TabIndex        =   5
      ToolTipText     =   "Cosine of an angle "
      Top             =   6720
      Width           =   375
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
      Left            =   4800
      TabIndex        =   4
      ToolTipText     =   "Tangent of an angle "
      Top             =   7680
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "IMAGE VIEWER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "MEMORY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   7560
      X2              =   7680
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   7560
      X2              =   7560
      Y1              =   5160
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   7200
      X2              =   7560
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "MODE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   43
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "COMPACT "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   42
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   " IMAGE VIEWER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   720
      TabIndex        =   37
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Gaurav Creations"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7680
      MouseIcon       =   "ddiary.frx":1F1C
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  DIGITAL DIARY            By  "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   34
      Top             =   120
      Width           =   2775
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "DATE/TIME"
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
      Left            =   8280
      TabIndex        =   33
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   7560
      Shape           =   2  'Oval
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "MEDIA PLAYER"
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
      Left            =   240
      TabIndex        =   32
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "CALCULATOR"
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
      Left            =   4800
      TabIndex        =   31
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   4680
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      Height          =   2295
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   4095
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6615
   End
   Begin VB.Shape Shape3 
      Height          =   4335
      Left            =   120
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6615
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   3495
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Width           =   4215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' This flag is set when the user chooses Cancel.
Dim CancelFlag
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

Private Sub Command1_Click()
memo = 1
Form3.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

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

Private Sub Command27_Click()
Dim mytime
mytime = InputBox("ENTER NEW TIME")
Time = mytime
End Sub
Private Sub Command28_Click()
Form8.Show
End Sub

Private Sub Command29_Click()
Unload Me
Form4.Show
End Sub

Private Sub Command3_Click()
For z = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height - z
Next z
Picture1.Visible = False
For v = 0 To 27 Step 0.1
Picture1.Height = Picture1.Height + v
Next v
On Error GoTo form1
CommonDialog1.DialogTitle = "Load Media"
CommonDialog1.CancelError = True
CommonDialog1.Filter = "MP3 Files|*.mp3|WAV Files|*.wav|MIDI files|*.mid|AVI Files|*.avi|WMV Files|*.wmv|All Files|*.*"
CommonDialog1.ShowOpen
MediaPlayer1.Open (CommonDialog1.FileName)
mnuPlay.Enabled = True
mnuStop.Enabled = True
mnuPause.Enabled = True
mnuRewind.Enabled = True
form1:
form1.Show
End Sub


Private Sub Command30_Click()
Form10.Show
End Sub

Private Sub Command31_Click()
Form9.Show
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
Private Sub Command33_Click()
    ' CancelFlag tells the Dial procedure to exit.
   CancelFlag = True
   Command33.Enabled = False
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Label6.Visible = False
CommonDialog2.Filter = "All Files|*.*|BMP Files|*.bmp|JPEG files|*.jpeg|GIF Files|*.gif|JPG Files|*.jpg|ICO files|*.ico|CUR files|*.cur"
CommonDialog2.ShowOpen
Image1.Picture = LoadPicture(CommonDialog2.FileName)
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

Private Sub Form_Load()
'Used for changing the screen resolution
'This program is bst viewed with the resolution given below
'If you would like to put it back in do so by removing the comments
'from the 2 lines given below
'Note: SOME PC's MAY NOT SUPPORT IT SO DO IT AT YOUR OWN RISK

'RememberScreenRes
'ChangeScreenSettings 1024, 768, 16
MSComm1.InputLen = 0
Picture1.Picture = LoadPicture(App.Path & "\dancer.gif")
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload Form7
End Sub

Private Sub Label1_Click()
Form8.Show
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = True
Text2.Text = Format(Now, "dddd,mmm dd,yyyy")
Text3.Text = Format(Now, "hh:mm:ss AM/PM")
End Sub
