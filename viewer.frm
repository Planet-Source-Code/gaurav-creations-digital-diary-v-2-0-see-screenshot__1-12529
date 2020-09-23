VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form10 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Viewer"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "viewer.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Zoom OUT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      LargeChange     =   100
      Left            =   7560
      Min             =   100
      TabIndex        =   3
      Top             =   0
      Value           =   100
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   120
      Min             =   100
      TabIndex        =   2
      Top             =   4800
      Value           =   100
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Image Viewer"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Zoom IN"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Image1.Stretch = True
Image1.Height = Image1.Height + 200
Image1.Width = Image1.Width + 200
 ' Position the horizontal scroll bar.
   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
   HScroll1.Width = Picture1.Width

   ' Position the vertical scroll bar.
   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width
   VScroll1.Height = Picture1.Height

   ' Set the Max property for the scroll bars.
   HScroll1.Max = Image1.Width - Picture1.Width
   VScroll1.Max = Image1.Height - Picture1.Height

   ' Determine if the child picture will fill up the
   ' screen.
   ' If so, there is no need to use scroll bars.
   VScroll1.Visible = (Picture1.Height < _
  Image1.Height)
   HScroll1.Visible = (Picture1.Width < _
   Image1.Width)

End Sub

Private Sub Command1_Click()
   tp = 0
   Label6.Visible = False
   Command2.Enabled = True
   Command3.Enabled = True
   Image1.Stretch = False
   Unload Form7
   ' Set ScaleMode to pixels.
   Form10.ScaleMode = vbPixels
   Picture1.ScaleMode = vbPixels

 
   ' Load the bitmap.
  CommonDialog1.Filter = "All Files|*.*|BMP Files|*.bmp|JPEG files|*.jpeg|GIF Files|*.gif|JPG Files|*.jpg|ICO files|*.ico|CUR files|*.cur"
  CommonDialog1.ShowOpen
  Image1.Picture = LoadPicture(CommonDialog1.FileName)
  preview = CommonDialog1.FileName
     ' Initialize location of both pictures.
   Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, _
   ScaleHeight - HScroll1.Height
  Image1.Move 0, 0

   ' Position the horizontal scroll bar.
   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
   HScroll1.Width = Picture1.Width

   ' Position the vertical scroll bar.
   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width
   VScroll1.Height = Picture1.Height

   ' Set the Max property for the scroll bars.
   HScroll1.Max = Image1.Width - Picture1.Width
   VScroll1.Max = Image1.Height - Picture1.Height

   ' Determine if the child picture will fill up the
   ' screen.
   ' If so, there is no need to use scroll bars.
   VScroll1.Visible = (Picture1.Height < _
  Image1.Height)
   HScroll1.Visible = (Picture1.Width < _
   Image1.Width)

If VScroll1.Visible = True Then
 tp = 1
End If
If HScroll1.Visible = True Then
 tp = 1
End If
 Form7.Show
End Sub

Private Sub Command3_Click()
Image1.Stretch = True
If Image1.Height > 200 Then
Image1.Height = Image1.Height - 200
Image1.Width = Image1.Width - 200
 ' Position the horizontal scroll bar.
   HScroll1.Top = Picture1.Height
   HScroll1.Left = 0
   HScroll1.Width = Picture1.Width

   ' Position the vertical scroll bar.
   VScroll1.Top = 0
   VScroll1.Left = Picture1.Width
   VScroll1.Height = Picture1.Height

   ' Set the Max property for the scroll bars.
   HScroll1.Max = Image1.Width - Picture1.Width
   VScroll1.Max = Image1.Height - Picture1.Height

   ' Determine if the child picture will fill up the
   ' screen.
   ' If so, there is no need to use scroll bars.
   VScroll1.Visible = (Picture1.Height < _
  Image1.Height)
   HScroll1.Visible = (Picture1.Width < _
   Image1.Width)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form7
Unload Me
End Sub

Private Sub HScroll1_Change()
   Image1.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_Change()
   Image1.Top = -VScroll1.Value
End Sub

