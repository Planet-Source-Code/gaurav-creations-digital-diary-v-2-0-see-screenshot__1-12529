VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Full Image Preview"
   ClientHeight    =   1590
   ClientLeft      =   11850
   ClientTop       =   6735
   ClientWidth     =   2910
   Icon            =   "preview.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2910
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If tp = 1 Then
Image1.Stretch = True
End If

Image1.Picture = LoadPicture(preview)
End Sub
