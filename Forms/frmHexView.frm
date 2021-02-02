VERSION 5.00
Begin VB.Form frmHexView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hex Viewer"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   Icon            =   "frmHexView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7560
      Top             =   2880
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hex Viewer"
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7335
      Begin VB.VScrollBar VScroll1 
         Height          =   1900
         LargeChange     =   16
         Left            =   6960
         Max             =   0
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox PicHex 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2025
         Left            =   120
         ScaleHeight     =   2025
         ScaleWidth      =   6840
         TabIndex        =   2
         Top             =   240
         Width           =   6840
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "frmHexView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Debug.Print UBound(GameImage)
    VScroll1.Max = (UBound(GameImage) - 256) \ 16
    VScroll1_Change
End Sub
Private Sub Timer1_Timer()
    VScroll1_Change
End Sub
Public Sub VScroll1_Change()
    Dim i As Integer
    Dim j As Long
    
    PicHex.Cls
    PicHex.CurrentY = 0
    j = VScroll1.Value 'Quick fix for strange bug
    PicHex.Print "       0  1  2  3  4  5  6  7   8  9  A  B  C  D  E  F"
    For i = 0 To 15
        RenderHexLines (j * 16) + (16 * i)
    Next i
End Sub
Private Sub RenderHexLines(AddY As Long)
    Dim StrPrint As String
    Dim i As Integer
    
    StrPrint = FixAddress(Hex(AddY)) & " - "
    For i = 0 To 15
        StrPrint = StrPrint & FixHex(Hex(GameImage(AddY + i))) & " "
        If i = 7 Then StrPrint = StrPrint & " "
    Next i
    StrPrint = StrPrint & " "
    For i = 0 To 15
        If GameImage(AddY + i) >= 32 And GameImage(AddY + i) <= 95 Or GameImage(AddY + i) > 96 And GameImage(AddY + i) < 127 Then
            StrPrint = StrPrint & Chr(GameImage(AddY + i))
        Else
            StrPrint = StrPrint & "."
        End If
    Next i
    PicHex.CurrentX = 0
    PicHex.Print StrPrint
End Sub
Private Sub CmdOK_Click()
    Unload Me
End Sub
