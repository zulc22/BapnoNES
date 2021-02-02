VERSION 5.00
Begin VB.Form frmRamEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RAM Editor"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   Icon            =   "frmRamEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Substituir"
      Height          =   855
      Left            =   3240
      TabIndex        =   10
      Top             =   2760
      Width           =   2535
      Begin VB.TextBox txtReplace 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtReplaceTo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Hex:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   400
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "=>"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   400
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ram Viewer"
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   7335
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
         TabIndex        =   8
         Top             =   240
         Width           =   6840
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1900
         LargeChange     =   16
         Left            =   6960
         Max             =   0
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7680
      Top             =   2760
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6000
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
      Begin VB.CommandButton CmdApply 
         Caption         =   "&OK"
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "00"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtRamAddrEd 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0000"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "=>"
         Height          =   195
         Left            =   1560
         TabIndex        =   3
         Top             =   405
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hex:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   400
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmRamEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdApply_Click()
    UpdateRam
End Sub
Private Sub Command1_Click()
    Dim i As Integer
    
    For i = 0 To UBound(Bank0)
        If Bank0(i) = Val("&h" & txtReplace.Text) Then
            Bank0(i) = Val("&h" & txtReplaceTo.Text)
        End If
    Next i
End Sub
Private Sub Form_Load()
    If Lang = 1 Then
        Frame2.Caption = "Edit"
        Frame3.Caption = "Replace"
    End If
    
    VScroll1.Max = (UBound(Bank0) - 256) \ 16
End Sub
Public Sub VScroll1_Change()
    Dim i As Integer
    
    PicHex.Cls
    PicHex.CurrentY = 0
    PicHex.Print "       0  1  2  3  4  5  6  7   8  9  A  B  C  D  E  F"
    For i = 0 To 15
        RenderHexLines CLng(VScroll1.Value * 16) + (16 * i)
    Next i
End Sub
Private Sub RenderHexLines(AddY As Long)
    Dim StrPrint As String
    Dim i As Integer
    
    StrPrint = FixAddress(Hex(AddY)) & " - "
    For i = 0 To 15
        StrPrint = StrPrint & FixHex(Hex(Bank0(AddY + i))) & " "
        If i = 7 Then StrPrint = StrPrint & " "
    Next i
    StrPrint = StrPrint & " "
    For i = 0 To 15
        If Bank0(AddY + i) >= 32 And Bank0(AddY + i) <= 95 Or Bank0(AddY + i) > 96 And Bank0(AddY + i) < 127 Then
            StrPrint = StrPrint & Chr(Bank0(AddY + i))
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
Private Sub Text1_Change()
    UpdateRam
End Sub
Private Sub Timer1_Timer()
    VScroll1_Change
End Sub
Sub UpdateRam()
    On Error Resume Next
    Bank0(Val("&h" & txtRamAddrEd.Text)) = Val("&h" & Text1.Text)
End Sub
