VERSION 5.00
Begin VB.Form frmRender 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6570
   ClientLeft      =   450
   ClientTop       =   450
   ClientWidth     =   7500
   Icon            =   "frmRender.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicScreen 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Sub Form_Load()
    Dim ZoomScale As Long
    Dim n As Long
    
    'Find the correct scale
    ZoomScale = Int((Screen.Height / Screen.TwipsPerPixelY) / 240)
    
    'Put the screen on correct position
    PicScreen.Move PicScreen.Left, PicScreen.Top, 256 * ZoomScale * Screen.TwipsPerPixelX, 240 * ZoomScale * Screen.TwipsPerPixelY
    
    'Put the form in FullScreen
    Caption = VERSION
    Top = 0
    Left = 0
    Width = Screen.Width
    Height = Screen.Height
    
    'Center the screen
    PicScreen.Top = (Me.Height) / 2 - PicScreen.Height / 2
    PicScreen.Left = Me.Width / 2 - PicScreen.Width / 2
    
    'Hide mouse pointer
    ShowCursor False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ShowCursor True
End Sub
Private Sub PicScreen_Paint()
    If CPUPaused Then BlitScreen
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        FScreen = False
        Unload Me
    ElseIf KeyCode = vbKeyF5 Then
        SaveState SlotIndex
    ElseIf KeyCode = vbKeyF7 Then
        LoadState SlotIndex
    Else
        Keyboard(KeyCode) = &H41
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Keyboard(KeyCode) = &H40
End Sub
