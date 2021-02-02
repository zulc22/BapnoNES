VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração de teclas"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   600
      TabIndex        =   53
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3240
      TabIndex        =   52
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controle 2 (Joystick)"
      Height          =   1935
      Left            =   240
      TabIndex        =   43
      Top             =   3600
      Width           =   4695
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   600
         TabIndex        =   51
         Top             =   1280
         Width           =   735
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         Height          =   255
         Left            =   600
         TabIndex        =   50
         Top             =   560
         Width           =   735
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         Height          =   255
         Left            =   2640
         TabIndex        =   49
         Top             =   1280
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   560
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle 2 (Teclado)"
      Height          =   1935
      Left            =   240
      TabIndex        =   26
      Top             =   3600
      Width           =   4695
      Begin VB.TextBox txt2Sta 
         Height          =   285
         Left            =   3240
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt2Sel 
         Height          =   285
         Left            =   3240
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt2B 
         Height          =   285
         Left            =   3240
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txt2A 
         Height          =   285
         Left            =   3240
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt2Right 
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt2Left 
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txt2Down 
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txt2Up 
         Height          =   285
         Left            =   1440
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   1480
         Width           =   735
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   1120
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   760
         Width           =   735
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   255
         Left            =   2640
         TabIndex        =   39
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Direita"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   1480
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Esquerda"
         Height          =   375
         Left            =   600
         TabIndex        =   37
         Top             =   1120
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Baixo"
         Height          =   375
         Left            =   600
         TabIndex        =   36
         Top             =   760
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Cima"
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   400
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Controle 1 (Joystick)"
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4695
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   560
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1280
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   560
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1280
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Controle 1 (Teclado)"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
      Begin VB.TextBox txtUp 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDown 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtLeft 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtRight 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtA 
         Height          =   285
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtB 
         Height          =   285
         Left            =   3240
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtSel 
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtSta 
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cima"
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   400
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Baixo"
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   760
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Esquerda"
         Height          =   375
         Left            =   600
         TabIndex        =   23
         Top             =   1120
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Direita"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1480
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   760
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Select"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   1120
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1480
         Width           =   735
      End
   End
   Begin VB.Image NesControl 
      Height          =   1440
      Left            =   705
      Picture         =   "frmConfig.frx":014A
      Top             =   10
      Width           =   3750
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdCancel_Click()
    If CPUPaused = True Then frmNES.mnuCPUPause_Click
    Unload frmConfig
End Sub
Private Sub CmdOK_Click()
    pad_ButA = Combo1.ListIndex: pad_ButB = Combo2.ListIndex: pad_ButSel = Combo3.ListIndex: pad_ButSta = Combo4.ListIndex
    nes_ButA = CByte(txtA.Text): nes_ButB = CByte(txtB.Text): nes_ButSel = CByte(txtSel.Text): nes_ButSta = CByte(txtSta.Text)
    nes_ButUp = CByte(txtUp.Text): nes_ButDn = CByte(txtDown.Text): nes_ButLt = CByte(txtLeft.Text): nes_ButRt = CByte(txtRight.Text)
    
    pad2_ButA = Combo5.ListIndex: pad2_ButB = Combo6.ListIndex: pad2_ButSel = Combo7.ListIndex: pad2_ButSta = Combo8.ListIndex
    nes2_ButA = CByte(txt2A.Text): nes2_ButB = CByte(txt2B.Text): nes2_ButSel = CByte(txt2Sel.Text): nes2_ButSta = CByte(txt2Sta.Text)
    nes2_ButUp = CByte(txt2Up.Text): nes2_ButDn = CByte(txt2Down.Text): nes2_ButLt = CByte(txt2Left.Text): nes2_ButRt = CByte(txt2Right.Text)
    
    CmdCancel_Click
End Sub
Private Sub Form_Load()
    Dim TmpStr As Long
    
    Dim NONE$
    Dim BTN$
    
    If Lang = 1 Then
        Caption = VERSION & ": Key Config"
        NONE$ = "None"
        BTN$ = "Button "
        Frame4.Caption = "Controller 1 (Joystick)"
        Frame2.Caption = "Controller 2 (Joystick)"
        Frame3.Caption = "Controller 1 (Keyboard)"
        Frame1.Caption = "Controller 2 (Keyboard)"
    Else
        Caption = VERSION & ": Configuração de teclas"
        NONE$ = "Nenhum"
        BTN$ = "Botão "
        Frame4.Caption = "Controle 1 (Joystick)"
        Frame2.Caption = "Controle 2 (Joystick)"
        Frame3.Caption = "Controle 1 (Teclado)"
        Frame1.Caption = "Controle 2 (Teclado)"
    End If
    
    ' This is why I hate Visual Basic
    
    Combo1.AddItem (NONE$)
    Combo2.AddItem (NONE$)
    Combo3.AddItem (NONE$)
    Combo4.AddItem (NONE$)
    Combo5.AddItem (NONE$)
    Combo6.AddItem (NONE$)
    Combo7.AddItem (NONE$)
    Combo8.AddItem (NONE$)
    
    For TmpStr = 1 To 12
        Combo1.AddItem (BTN$ & TmpStr)
        Combo2.AddItem (BTN$ & TmpStr)
        Combo3.AddItem (BTN$ & TmpStr)
        Combo4.AddItem (BTN$ & TmpStr)
        Combo5.AddItem (BTN$ & TmpStr)
        Combo6.AddItem (BTN$ & TmpStr)
        Combo7.AddItem (BTN$ & TmpStr)
        Combo8.AddItem (BTN$ & TmpStr)
    Next TmpStr
    
    Combo1.ListIndex = pad_ButA: Combo2.ListIndex = pad_ButB: Combo3.ListIndex = pad_ButSel: Combo4.ListIndex = pad_ButSta
    Combo5.ListIndex = pad2_ButA: Combo6.ListIndex = pad2_ButB: Combo7.ListIndex = pad2_ButSel: Combo8.ListIndex = pad2_ButSta
    If Gamepad1 > 0 Then
        Frame4.Visible = True
        Frame3.Visible = False
    Else
        Frame4.Visible = False
        Frame3.Visible = True
    End If
    If Gamepad2 > 0 Then
        Frame2.Visible = True
        Frame1.Visible = False
    Else
        Frame2.Visible = False
        Frame1.Visible = True
    End If
    
    ' WHY
    
    Combo1.ListIndex = pad_ButA: Combo2.ListIndex = pad_ButB: Combo3.ListIndex = pad_ButSel: Combo4.ListIndex = pad_ButSta
    txtUp.Text = nes_ButUp: txtDown.Text = nes_ButDn: txtLeft.Text = nes_ButLt: txtRight.Text = nes_ButRt
    txtA.Text = nes_ButA: txtB.Text = nes_ButB: txtSel.Text = nes_ButSel: txtSta.Text = nes_ButSta

    Combo5.ListIndex = pad2_ButA: Combo6.ListIndex = pad2_ButB: Combo7.ListIndex = pad2_ButSel: Combo8.ListIndex = pad2_ButSta
    txt2Up.Text = nes2_ButUp: txt2Down.Text = nes2_ButDn: txt2Left.Text = nes2_ButLt: txt2Right.Text = nes2_ButRt
    txt2A.Text = nes2_ButA: txt2B.Text = nes2_ButB: txt2Sel.Text = nes2_ButSel: txt2Sta.Text = nes2_ButSta

    Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CmdOK_Click
End Sub
Private Sub txtA_KeyDown(KeyCode As Integer, Shift As Integer) 'Control 1
    txtA.Text = KeyCode
End Sub
Private Sub txtA_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
    txtB.Text = KeyCode
End Sub
Private Sub txtB_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtDown_KeyDown(KeyCode As Integer, Shift As Integer)
    txtDown.Text = KeyCode
End Sub
Private Sub txtDown_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtLeft_KeyDown(KeyCode As Integer, Shift As Integer)
    txtLeft.Text = KeyCode
End Sub
Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtRight_KeyDown(KeyCode As Integer, Shift As Integer)
    txtRight.Text = KeyCode
End Sub
Private Sub txtRight_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtSel_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSel.Text = KeyCode
End Sub
Private Sub txtSel_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtSta_KeyDown(KeyCode As Integer, Shift As Integer)
    txtSta.Text = KeyCode
End Sub
Private Sub txtSta_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtUp_KeyDown(KeyCode As Integer, Shift As Integer)
    txtUp.Text = KeyCode
End Sub
Private Sub txtUp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2A_KeyDown(KeyCode As Integer, Shift As Integer)  'Control 2
    txt2A.Text = KeyCode
End Sub
Private Sub txt2A_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2B_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2B.Text = KeyCode
End Sub
Private Sub txt2B_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2Down_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2Down.Text = KeyCode
End Sub
Private Sub txt2Down_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2Left_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2Left.Text = KeyCode
End Sub
Private Sub txt2Left_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2Right_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2Right.Text = KeyCode
End Sub
Private Sub txt2Right_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2Sel_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2Sel.Text = KeyCode
End Sub
Private Sub txt2Sel_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2Sta_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2Sta.Text = KeyCode
End Sub
Private Sub txt2Sta_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txt2Up_KeyDown(KeyCode As Integer, Shift As Integer)
    txt2Up.Text = KeyCode
End Sub
Private Sub txt2Up_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
