VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About BapnoNES"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":014A
   ScaleHeight     =   1545
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image ImgIcon 
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":01B0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(STUFF)"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(VERSION)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    If CPUPaused = True Then frmNES.mnuCPUPause_Click
    Unload frmAbout
End Sub
Private Sub Form_Load()
    If Lang = 1 Then
        Caption = "About BapnoNES"
        Label2.Caption = "BapnoNES is a fork of YoshiNES maintained by zulc22."
    Else
        Caption = "Sobre o BapnoNES"
        Label2.Caption = "texto de espaço reservado"
    End If
    Label1.Caption = VERSION
    Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CmdOK_Click
End Sub
