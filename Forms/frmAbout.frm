VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o YoshiNES"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   4215
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1320
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtAbout 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1320
      Width           =   4020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Codename: Moltres"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Image ImgIcon 
      Height          =   720
      Left            =   240
      Picture         =   "frmAbout.frx":05F0
      Top             =   195
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":2132
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VERSÃO"
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
      Top             =   240
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
    Dim Message As String
    Message = " • Alguns mappers são cortesia de DarcNES e Pretendo." & vbCrLf & vbCrLf & " • Obrigado ao nyef e o Delta pela informação." & vbCrLf & vbCrLf
    Message = Message & " • Código do Open File Dialog cortesia de Matthew Leverton." & vbCrLf & vbCrLf
    Message = Message & " • Renderização baseada em linhas de Lothos." & vbCrLf & vbCrLf
    Message = Message & " • Idéia do código de paleta por FCE." & vbCrLf & vbCrLf & " • Obrigado ao The Quietust pela informação de mappers e ajuda." & vbCrLf & vbCrLf
    Message = Message & " • Ao Norix, é de seu emulador que adapto código de vários Mappers." & vbCrLf & vbCrLf
    Message = Message & " • A Chris Cowley pelas idéias de interface e muita ajuda com vários emuladores." & vbCrLf & vbCrLf
    Message = Message & " • YoshiNES é Copyright © 2011 Gabriel Dark 100 (a.k.a. Gabriel King)."
    Label1.Caption = VERSION
    txtAbout.Text = Message
    Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CmdOK_Click
End Sub
