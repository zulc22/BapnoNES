VERSION 5.00
Begin VB.Form frmMakePal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criar paleta"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   Icon            =   "frmMakePal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3000
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Saturação"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
      Begin VB.HScrollBar ScrollCon 
         Height          =   255
         Left            =   240
         Max             =   10
         TabIndex        =   4
         Top             =   360
         Value           =   5
         Width           =   3735
      End
      Begin VB.Label lblCon 
         Alignment       =   2  'Center
         Caption         =   "0.5"
         Height          =   195
         Left            =   4080
         TabIndex        =   5
         Top             =   400
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Matiz"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.HScrollBar ScrollHue 
         Height          =   255
         Left            =   240
         Max             =   400
         TabIndex        =   1
         Top             =   360
         Value           =   332
         Width           =   3735
      End
      Begin VB.Label lblHue 
         Alignment       =   2  'Center
         Caption         =   "332"
         Height          =   195
         Left            =   4080
         TabIndex        =   2
         Top             =   400
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmMakePal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Sub CmdOK_Click()
    NewPal ScrollCon.Value / 10, ScrollHue.Value
    Unload Me
End Sub
Private Sub ScrollCon_Change()
    lblCon.Caption = ScrollCon.Value / 10
End Sub
Private Sub ScrollHue_Change()
    lblHue.Caption = ScrollHue.Value
End Sub
