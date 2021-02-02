VERSION 5.00
Begin VB.Form frmVPal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de paleta"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmVPal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdNewPal 
      Caption         =   "Nova Paleta"
      Height          =   375
      Left            =   3120
      TabIndex        =   108
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3120
      TabIndex        =   107
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   4320
   End
   Begin VB.Frame Frame3 
      Caption         =   "Paleta"
      Height          =   1095
      Left            =   240
      TabIndex        =   74
      Top             =   240
      Width           =   4335
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   31
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   106
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   30
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   105
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   29
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   104
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   28
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   103
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   27
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   102
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   26
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   101
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   25
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   100
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   24
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   99
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   23
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   98
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   22
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   97
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   21
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   96
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   20
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   95
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   19
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   94
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   18
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   93
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   17
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   92
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   16
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   91
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   15
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   90
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   14
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   89
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   88
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   87
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   11
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   86
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   10
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   85
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   84
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   83
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   82
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   81
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   80
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   79
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   78
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   77
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   76
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox spColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   75
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cor selecionada"
      Height          =   1335
      Left            =   240
      TabIndex        =   65
      Top             =   3120
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Aplicar"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   72
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtB 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   71
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtG 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   70
         Top             =   600
         Width           =   735
      End
      Begin VB.PictureBox cView 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   855
         TabIndex        =   66
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
         Height          =   195
         Left            =   1320
         TabIndex        =   69
         Top             =   1000
         Width           =   150
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G:"
         Height          =   195
         Left            =   1320
         TabIndex        =   68
         Top             =   640
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
         Height          =   195
         Left            =   1320
         TabIndex        =   67
         Top             =   280
         Width           =   165
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cores da paleta"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   64
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   63
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   62
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   61
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   60
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   59
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   58
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   57
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   56
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   55
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   10
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   54
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   11
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   53
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   52
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   51
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   14
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   50
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   15
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   49
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   16
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   48
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   17
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   47
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   18
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   46
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   19
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   45
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   20
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   44
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   21
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   43
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   22
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   23
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   24
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   40
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   25
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   26
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   27
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   37
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   28
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   36
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   29
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   35
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   30
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   31
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   32
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   33
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   31
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   34
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   35
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   36
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   28
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   37
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   27
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   38
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   39
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   25
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   40
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   41
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   42
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   43
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   44
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   45
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   46
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   47
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   48
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   49
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   50
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   51
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   52
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   53
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   54
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   55
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   56
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   57
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   58
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   59
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   5
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   60
         Left            =   3120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   61
         Left            =   3360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   62
         Left            =   3600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox pColor 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   63
         Left            =   3840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   1
         Top             =   1080
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmVPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mRgb(2) As Byte
Dim RgbIndex As Long
Dim n As Long
Private Sub CmdNewPal_Click()
    frmMakePal.Show
End Sub
Private Sub CmdOK_Click()
    Unload frmVPal
End Sub
Private Sub Command2_Click()
    On Error GoTo ErrH
    cView.BackColor = RGB(txtR.Text, txtG.Text, txtB.Text)
    ModRgb txtR.Text, txtG.Text, txtB.Text, RgbIndex
ErrH:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CmdOK_Click
End Sub
Private Sub Form_Load()
    For n = 0 To 63
        MemCopy mRgb(0), pal(n), Len(pal(n))
        pColor(n).BackColor = RGB(mRgb(2), mRgb(1), mRgb(0))
    Next n
    DrawSPColors
End Sub
Sub DrawSPColors()
    For n = 0 To 31
        MemCopy mRgb(0), pal(VRAM(n + &H3F00)), Len(pal(VRAM(n + &H3F00)))
        spColor(n).BackColor = RGB(mRgb(2), mRgb(1), mRgb(0))
    Next n
End Sub
Private Sub pColor_Click(Index As Integer)
    RgbIndex = Index
    cView.BackColor = pColor(Index).BackColor
    GetRGB pColor(Index).BackColor
End Sub
Public Function GetRGB(ByVal lngColor As Long)
    MemCopy mRgb(0), lngColor, Len(lngColor)
    txtR.Text = mRgb(0)
    txtG.Text = mRgb(1)
    txtB.Text = mRgb(2)
End Function
Public Function ModRgb(R As Long, G As Long, B As Long, n As Long)
    SetPalVal R, G, B, n
End Function
Private Sub Timer1_Timer()
    DrawSPColors
End Sub
Private Sub txtR_Change()
    ChangeColor
End Sub
Private Sub txtG_Change()
    ChangeColor
End Sub
Private Sub txtB_Change()
    ChangeColor
End Sub
Sub ChangeColor()
    On Error GoTo ErrH
    cView.BackColor = RGB(txtR.Text, txtG.Text, txtB.Text)
ErrH:
End Sub
