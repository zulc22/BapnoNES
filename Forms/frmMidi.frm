VERSION 5.00
Begin VB.Form frmMidi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instrumentos MIDI"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   Icon            =   "frmMidi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Instrumentos MIDI (0 - 127)"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Text            =   "Combo4"
         Top             =   1440
         Width           =   4455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Text            =   "Combo3"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   720
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label8 
         Caption         =   "Noise"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Triangle"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Square 2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   765
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Square 1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   405
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMidi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    Unload Me
End Sub
Private Sub Combo1_Change()
    SelectInstrument 0, Combo1.ListIndex
End Sub
Private Sub Combo2_Change()
    SelectInstrument 1, Combo2.ListIndex
End Sub
Private Sub Combo3_Change()
    SelectInstrument 2, Combo3.ListIndex
End Sub
Private Sub Combo4_Change()
    SelectInstrument 3, Combo4.ListIndex
End Sub
Private Sub Form_Load()
    Show
    AddBoxes
    Combo1.ListIndex = Instrumental(0)
    Combo2.ListIndex = Instrumental(1)
    Combo3.ListIndex = Instrumental(2)
    Combo4.ListIndex = Instrumental(3)
End Sub
Sub AddBoxes()
    Dim AddMsg As String
    For i = 1 To 128
        If i = 1 Then AddMsg = "1 Piano"
        If i = 2 Then AddMsg = "2 Piano"
        If i = 3 Then AddMsg = "3 Piano eléctro-acústico"
        If i = 4 Then AddMsg = "4 Piano de cabaré"
        If i = 5 Then AddMsg = "5 Piano eléctrico (tipo Fender Rhodes)"
        If i = 6 Then AddMsg = "6 Piano eléctrico (sintético tipo DX7)"
        If i = 7 Then AddMsg = "7 Cravo"
        If i = 8 Then AddMsg = "8 Clavicórdio"
        If i = 9 Then AddMsg = "9 Celesta"
        If i = 10 Then AddMsg = "10 Glockenspiel"
        If i = 11 Then AddMsg = "11 Caixa de música"
        If i = 12 Then AddMsg = "12 Vibrafone"
        If i = 13 Then AddMsg = "13 Marimba"
        If i = 14 Then AddMsg = "14 Xilofone"
        If i = 15 Then AddMsg = "15 Carrilhão de orquestra"
        If i = 16 Then AddMsg = "16 Santur"
        If i = 17 Then AddMsg = "17 Órgão Hammond"
        If i = 18 Then AddMsg = "18 Órgão percussivo"
        If i = 19 Then AddMsg = "19 Órgão de rock"
        If i = 20 Then AddMsg = "20 Órgão de tubos"
        If i = 21 Then AddMsg = "21 Harmónio"
        If i = 22 Then AddMsg = "22 Acordeão"
        If i = 23 Then AddMsg = "23 Harmónica"
        If i = 24 Then AddMsg = "24 Bandoneón"
        If i = 25 Then AddMsg = "25 Guitarra de cordas de nylon"
        If i = 26 Then AddMsg = "26 Guitarra de cordas de aço"
        If i = 27 Then AddMsg = "27 Guitarra semi-acústica"
        If i = 28 Then AddMsg = "28 Guitarra elétrica"
        If i = 29 Then AddMsg = "29 Guitarra abafada"
        If i = 30 Then AddMsg = "30 Guitarra elétrica com saturação"
        If i = 31 Then AddMsg = "31 Guitarra elétrica com distorção"
        If i = 32 Then AddMsg = "32 Harmónicos"
        If i = 33 Then AddMsg = "33 Contrabaixo (dedilhado)"
        If i = 34 Then AddMsg = "34 Baixo elétrico dedilhado"
        If i = 35 Then AddMsg = "35 Baixo elétrico beliscado com palheta"
        If i = 36 Then AddMsg = "36 Baixo elétrico sem trastos"
        If i = 37 Then AddMsg = "37 Baixo elétrico percutido 1 (pop)"
        If i = 38 Then AddMsg = "38 Baixo elétrico percutido 2 (thump)"
        If i = 39 Then AddMsg = "39 Baixo sintético 1 (analógico)"
        If i = 40 Then AddMsg = "40 Baixo sintético 2 (digital)"
        If i = 41 Then AddMsg = "41 Violino"
        If i = 42 Then AddMsg = "42 Viola"
        If i = 43 Then AddMsg = "43 Violoncelo"
        If i = 44 Then AddMsg = "44 Contrabaixo"
        If i = 45 Then AddMsg = "45 Cordas em trêmulo"
        If i = 46 Then AddMsg = "46 Cordas em pizzicatto"
        If i = 47 Then AddMsg = "47 Harpa"
        If i = 48 Then AddMsg = "48 Tímpanos"
        If i = 49 Then AddMsg = "49 Orquestra de cordas 1"
        If i = 50 Then AddMsg = "50 Orquestra de cordas 2 (ataque lento)"
        If i = 51 Then AddMsg = "51 Cordas sintéticas 1"
        If i = 52 Then AddMsg = "52 Cordas sintéticas 2 (filtro ressonante)"
        If i = 53 Then AddMsg = "53 Coro"
        If i = 54 Then AddMsg = "54 Voz humana (solista)"
        If i = 55 Then AddMsg = "55 Voz humana (sintética)"
        If i = 56 Then AddMsg = "56 Batida orquestral"
        If i = 57 Then AddMsg = "57 Trompete"
        If i = 58 Then AddMsg = "58 Trombone"
        If i = 59 Then AddMsg = "59 Tuba"
        If i = 60 Then AddMsg = "60 Trompete com surdina"
        If i = 61 Then AddMsg = "61 Trompa"
        If i = 62 Then AddMsg = "62 Metais"
        If i = 63 Then AddMsg = "63 Metais sintéticos 1 (imitação de secção de trompetes e trombones)"
        If i = 64 Then AddMsg = "64 Metais sintéticos 2 (imitação de secção de trompas)"
        If i = 65 Then AddMsg = "65 Saxofone soprano"
        If i = 66 Then AddMsg = "66 Saxofone alto"
        If i = 67 Then AddMsg = "67 Saxofone tenor"
        If i = 68 Then AddMsg = "68 Saxofone barítono"
        If i = 69 Then AddMsg = "69 Oboé"
        If i = 70 Then AddMsg = "70 Corne inglês"
        If i = 71 Then AddMsg = "71 Fagote"
        If i = 72 Then AddMsg = "72 Clarinete"
        If i = 73 Then AddMsg = "73 Flautim"
        If i = 74 Then AddMsg = "74 Flauta transversal"
        If i = 75 Then AddMsg = "75 Flauta de bisel"
        If i = 76 Then AddMsg = "76 Flauta de Pã"
        If i = 77 Then AddMsg = "77 Sopro em gargalo de garrafa"
        If i = 78 Then AddMsg = "78 Shakuhachi"
        If i = 79 Then AddMsg = "79 Assobio"
        If i = 80 Then AddMsg = "80 Ocarina"
        If i = 81 Then AddMsg = "81 Onda quadrada"
        If i = 82 Then AddMsg = "82 Onda dente de serra"
        If i = 83 Then AddMsg = "83 Calliope (Órgão a vapor sintético)"
        If i = 84 Then AddMsg = "84 Chiff Lead"
        If i = 85 Then AddMsg = "85 Charango sintético"
        If i = 86 Then AddMsg = "86 Solo vox"
        If i = 87 Then AddMsg = "87 Onda dente de serra em quintas paralelas"
        If i = 88 Then AddMsg = "88 Baixo e solo"
        If i = 89 Then AddMsg = "89 Fundo New Age"
        If i = 90 Then AddMsg = "90 Fundo morno"
        If i = 91 Then AddMsg = "91 Polysynth"
        If i = 92 Then AddMsg = "92 Space voice"
        If i = 93 Then AddMsg = "93 Vidro friccionado"
        If i = 94 Then AddMsg = "94 Fundo metálico"
        If i = 95 Then AddMsg = "95 Fundo halo"
        If i = 96 Then AddMsg = "96 Fundo com abertura do filtro"
        If i = 97 Then AddMsg = "97 Chuva de gelo"
        If i = 98 Then AddMsg = "98 Trilha sonora"
        If i = 99 Then AddMsg = "99 Cristal"
        If i = 100 Then AddMsg = "100 Atmosfera"
        If i = 101 Then AddMsg = "101 Brilhos"
        If i = 102 Then AddMsg = "102 Goblins"
        If i = 103 Then AddMsg = "103 Ecos"
        If i = 104 Then AddMsg = "104 Ficção científica"
        If i = 105 Then AddMsg = "105 Sitar"
        If i = 106 Then AddMsg = "106 Banjo"
        If i = 107 Then AddMsg = "107 Shamisen"
        If i = 108 Then AddMsg = "108 Taishikoto"
        If i = 109 Then AddMsg = "109 Kalimba"
        If i = 110 Then AddMsg = "110 Gaita de foles"
        If i = 111 Then AddMsg = "111 Rabeca"
        If i = 112 Then AddMsg = "112 Shehnai"
        If i = 113 Then AddMsg = "113 Sino"
        If i = 114 Then AddMsg = "114 Agogô"
        If i = 115 Then AddMsg = "115 Tambor de aço"
        If i = 116 Then AddMsg = "116 Bloco de madeira"
        If i = 117 Then AddMsg = "117 Taiko"
        If i = 118 Then AddMsg = "118 Timbalões acústicos"
        If i = 119 Then AddMsg = "119 Timbalões sintéticos"
        If i = 120 Then AddMsg = "120 Prato revertido"
        If i = 121 Then AddMsg = "121 Corda de violão riscada"
        If i = 122 Then AddMsg = "122 Respiração"
        If i = 123 Then AddMsg = "123 Ondas do mar"
        If i = 124 Then AddMsg = "124 Pássaro piando"
        If i = 125 Then AddMsg = "125 Telefone tocando"
        If i = 126 Then AddMsg = "126 Helicóptero"
        If i = 127 Then AddMsg = "127 Aplausos"
        If i = 128 Then AddMsg = "128 Tiro (arma de fogo)"
        Combo1.AddItem AddMsg
        Combo2.AddItem AddMsg
        Combo3.AddItem AddMsg
        Combo4.AddItem AddMsg
    Next i
End Sub
