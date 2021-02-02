VERSION 5.00
Begin VB.Form frmROMInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Info. da ROM"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmROMInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblList 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmROMInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim MType As String
    
    Dim mYes As String
    Dim mNo As String
    
    If Lang = 1 Then
        mYes = "Yes"
        mNo = "No"
    Else
        mYes = "Sim"
        mNo = "Não"
    End If
    
    lblList.Caption = vbNullString

    lblList.Caption = lblList.Caption & "PRG-ROM:      " & PrgCount * 16 & "k" & " (" & PrgCount & QFix(PrgCount) & ")" & vbCrLf
    lblList.Caption = lblList.Caption & "CHR-ROM:      " & ChrCount * 8 & "k" & " (" & ChrCount & QFix(ChrCount) & ")" & vbCrLf

    lblList.Caption = lblList.Caption & "Mapper:       " & Mapper & " (" & MapperName(Mapper) & ")" & vbCrLf

    If Mirroring = 0 Then MType = "Horizontal"
    If Mirroring = 1 Then MType = "Vertical"
    If Mirroring = 2 Then MType = "One screen"
    If Mirroring = 4 Then MType = "Four screen"
    lblList.Caption = lblList.Caption & "Mirroring:    " & MType & vbCrLf

    lblList.Caption = lblList.Caption & "Trainer:      " & IIf(Trainer, mYes, mNo) & vbCrLf

    lblList.Caption = lblList.Caption & "Battery:      " & IIf(Batt, mYes, mNo) & vbCrLf
    Show
End Sub
Public Function MapperName(ByVal MapperNum As Long) As String
    Dim StrMapperName As String
    
    'Mapper Names
    Select Case MapperNum
        Case 0
            If Lang = 1 Then
                StrMapperName = "None"
            Else
                StrMapperName = "Nenhum"
            End If
        Case 1: StrMapperName = "Nitendo MMC1"
        Case 2: StrMapperName = "UNROM"
        Case 3: StrMapperName = "CNROM"
        Case 4: StrMapperName = "Nitendo MMC3"
        Case 5: StrMapperName = "Nitendo MMC5"
        Case 6: StrMapperName = "FFE F4xxx"
        Case 7: StrMapperName = "AOROM"
        Case 8: StrMapperName = "FFE F3xxx"
        Case 9: StrMapperName = "Nitendo MMC2"
        Case 10: StrMapperName = "Nitendo MMC4"
        Case 11: StrMapperName = "Colour Dreams"
        Case 15: StrMapperName = "Waixing"
        Case 16: StrMapperName = "Bandai"
        Case 17: StrMapperName = "FFE F8xxx"
        Case 18: StrMapperName = "Jaleco SS8806"
        Case 19: StrMapperName = "Namcot 106"
        Case 21: StrMapperName = "Konami VRC4"
        Case 22: StrMapperName = "Konami VRC2 A"
        Case 23: StrMapperName = "Konami VRC2 B"
        Case 24: StrMapperName = "Konami VRC6 A"
        Case 26: StrMapperName = "Konami VRC6 B"
        Case 32: StrMapperName = "Irem G101"
        Case 33: StrMapperName = "Taito TC0190"
        Case 34: StrMapperName = "Nina-1"
        Case 51: StrMapperName = "11-in-1 Ball Games"
        Case 52: StrMapperName = "Mario 7-in-1"
        Case 58: StrMapperName = "68-in-1 (Game Star)"
        Case 61: StrMapperName = "20-in-1"
        Case 64: StrMapperName = "Rambo-1"
        Case 65: StrMapperName = "Irem H3001"
        Case 66: StrMapperName = "GNROM"
        Case 68: StrMapperName = "Sunsoft 4"
        Case 69: StrMapperName = "Sunsoft 5"
        Case 71: StrMapperName = "Camerica"
        Case 73: StrMapperName = "Konami VRC3"
        Case 78: StrMapperName = "Irem 74HC161/32"
        Case 79: StrMapperName = "AVE"
        Case 83: StrMapperName = "Cony"
        Case 90: StrMapperName = "PCJY"
        Case 91: StrMapperName = "HK-SF3"
        Case 95: StrMapperName = "Namco 1xx"
        Case 99, 151: StrMapperName = "VS-Unisystem"
        Case 105: StrMapperName = "NWC 1990"
        Case 119: StrMapperName = "TQROM"
        Case 174: StrMapperName = "NTDec 5-in-1"
        Case 182: StrMapperName = "SDK/LK Pirate"
        Case 200: StrMapperName = "1200-in-1"
        Case 201: StrMapperName = "21-in-1"
        Case 211: StrMapperName = "JY Company"
        Case 212: StrMapperName = "Unchained Melody"
        Case 227: StrMapperName = "1200-in-1"
        Case 228: StrMapperName = "Active Enterprise"
        Case 231: StrMapperName = "20-in-1"
        Case 232: StrMapperName = "Quattro Games"
        Case 242: StrMapperName = "Wai Xing Zhan Shi"
        Case 250: StrMapperName = "Time Diver Avenger"
        Case 255: StrMapperName = "110-in-1"
    End Select
    If StrMapperName = vbNullString Then MapperName = "???" Else MapperName = StrMapperName
End Function
Public Function QFix(ByVal fCount As Long) As String
    If Lang = 1 Then
        If fCount = 1 Then QFix = " bank" Else QFix = " banks"
    Else
        If fCount = 1 Then QFix = " banco" Else QFix = " bancos"
    End If
End Function
