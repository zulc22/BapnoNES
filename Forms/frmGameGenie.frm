VERSION 5.00
Begin VB.Form frmGameGenie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Genie"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmGameGenie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de códigos"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.CommandButton Command6 
         Caption         =   "&Decodificar"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Apagar último"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Apagar primeiro"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "A&pagar tudo"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Adicionar"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox lstGg 
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
         ItemData        =   "frmGameGenie.frx":014A
         Left            =   240
         List            =   "frmGameGenie.frx":014C
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmGameGenie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOK_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    Dim TempValue As String
    Select Case Lang
        Case 1
            TempValue = InputBox("Type Game Genie code (6 characters):", "Add code")
        Case 0
            TempValue = InputBox("Digite o código Game Genie (6 caracteres):", "Adicionar código")
    End Select
    If Len(TempValue) = 6 Then ggCodes = ggCodes & UCase(TempValue): lstGg.AddItem (UCase(TempValue))
End Sub
Private Sub Command2_Click()
    ggCodes = vbNullString
    lstGg.Clear
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub
Private Sub Command4_Click()
    lstGg.RemoveItem (0)
    ggCodes = Right(ggCodes, Len(ggCodes) - 6)
End Sub
Private Sub Command5_Click()
    lstGg.RemoveItem (lstGg.ListCount - 1)
    ggCodes = Left(ggCodes, Len(ggCodes) - 6)
End Sub
Private Sub Command6_Click()
    On Error GoTo ErrH
    Dim StrMsg As String
    Dim lAdrress As Long
    Dim lValue As Long
    lAddress = GetNES6Address(lstGg.Text)
    lValue = GetNES6Value(lstGg.Text)
    If Lang = 1 Then
        StrMsg = "Address: " & CStr(lAddress) & " (0x" & Hex(lAddress) & ")" & vbCrLf
        StrMsg = StrMsg & "Value: " & CStr(lValue) & " (0x" & Hex(lValue) & ")"
        MsgBox StrMsg, vbInformation, "Results"
    Else
        StrMsg = "Endereço: " & CStr(lAddress) & " (0x" & Hex(lAddress) & ")" & vbCrLf
        StrMsg = StrMsg & "Valor: " & CStr(lValue) & " (0x" & Hex(lValue) & ")"
        MsgBox StrMsg, vbInformation, "Código Game Genie decodificado"
    End If
ErrH:
End Sub
Private Function GetNESValueByLetter(szChar As String) As String
    ' **********************************
    ' *   Thank you to Maul for this   *
    ' **********************************
    
    ' GetNESValueByLetter - Returns binary representation of a code letter
    Select Case UCase$(szChar$)
        Case "A": GetNESValueByLetter = "0000" 'Set #1
        Case "P": GetNESValueByLetter = "0001"
        Case "Z": GetNESValueByLetter = "0010"
        Case "L": GetNESValueByLetter = "0011"
        Case "G": GetNESValueByLetter = "0100"
        Case "I": GetNESValueByLetter = "0101"
        Case "T": GetNESValueByLetter = "0110"
        Case "Y": GetNESValueByLetter = "0111"
        Case "E": GetNESValueByLetter = "1000" 'Set #2
        Case "O": GetNESValueByLetter = "1001"
        Case "X": GetNESValueByLetter = "1010"
        Case "U": GetNESValueByLetter = "1011"
        Case "K": GetNESValueByLetter = "1100"
        Case "S": GetNESValueByLetter = "1101"
        Case "V": GetNESValueByLetter = "1110"
        Case "N": GetNESValueByLetter = "1111"
    End Select
End Function
Public Function GetNES6Value(szCode As String) As Integer
    ' **********************************
    ' *   Thank you to Maul for this   *
    ' **********************************
    
    ' GetNES6Value - Returns the value of a 6 letter code
    Dim szString As String, szValue As String
    Dim nLoop As Integer, vPos() As Variant
    ' Convert code to binary
    For nLoop = 1 To Len(szCode$)
        szString = szString & GetNESValueByLetter(Mid(szCode, nLoop, 1))
    Next nLoop
    ' Unscramble value
    vPos = Array(1, 6, 7, 8, 21, 2, 3, 4)
    For nLoop = LBound(vPos) To UBound(vPos)
        szValue = szValue & Mid(szString, vPos(nLoop), 1)
    Next nLoop
    GetNES6Value = CInt(BinToLong(szValue))
End Function
Public Function GetNES6Address(szCode As String) As Long
    ' **********************************
    ' *   Thank you to Maul for this   *
    ' **********************************
    
    ' GetNES6Address - Returns the address of a 6 letter code
    Dim szString As String, szAddress As String
    Dim nLoop As Integer, vPos() As Variant
    ' Convert code to binary
    For nLoop = 1 To Len(szCode$)
        szString = szString & GetNESValueByLetter(Mid(szCode, nLoop, 1))
    Next nLoop
    ' Unscramble address
    vPos = Array(14, 15, 16, 17, 22, 23, 24, 5, 10, 11, 12, 13, 18, 19, 20)
    For nLoop = LBound(vPos) To UBound(vPos)
        szAddress = szAddress & Mid(szString, vPos(nLoop), 1)
    Next nLoop
    GetNES6Address& = CLng(BinToLong(szAddress))
End Function
Private Function BinToLong(szString As String) As Long
    ' **********************************
    ' ****Thank you to Maul for this****
    ' **********************************
    
    ' BinToLong - Returns a long integer from a binary string
    Dim lBinary As Long, nLoop As Integer, nLen As Integer
    nLen = Len(szString)
    For nLoop = nLen To 1 Step -1
        If Mid(szString, nLoop, 1) = "1" Then
            lBinary = lBinary + (2 ^ (nLen - nLoop))
        End If
    Next nLoop
    BinToLong = lBinary
End Function
Private Sub Form_Load()
    Dim i As Long
    For i = 1 To Len(ggCodes)
        lstGg.AddItem (Mid(ggCodes, i, 6))
        i = i + 5
    Next i
    
    Select Case Lang
        Case 1
            Command1.Caption = "&Add"
            Command6.Caption = "&Decode"
            Command4.Caption = "Erase first"
            Command5.Caption = "Erase last"
            Command2.Caption = "&Clear"
            Frame2.Caption = "Code list"
        Case 0
            Command1.Caption = "&Adicionar"
            Command6.Caption = "&Decodificar"
            Command4.Caption = "Apagar primeiro"
            Command5.Caption = "A&pagar tudo"
            Command2.Caption = "&Clear"
            Frame2.Caption = "Code list"
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CPUPaused = True Then frmNES.mnuCPUPause_Click
    Unload frmGameGenie
End Sub
