VERSION 5.00
Begin VB.Form frmPattern 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pattern Tables"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   Icon            =   "frmPattern.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aplicar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox PicCols 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   5760
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox PicChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   3840
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "frmPattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mRgb(2) As Byte
Dim DrawPal(3) As Long
Dim OldCurVBank As Long

Dim SelX, SelY, SelCol As Integer
Private Sub CmdClose_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    ChangeVROM
End Sub
Private Sub Form_Load()
    PatternTables
End Sub
Sub PatternTables()
    'On Error Resume Next
    'Draw Pattern Tables
    Dim LoopCount As Long
    Dim Adc As Long
    Dim VBytes As String
    Dim Tmp(7) As Long
    Dim a As Long
    
    If ChrCount = 0 Then Exit Sub
    
    'Get the first 4 colors of selected pallete colors
    For i = 0 To 3
        MemCopy mRgb(0), pal(VRAM(i + &H3F00)), Len(pal(VRAM(i + &H3F00)))
        DrawPal(i) = RGB(mRgb(2), mRgb(1), mRgb(0))
    Next i
    
    Adc = CurVBank
    Adc = Adc * 8192
    OldCurVBank = Adc
    Do 'Load only necessary VROM values to save time
        LoopCount = LoopCount + 1
        If ChrCount Then
            VBytes = VBytes & Chr(VROM(Adc + LoopCount))
        Else
            VBytes = VBytes & Chr(GameImage(LoopCount))
        End If
        DoEvents
    Loop Until LoopCount >= 8192
    VBytes = " " & VBytes
    a = 1
    For Table = 0 To 1
        For TY = 0 To 120 Step 8
            For TX = 0 To 120 Step 8
               For Y = 0 To 7
                    For planes = 0 To 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 128) = 128 Then Tmp(0) = Tmp(0) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 64) = 64 Then Tmp(1) = Tmp(1) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 32) = 32 Then Tmp(2) = Tmp(2) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 16) = 16 Then Tmp(3) = Tmp(3) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 8) = 8 Then Tmp(4) = Tmp(4) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 4) = 4 Then Tmp(5) = Tmp(5) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 2) = 2 Then Tmp(6) = Tmp(6) + planes + 1
                        If (Asc(Mid$(VBytes, a + planes * 8, 1)) And 1) = 1 Then Tmp(7) = Tmp(7) + planes + 1
                    Next planes
                    For bl = 0 To 7
                        PSet (Table * 128 + TX + bl, TY + Y), DrawPal(Tmp(bl))
                        Tmp(bl) = 0
                    Next bl
                    a = a + 1
               Next Y
               a = a + 8
            Next TX
        Next TY
    Next Table
    
    For i = 0 To 3
        PicCols.Line (0, i * 32)-(32, (i * 32) + 32), DrawPal(i), BF
    Next i
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then PatternTables
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelX = X \ 8
    SelY = Y \ 8
    PicChar.PaintPicture Me.Image, 0, 0, 128, 128, SelX * 8, SelY * 8, 8, 8
End Sub
Private Sub PicChar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicChar_MouseDown Button, Shift, X, Y
End Sub
Private Sub PicChar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PicChar.Line ((X \ 16) * 16, (Y \ 16) * 16)-(((X \ 16) * 16) + 15, ((Y \ 16) * 16) + 15), DrawPal(SelCol), BF
        PaintPicture PicChar.Image, SelX * 8, SelY * 8, 8, 8, 0, 0, 128, 128
    End If
End Sub
Private Sub PicCols_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelCol = Y \ 32
End Sub
Private Sub ChangeVROM()
    'Need optimizations, but it works! :)
    Dim Table, TX, TY, SprX, SprY As Integer
    Dim BinVal1, BinVal2 As String
    Dim a As Integer
    
    For Table = 0 To 1
        For TY = 0 To 120 Step 8
            For TX = 0 To 120 Step 8
                For SprY = 0 To 7
                    BinVal1 = vbNullString
                    BinVal2 = vbNullString
                    For SprX = 0 To 7
                        If Me.Point((Table * 128) + TX + SprX, TY + SprY) = DrawPal(0) Then
                            BinVal1 = BinVal1 & "0"
                            BinVal2 = BinVal2 & "0"
                        End If
                        If Me.Point((Table * 128) + TX + SprX, TY + SprY) = DrawPal(1) Then
                            BinVal1 = BinVal1 & "1"
                            BinVal2 = BinVal2 & "0"
                        End If
                        If Me.Point((Table * 128) + TX + SprX, TY + SprY) = DrawPal(2) Then
                            BinVal1 = BinVal1 & "0"
                            BinVal2 = BinVal2 & "1"
                        End If
                        If Me.Point((Table * 128) + TX + SprX, TY + SprY) = DrawPal(3) Then
                            BinVal1 = BinVal1 & "1"
                            BinVal2 = BinVal2 & "1"
                        End If
                    Next SprX
                    VROM(OldCurVBank + a) = BinaryToDecimal(BinVal1)
                    VROM(OldCurVBank + a + 8) = BinaryToDecimal(BinVal2)
                    VRAM(a) = BinaryToDecimal(BinVal1) 'Immediate change!
                    VRAM(a + 8) = BinaryToDecimal(BinVal2)
                    a = a + 1
                Next SprY
                a = a + 8
            Next TX
        Next TY
    Next Table
End Sub
Public Function BinaryToDecimal(ByVal Binary As String) As Long
    Dim n As Long
    Dim s As Integer

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ _
            (s - 1)))
    Next s

    BinaryToDecimal = n
End Function
