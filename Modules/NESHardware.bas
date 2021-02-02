Attribute VB_Name = "NESHardware"
'+====================================================+
'| YoshiNES made by Gabriel Dark.                     |
'| Original code by Don Jarrett, 2000.                |
'| Big portions of graphics code done by David Finch. |
'+====================================================+

Option Explicit ' Option Explicit is important for avoiding mysterious, crippling bugs

DefLng A-Z 'this ensures that no variants are used unless we choose so

Public VERSION As String

Private Js As New jStick

Public TmpLatch As Byte, ppuLatch As Byte

Public Interlace As Long

Public MirrorTypes(4) As String

'DF: array to draw each frame to
Public vBuffer(256& * 241& - 1) As Byte '256*241 to allow for some overflow

Public nt(0 To 3, &H0 To &H3FF) As Byte
Public Mirror(3) As Byte

Public vBuffer16(256& * 240& - 1) As Integer
Public vBuffer32(256& * 240& - 1) As Long

Public oldvBuffer16(256& * 240& - 1) As Integer 'used for scaling modes that take advantage of unchanged pixels

Public vBuffer2x16(512& * 480& - 1) As Integer
Public vBuffer2x32(512& * 480& - 1) As Long

Public tLook(65536 * 8 - 1) As Byte

Public SpritesChanged As Boolean

Public Record As Boolean
Public Playing As Boolean

Public map17_irqon As Boolean
Public map17_irq As Long

Public IRQCounter As Long ' for mapper 6
Public map6_irqon As Byte
Public map225_psize As Byte
Public map225_psel As Byte
Public Train(&H1FF) As Byte
Public latch13 As Byte
Public MMC19_IRQCount As Long
Public MIRQOn As Byte

'DF: powers of 2
Public Pow2(31) As Long

' NES Hardware defines
Public PPU_Control1 As Byte ' $2000
Public PPU_Control2 As Byte ' $2001
Public PPU_Status As Byte ' $2002
Public SpriteAddress As Long ' $2003
Public PPUAddressHi As Long ' $2006, 1st write
Public PPUAddress As Long ' $2006
Public PPU_AddressIsHi As Boolean
Public VRAM(&H3FFF) As Byte, VROM() As Byte  ' Video RAM
Public SpriteRAM(&HFF) As Byte

Public Sound(0 To &H15) As Byte
Public SoundCtrl As Byte

Public VScroll2 As Long

Public PrgCount As Byte, PrgCount2 As Long, ChrCount As Byte, ChrCount2 As Long

Public Cmd As Byte, Prg As Byte, Chr1 As Byte

Public reg8 As Byte
Public regA As Byte
Public regC As Byte
Public regE As Byte

Public NESPal(&HF) As Byte

Public CPal() As Long

Public FrameSkip As Long 'Integer
Public Frames As Long

Public ScrollToggle As Byte
Public HScroll As Byte, VScroll As Long 'Integer ' $2005

Public map15_BankAddr As Byte
Public map15_SwapReg As Byte

'DF: these variables were undefined and therefore local:
Public swap As Boolean
Public map15_swapaddr As Long

' MMC3[Mapper #4] infos
Public MMC3_Command As Byte
Public MMC3_PrgAddr As Byte
Public MMC3_ChrAddr As Integer
Public MMC3_IrqVal As Byte
Public MMC3_TmpVal As Byte
Public MMC3_IrqOn As Boolean

Public PatternTable As Long
Public NameTable As Long

Public bank_regs(16) As Byte

Public Const PPU_InVBlank = &H80
Public Const PPU_Sprite0 = &H40
Public Const PPU_SpriteCount = &H20
Public Const PPU_Ignored = &H10

Public reg8000 As Byte ' Needed for mapper #69.

Public map24_IRQEnabled As Byte
Public map24_IRQLatch As Byte
Public map24_IRQCounter As Byte
Public map24_IRQEnOnWrite As Byte ' Needed for mapper #24.

Public AndIt As Byte, AndIt2 As Byte

Public data(4) As Byte, sequence As Long 'Integer
Public accumulator As Long 'Integer

Public Render As Boolean

Public MotionBlur As Boolean
Public mScanlines As Boolean
Public mCut As Boolean
Public Smooth2x As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public OpenFile As OPENFILENAME

Public PrgMark As Long

Public ChrCnt As Byte, PrgCnt As Byte
Public FirstRead As Boolean 'First read to $2007 is invalid
Public Joypad1(7) As Byte, Joypad1_Count As Byte
Public Joypad2(7) As Byte, Joypad2_Count As Byte

Public Mapper As Byte, Mirroring As Byte, Trainer As Byte, FourScreen As Byte, Batt As Byte
Public MirrorXor As Long 'Integer
Public UsesSRAM As Boolean

Public Keyboard(255) As Byte

Public CPURunning As Boolean

Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb&)
Public Declare Sub MemFill Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public MMC32_Switch As Byte

Public MMC16_Irq As Long 'Integer
Public MMC16_IrqOn As Byte

Public PPUAddress2 As Long
Public HScroll2 As Long

Public SpriteAddr As Long 'Integer

Public Pal(255) As Long
Public Pal16(255) As Integer
Public Pal15(255) As Integer

Public NoScroll2006 As Boolean

Public Mapper40_IRQEnabled As Byte, Mapper40_IRQCounter As Byte

Public Latch1 As Byte, Latch2 As Byte
Public Latch0FD As Byte, Latch0FE As Byte
Public Latch1FD As Byte, Latch1FE As Byte

'Mapper 5
Public map5_PrgSize As Byte
Public map5_ChrSize As Byte
Public map5_BGChrPage(3) As Byte
Public map5_ChrPage(7) As Byte

'Mapper 23
Public map23_HiChr(7) As Byte
Public map23_LoChr(7) As Byte
Public map23_IRQEnabled As Byte
Public map23_IRQLatch As Byte
Public map23_IRQLatchLo As Byte
Public map23_IRQLatchHi As Byte
Public map23_IRQCounter As Byte
Public map23_IRQEnOnWrite As Byte

'Mapper 27
Public map27_Chr(7) As Byte

'Mapper 75
Public MMC_Prg0 As Byte
Public MMC_Prg1 As Byte
Public MMC_Prg(1) As Byte

'Mapper 83
Public map83CHR As Long

'Mapper 85
Public map85_IRQEnabled As Byte
Public map85_IRQLatch As Byte
Public map85_IRQCounter As Byte
Public map85_IRQEnOnWrite As Byte

'Mapper 99 VS-Unisystem
Public VSCoin As Byte

'Mapper 114
Public MemIndex As Long

'Mapper 117
Public Irq_Latch As Long
Public Irq_Counter As Long
Public Irq_Enabled As Byte

'Mapper 182
Public Map182Reg As Long

'Mapper 230
Public ResetSwitch As Boolean

'Declares
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Game Genie Functions
Public ggCodes As String
Public ggConvert As String
Public CodeAddY As Long
Public CodeCompVal As Long
Public CodeVal As Long

'Controllers
' Joypad 1
Public nes_ButA As Byte
Public nes_ButB As Byte
Public nes_ButSel As Byte
Public nes_ButSta As Byte
Public nes_ButUp As Byte
Public nes_ButDn As Byte
Public nes_ButLt As Byte
Public nes_ButRt As Byte
' Joypad 2
Public nes2_ButA As Byte
Public nes2_ButB As Byte
Public nes2_ButSel As Byte
Public nes2_ButSta As Byte
Public nes2_ButUp As Byte
Public nes2_ButDn As Byte
Public nes2_ButLt As Byte
Public nes2_ButRt As Byte
' Gamepad 1
Public pad_ButA As Byte
Public pad_ButB As Byte
Public pad_ButSel As Byte
Public pad_ButSta As Byte
' Gamepad 2
Public pad2_ButA As Byte
Public pad2_ButB As Byte
Public pad2_ButSel As Byte
Public pad2_ButSta As Byte
' Zapper
Public ZapperTrigger As Byte
Public ZapperLight As Byte
Public ZapperX, ZapperY As Integer
Public Zapper As Boolean

Public DipSwitch As Byte

'Other
Public Lang As Integer

Public Recents(4) As String
Public RomName As String
Public SlotIndex As Long 'for the save-state
Public Gamepad1, Gamepad2 As Integer
Public TileBased As Boolean
Public FScreen As Boolean
'Fills color lookup table
Public Sub fillTLook()
    Dim b1, b2, c, X
    For b1 = 0 To 255
    For b2 = 0 To 255
        For X = 0 To 7
            If b1 And Pow2(X) Then c = 1 Else c = 0
            If b2 And Pow2(X) Then c = c + 2
            tLook(b1 * 2048 + b2 * 8 + X) = c
        Next X
    Next b2, b1
End Sub
Public Sub BlitScreen()
    Static mbt As Boolean
    Static n As Long
    mbt = Not mbt
    n = 1 - n
    Static count64 As Long
    count64 = (count64 + 17) And 63
    
    Dim fscan As Boolean
    
    Dim t As Long
    
    Dim p01 As Integer, p10 As Integer, p21 As Integer, p12 As Integer
    
    Dim p01b As Long, p10b As Long, p21b As Long, p12b As Long
    
    Dim X, Y, k, B, c
    
    Dim i As Long
    Static p16(31) As Integer
    Static p32(31) As Long
    Static ColorDepth
    
    Dim xCount, xCnt As Long
    
    Dim mRgb(2) As Byte
    Dim sFact As Long
    Dim rgbFact As Long
    Dim MaskR, MaskG, MaskB As Long
    sFact = 32 'The color remove of scanline
    
    If FScreen = False Then
        If ColorDepth = 0 Or (Frames And 63) = 0 Then ColorDepth = GetColorDepth(frmNES.PicScreen)
    Else
        If ColorDepth = 0 Or (Frames And 63) = 0 Then ColorDepth = GetColorDepth(frmRender.PicScreen)
    End If
    Select Case ColorDepth
    Case 8
            ' Code deleted by the original team (Why?)
            ' blit8 vBuffer, frmNES.PicScreen, 256, 240
    Case 16
            If MotionBlur Then
                For i = 0 To 31
                    p16(i) = (Pal(VRAM(i + &H3F00)) And &HFEFEFE) \ 2
                Next i
                For i = 0 To 61439
                    vBuffer16(i) = (vBuffer16(i) And &HFEFEFE) \ 2 + p16(vBuffer(i))
                Next i
            End If
            If mScanlines Then
                For i = 0 To 31
                    p16(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 1
                    If xCount < 256 Then
                        vBuffer16(i) = p16(vBuffer(i))
                    Else
                        MemCopy mRgb(0), p16(vBuffer(i)), Len(p16(vBuffer(i)))
                        If mRgb(0) >= sFact Then mRgb(0) = mRgb(0) - sFact Else mRgb(0) = sFact
                        If mRgb(1) >= sFact Then mRgb(1) = mRgb(1) - sFact Else mRgb(1) = sFact
                        If mRgb(2) >= sFact Then mRgb(2) = mRgb(2) - sFact Else mRgb(2) = sFact
                        vBuffer16(i) = RGB(mRgb(0), mRgb(1), mRgb(2))
                        If xCount = 256 * 2 Then xCount = 0
                    End If
                    xCount = xCount + 1
                Next i
            End If
            If Not MotionBlur And Not mScanlines Then
                For i = 0 To 31
                    p16(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 4
                    vBuffer16(i) = p16(vBuffer(i))
                    vBuffer16(i + 1) = p16(vBuffer(i + 1))
                    vBuffer16(i + 2) = p16(vBuffer(i + 2))
                    vBuffer16(i + 3) = p16(vBuffer(i + 3))
                Next i
            End If
            If Smooth2x Then
                If Smooth2x = 1 Then
                    For Y = 0 To 239
                    i = Y * 256&
                    k = Y * 1024&
                    c = 0
                    For X = 0 To 255
                        B = vBuffer16(i)
                        vBuffer2x16(k + 1) = B
                        B = (B And &HF7DE&) \ 2
                        vBuffer2x16(k) = B + c
                        c = B
                        i = i + 1
                        k = k + 2
                    Next X, Y
                    For Y = 0 To 238
                    k = Y * 1024 + 512
                    For X = 0 To 511
                        vBuffer2x16(k) = (vBuffer2x16(k - 512) And &HF7DE&) \ 2 + (vBuffer2x16(k + 512) And &HF7DE&) \ 2
                        k = k + 1
                    Next X, Y
                ElseIf Smooth2x = 2 Then
                    For Y = 1 To 238
                    i = Y * 256&
                    k = Y * 1024&
                    c = vBuffer16(i - 1)
                    p21 = vBuffer16(i)
                    fscan = (Y And 63) = count64
                    For X = IIf(X > 1, 0, 1) To 255
                        B = vBuffer16(i + 1)
                        If B <> oldvBuffer16(i + 1) Then
                            If t <= 0 Then
                                c = vBuffer16(i - 1)
                                p21 = vBuffer16(i)
                            End If
                            t = 3
                        End If
                        If t > 0 Or fscan Then
                            t = t - 1
                            p01 = c
                            c = p21
                            p21 = B
                            B = (c And &HF7DE&) \ 2
                            p10 = vBuffer16(i - 256)
                            p12 = vBuffer16(i + 256)
                            If p01 = p10 And p12 = p21 Then
                                If p10 = p12 And vBuffer16(i - 257) = c Then
                                    vBuffer2x16(k) = c
                                    vBuffer2x16(k + 1) = (p10 And &HF7DE&) \ 2 + B
                                    vBuffer2x16(k + 512) = (p12 And &HF7DE&) \ 2 + B
                                    vBuffer2x16(k + 513) = c
                                Else
                                    vBuffer2x16(k) = (p01 And &HF7DE&) \ 2 + B
                                    vBuffer2x16(k + 1) = c
                                    vBuffer2x16(k + 512) = c
                                    vBuffer2x16(k + 513) = (p21 And &HF7DE&) \ 2 + B
                                End If
                            ElseIf p10 = p21 And p01 = p12 Then
                                vBuffer2x16(k) = c
                                vBuffer2x16(k + 1) = (p10 And &HF7DE&) \ 2 + B
                                vBuffer2x16(k + 512) = (p12 And &HF7DE&) \ 2 + B
                                vBuffer2x16(k + 513) = c
                            Else
                                vBuffer2x16(k) = c
                                vBuffer2x16(k + 1) = c
                                vBuffer2x16(k + 512) = c
                                vBuffer2x16(k + 513) = c
                            End If
                        End If
                        i = i + 1
                        k = k + 2
                    Next X, Y
                    MemCopy oldvBuffer16(0), vBuffer16(0), 122880
                End If
                If FScreen = False Then
                    Blit16 vBuffer2x16, frmNES.PicScreen, 512, 480
                Else
                    Blit16 vBuffer2x16, frmRender.PicScreen, 512, 480
                End If
            Else
                If FScreen = False Then
                    Blit16 vBuffer16, frmNES.PicScreen, 256, 240
                Else
                    Blit16 vBuffer16, frmRender.PicScreen, 256, 240
                End If
            End If
    Case 15
            If MotionBlur Then
                For i = 0 To 31
                    p16(i) = (Pal(VRAM(i + &H3F00)) And &HFEFEFE) \ 2
                Next i
                For i = 0 To 61439
                    vBuffer16(i) = (vBuffer16(i) And &HFEFEFE) \ 2 + p16(vBuffer(i))
                Next i
            End If
            If mScanlines Then
                For i = 0 To 31
                    p16(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 1
                    If xCount < 256 Then
                        vBuffer16(i) = p16(vBuffer(i))
                    Else
                        MemCopy mRgb(0), p16(vBuffer(i)), Len(p16(vBuffer(i)))
                        If mRgb(0) >= sFact Then mRgb(0) = mRgb(0) - sFact Else mRgb(0) = sFact
                        If mRgb(1) >= sFact Then mRgb(1) = mRgb(1) - sFact Else mRgb(1) = sFact
                        If mRgb(2) >= sFact Then mRgb(2) = mRgb(2) - sFact Else mRgb(2) = sFact
                        vBuffer16(i) = RGB(mRgb(0), mRgb(1), mRgb(2))
                        If xCount = 256 * 2 Then xCount = 0
                    End If
                    xCount = xCount + 1
                Next i
            End If
            If Not MotionBlur And Not mScanlines Then
                For i = 0 To 31
                    p16(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 4
                    vBuffer16(i) = p16(vBuffer(i))
                    vBuffer16(i + 1) = p16(vBuffer(i + 1))
                    vBuffer16(i + 2) = p16(vBuffer(i + 2))
                    vBuffer16(i + 3) = p16(vBuffer(i + 3))
                Next i
            End If
            If Smooth2x Then
                If Smooth2x = 1 Then
                    For Y = 0 To 239
                    i = Y * 256&
                    k = Y * 1024&
                    c = 0
                    For X = 0 To 255
                        B = vBuffer16(i)
                        vBuffer2x16(k + 1) = B
                        B = (B And &H7BDE&) \ 2
                        vBuffer2x16(k) = B + c
                        c = B
                        i = i + 1
                        k = k + 2
                    Next X, Y
                    For Y = 0 To 238
                    k = Y * 1024 + 512
                    For X = 0 To 511
                        vBuffer2x16(k) = (vBuffer2x16(k - 512) And &H7BDE&) \ 2 + (vBuffer2x16(k + 512) And &H7BDE&) \ 2
                        k = k + 1
                    Next X, Y
                ElseIf Smooth2x = 2 Then
                    For Y = 1 To 238
                    i = Y * 256&
                    k = Y * 1024&
                    c = vBuffer16(i - 1)
                    p21 = vBuffer16(i)
                    fscan = (Y And 63) = count64
                    For X = IIf(X > 1, 0, 1) To 255
                        If vBuffer16(i + 2) <> oldvBuffer16(i + 2) Then
                            If t <= 0 Then
                                c = vBuffer16(i - 1)
                                p21 = vBuffer16(i)
                            End If
                            t = 5
                        End If
                        If t > 0 Or fscan Then
                            t = t - 1
                            p01 = c
                            c = p21
                            p21 = vBuffer16(i + 1)
                            B = (c And &H7BDE&) \ 2
                            p10 = vBuffer16(i - 256)
                            p12 = vBuffer16(i + 256)
                            If p01 = p10 And p12 = p21 Then
                                If p10 = p12 And vBuffer16(i - 257) = c Then
                                    vBuffer2x16(k) = c
                                    vBuffer2x16(k + 1) = (p10 And &H7BDE&) \ 2 + B
                                    vBuffer2x16(k + 512) = (p12 And &H7BDE&) \ 2 + B
                                    vBuffer2x16(k + 513) = c
                                Else
                                    vBuffer2x16(k) = (p01 And &H7BDE&) \ 2 + B
                                    vBuffer2x16(k + 1) = c
                                    vBuffer2x16(k + 512) = c
                                    vBuffer2x16(k + 513) = (p21 And &H7BDE&) \ 2 + B
                                End If
                            ElseIf p10 = p21 And p01 = p12 Then
                                vBuffer2x16(k) = c
                                vBuffer2x16(k + 1) = (p10 And &H7BDE&) \ 2 + B
                                vBuffer2x16(k + 512) = (p12 And &H7BDE&) \ 2 + B
                                vBuffer2x16(k + 513) = c
                            Else
                                vBuffer2x16(k) = c
                                vBuffer2x16(k + 1) = c
                                vBuffer2x16(k + 512) = c
                                vBuffer2x16(k + 513) = c
                            End If
                        End If
                        i = i + 1
                        k = k + 2
                    Next X, Y
                    MemCopy oldvBuffer16(0), vBuffer16(0), 122880
                End If
                If FScreen = False Then
                    Blit15 vBuffer2x16, frmNES.PicScreen, 512, 480
                Else
                    Blit15 vBuffer2x16, frmRender.PicScreen, 512, 480
                End If
            Else
                If FScreen = False Then
                    Blit15 vBuffer16, frmNES.PicScreen, 256, 240
                Else
                    Blit15 vBuffer16, frmRender.PicScreen, 256, 240
                End If
            End If
    Case Else
            If mCut Then
                For i = 0 To 31
                    p32(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 1
                    If xCnt < 8 Then
                        vBuffer32(i) = vbBlack
                    Else
                        vBuffer32(i) = p32(vBuffer(i))
                    End If
                    xCnt = xCnt + 1
                    If xCnt = 256 Then xCnt = 0
                Next i
            End If
            If MotionBlur Then
                For i = 0 To 31
                    p32(i) = (Pal(VRAM(i + &H3F00)) And &HFEFEFE) \ 2
                Next i
                For i = 0 To 61439
                    vBuffer32(i) = (vBuffer32(i) And &HFEFEFE) \ 2 + p32(vBuffer(i))
                Next i
            End If
            If mScanlines Then
                For i = 0 To 31
                    p32(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 1
                    If xCount < 256 Then
                        vBuffer32(i) = p32(vBuffer(i))
                    Else
                        MemCopy mRgb(0), p32(vBuffer(i)), Len(p32(vBuffer(i)))
                        If mRgb(0) >= sFact Then mRgb(0) = mRgb(0) - sFact Else mRgb(0) = sFact
                        If mRgb(1) >= sFact Then mRgb(1) = mRgb(1) - sFact Else mRgb(1) = sFact
                        If mRgb(2) >= sFact Then mRgb(2) = mRgb(2) - sFact Else mRgb(2) = sFact
                        vBuffer32(i) = RGB(mRgb(0), mRgb(1), mRgb(2))
                        If xCount = 256 * 2 Then xCount = 0
                    End If
                    xCount = xCount + 1
                Next i
            End If
            If Not MotionBlur And Not mScanlines And Not mCut Then
                For i = 0 To 31
                    p32(i) = Pal(VRAM(i + &H3F00))
                Next i
                For i = 0 To 61439 Step 4
                    vBuffer32(i) = p32(vBuffer(i))
                    vBuffer32(i + 1) = p32(vBuffer(i + 1))
                    vBuffer32(i + 2) = p32(vBuffer(i + 2))
                    vBuffer32(i + 3) = p32(vBuffer(i + 3))
                Next i
            End If
            If Smooth2x Then
                If Smooth2x = 1 Then
                    For Y = 0 To 239
                    i = Y * 256&
                    k = Y * 1024&
                    c = 0
                    For X = 0 To 255
                        B = vBuffer32(i)
                        vBuffer2x32(k + 1) = B
                        B = (B And &HFEFEFE) \ 2
                        vBuffer2x32(k) = B + c
                        c = B
                        i = i + 1
                        k = k + 2
                    Next X, Y
                    For Y = 0 To 238
                    k = Y * 1024 + 512
                    For X = 0 To 511
                        vBuffer2x32(k) = (vBuffer2x32(k - 512) And &HFEFEFE) \ 2 + (vBuffer2x32(k + 512) And &HFEFEFE) \ 2
                        k = k + 1
                    Next X, Y
                ElseIf Smooth2x = 2 Then
                    t = n
                    For Y = 1 To 238
                    i = Y * 256&
                    k = Y * 1024&
                    c = vBuffer32(i - 1)
                    p21b = vBuffer32(i)
                    t = 1 - t
                    For X = IIf(X > 1, 0, 1) To 255
                        t = 1 - t
                        p01b = c
                        c = p21b
                        p21b = vBuffer32(i + 1)
                        If t Or c <> vBuffer2x32(k) Then
                            B = (c And &HFEFEFE) \ 2
                            p10b = vBuffer32(i - 256)
                            p12b = vBuffer32(i + 256)
                            If p01b = p10b And p12b = p21b Then
                                vBuffer2x32(k) = (p01b And &HFEFEFE) \ 2 + B
                                vBuffer2x32(k + 1) = c
                                vBuffer2x32(k + 512) = c
                                vBuffer2x32(k + 513) = (p21b And &HFEFEFE) \ 2 + B
                            ElseIf p10b = p21b And p01b = p12b Then
                                vBuffer2x32(k) = c
                                vBuffer2x32(k + 1) = (p10b And &HFEFEFE) \ 2 + B
                                vBuffer2x32(k + 512) = (p12b And &HFEFEFE) \ 2 + B
                                vBuffer2x32(k + 513) = c
                            Else
                                vBuffer2x32(k) = c
                                vBuffer2x32(k + 1) = c
                                vBuffer2x32(k + 512) = c
                                vBuffer2x32(k + 513) = c
                            End If
                        End If
                        i = i + 1
                        k = k + 2
                    Next X, Y
                End If
                If FScreen = False Then
                    Blit vBuffer2x32, frmNES.PicScreen, 512, 480
                Else
                    Blit vBuffer2x32, frmRender.PicScreen, 512, 480
                End If
            Else
                If FScreen = False Then
                    Blit vBuffer32, frmNES.PicScreen, 256, 240
                Else
                    Blit vBuffer32, frmRender.PicScreen, 256, 240
                End If
            End If
    End Select
End Sub
Public Sub RenderScanline(ByVal Scanline As Long)
    DoMirror ' set the mirroring
    'scanline based sprite rendering
    If Scanline > 239 Then Exit Sub
    
    If Scanline = 0 Then
        If Render Then MemFill vBuffer(0), 256& * 240&, 16
    
        'temporary measure until the mirroring problems can be fixed
        Static pm, pmx
        If MirrorXor <> pmx Then
            If MirrorXor = 0 Then
                'Mirroring = 2
            ElseIf MirrorXor = &H400& Then
                'Mirroring = 0
            Else
                'Mirroring = 1
            End If
        ElseIf pm <> Mirroring Then
            If Mirroring = 0 And MirrorXor <> &H400& Then
                    MirrorXor = &H400&
            ElseIf Mirroring = 1 And MirrorXor <> &H800& Then
                    MirrorXor = &H800&
            ElseIf Mirroring = 2 And MirrorXor <> 0 Then
                    MirrorXor = &H800&
            End If
        End If
        pm = Mirroring
        pmx = MirrorXor
    End If
    
    'Quick fix for scroll problem
    If ((PPU_Control2 And 8) = 0) Or Not Render Then  '((PPU_Control2 And 16) = 0)
        If Scanline > SpriteRAM(0) + 8 Then PPU_Status = PPU_Status Or 64
        Exit Sub
    End If
    
    Dim v As Long
    Dim nt2 As Byte
    'still some bugs in Little Nemo, Kirby.

    If Scanline = 0 Then
        PPUAddress = PPUAddress2
    Else
        PPUAddress = (PPUAddress And &HFBE0&) Or (PPUAddress2 And &H41F&)
    End If
   
    NameTable = &H2000& + (PPUAddress And &HC00)
    nt2 = (NameTable And &HC00&) \ &H400&

    HScroll = (PPUAddress And 31) * 8 + HScroll2
    VScroll = (PPUAddress \ 32 And 31) * 8 Or ((PPUAddress \ &H1000&) And 7)
    
    'If PPUAddress And 8192 Then VScroll = VScroll + 240
    
    VScroll = VScroll - Scanline
    
    v = PPUAddress
    
    ' the following "if" and contents were ported from Nester
    If (v And &H7000&) = &H7000& Then '/* is subtile y offset == 7? */
        v = v And &H8FFF& '/* subtile y offset = 0 */
        If (v And &H3E0&) = &H3A0& Then '/* name_tab line == 29? */
            v = v Xor &H800&   '/* switch nametables (bit 11) */
            v = v And &HFC1F&  '/* name_tab line = 0 */
        Else
            If (v And &H3E0&) = &H3E0& Then  '/* line == 31? */
                v = v And &HFC1F&  '/* name_tab line = 0 */
            Else
                v = v + &H20&
            End If
        End If
    Else
        v = v + &H1000&
    End If
    
    PPUAddress = v And &HFFFF&

    If Keyboard(219) And 1 Then VScroll = Frames Mod 240
    If Keyboard(221) And 1 Then HScroll = Frames And 255
    
    
    If Scanline = 239 Then PPU_Status = PPU_Status Or &H80
    If Not Render Then
        If PPU_Status And &H40 Then Exit Sub
        If Scanline > SpriteRAM(0) + 8 Then PPU_Status = PPU_Status Or &H40
        Exit Sub
    End If
    
    Dim TileRow As Byte, TileYOffset As Long 'Integer
    Dim TileCounter As Long 'Integer
    Dim Color As Long
    Dim TileIndex As Byte, Byte1 As Byte, Byte2 As Byte
    Dim LookUp As Byte, addToCol As Long
    Dim pixel As Long 'Integer
    Dim X As Long, Aa As Long
    Dim m As Long
    Dim sc As Long
    Dim atrtab As Long
    Static phs As Long, pvs As Long
    Dim Y As Long
    
    If TileBased Then
        Dim h As Long
        If PPU_Control1 And &H20 Then h = 16 Else h = 8
        If (PPU_Status And &H40) = 0 Then If Scanline > SpriteRAM(0) + h Then PPU_Status = PPU_Status Or &H40
        
        If Scanline = 0 Then
            DrawSprites True
        ElseIf Scanline = 236 Then
            DrawSprites False
        End If
    End If
    
    sc = Scanline + VScroll
    If sc > 480 Then sc = sc - 480
    
    'draw background
    TileRow = (sc \ 8) Mod 30
    TileYOffset = sc And 7
    
    If (Not TileBased) Or VScroll <> pvs Or HScroll <> phs Or TileYOffset = 0 Then
    If TileYOffset = 0 Then
        pvs = VScroll
        phs = HScroll
    End If
    
    atrtab = &H3C0
    PatternTable = (PPU_Control1 And &H10) * &H100&
    For TileCounter = HScroll \ 8 To 31
        TileIndex = nt(Mirror(nt2), TileCounter + TileRow * 32)
        'TileIndex = VRAM(NameTable + TileCounter + TileRow * 32)
        If Mapper = 9 Or Mapper = 10 Then
            If PatternTable = &H0 Then
                map9_latch TileIndex, False
            ElseIf PatternTable = &H1000& Then
                map9_latch TileIndex, True
            End If
        End If
        X = TileCounter * 8 - HScroll + 7
        If X < 7 Then m = X Else m = 7
        X = X + Scanline * 256&
        LookUp = nt(Mirror(nt2), (&H3C0& + TileCounter \ 4 + (TileRow \ 4) * &H8&))
        Select Case (TileCounter And 2) Or (TileRow And 2) * 2
            Case 0
                addToCol = LookUp * 4 And 12
            Case 2
                addToCol = LookUp And 12
            Case 4
                addToCol = LookUp \ 4 And 12
            Case 6
                addToCol = LookUp \ 16 And 12
        End Select
        If TileBased And TileYOffset = 0 Then
            For Y = 0 To 7
                Byte1 = VRAM(PatternTable + TileIndex * 16 + Y)
                Byte2 = VRAM(PatternTable + TileIndex * 16 + 8 + Y)
                Aa = Byte1 * 2048& + Byte2 * 8
                For pixel = m To 0 Step -1
                    Color = tLook(Aa + pixel)
                    If Color Then vBuffer(X - pixel) = Color Or addToCol
                Next pixel
                X = X + 256
            Next Y
        Else
            Byte1 = VRAM(PatternTable + TileIndex * 16 + TileYOffset)
            Byte2 = VRAM(PatternTable + TileIndex * 16 + 8 + TileYOffset)
            Aa = Byte1 * 2048& + Byte2 * 8
            For pixel = m To 0 Step -1
                Color = tLook(Aa + pixel)
                If Color Then vBuffer(X - pixel) = Color Or addToCol
            Next pixel
        End If
    Next TileCounter
    
    'NameTable = &H2000 + (PPUAddress And &H800) Xor &H400
    NameTable = NameTable Xor &H400
    nt2 = (NameTable And &HC00&) \ &H400&
    
    atrtab = &H3C0
    For TileCounter = 0 To HScroll \ 8
        TileIndex = nt(Mirror(nt2), TileCounter + TileRow * 32)
        'TileIndex = VRAM(NameTable + (TileCounter + TileRow * 32))
        If Mapper = 9 Or Mapper = 10 Then
            If PatternTable = &H0 Then
                map9_latch TileIndex, False
            ElseIf PatternTable = &H1000& Then
                map9_latch TileIndex, True
            End If
        End If
        X = TileCounter * 8 + 256 - HScroll + 7
        If X > 255 Then m = X - 255 Else m = 0
        X = X + Scanline * 256&
        LookUp = nt(Mirror(nt2), (&H3C0& + TileCounter \ 4 + (TileRow \ 4) * &H8&))
        Select Case (TileCounter And 2) Or (TileRow And 2) * 2
            Case 0
                addToCol = LookUp * 4 And 12
            Case 2
                addToCol = LookUp And 12
            Case 4
                addToCol = LookUp \ 4 And 12
            Case 6
                addToCol = LookUp \ 16 And 12
        End Select
        If TileBased And TileYOffset = 0 Then
            For Y = 0 To 7
                Byte1 = VRAM(PatternTable + TileIndex * 16 + Y)
                Byte2 = VRAM(PatternTable + TileIndex * 16 + 8 + Y)
                Aa = Byte1 * 2048& + Byte2 * 8
                For pixel = 7 To m Step -1
                    Color = tLook(Aa + pixel)
                    If Color Then vBuffer(X - pixel) = Color Or addToCol
                Next pixel
                X = X + 256
            Next Y
        Else
            Byte1 = VRAM(PatternTable + (TileIndex * 16) + TileYOffset)
            Byte2 = VRAM(PatternTable + (TileIndex * 16) + 8 + TileYOffset)
            Aa = Byte1 * 2048& + Byte2 * 8
            For pixel = 7 To m Step -1
                Color = tLook(Aa + pixel)
                If Color Then vBuffer(X - pixel) = Color Or addToCol
            Next pixel
        End If
    Next TileCounter
    End If
    
    If Not TileBased Then RenderSprites Scanline - 1
End Sub
Public Sub RenderSprites(ByVal Scanline As Long)
    Dim solid(264) As Boolean
    
    If (PPU_Control2 And 16) = 0 Or frmNES.mLayer2.Checked = False Then Exit Sub
       
    Dim TileRow As Byte, TileYOffset As Long 'Integer
    Dim TileCounter As Long 'Integer
    Dim Color As Byte
    Dim TileIndex As Byte, Byte1 As Byte, Byte2 As Byte
    Dim addToCol As Long
    Dim h As Byte
    Dim minX As Long
    Dim X1 As Long, Y1 As Long
    Dim ptable As Long
    
    TileRow = Scanline \ 8
    TileYOffset = Scanline And 7
    If PPU_Control1 And &H20 Then h = 16 Else h = 8
    If PPU_Control1 And &H8 Then
        If h = 8 Then PatternTable = &H1000&
    Else
        If h = 8 Then PatternTable = &H0
    End If
    If PPU_Control2 And &H8 Then minX = 0 Else minX = 8
    Dim spr As Long 'Integer
    Dim SpriteAddr As Integer
    Dim i As Long, X As Long, Aa As Long, v As Long
    Dim OnTop As Boolean
    Dim attr As Byte
    i = Scanline * 256&
    For spr = 0 To 63
        SpriteAddr = 4 * spr
        Y1 = SpriteRAM(SpriteAddr) + 1
        If Y1 <= Scanline And Y1 > Scanline - h Then
            attr = SpriteRAM(SpriteAddr + 2)
            OnTop = (attr And 32) = 0
            'If (attr And 32) = 0 Xor topLayer Then
                X1 = SpriteRAM(SpriteAddr + 3)
                If X1 >= minX Then
                    If Render Then
                        addToCol = &H10 + (attr And 3) * 4
                        TileIndex = SpriteRAM(SpriteAddr + 1)
                        If Mapper = 9 Or Mapper = 10 Then
                            If PatternTable = &H0 Then
                                map9_latch TileIndex, False
                            ElseIf PatternTable = &H1000& Then
                                map9_latch TileIndex, True
                            End If
                        End If
                        If h = 16 Then
                            If TileIndex And 1 Then
                                PatternTable = &H1000
                                TileIndex = TileIndex Xor 1
                            Else
                                PatternTable = 0
                            End If
                        End If
                        If attr And 128 Then 'vertical flip
                            v = Y1 - Scanline - 1
                        Else
                            v = Scanline - Y1
                        End If
                        v = v And h - 1
                        If v >= 8 Then v = v + 8
                        Byte1 = VRAM(PatternTable + (TileIndex * 16) + v)
                        Byte2 = VRAM(PatternTable + (TileIndex * 16) + 8 + v)
                        'real sprite 0 detection
                        If spr = 0 And (PPU_Status And 64) = 0 Then
                            If attr And 64 Then 'horizontal flip
                                Aa = i + X1
                                For X = 0 To 7
                                    If Byte1 And Pow2(X) Then Color = 1 Else Color = 0
                                    If Byte2 And Pow2(X) Then Color = Color + 2
                                    If Color Then
                                        If vBuffer(Aa + X) And 3 And (PPU_Status And 64) = 0 Then
                                            PPU_Status = PPU_Status Or 64
                                            If OnTop Then vBuffer(Aa + X) = addToCol Or Color
                                        Else
                                            vBuffer(Aa + X) = addToCol Or Color
                                        End If
                                        solid(X1 + X) = True
                                    End If
                                Next X
                            Else
                                Aa = i + X1 + 7
                                For X = 7 To 0 Step -1
                                    If Byte1 And Pow2(X) Then Color = 1 Else Color = 0
                                    If Byte2 And Pow2(X) Then Color = Color + 2
                                    If Color Then
                                        If vBuffer(Aa - X) And 3 And (PPU_Status And 64) = 0 Then
                                            PPU_Status = PPU_Status Or 64
                                            If OnTop Then vBuffer(Aa - X) = addToCol Or Color
                                        Else
                                            vBuffer(Aa - X) = addToCol Or Color
                                        End If
                                        solid(X1 + 7 - X) = True
                                    End If
                                Next X
                            End If
                        Else
                            If attr And 64 Then 'horizontal flip
                                Aa = i + X1
                                For X = 0 To 7 'draw yellow block for now.
                                    If Byte1 And Pow2(X) Then Color = 1 Else Color = 0
                                    If Byte2 And Pow2(X) Then Color = Color + 2
                                    If Color Then
                                        If Not solid(X1 + X) Then
                                            If OnTop Then
                                                vBuffer(Aa + X) = addToCol Or Color
                                            ElseIf (vBuffer(Aa + X) And 3) = 0 Then
                                                vBuffer(Aa + X) = addToCol Or Color
                                            End If
                                            solid(X1 + X) = True
                                        End If
                                    End If
                                Next X
                            Else
                                Aa = i + X1 + 7
                                For X = 7 To 0 Step -1 'draw yellow block for now.
                                    If Byte1 And Pow2(X) Then Color = 1 Else Color = 0
                                    If Byte2 And Pow2(X) Then Color = Color + 2
                                    If Color Then
                                        If Not solid(X1 + 7 - X) Then
                                            If OnTop Then
                                                vBuffer(Aa - X) = addToCol Or Color
                                            ElseIf (vBuffer(Aa - X) And 3) = 0 Then
                                                vBuffer(Aa - X) = addToCol Or Color
                                            End If
                                            solid(X1 + 7 - X) = True
                                        End If
                                    End If
                                Next X
                            End If
                        End If
                    End If
                    If spr = 0 Then If Scanline = Y1 + h - 1 Then PPU_Status = PPU_Status Or &H40 'claim we hit sprite #0
                End If
            'End If
        End If
    Next spr
End Sub
Public Sub map6_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H42FC& To &H42FD&: ' Unknown
        Case &H42FE& ' Page Select
            If Value And &H20 Then
                NameTable = &H2400&
            Else
                NameTable = &H2000&
            End If
        Case &H42FF& ' Mirroring
            Mirroring = (Value And &H20) \ &H20
            DoMirror
        Case &H4501&
            map6_irqon = 0
        Case &H4502&
            TmpLatch = Value
        Case &H4503&
            IRQCounter = (Value * &H100&) + TmpLatch
            map6_irqon = 1
        Case &H8000& To &HFFFF&
            reg8 = (Value And &HF) * 2
            regA = reg8 + 1
            regC = reg8 + 1
            regE = reg8 + 1
            Select8KVROM Value And &H3
            SetupBanks
    End Select
End Sub
Public Function map6_hblank(Scanline) As Byte
    If (map6_irqon <> 0) Then
        IRQCounter = IRQCounter + 1
        If (IRQCounter >= &HFFFF&) Then
            IRQCounter = 0
            irq6502
        End If
    End If
End Function
Public Sub map5_write(Address As Long, Value As Byte)
    'IT WORKS! Finally!
    'The MMC5 is the more complex Mapper i ever seen
    
    Select Case Address
        Case &H5100&: map5_PrgSize = Value And &H3
        Case &H5101&: map5_ChrSize = Value And &H3
        Case &H5105&
            Dim i As Integer
            
            For i = 0 To 3
                Select1KVROM Value And &H3, i
                Select1KVROM Value And &H3, i + 4
                Value = Value / 2
            Next i
        Case &H5114& To &H5117&
            If Value And &H80 Then
                Select Case Address And &H7
                    Case 4: If map5_PrgSize = 3 Then reg8 = Value And &H7F
                    Case 5
                        If map5_PrgSize = 1 Or map5_PrgSize = 2 Then
                            reg8 = Value And &H7F
                            regA = reg8 + 1
                        ElseIf map5_PrgSize = 3 Then
                            regA = Value And &H7F
                        End If
                    Case 6
                        If map5_PrgSize = 2 Or map5_PrgSize = 3 Then
                            regC = Value And &H7F
                        End If
                    Case 7
                        If map5_PrgSize = 0 Then
                            reg8 = RShift((Value And &H7F), 2)
                            regA = reg8 + 1
                            regC = reg8 + 2
                            regE = reg8 + 3
                        ElseIf map5_PrgSize = 1 Then
                            regC = Value And &H7F
                            regE = reg8 + 1
                        ElseIf map5_PrgSize = 2 Or map5_PrgSize = 3 Then
                            regE = Value And &H7F
                        End If
                End Select
                SetupBanks
            Else
                'WRAM Bank
            End If
        Case &H5120& To &H5127&
            'SP
            map5_ChrPage(Address And &H7) = Value
            
            Select Case map5_ChrSize
                Case 0
                    Select8KVROM map5_ChrPage(7)
                Case 1
                    Select4KVROM map5_ChrPage(3), 0
                    Select4KVROM map5_ChrPage(7), 4
                Case 2
                    Select2KVROM map5_ChrPage(1), 0
                    Select2KVROM map5_ChrPage(3), 2
                    Select2KVROM map5_ChrPage(5), 4
                    Select2KVROM map5_ChrPage(7), 6
                Case 3
                    Select1KVROM map5_ChrPage(0), 0
                    Select1KVROM map5_ChrPage(1), 1
                    Select1KVROM map5_ChrPage(2), 2
                    Select1KVROM map5_ChrPage(3), 3
                    Select1KVROM map5_ChrPage(4), 4
                    Select1KVROM map5_ChrPage(5), 5
                    Select1KVROM map5_ChrPage(6), 6
                    Select1KVROM map5_ChrPage(7), 7
            End Select
        Case &H5128 To &H512B&
            'BG
            map5_BGChrPage(Address And &H3) = Value
            
            Select Case map5_ChrSize
                Case 3
                    Select1KVROM map5_BGChrPage(0), 4
                    Select1KVROM map5_BGChrPage(1), 5
                    Select1KVROM map5_BGChrPage(2), 6
                    Select1KVROM map5_BGChrPage(3), 7
            End Select
    End Select
End Sub
Public Sub map90_write(Address As Long, Value As Byte)
    Dim TmpBank(4), i As Integer
    
    'Mapper 211 and 90, Tiny Toon 6 & DKC 4: garbled graphics and emulator hang
    Select Case Address
        Case &H8000&: reg8 = Value
        Case &H8001&: regA = Value
        Case &H8002&: regC = Value
        Case &H8003&: regE = Value
        Case &H9000& To &H9007&: Select1KVROM Value, Address And 7
        Case &HA000& To &HA007&: Select1KVROM Value, Address And 7
        Case &HC000 To &HC007& 'IRQ Regs
        Case &HD001& ' Mirroring
            'Control Regs (Bank Modes, mirroring) D000 to D003
            If Value = 0 Then Mirroring = 1 Else Mirroring = 0
            DoMirror
    End Select
    SetupBanks
End Sub
Public Sub map24_write(Address As Long, Value As Byte)
    If Mapper = 26 Then
        If (Address And &H3) = 1 Then
            Address = Address + 1
        ElseIf (Address And &H3) = 2 Then
            Address = Address - 1
        End If
    End If
    Select Case Address
        Case &H8000&
            reg8 = Value * 2
            regA = reg8 + 1
            Call SetupBanks
        Case &HB003&
            Mirroring = ((Value And &HC) \ &H4)
            If Mirroring = 0 Then
                Mirroring = 1
            Else
                Mirroring = 0
            End If
            DoMirror
        Case &HC000&
            regC = Value
            SetupBanks
        Case &HD000& To &HD003&
            Select1KVROM Value, (Address And &H3)
        Case &HE000& To &HE003&
            Select1KVROM Value, (Address And &H3) + 4
        Case &HF000&
            map24_IRQLatch = Value
        Case &HF001&
            map24_IRQEnabled = (Value And &H2)
            map24_IRQEnOnWrite = (Value And &H1)
            If (map24_IRQEnabled) Then
                map24_IRQCounter = map24_IRQLatch
            End If
        Case &HF002&
            map24_IRQEnabled = map24_IRQEnOnWrite
    End Select
End Sub
Public Sub map24_irq()
    If map24_IRQEnabled <> 0 Then
        If map24_IRQCounter = &HFF Then
            map24_IRQCounter = map24_IRQLatch
            irq6502
        Else
            map24_IRQCounter = map24_IRQCounter + 1
        End If
    End If
End Sub
Public Sub map13_write(Address As Long, Value As Byte)
    Dim prg_bank As Byte
    
    prg_bank = (Value And &H30) \ 16
            
    reg8 = prg_bank * 4: regA = reg8 + 1: regC = reg8 + 2: regE = reg8 + 3: SetupBanks
    
    Select4KVROM 0, 2
    Select4KVROM (Value And &H3), 3
   
    latch13 = Value
    
End Sub
Public Sub map16_write(Address As Long, Value As Byte)
    Select Case (Address And &HD)
        Case &H0 To &H7: Select1KVROM Value, Address And &H7
        Case &H8: reg8 = Value * 2: regA = reg8 + 1: SetupBanks
        Case &H9: Mirroring = (Value And &H1)
        Case &HA: If Value Then MMC16_IrqOn = 1 Else MMC16_IrqOn = 0
        Case &HB: TmpLatch = Value
        Case &HC: MMC16_Irq = (Value * &H100&) + TmpLatch
        Case &HD: ' Unknown
    End Select
End Sub
Public Sub map16_irq()
    If (MMC16_IrqOn <> 0) Then
        MMC16_Irq = MMC16_Irq - 1
        If (MMC16_Irq = 0) Then
            irq6502
            MMC16_IrqOn = 0
        End If
    End If
End Sub
Public Sub map65_write(Address As Long, Value As Byte)
' Mapper #65 - Irem H-3001
    Select Case Address
        Case &H8000&
            reg8 = Value
            SetupBanks
        Case &H9003& ' Mirroring
        Case &H9005& ' IRQ Control 1
        Case &H9006& ' IRQ Control 2
        Case &HA000&
            regA = Value
            SetupBanks
        Case &HB000& To &HB007&
            Select1KVROM Value, (Address And &H7)
        Case &HC000&
            regC = Value
            SetupBanks
    End Select
End Sub
Public Sub map19_write(Address As Long, Value As Byte)
    If Address < &H5000& Then Exit Sub
    Select Case Address
        Case &H5000& To &H57FF&
            TmpLatch = Value
        Case &H5800& To &H5FFF&
            If Value And &H80 Then MIRQOn = 1 Else MIRQOn = 0
            MMC19_IRQCount = ((Value And &H7F) * &H100&) + TmpLatch
        Case &H8000& To &H87FF&
            Select1KVROM Value, 0
        Case &H8800& To &H8FFF&
            Select1KVROM Value, 1
        Case &H9000& To &H97FF&
            Select1KVROM Value, 2
        Case &H9800& To &H9FFF&
            Select1KVROM Value, 3
        Case &HA000& To &HA7FF&
            Select1KVROM Value, 4
        Case &HA800& To &HAFFF&
            Select1KVROM Value, 5
        Case &HB000& To &HB7FF&
            Select1KVROM Value, 6
        Case &HB800& To &HBFFF&
            Select1KVROM Value, 7
        Case &HC000& To &HC7FF&
            If Value < &HE0 Then Select1KVROM Value, 8
        Case &HC800& To &HC8FF&
            If Value < &HE0 Then Select1KVROM Value, 9
        Case &HD000& To &HD7FF&
            If Value < &HE0 Then Select1KVROM Value, 10
        Case &HD800& To &HD8FF&
            If Value < &HE0 Then Select1KVROM Value, 11
        Case &HE000& To &HE7FF&
            reg8 = Value
            SetupBanks
        Case &HE800& To &HEFFF&
            regA = Value
            SetupBanks
        Case &HF000& To &HF7FF&
            regC = Value
            SetupBanks
    End Select
End Sub
Public Sub map19_irq()
    If MIRQOn = 1 Then
        MMC19_IRQCount = MMC19_IRQCount + 1
        If MMC19_IRQCount = &H7FFF& Then
            irq6502
        End If
    End If
End Sub
Public Sub DoMirror()
    MirrorXor = (((Mirroring + 1) Mod 3) * &H400&)
    If Mirroring = 0 Then
        Mirror(0) = 0: Mirror(1) = 0: Mirror(2) = 1: Mirror(3) = 1
    ElseIf Mirroring = 1 Then
        Mirror(0) = 0: Mirror(1) = 1: Mirror(2) = 0: Mirror(3) = 1
    ElseIf Mirroring = 2 Then
        Mirror(0) = 0: Mirror(1) = 0: Mirror(2) = 0: Mirror(3) = 0
    ElseIf Mirroring = 4 Then
        Mirror(0) = 0: Mirror(1) = 1: Mirror(2) = 2: Mirror(3) = 3
    End If
End Sub
Public Function Read6502(ByVal Address As Long) As Byte
    Dim Tmp As Byte
    Select Case Address
        Case &H0 To &H1FFF&: Read6502 = Bank0(Address And &H7FF&) ' NES ram 0-7ff mirrored at 800 1000 1800
        Case &H2000& To &H3FFF&
            Select Case (Address And &H7&)
                Case &H0&: Read6502 = ppuLatch
                Case &H1&: Read6502 = ppuLatch
                Case &H2&
                    Dim ret As Byte
                    ScrollToggle = 0
                    PPU_AddressIsHi = True
                    ret = (ppuLatch And &H1F) Or PPU_Status
                    If (ret And &H80) Then PPU_Status = (PPU_Status And &H60)
                    Read6502 = ret
                Case &H4&
                    Tmp = ppuLatch
                    ppuLatch = SpriteRAM(SpriteAddress)
                    SpriteAddress = (SpriteAddress + 1) And &HFF
                    Read6502 = Tmp
                Case &H5&: Read6502 = ppuLatch
                Case &H6&: Read6502 = ppuLatch
                Case &H7&
                    Tmp = ppuLatch
                    If Mapper = 9 Then
                        If PPUAddress < &H2000& Then
                            map9_latch Tmp, (PPUAddress And &H1000&)
                        End If
                    End If
                    If PPUAddress >= &H2000& And PPUAddress <= &H2FFF& Then
                        ppuLatch = nt(Mirror((PPUAddress And &HC00&) \ &H400&), PPUAddress And &H3FF&)
                    Else
                        ppuLatch = VRAM(PPUAddress And &H3FFF&)
                    End If
                    If (PPU_Control1 And &H4) Then
                        PPUAddress = PPUAddress + 32
                    Else
                        PPUAddress = PPUAddress + 1
                    End If
                    PPUAddress = (PPUAddress And &H3FFF&)
                    Read6502 = Tmp
            End Select
        Case &H4000& To &H4013&, &H4015&
            Read6502 = Sound(Address - &H4000&)
        Case &H4016& ' Joypad1
            Read6502 = Joypad1(Joypad1_Count)
            Joypad1_Count = (Joypad1_Count + 1) And 7
        Case &H4017& ' Joypad2
            If Zapper Then
                If Joypad2_Count = 3 Then
                    Read6502 = ZapperTrigger 'Zapper Trigger
                    'D1 - Pressed; 0 - Released
                ElseIf Joypad2_Count = 4 Then
                    Read6502 = ZapperLight
                    'D1 - No light detected; 0 - Light detected! ???
                End If
            Else
                Read6502 = Joypad2(Joypad2_Count)
            End If
            Joypad2_Count = (Joypad2_Count + 1) And 7
        Case &H4020&: If Mapper = 99 Then Read6502 = VSCoin
        Case &H5000&: Read6502 = DipSwitch
        Case &H6000& To &H7FFF&: Read6502 = Bank6(Address And &H1FFF&)
        Case &H8000& To &H9FFF&: Read6502 = Bank8(Address And &H1FFF&)
        Case &HA000& To &HBFFF&: Read6502 = BankA(Address And &H1FFF&)
        Case &HC000& To &HDFFF&: Read6502 = BankC(Address And &H1FFF&)
        Case &HE000& To &HFFFF&: Read6502 = BankE(Address And &H1FFF&)
    End Select
End Function
Public Sub Write6502(ByVal Address As Long, ByVal Value As Byte)
    On Error Resume Next 'Ignore Rom or Mapper errors
    If Address >= &H2000& And Address <= &H3FFF& Then ppuLatch = Value
    Select Case Address
        Case &H0& To &H1FFF&: Bank0(Address And &H7FF&) = Value
        Case &H2000& To &H3FFF&
            Select Case (Address And &H7)
                Case &H0&
                    PPU_Control1 = Value
                    PPUAddress2 = (PPUAddress2 And &HF3FF&) Or (Value And 3) * &H400&
                Case &H1&: PPU_Control2 = Value
                Case &H2&: ppuLatch = Value
                Case &H3&: SpriteAddress = Value
                Case &H4&
                    SpriteRAM(SpriteAddress) = Value
                    SpriteAddress = (SpriteAddress + 1) And &HFF
                    SpritesChanged = True
                Case &H5&
                    If PPU_AddressIsHi Then
                        HScroll2 = Value And 7
                        PPUAddress2 = (PPUAddress2 And &HFFE0&) Or Value \ 8
                        PPU_AddressIsHi = False
                    Else
                        PPUAddress2 = (PPUAddress2 And &H8C1F&) Or (Value And &HF8) * 4 Or (Value And 7) * &H1000&
                        PPU_AddressIsHi = True
                    End If
                Case &H6&
                    If PPU_AddressIsHi Then
                        PPUAddress = (PPUAddress And &HFF) Or ((Value And &H3F) * &H100&)
                        PPU_AddressIsHi = False
                    Else
                        PPUAddress = (PPUAddress And &H7F00&) Or Value
                        PPU_AddressIsHi = True
                    End If
                Case &H7&
                    ppuLatch = Value
                    If Mapper = 9 Then
                        If PPUAddress <= &H1FFF& And PPUAddress >= &H0& Then
                            map9_latch Value, (PPUAddress And &H1000&)
                        End If
                    End If
                    If PPUAddress >= &H2000& And PPUAddress <= &H2FFF& Then
                        nt(Mirror((PPUAddress And &HC00&) \ &H400&), PPUAddress And &H3FF&) = Value
                    Else
                        If PPUAddress <= 16383 Then VRAM(PPUAddress) = Value
                        If (PPUAddress And &HFFEF&) = &H3F00& Then VRAM(PPUAddress Xor 16) = Value
                    End If
                    If (PPU_Control1 And &H4) Then
                        PPUAddress = (PPUAddress + 32)
                    Else
                        PPUAddress = (PPUAddress + 1)
                    End If
                    PPU_AddressIsHi = True
                    PPUAddress = (PPUAddress And &H3FFF&)
            End Select
        Case &H4000& To &H4013&
            Sound(Address - &H4000&) = Value
            Dim n
            n = (Address - &H4000&) \ 4
            If n < 4 Then ChannelWrite(n) = True
        Case &H4014&
            MemCopy SpriteRAM(0), Bank0(Value * &H100&), 256 '&HFF FIXED
            SpritesChanged = True
        Case &H4015&: SoundCtrl = Value
        Case &H4016&
            'VS-Unisystem
            If Mapper = 99 Then If Value And &H4 Then Select8KVROM 1 Else Select8KVROM 0
        Case &H4020& To &H5FFF&
            If Address = &H4020& And Mapper = 99 Then VSCoin = Value 'VS-Unisystem
            Select Case Mapper 'Ex Write
                Case 5, 79, 164: MapperWrite Address, Value
            End Select
        Case &H6000& To &H7FFF&
            If SpecialWrite6000 Then
                MapperWrite Address, Value
            Else
                Bank6(Address And &H1FFF&) = Value
            End If
        Case &H8000& To &HFFFF&: MapperWrite Address, Value
    End Select
End Sub
Public Sub map11_write(Address As Long, Value As Byte)
    reg8 = 4 * (Value And &HF)
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    Select8KVROM (Value And &HF0)
End Sub
Public Sub map15_write(Address As Long, Value As Byte)
        map15_BankAddr = (Value And &H3F) * 2
    Select Case (Address And &H3)
        Case &H0
                map15_swapaddr = (Value And &H80)
                reg8 = map15_BankAddr
                regA = map15_BankAddr + 1
                regC = map15_BankAddr + 2
                regE = map15_BankAddr + 3
                SetupBanks
                Mirroring = (Value And &H40) \ &H40
        Case &H1
                map15_SwapReg = (Value And &H80)
                regC = map15_BankAddr
                regE = map15_BankAddr + 1
                SetupBanks
        Case &H2
            If (Value And &H80) Then
                reg8 = map15_BankAddr + 1
                regA = map15_BankAddr + 1
                regC = map15_BankAddr + 1
                regE = map15_BankAddr + 1
            Else
                reg8 = map15_BankAddr
                regA = reg8
                regC = regA
                regE = regC
            End If
            SetupBanks
        Case &H3
            map15_SwapReg = (Value And &H80)
            regC = map15_BankAddr
            regE = map15_BankAddr + 1
            SetupBanks
            Mirroring = (Value And &H40&) \ &H40&
            ' TODO: Add mirroring
    End Select
End Sub
Public Sub map18_write(Address As Long, Value As Byte)
    'Address = (Address And &HF003&)
    Select Case Address
        Case &H8000& To &H8001&: reg8 = Value
        Case &H8002& To &H8003&: regA = Value
        Case &H9000& To &H9001&: regC = Value
        Case &HA000& To &HA001&: Select1KVROM Value, 0
        Case &HA002& To &HA003&: Select1KVROM Value, 1
        Case &HB000& To &HB001&: Select1KVROM Value, 2
        Case &HB002& To &HB003&: Select1KVROM Value, 3
        Case &HC000& To &HC001&: Select1KVROM Value, 4
        Case &HC002& To &HC003&: Select1KVROM Value, 5
        Case &HD000& To &HD001&: Select1KVROM Value, 6
        Case &HD002& To &HD003&: Select1KVROM Value, 7
        Case &HF002&: If Value = 0 Then Mirroring = 0 Else Mirroring = 1: DoMirror
    End Select
    SetupBanks
End Sub
Public Sub map32_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000& To &H8FFF&
            If (MMC32_Switch And &H2) = 2 Then
                regC = Value
                SetupBanks
            Else
                reg8 = Value
                SetupBanks
            End If
        Case &H9000& To &H9FFF&
            Mirroring = (Value And &H1)
            MMC32_Switch = Value
            DoMirror
        Case &HA000& To &HAFFF&
            regA = Value
            SetupBanks
        Case &HBFF0&: Select1KVROM Value, 0
        Case &HBFF1&: Select1KVROM Value, 1
        Case &HBFF2&: Select1KVROM Value, 2
        Case &HBFF3&: Select1KVROM Value, 3
        Case &HBFF4&: Select1KVROM Value, 4
        Case &HBFF5&: Select1KVROM Value, 5
        Case &HBFF6&: Select1KVROM Value, 6
        Case &HBFF7&: Select1KVROM Value, 7
    End Select
End Sub
Public Sub map33_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&: reg8 = Value: SetupBanks
        Case &H8001&: regA = Value: SetupBanks
        Case &H8002&: Select2KVROM Value, 0
        Case &H8003&: Select2KVROM Value, 2
        Case &HA000&: Select1KVROM Value, 4
        Case &HA001&: Select1KVROM Value, 5
        Case &HA002&: Select1KVROM Value, 6
        Case &HA003&: Select1KVROM Value, 7
        Case &HC000, &HC001, &HE000&
    End Select
End Sub
Public Sub map34_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H7FFD&
            reg8 = 4 * (Value)
            regA = reg8 + 1
            regC = regA + 1
            regE = regC + 1
            SetupBanks
        Case &H7FFE&
            Select4KVROM Value, 0
        Case &H7FFF&
            Select4KVROM Value, 1
        Case &H8000& To &HFFFF&
            reg8 = 4 * (Value)
            regA = reg8 + 1
            regC = regA + 1
            regE = regC + 1
            SetupBanks
    End Select
End Sub
Public Sub map40_write(Address As Long, Value As Byte)
    Select Case (Address And &HE000&)
        Case &H8000&
            Mapper40_IRQEnabled = 0
            Mapper40_IRQCounter = 36
        Case &HA000&
            ' IRQ enable
            Mapper40_IRQEnabled = 1
        Case &HE000&
            regC = Value
            SetupBanks
    End Select
End Sub
Public Sub map40_irq()
    If Mapper40_IRQEnabled = 1 Then
        Mapper40_IRQCounter = (Mapper40_IRQCounter - 1) And 36
        If Mapper40_IRQCounter = 0 Then irq6502
    End If
End Sub
Public Sub map64_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&
            Cmd = Value And &HF
            Prg = (Value And &H40)
            Chr1 = (Value And &H80)
        Case &H8001&
            Select Case Cmd
                Case 0
                    If (Chr1) Then
                        Call Select1KVROM(Value, 4)
                        Call Select1KVROM(Value + 1, 5)
                    Else
                        Call Select1KVROM(Value, 0)
                        Call Select1KVROM(Value, 1)
                    End If
                Case 1
                    If (Chr1) Then
                        Call Select1KVROM(Value, 6)
                        Call Select1KVROM(Value + 1, 7)
                    Else
                        Call Select1KVROM(Value, 2)
                        Call Select1KVROM(Value + 1, 3)
                    End If
                Case 2
                    If (Chr1) Then
                        Call Select1KVROM(Value, 0)
                    Else
                        Call Select1KVROM(Value, 4)
                    End If
                Case 3
                    If (Chr1) Then
                        Call Select1KVROM(Value, 1)
                    Else
                        Call Select1KVROM(Value, 5)
                    End If
                Case 4
                    If (Chr1) Then
                        Call Select1KVROM(Value, 2)
                    Else
                        Call Select1KVROM(Value, 6)
                    End If
                Case 5
                    If (Chr1) Then
                        Call Select1KVROM(Value, 3)
                    Else
                        Call Select1KVROM(Value, 7)
                    End If
                Case 6
                    If (Prg) Then
                        regA = Value
                    Else
                        reg8 = Value
                    End If
                    SetupBanks
                Case 7
                    If (Prg) Then regC = Value Else reg8 = Value
                    SetupBanks
                Case 8: Call Select1KVROM(Value, 1)
                Case 9: Call Select1KVROM(Value, 3)
                Case &HF
                    If (Prg) Then reg8 = Value Else regC = Value
                    SetupBanks
            End Select
        Case &HA000&
            Mirroring = (Value And &H1)
            DoMirror
    End Select
End Sub
Public Sub map66_write(Address As Long, Value As Byte)
    reg8 = (Value * 4)
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    Select8KVROM Value
End Sub
Public Sub map68_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&: Select2KVROM Value, 0
        Case &H9000&: Select2KVROM Value, 1
        Case &HA000&: Select2KVROM Value, 2
        Case &HB000&: Select2KVROM Value, 3
        Case &HE000&: Mirroring = (Value And &H3): DoMirror
        Case &HF000&: reg8 = (Value * 2): regA = (Value * 2) + 1: SetupBanks
    End Select
End Sub
Public Sub map69_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&: reg8000 = Value And &HF
        Case &HA000&
            Select Case reg8000
                Case 0 To 7: Select1KVROM Value, reg8000
                Case 8: Value = MaskBankAddress(Value): MemCopy Bank6(0), GameImage(Value * &H2000&), &H2000&
                Case 9: reg8 = Value
                Case 10: regA = Value
                Case 11: regC = Value
                Case 12: Mirroring = (Value And &H3): DoMirror
                Case 13: ' IRQ
                Case 14: ' Low byte of scanline
                Case 15: ' High byte of scanline
            End Select
    End Select
End Sub
Public Sub map71_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000& To &HBFFF&: ' Unknown
        Case &HC000& To &HFFFF&: reg8 = (Value * 2): regA = (Value * 2) + 1: SetupBanks
    End Select
End Sub
Public Sub map78_write(Address As Long, Value As Byte)
    Dim VRomPtr As Byte, RomPtr As Byte
    VRomPtr = (Value \ &H10&) And &HF
    RomPtr = Value And &HF
    Select8KVROM VRomPtr
    reg8 = (RomPtr * 2)
    regA = (RomPtr * 2) + 1
    SetupBanks
End Sub
Public Sub map91_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H6000& To &H6003&: Select2KVROM Value, (Address And &H3)
        Case &H7000&: reg8 = Value: SetupBanks
        Case &H7001&: regA = Value: SetupBanks
    End Select
End Sub
Public Sub map227_write(Address As Long, Value As Byte)
    Dim mBank As Byte
    
    mBank = (Address And &H100) + 4
    If Address And &H1 Then
        reg8 = mBank
        regA = reg8 + 1
        regC = reg8 + 2
        regE = reg8 + 3
    Else
        If Address And &H4 Then
            reg8 = (mBank * 4) + 2
            regA = (mBank * 4) + 3
            regC = (mBank * 4) + 2
            regE = (mBank * 4) + 3
        Else
            reg8 = (mBank * 4)
            regA = (mBank * 4) + 1
            regC = (mBank * 4)
            regE = (mBank * 4) + 1
        End If
    End If
    If Address And &H80 Then
        If Address And &H200 Then
            regC = ((mBank And &H1C) * 4) + 14
            regE = ((mBank And &H1C) * 4) + 15
        Else
            regC = ((mBank And &H1C) * 4)
            regE = ((mBank And &H1C) * 4) + 1
        End If
    End If
    If Address And &H2 Then Mirroring = 0 Else Mirroring = 1: DoMirror
End Sub
Public Sub map228_write(Address As Long, Value As Byte)
    Dim mPrg As Long
    
    mPrg = RShift((Address And &H780), 7)
    If InStr(LCase(RomName), "action 52") Then 'Quick fix to run Cheetah Man 2
        Select Case RShift((Address And &H1800), 11)
            Case 1: mPrg = &H10
            Case 3: mPrg = &H20
        End Select
    End If
    If Address And &H20 Then
        mPrg = LShift(mPrg, 1)
        If Address And &H40 Then mPrg = mPrg + 1
        reg8 = (mPrg * 4) Mod LShift(PrgCount, 1)
        regA = (reg8 + 1) Mod LShift(PrgCount, 1)
        regC = (mPrg * 4) Mod LShift(PrgCount, 1)
        regE = (regC + 1) Mod LShift(PrgCount, 1)
    Else
        reg8 = (mPrg * 4) Mod LShift(PrgCount, 1)
        regA = (reg8 + 1) Mod LShift(PrgCount, 1)
        regC = (reg8 + 2) Mod LShift(PrgCount, 1)
        regE = (reg8 + 3) Mod LShift(PrgCount, 1)
    End If
    
    Select8KVROM LShift((Address And &HF), 2) + (Value And &H3)
    
    If (Address And &H2000) Then Mirroring = 1 Else Mirroring = 0
    DoMirror
    
    SetupBanks
End Sub
Public Sub map230_write(Address As Long, Value As Byte)
    If ResetSwitch Then
        reg8 = (Value And &H7) '* 2
        regA = reg8 + 1
        regC = 7
        regE = regC
        Mirroring = 1
    Else
        If Value And &H20 Then
            reg8 = (Value And &H1F) + 8
            regA = reg8 + 1
            regC = regA + 1
            regE = regC + 1
        Else
            reg8 = (Value And &H1F) + 4
            regA = reg8 + 1
            regC = regA + 1
            regE = regC + 1
        End If
        If Value And &H40 Then Mirroring = 1 Else Mirroring = 0
        DoMirror
    End If
End Sub
Public Sub map231_write(Address As Long, Value As Byte)
    Dim mBank As Byte
    If Address And &H20 Then
        reg8 = Address + 1
        regA = reg8 + 1
        regC = reg8 + 2
        regE = reg8 + 3
    Else
        mBank = Address And &H1E
        reg8 = mBank * 2
        regA = (mBank * 2) + 1
        regC = mBank * 2
        regE = (mBank * 2) + 1
    End If
    If Address And &H80 Then Mirroring = 0 Else Mirroring = 1: DoMirror
End Sub
Public Sub MapperWrite(ByVal Address As Long, ByVal Value As Byte)
    '===================================='
    '       MapperWrite(Address,value)   '
    ' Selects/Switches Chr-ROM and Prg-  '
    ' ROM depending on the mapper. Based '
    ' on DarcNES.                        '
    '===================================='
    Select Case Mapper
        Case 1, 185: map1_write Address, Value
        Case 2, 73, 130, 188, 232: map2_write Address, Value
        Case 3: map3_write Address, Value
        Case 4, 45, 47, 118, 119, 115, 95, 12, 64, 158, 74, 245: map4_write Address, Value
        Case 5: map5_write Address, Value
        Case 6: map6_write Address, Value
        Case 7, 51, 53, 226: map7_write Address, Value
        Case 8, 66, 107, 229, 255: map66_write Address, Value
        Case 9, 10: map9_write Address, Value
        Case 11, 234, 46: map11_write Address, Value
        Case 13: map13_write Address, Value
        Case 15: map15_write Address, Value
        Case 16: map16_write Address, Value
        Case 17: map17_write Address, Value
        Case 18: map18_write Address, Value
        Case 19, 86, 87, 113, 184: map19_write Address, Value
        Case 21: map21_write Address, Value
        Case 22: map22_write Address, Value
        Case 23: map23_write Address, Value
        Case 24, 26: map24_write Address, Value
        Case 32: map32_write Address, Value
        Case 33: map33_write Address, Value
        Case 34: map34_write Address, Value
        Case 40, 61, 200: map40_write Address, Value
        Case 57: map57_write Address, Value
        Case 58, 174: map58_write Address, Value
        Case 64: map64_write Address, Value
        Case 65: map65_write Address, Value
        Case 68: map68_write Address, Value
        Case 69: map69_write Address, Value
        Case 71: map71_write Address, Value
        Case 75: map75_write Address, Value
        Case 78: map78_write Address, Value
        Case 79, 146: map79_write Address, Value
        Case 83: map83_write Address, Value
        Case 85: map85_write Address, Value
        Case 90, 160, 211: map90_write Address, Value
        Case 91: map91_write Address, Value
        Case 94: map94_write Address, Value
        Case 99: If InStr(LCase(RomName), "castlevania") Then map2_write Address, Value
        Case 117: map117_write Address, Value
        Case 151: map151_write Address, Value
        Case 182: map182_write Address, Value
        Case 201: map201_write Address, Value
        Case 212: map212_write Address, Value
        Case 227: map227_write Address, Value
        Case 228: map228_write Address, Value
        Case 230: map230_write Address, Value
        Case 231: map231_write Address, Value
        Case 250: map250_write Address, Value
        Case 255: map255_write Address, Value
    End Select
End Sub
Public Sub map22_write(Address As Long, Value As Byte)
    ' Konami VRC2 Type A
    ' This mapper was a breeze.
    Select Case Address
        Case &H8000&
            reg8 = Value
            SetupBanks
        Case &H9000&
            Mirroring = (Value And &H3)
        Case &HA000&
            regA = Value
            SetupBanks
        Case &HB000&
            Select1KVROM Value \ 2, 0
        Case &HB001&
            Select1KVROM Value \ 2, 1
        Case &HC000&
            Select1KVROM Value \ 2, 2
        Case &HC001&
            Select1KVROM Value \ 2, 3
        Case &HD000&
            Select1KVROM Value \ 2, 4
        Case &HD001&
            Select1KVROM Value \ 2, 5
        Case &HE000&
            Select1KVROM Value \ 2, 6
        Case &HE001&
            Select1KVROM Value \ 2, 7
    End Select
End Sub
Public Sub map57_write(Address As Long, Value As Byte)
    Select Case Address
    Case &H8000& To &H8003&
        If Value And &H40& Then
            Select8KVROM (Value And &H3&) + (&H10& \ &H2&) + 7
        End If
    Case &H8800&
        If Value And &H80& Then
            reg8 = ((Value And &H40&) \ &H40&) * 4 + 8
            regA = reg8 + 1
            regC = reg8 + 2
            regE = reg8 + 3
        Else
            reg8 = ((Value And &H60&) \ &H20&) * 2
            regA = reg8 + 1
            regC = reg8
            regE = regA
        End If
        SetupBanks
        Select8KVROM (Value And 7) + ((Value And &H10&) \ &H2&)
        If Value And 8 Then Mirroring = 0 Else Mirroring = 1
    End Select
End Sub
Public Sub map58_write(Address As Long, Value As Byte)
    If Address And &H40 Then
        reg8 = Address And &H7
        regA = reg8 + 1
        regC = reg8
        regE = regA
    Else
        reg8 = RShift(Address And &H6, 1)
        regA = reg8 + 1
        regC = reg8 + 2
        regE = reg8 + 3
    End If
    SetupBanks
    Select8KVROM RShift(Address And &H38, 3)
End Sub
Public Sub map75_write(Address As Long, Value As Byte)
    Select Case (Address And &HF000)
        Case &H8000
            reg8 = Value
            Call SetupBanks
        Case &H9000
            If Value And 1 Then Mirroring = 0 Else Mirroring = 1
            DoMirror
            MMC_Prg(0) = (MMC_Prg0 And &HF) Or ((Value And 2) * 8)
            MMC_Prg(1) = (MMC_Prg1 And &HF) Or ((Value And 4) * 4)
            Select4KVROM MMC_Prg(0), 0
            Select4KVROM MMC_Prg(1), 1
        Case &HA000
            regA = Value
            Call SetupBanks
        Case &HC000
            regC = Value
            Call SetupBanks
        Case &HE000
            MMC_Prg(0) = (MMC_Prg(0) And &H10) Or (Value And &HF)
            Select4KVROM MMC_Prg(0), 0
        Case &HF000
            MMC_Prg(1) = (MMC_Prg1 And &H10) Or (Value And &HF)
            Select4KVROM MMC_Prg(1), 1
    End Select
End Sub
Public Sub map94_write(Address As Long, Value As Byte)
    If (Address And &HFFF0) = &HFF00 Then
        reg8 = ((Value \ 4) And 7) * 2
        regA = reg8 + 1
        SetupBanks
    End If
End Sub
Public Function LoadNES(ByVal FileName As String) As Byte
    TmpLatch = 0
    '===================================='
    '           LoadNES(filename)        '
    ' Used to Load the NES ROM/VROM to   '
    ' specified arrays, GameImage and    '
    ' VROM, then figures out what to do  '
    ' based on the mapper number.        '
    '===================================='
    Dim Header As String * 3
    Dim FileNum As Integer
    FileNum = FreeFile
    
    Dim i As Long
    Dim ROMCtrl As Byte, ROMCtrl2 As Byte
    
    Erase VRAM: Erase VROM: Erase GameImage: Erase Bank8: Erase BankA
    Erase BankC: Erase BankE: Erase Bank0: Erase Bank6
    
    SpecialWrite6000 = False
    
    If Dir$(FileName) = vbNullString Then MsgBox "Arquivo não encontrado!", vbCritical, "YoshiNES": LoadNES = 0: Exit Function
    
    Close #1
    
    PrgCount = 0: PrgCount2 = 0: ChrCount = 0: ChrCount2 = 0
    
    Open FileName For Binary As #FileNum
        Get #FileNum, , Header
        If Header <> "NES" Then
            MsgBox "Cabeçalho da ROM inválido!", vbCritical, "YoshiNES"
            LoadNES = 0
            Close #FileNum
            Exit Function
        End If
        
        Get #FileNum, 5, PrgCount: PrgCount2 = PrgCount
        Get #FileNum, 6, ChrCount: ChrCount2 = ChrCount
        Get #FileNum, 7, ROMCtrl
        Get #FileNum, 8, ROMCtrl2
        
        Mapper = (ROMCtrl And &HF0) \ 16
        Mapper = Mapper + ROMCtrl2
        
        If Mapper <> 0 And (ChrCount = 1 And PrgCount = 2) Then Mapper = 0
        If Mapper <> 0 And (ChrCount = 1 And PrgCount = 1) Then Mapper = 0
        
        Batt = ROMCtrl And &H2
        Trainer = ROMCtrl And &H4
        Mirroring = ROMCtrl And &H1
        FourScreen = ROMCtrl And &H8
        If ROMCtrl And &H2 Then UsesSRAM = True
        PrgMark = (PrgCount2 * &H4000&) - 1
        If Trainer Then
            Get #FileNum, 17, Train
        End If
        
        ReDim GameImage(PrgMark) As Byte
        Dim StartAt As Integer
        If Trainer Then StartAt = 529 Else StartAt = 17
        Get #FileNum, StartAt, GameImage
    
        ReDim VROM(ChrCount2 * &H2000&) As Byte
        PrgMark = &H4000& * PrgCount2 + StartAt
        If ChrCount2 Then
            Get #FileNum, PrgMark, VROM
        End If
        
        'Add the Game Genie codes
        If Len(ggCodes) <> 0 Then
            For i = 1 To Len(ggCodes)
                ggConvert = Mid(ggCodes, i, 6)
                CodeAddY = ggAddY(ggConvert)
                CodeVal = ggVal(ggConvert)
                GameImage(CodeAddY) = CByte(CodeVal)
                i = i + 5
            Next i
        End If
        If Dir(App.Path & "\" & PalName) = PalName Then LoadPal PalName Else NewPal
        
        mmc_reset
        
        ResetSwitch = Not ResetSwitch
        Select Case Mapper
            Case 0, 2, 3, 18, 21, 22, 23, 32, 33, 47, 73, 75, 85, 88, 90, 95, 130, 185, 188, 200, 232, 255
                If ChrCount Then Select8KVROM 0
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
            Case 1
                Select8KVROM 0
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
                sequence = 0: accumulator = 0
                Erase data
                data(0) = &H1F: data(3) = 0
            Case 4, 12, 45, 64, 74, 95, 115, 118, 158, 245, 250
                swap = False
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                map4_sync
                MMC3_IrqVal = 0: MMC3_IrqOn = False: MMC3_TmpVal = 0
                If ChrCount Then Select8KVROM 0
            Case 5
                reg8 = &HFC
                regA = &HFD
                regC = &HFE
                regE = &HFF
                SetupBanks
                Select8KVROM 0
                TileBased = True
            Case 6
                reg8 = &H0
                regA = &H1
                regC = &H7
                regE = &H8
                SetupBanks
                If ChrCount Then Select8KVROM 0
            Case 7, 51, 53, 226
                reg8 = 0
                regA = 1
                regC = 2
                regE = 3
                SetupBanks
            Case 8, 58, 107, 174, 229, 255
                reg8 = 0: regA = 1: regC = 2: regE = 3: SetupBanks
                Select8KVROM 0
                SetupBanks
            Case 9, 10 ' MMC2, MMC4
                reg8 = 0
                regA = &HFD
                regC = &HFE
                regE = &HFF
                SetupBanks
                Latch1 = &HFE
                Latch2 = &HFE
                Select8KVROM 0
            Case 11, 234, 201, 46
                reg8 = 0
                regA = 1
                regC = 2
                regE = 3
                SetupBanks
                Select8KVROM 0
            Case 13
                reg8 = 0: regA = 1: regC = 2: regE = 3: SetupBanks
                Select4KVROM 0, 0
                Select4KVROM 0, 1
                latch13 = 0
            Case 15
                reg8 = 0
                regA = 1
                regC = 2
                regE = 3
                SetupBanks
            Case 16
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
                MMC16_Irq = 0: MMC16_IrqOn = 0
                SpecialWrite6000 = True
            Case 17
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
                map17_irq = 0: map17_irqon = False
            Case 19, 86, 87, 113, 184
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
                Select8KVROM ChrCount - 1
            Case 23, 24, 26
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
                map24_IRQCounter = 0
                map24_IRQEnabled = 0
                map24_IRQLatch = 0
                map24_IRQEnOnWrite = 0
            Case 34
                SpecialWrite6000 = True
                reg8 = 0
                regA = 1
                regC = 2
                regE = 3
                SetupBanks
                If ChrCount Then Select8KVROM 0
            Case 40, 61
                Dim ChoseIt As Byte
                ChoseIt = 6
                ChoseIt = MaskBankAddress(ChoseIt)
                UsesSRAM = True
                MemCopy Bank6(0), GameImage(ChoseIt * &H2000&), &H2000&
                reg8 = &HFC
                regA = &HFD
                regC = &HFE
                regE = &HFF
                SetupBanks
                Select8KVROM 0
                Mapper40_IRQEnabled = 0
                Mapper40_IRQCounter = 0
            Case 57
                reg8 = 0
                regA = 1
                regC = 0
                regE = 1
                SetupBanks
                Select8KVROM 0
            Case 64
                Dim banks As Byte: banks = PrgCount * 2
                reg8 = &HFF: regA = reg8: regC = reg8: regE = reg8
                SetupBanks
                If ChrCount Then Select8KVROM 0
                Cmd = 0: Chr1 = 0: Prg = 0
            Case 65
                reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
                SetupBanks
            Case 66
                reg8 = 0: regA = 1: regC = 2: regE = 3
                Select8KVROM 0
                SetupBanks
            Case 68
                reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
                Select8KVROM 0
                SetupBanks
                UsesSRAM = True
            Case 69
                reg8 = 0
                regA = 1
                regC = 2
                regE = &HFF
                SetupBanks
                Select8KVROM 0
            Case 71
                reg8 = 0
                regA = 1
                regC = &HFE
                regE = &HFF
                SetupBanks
            Case 78
                regC = &HFE
                regE = &HFF
                SetupBanks
                map78_write 0, 0
            Case 79, 146
                reg8 = 0
                regA = 1
                regC = 2
                regE = 3
                SetupBanks
                Select8KVROM 1
            Case 83
                reg8 = 0
                regA = 1
                regC = 30
                regE = 31
                SetupBanks
            Case 91
                SpecialWrite6000 = True
                reg8 = &HFE
                regA = &HFF
                regC = &HFE
                regE = &HFF
                Select8KVROM 0
                SetupBanks
            Case 94
                reg8 = 0
                regA = 1
                regC = PrgCount * 2 - 2
                regE = PrgCount * 2 - 1
                SetupBanks
            Case 99, 151
                SetVSPal
                reg8 = 0: regA = 1: regC = &HFE: regE = &HFF
                SetupBanks
                If ChrCount Then Select8KVROM 0
            Case 117
                reg8 = 0
                regA = 1
                regC = PrgCount * 2 - 2
                regE = PrgCount * 2 - 1
                SetupBanks
                If ChrCount Then Select8KVROM 0
                Irq_Enabled = 0: Irq_Counter = 0
            Case 160, 182, 211, 212, 231
                reg8 = &HFC
                regA = &HFD
                regC = &HFE
                regE = &HFF
                SetupBanks
            Case 227
                reg8 = 0
                regA = 1
                regC = 0
                regE = 1
                SetupBanks
            Case 228
                reg8 = 0
                regA = 1
                regC = 2
                regE = 3
                SetupBanks
            Case 230
                If ResetSwitch Then
                    reg8 = 0
                    regA = 1
                    regC = &HE
                    regE = &HF
                Else
                    reg8 = &H10
                    regA = &H11
                    regC = PrgCount - 2
                    regE = PrgCount - 1
                End If
                SetupBanks
            Case Else
                MsgBox "O mapper #" & Mapper & " não é suportado.", vbCritical, VERSION
                Erase GameImage: Erase VROM
                Close #FileNum: LoadNES = 0: Exit Function
        End Select
        Debug.Print "Successfully loaded " & FileName
        reset6502
        If Mirroring = 1 Then MirrorXor = &H800& Else MirrorXor = &H400&
        If FourScreen Then Mirroring = 4
        DoMirror
    Close #FileNum
    
    CurrentLine = 0
    For i = 0 To 7
        Joypad1(i) = &H40
        Joypad2(i) = &H40
    Next i
    
    FileNum = FreeFile
    If UsesSRAM = True Then ' save the SRAM to a file.
        If Dir(App.Path & "\Srams\" & RomName & ".wrm") <> "" Then
            Open App.Path & "\Srams\" & RomName & ".wrm" For Binary Access Read As #FileNum
                Get #FileNum, , Bank6
            Close #FileNum
        End If
    End If
    
    Frames = 0
    CPUPaused = False
    ScrollToggle = 1
    frmNES.mnuFileRomInfo.Enabled = True
    frmNES.mnuFileFree.Enabled = True
    LoadNES = 1
End Function
Public Sub map17_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H42FE:
        Case &H42FF:
        Case &H4501: map17_irqon = (Value And &H1)
        Case &H4502: map17_irq = &HFF00&: map17_irq = map17_irq Or Value
        Case &H4503: map17_irq = &HFF: map17_irq = map17_irq Or Value * &H100&: map17_irqon = 1
        Case &H4504: reg8 = Value: SetupBanks
        Case &H4505: regA = Value: SetupBanks
        Case &H4506: regC = Value: SetupBanks
        Case &H4507: regE = Value: SetupBanks
        Case &H4510 To &H4517: Select1KVROM Value, (Address - &H4510)
    End Select
End Sub
Public Sub map17_doirq()
    If map17_irqon Then
        map17_irq = (map17_irq + 1)
        If map17_irq = &H10000 Then
            irq6502
            map17_irqon = False
            map17_irq = 0
        End If
    End If
End Sub
Public Sub map21_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&
            reg8 = Value
            SetupBanks
        Case &H9000&: MirrorXor = Pow2((Value And &H3) + 10)
        Case &HA000&
            regA = Value
            SetupBanks
        Case &HB000&: Select1KVROM Value, 0
        Case &HB001&: Select1KVROM Value, 1
        Case &HB004&: Select1KVROM Value, 1
        Case &HC000&: Select1KVROM Value, 2
        Case &HC001&: Select1KVROM Value, 3
        Case &HC004&: Select1KVROM Value, 3
        Case &HD000&: Select1KVROM Value, 4
        Case &HD001&: Select1KVROM Value, 5
        Case &HD004&: Select1KVROM Value, 5
        Case &HE000&: Select1KVROM Value, 6
        Case &HE001&: Select1KVROM Value, 7
        Case &HE004&: Select1KVROM Value, 7
    End Select
End Sub
Public Sub map23_write(Address As Long, Value As Byte)
    Dim page As Byte
    
    Address = ((Address / 4) And 3) Or Address
    Address = Address And &HF003
    Select Case Address
        Case &H8000& To &H800F&
            PrgSwitch1 = Value
            map4_sync
        Case &H9000& To &H9001&
            Mirroring = (Value And &H3)
            If Mirroring = 1 Then
                Mirroring = 0
            ElseIf Mirroring = 0 Then
                Mirroring = 1
            End If
            DoMirror
        Case &H9002& To &H9003&
            If (Value And &H2) Then swap = True Else swap = False
            map4_sync
        Case &HA000& To &HA00F&
            PrgSwitch2 = Value
            map4_sync
        Case &HB000&
            map23_LoChr(0) = Value
            page = (map23_LoChr(0) And &HF) + ((map23_HiChr(0) And &HF) * 16)
            Call Select1KVROM(page, 0)
        Case &HB001&
            map23_HiChr(0) = Value
            page = (map23_LoChr(0) And &HF) + ((map23_HiChr(0) And &HF) * 16)
            Call Select1KVROM(page, 0)
        Case &HB002&
            map23_LoChr(1) = Value
            page = (map23_LoChr(1) And &HF) + ((map23_HiChr(1) And &HF) * 16)
            Call Select1KVROM(page, 1)
        Case &HB003&
            map23_HiChr(1) = Value
            page = (map23_LoChr(1) And &HF) + ((map23_HiChr(1) And &HF) * 16)
            Call Select1KVROM(page, 1)
        Case &HC000&
            map23_LoChr(2) = Value
            page = (map23_LoChr(2) And &HF) + ((map23_HiChr(2) And &HF) * 16)
            Call Select1KVROM(page, 2)
        Case &HC001&
            map23_HiChr(2) = Value
            page = (map23_LoChr(2) And &HF) + ((map23_HiChr(2) And &HF) * 16)
            Call Select1KVROM(page, 2)
        Case &HC002&
            map23_LoChr(3) = Value
            page = (map23_LoChr(3) And &HF) + ((map23_HiChr(3) And &HF) * 16)
            Call Select1KVROM(page, 3)
        Case &HC003&
            map23_HiChr(3) = Value
            page = (map23_LoChr(3) And &HF) + ((map23_HiChr(3) And &HF) * 16)
            Call Select1KVROM(page, 3)
        Case &HD000&
            map23_LoChr(4) = Value
            page = (map23_LoChr(4) And &HF) + ((map23_HiChr(4) And &HF) * 16)
            Call Select1KVROM(page, 4)
        Case &HD001&
            map23_HiChr(4) = Value
            page = (map23_LoChr(4) And &HF) + ((map23_HiChr(4) And &HF) * 16)
            Call Select1KVROM(page, 4)
        Case &HD002&
            map23_LoChr(5) = Value
            page = (map23_LoChr(5) And &HF) + ((map23_HiChr(5) And &HF) * 16)
            Call Select1KVROM(page, 5)
        Case &HD003&
            map23_HiChr(5) = Value
            page = (map23_LoChr(5) And &HF) + ((map23_HiChr(5) And &HF) * 16)
            Call Select1KVROM(page, 5)
        Case &HE000&
            map23_LoChr(6) = Value
            page = (map23_LoChr(6) And &HF) + ((map23_HiChr(6) And &HF) * 16)
            Call Select1KVROM(page, 6)
        Case &HE001&
            map23_HiChr(6) = Value
            page = (map23_LoChr(6) And &HF) + ((map23_HiChr(6) And &HF) * 16)
            Call Select1KVROM(page, 6)
        Case &HE002&
            map23_LoChr(7) = Value
            page = (map23_LoChr(7) And &HF) + ((map23_HiChr(7) And &HF) * 16)
            Call Select1KVROM(page, 7)
        Case &HE003&
            map23_HiChr(7) = Value
            page = (map23_LoChr(7) And &HF) + ((map23_HiChr(7) And &HF) * 16)
            Call Select1KVROM(page, 7)
        Case &HF000&
            map23_IRQLatchLo = Value
            map23_IRQLatch = (map23_IRQLatchLo And &HF) + ((map23_IRQLatchHi And &HF) * 16)
        Case &HF001&
            map23_IRQLatchHi = Value
            map23_IRQLatch = (map23_IRQLatchLo And &HF) + ((map23_IRQLatchHi And &HF) * 16)
        Case &HF002&
            map23_IRQEnabled = (Value And &H2)
            map23_IRQEnOnWrite = (Value And &H1)
            If (map23_IRQEnabled) Then map23_IRQCounter = map23_IRQLatch
        Case &HF003&
            map23_IRQEnabled = map23_IRQEnOnWrite
    End Select
End Sub
Public Sub map23_irq()
    If map23_IRQEnabled <> 0 Then
        If map23_IRQCounter = &HFF Then
            map23_IRQCounter = map23_IRQLatch
            irq6502
        Else
            map23_IRQCounter = map23_IRQCounter + 1
        End If
    End If
End Sub
Public Sub map79_write(Address As Long, Value As Byte)
    'Need special write 4020 - 5fff
    Select8KVROM Value And &H7
End Sub
Public Sub map83_write(Address As Long, Value As Byte)
    'Cony (Weird Stuff)
    'A LOT better, but still needing some fixes
    Select Case Address
        Case &H8000&, &HB000&, &HB0FF&, &HB1FF&
            map83CHR = LShift(Value And &H30, 4)
            reg8 = Value
            regA = reg8 + 1
            regC = (Value And &H30) Or &HF
            regE = regC + 1
        Case &H8100&
            Value = Value And &H3
            Mirroring = Value
            DoMirror
        Case &H8300&: reg8 = Value
        Case &H8301&: regA = Value
        Case &H8302&: regC = Value
        Case &H8310&: Select1KVROM Value Or map83CHR, 0
        Case &H8311&: Select1KVROM Value Or map83CHR, 1
        Case &H8312&: Select1KVROM Value Or map83CHR, 2
        Case &H8313&: Select1KVROM Value Or map83CHR, 3
        Case &H8314&: Select1KVROM Value Or map83CHR, 4
        Case &H8315&: Select1KVROM Value Or map83CHR, 5
        Case &H8316&: Select1KVROM Value Or map83CHR, 6
        Case &H8317&: Select1KVROM Value Or map83CHR, 7
        Case &H8318&
            reg8 = Value
            regA = reg8 + 1
    End Select
    SetupBanks
End Sub
Public Sub map85_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&
            reg8 = Value
            SetupBanks
        Case &H8008&
            regA = Value
            SetupBanks
        Case &H9000&
            regC = Value
            SetupBanks
        Case &HE000&
            Mirroring = (Value And &H3)
            If Mirroring = 0 Then
                Mirroring = 1
            ElseIf Mirroring = 1 Then
                Mirroring = 0
            End If
            DoMirror
        Case &HA000&
            Call Select1KVROM(Value, 0)
        Case &HA008&
            Call Select1KVROM(Value, 1)
        Case &HB000&
            Call Select1KVROM(Value, 2)
        Case &HB008&
            Call Select1KVROM(Value, 3)
        Case &HC000&
            Call Select1KVROM(Value, 4)
        Case &HC008&
            Call Select1KVROM(Value, 5)
        Case &HD000&
            Call Select1KVROM(Value, 6)
        Case &HD008&
            Call Select1KVROM(Value, 7)
        Case &HE008&
            map85_IRQLatch = Value
        Case &HF000&
            map85_IRQEnabled = (Value And &H2)
            map85_IRQEnOnWrite = (Value And &H1)
            If (map85_IRQEnabled) Then map85_IRQCounter = map85_IRQLatch
        Case &HF008&
            map85_IRQEnabled = map85_IRQEnOnWrite
    End Select
End Sub
Public Sub map85_irq()
    If map85_IRQEnabled <> 0 Then
        If map85_IRQCounter = &HFF Then
            map85_IRQCounter = map85_IRQLatch
            irq6502
        Else
            map85_IRQCounter = map85_IRQCounter + 1
        End If
    End If
End Sub
Public Sub map117_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&: reg8 = Value: SetupBanks
        Case &H8001&: regA = Value: SetupBanks
        Case &H8002&: regC = Value: SetupBanks
        Case &HA000& To &HA007&: Select1KVROM Value, (Address And 7)
        Case &HC001& To &HC003&: Irq_Counter = Value
        Case &HE000&: Irq_Enabled = Value And 1
    End Select
End Sub
Public Function map117_hblank(Scanline) As Byte
    If (Scanline >= 0) And (Scanline <= 239) Then
    If (PPU_Control2 And 8) Or (PPU_Control2 And 16) Then
        If Irq_Enabled Then
            If Irq_Counter = Scanline Then
                Irq_Counter = 0
                irq6502
            End If
        End If
    End If
    End If
End Function
Public Sub map151_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&: reg8 = Value: SetupBanks
        Case &HA000&: regA = Value: SetupBanks
        Case &HC000&: regC = Value: SetupBanks
        Case &HE000&: Select4KVROM Value, 0
        Case &HF000&: Select4KVROM Value, 4
    End Select
End Sub
Public Sub map182_write(Address As Long, Value As Byte)
    Select Case Address And &HF003&
        Case &H8001&
            If Value And &H1 Then Mirroring = 0 Else Mirroring = 1
            DoMirror
        Case &HA000&: Map182Reg = Value And &H7
        Case &HC000&
            Select Case Map182Reg
                Case 0
                    Select1KVROM (Value And &HFE), 0
                    Select1KVROM (Value And &HFE) + 1, 1
                Case 1: Select1KVROM Value, 5
                Case 2
                    Select1KVROM (Value And &HFE), 2
                    Select1KVROM (Value And &HFE) + 1, 3
                Case 3: Select1KVROM Value, 7
                Case 4: reg8 = Value
                Case 5: regA = Value
                Case 6: Select1KVROM Value, 4
                Case 7: Select1KVROM Value, 6
            End Select
        Case &HE003&: 'IRQ
    End Select
    SetupBanks
End Sub
Public Sub map201_write(Address As Long, Value As Byte)
    Dim mBank As Byte
    
    mBank = Address And &H3
    If Address And &H8 Then mBank = 0
    reg8 = mBank
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    Select8KVROM mBank
End Sub
Public Sub map212_write(Address As Long, Value As Byte)
    Dim mPrg(2) As Integer
    
    '10000000 in 1 and other multicarts
    If Address = &H4000& Then
        mPrg(0) = Address And &H7
        mPrg(1) = &HFF
    Else
        mPrg(0) = Address And &H7
        mPrg(1) = Address And &H7
    End If
    
    'Mirroring
    Mirroring = Address And &H8
    DoMirror
    
    If Address And &H7 Then Select8KVROM 1 Else Select8KVROM 0
    
    'Bank select
    If mPrg(1) < &HFF Then
        reg8 = mPrg(0)
        regA = reg8 + 1
        regC = mPrg(1)
        regE = regC + 1
    Else
        reg8 = mPrg(0)
        regA = reg8 + 1
        regC = reg8 + 2
        regE = reg8 + 3
    End If
    SetupBanks
End Sub
Public Sub map250_write(Address As Long, Value As Byte)
    Dim TempAdd As Long
    Dim TempVal As Byte
    
    TempAdd = (Address And &HE000) Or ((Address And &H400) \ &H400)
    TempVal = Address And &HFF
    map4_write TempAdd, TempVal
End Sub
Public Sub map255_write(Address As Long, Value As Byte)
    Dim vPage As Byte
    Dim mBank, mChr, mPrg As Long
    
    'Set bank and prg values
    mPrg = (Address And &HF80) + 7
    mChr = (Address And &H3F)
    mBank = (Address And &H4000) + 14
    
    'Mirroring
    If (Address And &H2000) Then Mirroring = 1: DoMirror Else Mirroring = 0: DoMirror
    
    If Address And &H1000 Then
        If Address And &H40 Then
            reg8 = &H80 * mBank + mPrg * 4 + 2
            regA = &H80 * mBank + mPrg * 4 + 3
            regC = &H80 * mBank + mPrg * 4 + 2
            regE = &H80 * mBank + mPrg * 4 + 3
        Else
            reg8 = &H80 * mBank + mPrg * 4
            regA = &H80 * mBank + mPrg * 4 + 1
            regC = &H80 * mBank + mPrg * 4
            regE = &H80 * mBank + mPrg * 4 + 1
        End If
    Else
        reg8 = &H80 * mBank + mPrg * 4
        regA = reg8 + 1
        regC = reg8 + 2
        regE = reg8 + 3
    End If
    SetupBanks
    For vPage = 0 To 7
        Select1KVROM &H200 * mBank + mChr * 8 + vPage, vPage
    Next vPage
End Sub
Public Sub map1_write(Address As Long, Value As Byte)
    Dim bank_select As Long
    
    If (Value And &H80) Then
        data(0) = data(0) Or &HC
        accumulator = data((Address \ &H2000&) And 3)
        sequence = 5
    Else
        If Value And 1 Then accumulator = accumulator Or Pow2(sequence)
        sequence = sequence + 1
    End If
    
    If (sequence = 5) Then
        data(Address \ &H2000& And 3) = accumulator
        sequence = 0
        accumulator = 0
        
        If (PrgCount = &H20) Then '/* 512k cart */'
            bank_select = (data(1) And &H10) * 2
        Else '/* other carts */'
            bank_select = 0
        End If
        
        If data(0) And 2 Then 'enable panning
            Mirroring = (data(0) And 1) Xor 1
        Else 'disable panning
            Mirroring = 2
        End If
        DoMirror
        Select Case Mirroring
        Case 0
            MirrorXor = &H400
        Case 1
            MirrorXor = &H800
        Case 2
            MirrorXor = 0
        End Select
        
        If (data(0) And 8) = 0 Then 'base boot select $8000?
            reg8 = 4 * (data(3) And 15) + bank_select
            regA = 4 * (data(3) And 15) + bank_select + 1
            regC = 4 * (data(3) And 15) + bank_select + 2
            regE = 4 * (data(3) And 15) + bank_select + 3
            SetupBanks
        ElseIf (data(0) And 4) Then '16k banks
            reg8 = ((data(3) And 15) * 2) + bank_select
            regA = ((data(3) And 15) * 2) + bank_select + 1
            regC = &HFE
            regE = &HFF
            SetupBanks
        Else '32k banks
            reg8 = 0
            regA = 1
            regC = ((data(3) And 15) * 2) + bank_select
            regE = ((data(3) And 15) * 2) + bank_select + 1
            SetupBanks
        End If
        
        If (data(0) And &H10) Then '4k
            Select4KVROM data(1), 0
            Select4KVROM data(2), 1
        Else '8k
            Select8KVROM data(1) \ 2
        End If
    End If
End Sub
Public Sub map2_write(Address As Long, Value As Byte)
    reg8 = (Value * 2)
    regA = reg8 + 1
    SetupBanks
End Sub
Public Sub map3_write(Address As Long, Value As Byte)
    Select8KVROM Value
End Sub
Public Sub map4_write(Address As Long, Value As Byte)
    Select Case Address
        Case &H8000&
            MMC3_Command = Value And &H7
            If Value And &H80 Then MMC3_ChrAddr = &H1000& Else MMC3_ChrAddr = 0
            If Value And &H40 Then swap = 1 Else swap = 0
        Case &H8001&
            Select Case MMC3_Command
                Case 0: Select1KVROM Value, 0: Select1KVROM Value + 1, 1
                Case 1: Select1KVROM Value, 2: Select1KVROM Value + 1, 3
                Case 2: Select1KVROM Value, 4
                Case 3: Select1KVROM Value, 5
                Case 4: Select1KVROM Value, 6
                Case 5: Select1KVROM Value, 7
                Case 6: PrgSwitch1 = Value: map4_sync
                Case 7: PrgSwitch2 = Value: map4_sync
            End Select
        Case &HA000&
            If (Value And &H1) Then Mirroring = 0 Else Mirroring = 1
            DoMirror
        Case &HA001&: If Value Then UsesSRAM = True Else UsesSRAM = False
        Case &HC000&: MMC3_IrqVal = Value
        Case &HC001&: MMC3_TmpVal = Value
        Case &HE000&: MMC3_IrqOn = False: MMC3_IrqVal = MMC3_TmpVal
        Case &HE001&: MMC3_IrqOn = True
    End Select
End Sub
Public Function map4_hblank(Scanline, two As Byte) As Boolean
    If Scanline = 0 Then
        MMC3_IrqVal = MMC3_TmpVal
    ElseIf Scanline > 239 Then
        Exit Function
    ElseIf MMC3_IrqOn And (two And &H18) Then
        MMC3_IrqVal = (MMC3_IrqVal - 1) And &HFF
        If (MMC3_IrqVal = 0) Then
            irq6502
            MMC3_IrqVal = MMC3_TmpVal
        End If
    End If
End Function
Public Sub map7_write(Address As Long, Value As Byte)
    reg8 = 4 * (Value And &HF)
    regA = reg8 + 1
    regC = reg8 + 2
    regE = reg8 + 3
    SetupBanks
    If (Value And &H10&) <> 0 Then Mirroring = 4 Else Mirroring = 2 '4 2
    DoMirror
End Sub
Public Sub map9_write(Address As Long, Value As Byte)
    Static bnk As Long
    Select Case (Address And &HF000&)
        Case &HA000&
            If Mapper = 9 Then
                reg8 = Value
            ElseIf Mapper = 10 Then
                reg8 = Value * 2
                regA = reg8 + 1
            End If
            SetupBanks
        Case &HB000&
            Latch0FD = Value
            If Latch1 = &HFD Then Select4KVROM Value, 0
        Case &HC000&
            Latch0FE = Value
            If Latch1 = &HFE Then Select4KVROM Value, 0
        Case &HD000&
            Latch1FD = Value
            If Latch2 = &HFD Then Select4KVROM Value, 1
        Case &HE000&
            Latch1FE = Value
            If Latch2 = &HFE Then Select4KVROM Value, 1
        Case &HF000&
            If (Value And 1) Then
                Mirroring = 0
            ElseIf (Value And 1) = 0 Then
                Mirroring = 1
            End If
    End Select
End Sub
Public Sub map9_latch(TileNum As Byte, Hi As Boolean)
    If Mapper <> 9 Then Exit Sub
    If (TileNum = &HFD) Then
        If (Hi = False) Then
            Select4KVROM Latch0FD, 0
            Latch1 = &HFD
        ElseIf (Hi = True) Then
            Select4KVROM Latch1FD, 1
            Latch2 = &HFD
        End If
    ElseIf (TileNum = &HFE) Then
        If (Hi = False) Then
            Select4KVROM Latch0FE, 0
            Latch1 = &HFE
        ElseIf (Hi = True) Then
            Select4KVROM Latch1FE, 1
            Latch2 = &HFE
        End If
    End If
End Sub
'Tile based sprite render
Public Sub DrawSprites(OnTop As Boolean)
    If (PPU_Control2 And 16) = 0 Or frmNES.mLayer2.Checked = False Then Exit Sub
    
    Dim SpritePattern As Long 'Integer
    SpritePattern = (PPU_Control1 And &H8) * &H200&
    Dim spr As Long 'Integer
    Dim X1 As Long, Y1 As Long
    Dim Byte1 As Byte, Byte2 As Byte
    Dim Color As Byte
    Dim sa As Long
    Dim h As Long
    Dim i As Long
    Dim X As Long, Y As Long, attrib As Long, tileno As Long, Pal As Long, Aa As Long
    
    If PPU_Control1 And &H20 Then
        h = 16
    Else
        h = 8
    End If
    
    SpriteAddr = 0
    
    For spr = 63 To 0 Step -1
        SpriteAddr = 4 * spr
        attrib = SpriteRAM(SpriteAddr + 2)
        If (attrib And 32) = 0 Xor OnTop Then
            X = SpriteRAM(SpriteAddr + 3)
            Y = SpriteRAM(SpriteAddr)
            If Y < 239 And X < 248 Then
                tileno = SpriteRAM(SpriteAddr + 1)
                If h = 16 Then
                    SpritePattern = (tileno And 1) * &H1000
                    tileno = tileno Xor (tileno And 1)
                End If
                sa = SpritePattern + 16 * tileno
                i = Y * 256& + X + 256
                Pal = 16 + (attrib And 3) * 4
                
                If attrib And 128 Then
                    If attrib And 64 Then
                        For Y1 = h - 1 To 0 Step -1
                            If Y1 >= 8 Then
                                Byte1 = VRAM(sa + 8 + Y1)
                                Byte2 = VRAM(sa + 16 + Y1)
                            Else
                                Byte1 = VRAM(sa + Y1)
                                Byte2 = VRAM(sa + Y1 + 8)
                            End If
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 0 To 7
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i + X1) = Color Or Pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next Y1
                    Else
                        i = i + 7
                        For Y1 = h - 1 To 0 Step -1
                            If Y1 >= 8 Then
                                Byte1 = VRAM(sa + 8 + Y1)
                                Byte2 = VRAM(sa + 16 + Y1)
                            Else
                                Byte1 = VRAM(sa + Y1)
                                Byte2 = VRAM(sa + Y1 + 8)
                            End If
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 7 To 0 Step -1
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i - X1) = Color Or Pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next Y1
                    End If
                Else
                    If attrib And 64 Then
                        For Y1 = 0 To h - 1
                            If Y1 = 8 Then sa = sa + 8
                            Byte1 = VRAM(sa + Y1)
                            Byte2 = VRAM(sa + Y1 + 8)
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 0 To 7
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i + X1) = Color Or Pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next Y1
                    Else
                        i = i + 7
                        For Y1 = 0 To h - 1
                            If Y1 = 8 Then sa = sa + 8
                            Byte1 = VRAM(sa + Y1)
                            Byte2 = VRAM(sa + Y1 + 8)
                            Aa = Byte1 * 2048& + Byte2 * 8
                            For X1 = 7 To 0 Step -1
                                Color = tLook(X1 + Aa)
                                If Color Then vBuffer(i - X1) = Color Or Pal
                            Next X1
                            i = i + 256
                            If i >= 256& * 240 Then Exit For
                        Next Y1
                    End If
                End If
            End If
        End If
    Next spr
End Sub
Public Sub CheckJoy()
    If Gamepad1 = 1 Then
        ' Nes control 1 as Gamepad 1
        If Js.JoyY < 16000 Then JoyDown 4 Else JoyUp 4
        If Js.JoyY > 48000 Then JoyDown 5 Else JoyUp 5
        If Js.JoyX < 16000 Then JoyDown 6 Else JoyUp 6
        If Js.JoyX > 48000 Then JoyDown 7 Else JoyUp 7
        If ReturnBtnNum(Js.CurButton) = pad_ButA Then JoyDown 0 Else JoyUp 0
        If ReturnBtnNum(Js.CurButton) = pad_ButB Then JoyDown 1 Else JoyUp 1
        If ReturnBtnNum(Js.CurButton) = pad_ButSel Then JoyDown 2 Else JoyUp 2
        If ReturnBtnNum(Js.CurButton) = pad_ButSta Then JoyDown 3 Else JoyUp 3
    ElseIf Gamepad1 = 2 Then
        ' Nes control 1 as Gamepad 2
        If Js.Joy2Y < 16000 Then JoyDown 4 Else JoyUp 4
        If Js.Joy2Y > 48000 Then JoyDown 5 Else JoyUp 5
        If Js.Joy2X < 16000 Then JoyDown 6 Else JoyUp 6
        If Js.Joy2X > 48000 Then JoyDown 7 Else JoyUp 7
        If ReturnBtnNum(Js.Joy2CurButton) = pad_ButA Then JoyDown 0 Else JoyUp 0
        If ReturnBtnNum(Js.Joy2CurButton) = pad_ButB Then JoyDown 1 Else JoyUp 1
        If ReturnBtnNum(Js.Joy2CurButton) = pad_ButSel Then JoyDown 2 Else JoyUp 2
        If ReturnBtnNum(Js.Joy2CurButton) = pad_ButSta Then JoyDown 3 Else JoyUp 3
    End If
    
    If Gamepad2 = 1 Then
        ' Nes control 2 as Gamepad 1
        If Js.JoyY < 16000 Then JoyDown 4, 1 Else JoyUp 4, 1
        If Js.JoyY > 48000 Then JoyDown 5, 1 Else JoyUp 5, 1
        If Js.JoyX < 16000 Then JoyDown 6, 1 Else JoyUp 6, 1
        If Js.JoyX > 48000 Then JoyDown 7, 1 Else JoyUp 7, 1
        If ReturnBtnNum(Js.CurButton) = pad2_ButA Then JoyDown 0, 1 Else JoyUp 0, 1
        If ReturnBtnNum(Js.CurButton) = pad2_ButB Then JoyDown 1, 1 Else JoyUp 1, 1
        If ReturnBtnNum(Js.CurButton) = pad2_ButSel Then JoyDown 2, 1 Else JoyUp 2, 1
        If ReturnBtnNum(Js.CurButton) = pad2_ButSta Then JoyDown 3, 1 Else JoyUp 3, 1
    ElseIf Gamepad2 = 2 Then
        ' Nes control 2 as Gamepad 2
        If Js.Joy2Y < 16000 Then JoyDown 4, 1 Else JoyUp 4, 1
        If Js.Joy2Y > 48000 Then JoyDown 5, 1 Else JoyUp 5, 1
        If Js.Joy2X < 16000 Then JoyDown 6, 1 Else JoyUp 6, 1
        If Js.Joy2X > 48000 Then JoyDown 7, 1 Else JoyUp 7, 1
        If ReturnBtnNum(Js.Joy2CurButton) = pad2_ButA Then JoyDown 0, 1 Else JoyUp 0, 1
        If ReturnBtnNum(Js.Joy2CurButton) = pad2_ButB Then JoyDown 1, 1 Else JoyUp 1, 1
        If ReturnBtnNum(Js.Joy2CurButton) = pad2_ButSel Then JoyDown 2, 1 Else JoyUp 2, 1
        If ReturnBtnNum(Js.Joy2CurButton) = pad2_ButSta Then JoyDown 3, 1 Else JoyUp 3, 1
    End If
End Sub
Private Function ReturnBtnNum(JoyBut As Integer) As Integer
    Select Case JoyBut
        Case 1: ReturnBtnNum = 1
        Case 2: ReturnBtnNum = 2
        Case 4: ReturnBtnNum = 3
        Case 8: ReturnBtnNum = 4
        Case 16: ReturnBtnNum = 5
        Case 32: ReturnBtnNum = 6
        Case 64: ReturnBtnNum = 7
        Case 128: ReturnBtnNum = 8
        Case 256: ReturnBtnNum = 9
        Case 512: ReturnBtnNum = 10
        Case 1024: ReturnBtnNum = 11
        Case 2048: ReturnBtnNum = 12
    End Select
End Function
Public Sub JoyDown(ByVal PadNum As Integer, Optional GP As Integer = 0)
    If GP = 0 Then Joypad1(PadNum) = &H41 Else If GP = 1 Then Joypad2(PadNum) = &H41
End Sub
Public Sub JoyUp(ByVal PadNum As Integer, Optional GP As Integer = 0)
    If GP = 0 Then Joypad1(PadNum) = &H40 Else If GP = 1 Then Joypad2(PadNum) = &H40
End Sub
'*****************
'*  GAME GENIE   *
'*****************
Public Function BinToHex(BinNum As String) As String
    Dim BinLen As Integer, i As Integer, hexNum As Variant
    On Error GoTo ErrH
    BinLen = Len(BinNum)
    For i = BinLen To 1 Step -1
        If Asc(Mid(BinNum, i, 1)) < 48 Or Asc(Mid(BinNum, i, 1)) > 49 Then hexNum = vbNullString
        If Mid(BinNum, i, 1) And 1 Then hexNum = hexNum + 2 ^ Abs(i - BinLen)
    Next i
    BinToHex = Hex(hexNum)
ErrH:
End Function
Public Function ggAddY(ggCode As String) As Long
    On Error GoTo ErrH
    Dim ggVal As String
    ggVal = ggDecode(ggCode)
    ggCode = Right(Left(ggVal, 14), 1)
    ggCode = ggCode & Right(Left(ggVal, 15), 1)
    ggCode = ggCode & Right(Left(ggVal, 16), 1)
    ggCode = ggCode & Right(Left(ggVal, 17), 1)
    ggCode = ggCode & Right(Left(ggVal, 22), 1)
    ggCode = ggCode & Right(Left(ggVal, 23), 1)
    ggCode = ggCode & Right(Left(ggVal, 24), 1)
    ggCode = ggCode & Right(Left(ggVal, 5), 1)
    ggCode = ggCode & Right(Left(ggVal, 10), 1)
    ggCode = ggCode & Right(Left(ggVal, 11), 1)
    ggCode = ggCode & Right(Left(ggVal, 12), 1)
    ggCode = ggCode & Right(Left(ggVal, 13), 1)
    ggCode = ggCode & Right(Left(ggVal, 18), 1)
    ggCode = ggCode & Right(Left(ggVal, 19), 1)
    ggCode = ggCode & Right(Left(ggVal, 20), 1)
    ggAddY = CLng("&h" & BinToHex(ggCode))
ErrH:
End Function
Public Function ggVal(ggCode As String) As Long
    On Error GoTo ErrH
    Dim ggVral As String
    ggVral = ggDecode(ggCode)
    ggCode = Left(ggVral, 1)
    ggCode = ggCode & Right(Left(ggVral, 6), 1)
    ggCode = ggCode & Right(Left(ggVral, 7), 1)
    ggCode = ggCode & Right(Left(ggVral, 8), 1)
    ggCode = ggCode & Right(Left(ggVral, 21), 1)
    ggCode = ggCode & Right(Left(ggVral, 2), 1)
    ggCode = ggCode & Right(Left(ggVral, 3), 1)
    ggCode = ggCode & Right(Left(ggVral, 4), 1)
    ggVal = CLng("&h" & BinToHex(ggCode))
ErrH:
End Function
Public Function ggDecode(ggCode As String) As String
    On Error GoTo ErrH
    ggCode = UCase(ggCode)
    ggCode = Replace(ggCode, "A", "0000") 'Set #1
    ggCode = Replace(ggCode, "P", "0001")
    ggCode = Replace(ggCode, "Z", "0010")
    ggCode = Replace(ggCode, "L", "0011")
    ggCode = Replace(ggCode, "G", "0100")
    ggCode = Replace(ggCode, "I", "0101")
    ggCode = Replace(ggCode, "T", "0110")
    ggCode = Replace(ggCode, "Y", "0111")
    ggCode = Replace(ggCode, "E", "1000") 'Set #2
    ggCode = Replace(ggCode, "O", "1001")
    ggCode = Replace(ggCode, "X", "1010")
    ggCode = Replace(ggCode, "U", "1011")
    ggCode = Replace(ggCode, "K", "1100")
    ggCode = Replace(ggCode, "S", "1101")
    ggCode = Replace(ggCode, "V", "1110")
    ggCode = Replace(ggCode, "N", "1111")
    ggDecode = ggCode
ErrH:
End Function
Public Function FixHex(hCode As String) As String
    If Len(hCode) = 1 Then FixHex = "0" & hCode Else FixHex = hCode
End Function
Public Function FixAddress(hCode As String) As String
    If Len(hCode) = 1 Then FixAddress = "000" & hCode: Exit Function
    If Len(hCode) = 2 Then FixAddress = "00" & hCode: Exit Function
    If Len(hCode) = 3 Then FixAddress = "0" & hCode: Exit Function
    FixAddress = hCode
End Function
'*****************
'*   COLOR PAL   *
'*****************
Sub SetVSPal()
    Dim n As Integer
    Dim VSPal As Variant
    Dim TmpPal(63) As Long
    
    If InStr(LCase(RomName), "castlevania") Then
        VSPal = Array("0f", "27", "18", "3f", "3f", "25", "3f", "34", "16", "13", "3f", "3f", "20", "23", "3f", "0b", "3f", "3f", "06", "3f", "1b", "3f", "3f", "22", "3f", "24", "3f", "3f", "32", "3f", "3f", "03", "3f", "37", "26", "33", "11", "3f", "10", "3f", "14", "3f", "00", "09", "12", "0f", "3f", "30", "3f", "3f", "2a", "17", "0c", "01", "15", "19", "3f", "3f", "07", "37", "3f", "05", "3f", "3f")
    ElseIf InStr(LCase(RomName), "mario") Then
        VSPal = Array("18", "3f", "1c", "3f", "3f", "3f", "0b", "17", "10", "3f", "14", "3f", "36", "37", "1a", "3f", "25", "3f", "12", "3f", "0f", "3f", "3f", "3f", "3f", "3f", "22", "19", "3f", "3f", "3a", "21", "3f", "3f", "07", "3f", "3f", "3f", "00", "15", "0c", "3f", "3f", "3f", "3f", "3f", "3f", "3f", "3f", "3f", "07", "16", "3f", "3f", "30", "3c", "3f", "27", "3f", "3f", "29", "3f", "1b", "09")
    ElseIf InStr(LCase(RomName), "goonies") Then
        VSPal = Array("0f", "3f", "3f", "10", "3f", "30", "31", "3f", "01", "0f", "36", "3f", "3f", "3f", "3f", "3c", "3f", "3f", "3f", "12", "19", "3f", "17", "3f", "00", "3f", "3f", "02", "16", "3f", "3f", "3f", "3f", "3f", "3f", "37", "3f", "27", "26", "20", "3f", "04", "22", "3f", "11", "3f", "3f", "3f", "2c", "3f", "3f", "3f", "07", "2a", "3f", "3f", "3f", "3f", "3f", "38", "13", "3f", "3f", "0c")
    ElseIf InStr(LCase(RomName), "climber") Then
        VSPal = Array("18", "3f", "1c", "3f", "3f", "3f", "01", "17", "10", "3f", "2a", "3f", "36", "37", "1a", "39", "25", "3f", "12", "3f", "0f", "3f", "3f", "26", "3f", "3f", "22", "19", "3f", "0f", "3a", "21", "3f", "0a", "07", "06", "13", "3f", "00", "15", "0c", "3f", "11", "3f", "3f", "38", "3f", "3f", "3f", "3f", "07", "16", "3f", "3f", "30", "3c", "0f", "27", "3f", "31", "29", "3f", "11", "09")
    ElseIf InStr(LCase(RomName), "excite") Then
        VSPal = Array("3f", "3f", "3f", "3f", "1a", "30", "3c", "09", "0f", "0f", "3f", "0f", "3f", "3f", "3f", "30", "32", "1c", "3f", "12", "3f", "18", "17", "3f", "0c", "3f", "3f", "02", "16", "3f", "3f", "3f", "3f", "3f", "0f", "37", "3f", "28", "27", "3f", "29", "3f", "21", "3f", "11", "3f", "0f", "3f", "31", "3f", "3f", "3f", "0f", "2a", "28", "3f", "3f", "3f", "3f", "3f", "13", "3f", "3f", "3f")
    ElseIf InStr(LCase(RomName), "alley") Or InStr(LCase(RomName), "gradius") Then
        VSPal = Array("35", "3f", "16", "22", "1c", "3f", "3f", "15", "3f", "00", "27", "05", "04", "3f", "3f", "30", "21", "3f", "3f", "29", "3c", "3f", "36", "12", "3f", "2b", "3f", "3f", "3f", "3f", "3f", "01", "3f", "31", "3f", "2a", "2c", "0c", "3f", "3f", "3f", "07", "34", "06", "3f", "25", "26", "0f", "3f", "19", "10", "3f", "3f", "3f", "3f", "17", "3f", "11", "3f", "3f", "3f", "25", "18", "3f")
    Else
        Exit Sub
    End If
    
    For n = 0 To 63
        If VSPal(n) = "00" And InStr(LCase(RomName), "castlevania") Then VSPal(n) = "0f"
        TmpPal(n) = Pal(Val("&H" & VSPal(n)))
    Next n
    For n = 0 To 63
        Pal(n) = TmpPal(n)
    Next n
End Sub
Function Rgb2(ByVal B As Long, ByVal G As Long, ByVal R As Long) As Long
    'DF: had to reverse colors in palette for it to look right with new gfx.
    'later changed for 16bit color, then back to 32bit
    Rgb2 = RGB(R, G, B)
End Function
Function Rgb16(ByVal B As Long, ByVal G As Long, ByVal R As Long) As Long
    Dim c As Long
    c = (R \ 8) + (G \ 4) * 32 + (B \ 8) * 2048
    If c > 32767 Then c = c - 65536
    Rgb16 = c
End Function
Function Rgb15(ByVal B As Long, ByVal G As Long, ByVal R As Long) As Long
    Rgb15 = (R \ 8) + (G \ 8) * 32 + (B \ 8) * 1024
End Function
Public Sub LoadPal(File As String)
    Dim n As Long
    Dim R As Byte, G As Byte, B As Byte
    Dim FileNum As Integer
    FileNum = FreeFile
    
    If Dir(App.Path & "\" & File) = vbNullString Then Exit Sub
    Open App.Path & "\" & File For Binary As #FileNum
        For n = 0 To 63
            Get #FileNum, , R
            Get #FileNum, , G
            Get #FileNum, , B
            SetPalVal R, G, B, n
        Next n
    Close #FileNum
End Sub
Public Function SetPalVal(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal n As Long)
    Pal(n) = Rgb2(R, G, B)
    Pal16(n) = Rgb16(R, G, B)
    Pal15(n) = Rgb15(R, G, B)
    Pal(n + 64) = Pal(n)
    Pal(n + 128) = Pal(n)
    Pal(n + 192) = Pal(n)
    Pal16(n + 64) = Pal16(n)
    Pal16(n + 128) = Pal16(n)
    Pal16(n + 192) = Pal16(n)
    Pal15(n + 64) = Pal15(n)
    Pal15(n + 128) = Pal15(n)
    Pal15(n + 192) = Pal15(n)
End Function
Public Function NewPal(Optional Tint As Double = 0.5, Optional Hue As Integer = 332)
    Const Pi = 3.14159265
    Dim Theta As Double
    
    Dim R As Double
    Dim G As Double
    Dim B As Double
    
    Dim Br1(3) As Double
    Dim Br2(3) As Double
    Dim Br3(3) As Double
    
    Dim Y As Double
    Dim s As Double
    
    Dim Cols(16) As Integer
    Dim X, z, j As Integer
    
    Cols(2) = 240
    Cols(3) = 210
    Cols(4) = 180
    Cols(5) = 150
    Cols(6) = 120
    Cols(7) = 90
    Cols(8) = 60
    Cols(9) = 30
    Cols(11) = 330
    Cols(12) = 300
    Cols(13) = 270
    
    Br1(0) = 0.5: Br1(1) = 0.75: Br1(2) = 1!: Br1(3) = 1!
    Br2(0) = 0.29: Br2(1) = 0.45: Br2(2) = 0.73: Br2(3) = 0.9
    Br3(0) = 0: Br3(1) = 0.24: Br3(2) = 0.47: Br3(3) = 0.77
    
    For X = 0 To 3
        For z = 1 To 16                       'two loops
            s = Tint                              'grab tint
            Y = Br2(X)                            'grab default luminance
            If z = 1 Then s = 0: Y = Br1(X)       'is it colour XDh? if so, get luma
            If z = 14 Then s = 0: Y = Br3(X)      'is it colour X0h? if so, get luma
            If z = 15 Then Y = 0: s = 0           'is it colour XEh? if so, set to black
            If z = 16 Then Y = 0: s = 0           'is it colour XFh? if so, set to black
                
            Theta = Pi * ((Cols(z) + Hue) / 180)
                                
            R = Y + s * Sin(Theta)
            G = Y - (27 / 53) * s * Sin(Theta) + (10 / 53) * s * Cos(Theta)
            B = Y - s * Cos(Theta)
                
            R = R * 256
            B = B * 256
            G = G * 256
                
            If R > 255 Then R = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            If R < 0 Then R = 0
            If G < 0 Then G = 0
            If B < 0 Then B = 0
                
            SetPalVal Int(R), Int(G), Int(B), j
            j = j + 1
        Next
    Next
End Function
Function LShift(ByVal W As Integer, ByVal c As Integer) As Integer
    LShift = W * (2 ^ c)
End Function
Function RShift(ByVal W As Integer, ByVal c As Integer) As Integer
    RShift = W \ (2 ^ c)
End Function
