Attribute VB_Name = "Misc"
'Buffered IO. faster.
Private ReadBuffer(4095) As Byte
Private Writebuffer(4095) As Byte
Private ReadPtr As Long
Private WritePtr As Long
Private ReadSize As Long

Private RunLength As Byte
Private RunChar As Byte
Private NextChar As Byte
Private Temp As Byte
Public Gif As IPictureDisp
Private ind As Long

Public PalName As String
Public StrMsg As String
Private Function ReadChar() As Byte
    If ReadPtr = 0 Then
        If ReadSize >= 4096 Then
            Get #1, , ReadBuffer
        Else
            Dim B() As Byte, i As Long
            ReDim B(ReadSize - 1)
            Get #1, , B
            For i = 0 To ReadSize - 1
                ReadBuffer(i) = B(i)
            Next i
        End If
    End If
    ReadChar = ReadBuffer(ReadPtr)
    ReadPtr = (ReadPtr + 1) And 4095
    ReadSize = ReadSize - 1
End Function
Private Sub WriteChar(c As Byte)
    Writebuffer(WritePtr) = c
    WritePtr = (WritePtr + 1) And 4095
    If WritePtr = 0 Then Put #2, , Writebuffer
End Sub
Private Sub PreClose()
    If WritePtr > 0 Then
        Dim B() As Byte, i As Long
        ReDim B(WritePtr - 1)
        For i = 0 To WritePtr - 1: B(i) = Writebuffer(i): Next i
        Put #2, , B
    End If
End Sub
Public Sub Delete(f As String)
    On Error Resume Next
    Kill f
End Sub
Private Sub ScanRun()
    RunChar = NextChar
    RunLength = 0
    Do
        RunLength = RunLength + 1
        NextChar = ReadChar
    Loop Until NextChar <> RunChar Or RunLength = 255 Or ReadSize = 0
End Sub
Private Sub WriteRun()
    Dim i As Long
    For i = 1 To RunLength
    WriteChar RunChar
    Next i
End Sub
Private Sub EncodeRun()
    Dim i As Long
    If RunLength > 3 Then
        Temp = 207
        WriteChar Temp
        WriteChar RunLength
        WriteChar RunChar
    Else
        For i = 1 To RunLength
            WriteChar RunChar
            If RunChar = 207 Then
                Temp = 0
                WriteChar Temp
            End If
        Next i
    End If
End Sub
Private Sub DecodeRun()
    RunChar = ReadChar
    If RunChar = 207 Then
        RunLength = ReadChar
        If RunLength > 0 Then
            RunChar = ReadChar
        Else
            RunLength = 1
        End If
    Else
        RunLength = 1
    End If
End Sub
'very simple RLE file compression
Public Sub RLECompress(InFile As String, OutFile As String)
    Delete OutFile
    Open InFile For Binary As #1
    Open OutFile For Binary As #2
    ReadSize = LOF(1)
    ReadPtr = 0
    WritePtr = 0
    
    Get #1, , NextChar
    While ReadSize > 0
        ScanRun
        If ReadSize = 0 Then
            If NextChar = RunChar And RunLength < 255 Then
                RunLength = RunLength + 1
                EncodeRun
            Else
                EncodeRun
                RunLength = 1
                RunChar = NextChar
                EncodeRun
            End If
        Else
            EncodeRun
        End If
    Wend
    PreClose
    Close
End Sub
Public Sub RLEDecompress(InFile As String, OutFile As String)
    Delete OutFile
    Open InFile For Binary As #1
    Open OutFile For Binary As #2
    ReadSize = LOF(1)
    ReadPtr = 0
    WritePtr = 0
    
    While ReadSize > 0
        DecodeRun
        WriteRun
    Wend
    PreClose
    Close
End Sub
Public Sub SaveState(Index As Long)
    ' 12/3/01 - standardized for mapper support.
    ChDir App.Path & "\States\"
    Dim f As String
    f = RomName & ".sv" & CStr(Index)
    Open App.Path & "\States\nessave.tmp" For Binary As #1
        Put #1, , a
        Put #1, , X
        Put #1, , Y
        Put #1, , PC
        Put #1, , SavePC
        Put #1, , P
        Put #1, , s
        Put #1, , Value
        Put #1, , Sum
        Put #1, , FirstRead
        Put #1, , Joypad1
        Put #1, , Joypad1_Count
        Put #1, , Joypad2
        Put #1, , Joypad2_Count
        Put #1, , Mirroring
        Put #1, , Mapper
        Put #1, , Trainer
        Put #1, , MirrorXor
        Put #1, , HScroll
        Put #1, , VScroll
        Put #1, , bank_regs
        Put #1, , PPU_Control1
        Put #1, , PPU_Control2
        Put #1, , PPU_Status
        Put #1, , SpriteAddress
        Put #1, , PPUAddressHi
        Put #1, , PPUAddress
        Put #1, , PPU_AddressIsHi
        Put #1, , PatternTable
        Put #1, , NameTable
        Put #1, , reg8
        Put #1, , regA
        Put #1, , regC
        Put #1, , regE
        Put #1, , CurrentLine
        Put #1, , Bank0
        Put #1, , Bank6
        Put #1, , VRAM
        Put #1, , SpriteRAM
        Select Case Mapper
            Case 0, 2, 3, 5, 6, 7, 8, 11, 66, 68, 71, 78, 91: 'nothing to do here
            Case 1
                Put #1, , data
                Put #1, , accumulator
                Put #1, , sequence
            Case 4
                Put #1, , MMC3_Command
                Put #1, , MMC3_PrgAddr
                Put #1, , MMC3_ChrAddr
                Put #1, , MMC3_IrqVal
                Put #1, , MMC3_TmpVal
                Put #1, , MMC3_IrqOn
                Put #1, , swap
                Put #1, , PrgSwitch1
                Put #1, , PrgSwitch2
            Case 9, 10
                Put #1, , Latch0FD
                Put #1, , Latch0FE
                Put #1, , Latch1FD
                Put #1, , Latch1FE
            Case 13
                Put #1, , latch13
            Case 15
                Put #1, , map15_BankAddr
                Put #1, , map15_SwapReg
            Case 16
                Put #1, , TmpLatch
                Put #1, , MMC16_IrqOn
                Put #1, , MMC16_Irq
            Case 17
                Put #1, , map17_irqon
                Put #1, , map17_irq
            Case 19
                Put #1, , TmpLatch
                Put #1, , MIRQOn
                Put #1, , MMC19_IRQCount
            Case 24
                Put #1, , map24_IRQCounter
                Put #1, , map24_IRQEnabled
                Put #1, , map24_IRQLatch
                Put #1, , map24_IRQEnOnWrite
            Case 32
                Put #1, , MMC32_Switch
            Case 40
                Put #1, , Mapper40_IRQEnabled
                Put #1, , Mapper40_IRQCounter
            Case 64
                Put #1, , Cmd
                Put #1, , Prg
                Put #1, , Chr1
            Case 69
                Put #1, , reg8000
        End Select
        Put #1, , nt
    Close #1
    RLECompress "nessave.tmp", f
    Delete "nessave.tmp"
    ExibeMsg "Jogo salvo"
End Sub
Public Sub LoadState(Index As Long)
    ChDir App.Path & "\States\"
    Dim f As String
    f = RomName & ".sv" & CStr(Index)
    If Dir$(App.Path & "\States\" & f) = vbNullString Then Exit Sub
    RLEDecompress f, "nesload.tmp"
    Open App.Path & "\States\nesload.tmp" For Binary As #1
        Get #1, , a
        Get #1, , X
        Get #1, , Y
        Get #1, , PC
        Get #1, , SavePC
        Get #1, , P
        Get #1, , s
        Get #1, , Value
        Get #1, , Sum
        Get #1, , FirstRead
        Get #1, , Joypad1
        Get #1, , Joypad1_Count
        Get #1, , Joypad2
        Get #1, , Joypad2_Count
        Get #1, , Mirroring
        Get #1, , Mapper
        Get #1, , Trainer
        Get #1, , MirrorXor
        Get #1, , HScroll
        Get #1, , VScroll
        Get #1, , bank_regs
        Get #1, , PPU_Control1
        Get #1, , PPU_Control2
        Get #1, , PPU_Status
        Get #1, , SpriteAddress
        Get #1, , PPUAddressHi
        Get #1, , PPUAddress
        Get #1, , PPU_AddressIsHi
        Get #1, , PatternTable
        Get #1, , NameTable
        Get #1, , reg8
        Get #1, , regA
        Get #1, , regC
        Get #1, , regE
        Get #1, , CurrentLine
        Get #1, , Bank0
        Get #1, , Bank6
        Get #1, , VRAM
        Get #1, , SpriteRAM
        Select Case Mapper
            Case 0: 'nothing to do here
            Case 1
                Get #1, , data
                Get #1, , accumulator
                Get #1, , sequence
            Case 4
                Get #1, , MMC3_Command
                Get #1, , MMC3_PrgAddr
                Get #1, , MMC3_ChrAddr
                Get #1, , MMC3_IrqVal
                Get #1, , MMC3_TmpVal
                Get #1, , MMC3_IrqOn
                Get #1, , swap
                Get #1, , PrgSwitch1
                Get #1, , PrgSwitch2
            Case 9, 10
                Get #1, , Latch0FD
                Get #1, , Latch0FE
                Get #1, , Latch1FD
                Get #1, , Latch1FE
            Case 13
                Get #1, , latch13
            Case 15
                Get #1, , map15_BankAddr
                Get #1, , map15_SwapReg
            Case 16
                Get #1, , TmpLatch
                Get #1, , MMC16_IrqOn
                Get #1, , MMC16_Irq
            Case 17
                Get #1, , map17_irqon
                Get #1, , map17_irq
            Case 19
                Get #1, , TmpLatch
                Get #1, , MIRQOn
                Get #1, , MMC19_IRQCount
            Case 24
                Get #1, , map24_IRQCounter
                Get #1, , map24_IRQEnabled
                Get #1, , map24_IRQLatch
                Get #1, , map24_IRQEnOnWrite
            Case 32
                Get #1, , MMC32_Switch
            Case 40
                Get #1, , Mapper40_IRQEnabled
                Get #1, , Mapper40_IRQCounter
            Case 64
                Get #1, , Cmd
                Get #1, , Prg
                Get #1, , Chr1
            Case 69
                Get #1, , reg8000
        End Select
        Get #1, , nt
    Close #1
    SetupBanks
    Delete "nesload.tmp"
    ExibeMsg "Jogo carregado"
End Sub
Public Sub PlayMovie(Index As Integer)
    ' 12/3/01 - standardized for mapper support.
    ChDir App.Path & "\Movies\"
    Dim f As String
    f = RomName & ".mv" & CStr(Index)
    Open f For Binary As #1
        Get #1, , a
        Get #1, , X
        Get #1, , Y
        Get #1, , PC
        Get #1, , SavePC
        Get #1, , P
        Get #1, , s
        Get #1, , Value
        Get #1, , Sum
        Get #1, , FirstRead
        Get #1, , Joypad1
        Get #1, , Joypad1_Count
        Get #1, , Joypad2
        Get #1, , Joypad2_Count
        Get #1, , Mirroring
        Get #1, , Mapper
        Get #1, , Trainer
        Get #1, , MirrorXor
        Get #1, , HScroll
        Get #1, , VScroll
        Get #1, , bank_regs
        Get #1, , PPU_Control1
        Get #1, , PPU_Control2
        Get #1, , PPU_Status
        Get #1, , SpriteAddress
        Get #1, , PPUAddressHi
        Get #1, , PPUAddress
        Get #1, , PPU_AddressIsHi
        Get #1, , PatternTable
        Get #1, , NameTable
        Get #1, , reg8
        Get #1, , regA
        Get #1, , regC
        Get #1, , regE
        Get #1, , CurrentLine
        Get #1, , Bank0
        Get #1, , Bank6
        Get #1, , VRAM
        Get #1, , SpriteRAM
        
        Select Case Mapper
            Case 0: 'nothing to do here
            Case 1
                Get #1, , data
                Get #1, , accumulator
                Get #1, , sequence
            Case 4
                Get #1, , MMC3_Command
                Get #1, , MMC3_PrgAddr
                Get #1, , MMC3_ChrAddr
                Get #1, , MMC3_IrqVal
                Get #1, , MMC3_TmpVal
                Get #1, , MMC3_IrqOn
                Get #1, , swap
                Get #1, , PrgSwitch1
                Get #1, , PrgSwitch2
            Case 9, 10
                Get #1, , Latch0FD
                Get #1, , Latch0FE
                Get #1, , Latch1FD
                Get #1, , Latch1FE
            Case 13
                Get #1, , latch13
            Case 15
                Get #1, , map15_BankAddr
                Get #1, , map15_SwapReg
            Case 16
                Get #1, , TmpLatch
                Get #1, , MMC16_IrqOn
                Get #1, , MMC16_Irq
            Case 17
                Get #1, , map17_irqon
                Get #1, , map17_irq
            Case 19
                Get #1, , TmpLatch
                Get #1, , MIRQOn
                Get #1, , MMC19_IRQCount
            Case 24
                Get #1, , map24_IRQCounter
                Get #1, , map24_IRQEnabled
                Get #1, , map24_IRQLatch
                Get #1, , map24_IRQEnOnWrite
            Case 32
                Get #1, , MMC32_Switch
            Case 40
                Get #1, , Mapper40_IRQEnabled
                Get #1, , Mapper40_IRQCounter
            Case 64
                Get #1, , Cmd
                Get #1, , Prg
                Get #1, , Chr1
        End Select
        Get #1, , nt
        Playing = True
    ExibeMsg "Tocando filme"
End Sub
Public Sub StopPlaying()
    Playing = False
    Close #1
End Sub
Public Sub RecordMovie(Index As Long)
    ChDir App.Path & "\Movies\"
    Dim f As String
    f = RomName & ".mv" & CStr(Index)
    Open f For Binary As #1
        Put #1, , a
        Put #1, , X
        Put #1, , Y
        Put #1, , PC
        Put #1, , SavePC
        Put #1, , P
        Put #1, , s
        Put #1, , Value
        Put #1, , Sum
        Put #1, , FirstRead
        Put #1, , Joypad1
        Put #1, , Joypad1_Count
        Put #1, , Joypad2
        Put #1, , Joypad2_Count
        Put #1, , Mirroring
        Put #1, , Mapper
        Put #1, , Trainer
        Put #1, , MirrorXor
        Put #1, , HScroll
        Put #1, , VScroll
        Put #1, , bank_regs
        Put #1, , PPU_Control1
        Put #1, , PPU_Control2
        Put #1, , PPU_Status
        Put #1, , SpriteAddress
        Put #1, , PPUAddressHi
        Put #1, , PPUAddress
        Put #1, , PPU_AddressIsHi
        Put #1, , PatternTable
        Put #1, , NameTable
        Put #1, , reg8
        Put #1, , regA
        Put #1, , regC
        Put #1, , regE
        Put #1, , CurrentLine
        Put #1, , Bank0
        Put #1, , Bank6
        Put #1, , VRAM
        Put #1, , SpriteRAM
        Select Case Mapper
            Case 0: 'nothing to do here
            Case 1
                Put #1, , data
                Put #1, , accumulator
                Put #1, , sequence
            Case 4
                Put #1, , MMC3_Command
                Put #1, , MMC3_PrgAddr
                Put #1, , MMC3_ChrAddr
                Put #1, , MMC3_IrqVal
                Put #1, , MMC3_TmpVal
                Put #1, , MMC3_IrqOn
                Put #1, , swap
                Put #1, , PrgSwitch1
                Put #1, , PrgSwitch2
            Case 9, 10
                Put #1, , Latch0FD
                Put #1, , Latch0FE
                Put #1, , Latch1FD
                Put #1, , Latch1FE
            Case 13
                Put #1, , latch13
            Case 15
                Put #1, , map15_BankAddr
                Put #1, , map15_SwapReg
            Case 16
                Put #1, , TmpLatch
                Put #1, , MMC16_IrqOn
                Put #1, , MMC16_Irq
            Case 17
                Put #1, , map17_irqon
                Put #1, , map17_irq
            Case 19
                Put #1, , TmpLatch
                Put #1, , MIRQOn
                Put #1, , MMC19_IRQCount
            Case 24
                Put #1, , map24_IRQCounter
                Put #1, , map24_IRQEnabled
                Put #1, , map24_IRQLatch
                Put #1, , map24_IRQEnOnWrite
            Case 32
                Put #1, , MMC32_Switch
            Case 40
                Put #1, , Mapper40_IRQEnabled
                Put #1, , Mapper40_IRQCounter
            Case 64
                Put #1, , Cmd
                Put #1, , Prg
                Put #1, , Chr1
        End Select
        Put #1, , nt
        Record = True
    ExibeMsg "Gravando filme"
End Sub
Public Sub StopRecording()
    Close #1
    Record = False
    ExibeMsg "Gravação parada"
End Sub
Public Function LoadCfg() As Boolean
    Gamepad1 = Val(GetSetting("YoshiNES", "Control", "Gamepad 1"))
    Gamepad2 = Val(GetSetting("YoshiNES", "Control", "Gamepad 2"))
    pad_ButA = Val(GetSetting("YoshiNES", "Control", "Gamepad1 BUT A"))
    pad_ButB = Val(GetSetting("YoshiNES", "Control", "Gamepad1 BUT B"))
    pad_ButSel = Val(GetSetting("YoshiNES", "Control", "Gamepad1 BUT Select"))
    pad_ButSta = Val(GetSetting("YoshiNES", "Control", "Gamepad1 BUT Start"))
    pad2_ButA = Val(GetSetting("YoshiNES", "Control", "Gamepad2 BUT A"))
    pad2_ButB = Val(GetSetting("YoshiNES", "Control", "Gamepad2 BUT B"))
    pad2_ButSel = Val(GetSetting("YoshiNES", "Control", "Gamepad2 BUT Select"))
    pad2_ButSta = Val(GetSetting("YoshiNES", "Control", "Gamepad2 BUT Start"))
    nes_ButA = Val(GetSetting("YoshiNES", "Control", "NES1 BUT A KeyCode", vbKeyZ))
    nes_ButB = Val(GetSetting("YoshiNES", "Control", "NES1 BUT B KeyCode", vbKeyX))
    nes_ButSel = Val(GetSetting("YoshiNES", "Control", "NES1 BUT Select KeyCode", vbKeyControl))
    nes_ButSta = Val(GetSetting("YoshiNES", "Control", "NES1 BUT Start KeyCode", vbKeyReturn))
    nes_ButUp = Val(GetSetting("YoshiNES", "Control", "NES1 BUT Up KeyCode", vbKeyUp))
    nes_ButDn = Val(GetSetting("YoshiNES", "Control", "NES1 BUT Down KeyCode", vbKeyDown))
    nes_ButLt = Val(GetSetting("YoshiNES", "Control", "NES1 BUT Left KeyCode", vbKeyLeft))
    nes_ButRt = Val(GetSetting("YoshiNES", "Control", "NES1 BUT Right KeyCode", vbKeyRight))
    nes2_ButA = Val(GetSetting("YoshiNES", "Control", "NES2 BUT A KeyCode", vbKeyNumpad7))
    nes2_ButB = Val(GetSetting("YoshiNES", "Control", "NES2 BUT B KeyCode", vbKeyNumpad9))
    nes2_ButSel = Val(GetSetting("YoshiNES", "Control", "NES2 BUT Select KeyCode", vbKeyNumpad0))
    nes2_ButSta = Val(GetSetting("YoshiNES", "Control", "NES2 BUT Start KeyCode", vbKeyNumpad5))
    nes2_ButUp = Val(GetSetting("YoshiNES", "Control", "NES2 BUT Up KeyCode", vbKeyNumpad8))
    nes2_ButDn = Val(GetSetting("YoshiNES", "Control", "NES2 BUT Down KeyCode", vbKeyNumpad2))
    nes2_ButLt = Val(GetSetting("YoshiNES", "Control", "NES2 BUT Left KeyCode", vbKeyNumpad4))
    nes2_ButRt = Val(GetSetting("YoshiNES", "Control", "NES2 BUT Right KeyCode", vbKeyNumpad6))
    frmNES.mnuCh1.Checked = CBool(Val(GetSetting("YoshiNES", "Audio", "Channel 1", 1)))
    frmNES.mnuCh2.Checked = CBool(Val(GetSetting("YoshiNES", "Audio", "Channel 2", 1)))
    frmNES.mnuCh3.Checked = CBool(Val(GetSetting("YoshiNES", "Audio", "Channel 3", 1)))
    frmNES.mnuCh4.Checked = CBool(Val(GetSetting("YoshiNES", "Audio", "Channel 4", 0)))
    Instrumental(0) = Val(GetSetting("YoshiNES", "Audio", "Instrumental 1", 80))
    Instrumental(1) = Val(GetSetting("YoshiNES", "Audio", "Instrumental 2", 80))
    Instrumental(2) = Val(GetSetting("YoshiNES", "Audio", "Instrumental 3", 79))
    Instrumental(3) = Val(GetSetting("YoshiNES", "Audio", "Instrumental 4", 127))
    frmNES.mLayer1.Checked = CBool(Val(GetSetting("YoshiNES", "Video", "Layer 1", 1)))
    frmNES.mLayer2.Checked = CBool(Val(GetSetting("YoshiNES", "Video", "Layer 2", 1)))
    frmNES.mnuShowStatus.Checked = CBool(Val(GetSetting("YoshiNES", "Video", "Show FPS")))
    frmNES.PicScreen.Width = Val(GetSetting("YoshiNES", "Video", "Width", 3840 * 2))
    frmNES.PicScreen.Height = Val(GetSetting("YoshiNES", "Video", "Height", 3600 * 2))
    Recents(0) = GetSetting("YoshiNES", "Misc", "Recent 1")
    Recents(1) = GetSetting("YoshiNES", "Misc", "Recent 2")
    Recents(2) = GetSetting("YoshiNES", "Misc", "Recent 3")
    Recents(3) = GetSetting("YoshiNES", "Misc", "Recent 4")
    Recents(4) = GetSetting("YoshiNES", "Misc", "Recent 5")
    PalName = GetSetting("YoshiNES", "Misc", "Pal Name")
    Lang = GetSetting("YoshiNES", "Misc", "Language", 0)
    frmNES.ResScreen
End Function
Public Function SaveCfg()
    SaveSetting "YoshiNES", "Control", "Gamepad 1", Gamepad1
    SaveSetting "YoshiNES", "Control", "Gamepad 2", Gamepad2
    SaveSetting "YoshiNES", "Control", "Gamepad1 BUT A", pad_ButA
    SaveSetting "YoshiNES", "Control", "Gamepad1 BUT B", pad_ButB
    SaveSetting "YoshiNES", "Control", "Gamepad1 BUT Select", pad_ButSel
    SaveSetting "YoshiNES", "Control", "Gamepad1 BUT Start", pad_ButSta
    SaveSetting "YoshiNES", "Control", "Gamepad2 BUT A", pad2_ButA
    SaveSetting "YoshiNES", "Control", "Gamepad2 BUT B", pad2_ButB
    SaveSetting "YoshiNES", "Control", "Gamepad2 BUT Select", pad2_ButSel
    SaveSetting "YoshiNES", "Control", "Gamepad2 BUT Start", pad2_ButSta
    SaveSetting "YoshiNES", "Control", "NES1 BUT A KeyCode", nes_ButA
    SaveSetting "YoshiNES", "Control", "NES1 BUT B KeyCode", nes_ButB
    SaveSetting "YoshiNES", "Control", "NES1 BUT Select KeyCode", nes_ButSel
    SaveSetting "YoshiNES", "Control", "NES1 BUT Start KeyCode", nes_ButSta
    SaveSetting "YoshiNES", "Control", "NES1 BUT Up KeyCode", nes_ButUp
    SaveSetting "YoshiNES", "Control", "NES1 BUT Down KeyCode", nes_ButDn
    SaveSetting "YoshiNES", "Control", "NES1 BUT Left KeyCode", nes_ButLt
    SaveSetting "YoshiNES", "Control", "NES1 BUT Right KeyCode", nes_ButRt
    SaveSetting "YoshiNES", "Control", "NES2 BUT A KeyCode", nes2_ButA
    SaveSetting "YoshiNES", "Control", "NES2 BUT B KeyCode", nes2_ButB
    SaveSetting "YoshiNES", "Control", "NES2 BUT Select KeyCode", nes2_ButSel
    SaveSetting "YoshiNES", "Control", "NES2 BUT Start KeyCode", nes2_ButSta
    SaveSetting "YoshiNES", "Control", "NES2 BUT Up KeyCode", nes2_ButUp
    SaveSetting "YoshiNES", "Control", "NES2 BUT Down KeyCode", nes2_ButDn
    SaveSetting "YoshiNES", "Control", "NES2 BUT Left KeyCode", nes2_ButLt
    SaveSetting "YoshiNES", "Control", "NES2 BUT Right KeyCode", nes2_ButRt
    SaveSetting "YoshiNES", "Audio", "Channel 1", CInt(frmNES.mnuCh1.Checked)
    SaveSetting "YoshiNES", "Audio", "Channel 2", CInt(frmNES.mnuCh2.Checked)
    SaveSetting "YoshiNES", "Audio", "Channel 3", CInt(frmNES.mnuCh3.Checked)
    SaveSetting "YoshiNES", "Audio", "Channel 4", CInt(frmNES.mnuCh4.Checked)
    SaveSetting "YoshiNES", "Audio", "Instrumental 1", Instrumental(0)
    SaveSetting "YoshiNES", "Audio", "Instrumental 2", Instrumental(1)
    SaveSetting "YoshiNES", "Audio", "Instrumental 3", Instrumental(2)
    SaveSetting "YoshiNES", "Audio", "Instrumental 4", Instrumental(3)
    SaveSetting "YoshiNES", "Video", "Layer 1", CInt(frmNES.mLayer1.Checked)
    SaveSetting "YoshiNES", "Video", "Layer 2", CInt(frmNES.mLayer2.Checked)
    SaveSetting "YoshiNES", "Video", "Show FPS", CInt(frmNES.mnuShowStatus.Checked)
    SaveSetting "YoshiNES", "Video", "Width", frmNES.PicScreen.Width
    SaveSetting "YoshiNES", "Video", "Height", frmNES.PicScreen.Height
    SaveSetting "YoshiNES", "Misc", "Recent 1", Recents(0)
    SaveSetting "YoshiNES", "Misc", "Recent 2", Recents(1)
    SaveSetting "YoshiNES", "Misc", "Recent 3", Recents(2)
    SaveSetting "YoshiNES", "Misc", "Recent 4", Recents(3)
    SaveSetting "YoshiNES", "Misc", "Recent 5", Recents(4)
    SaveSetting "YoshiNES", "Misc", "Pal Name", PalName
    SaveSetting "YoshiNES", "Misc", "Language", Lang
End Function
Public Sub ExibeMsg(ByVal Msg As String)
    StrMsg = Msg
    frmNES.MsgTimer.Enabled = True
    frmNES.UpdateFPS
End Sub
