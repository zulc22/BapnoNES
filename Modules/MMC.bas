Attribute VB_Name = "MMC"
Option Explicit

DefLng A-Z

'Functions for emulating MMCs. Select8KVROM and the like
Public CurrVr As Byte
Public PrgSwitch1 As Byte
Public PrgSwitch2 As Byte
Public SpecialWrite6000 As Boolean

Public Bank0(2047) As Byte ' RAM
Public Bank6(8191) As Byte ' SaveRAM
Public Bank8(8191) As Byte '8-E are PRG-ROM.
Public BankA(8191) As Byte
Public BankC(8191) As Byte
Public BankE(8191) As Byte

Private P8, PA, RC, PE 'Addresses of Prg-Rom banks currently selected.
Private prevBSSrc(7) As Long 'Used to ensure that it doesn't bankswitch when the correct bank is already selected
Private allowXor As Boolean

Public CurVBank As Integer
Private Sub CopyBanks(dest, src, count)
    On Error Resume Next
    If Mapper = 4 And allowXor Then
        Dim i
        For i = 0 To count - 1
            MemCopy VRAM(MMC3_ChrAddr Xor (dest + i) * &H400), VROM((src + i) * &H400), &H400
        Next i
    Else
        MemCopy VRAM(dest * &H400), VROM(src * &H400), count * &H400
    End If
End Sub
'doesn't bankswitch when not needed
Private Sub BankSwitch(ByVal dest, ByVal src, ByVal count)
    Dim Aa, B, c
    Aa = 0
    c = 0
    allowXor = count <= 2 'only xor with MMC3_ChrAddr with banks of 1 or 2k
    For B = 0 To count - 1
        If prevBSSrc(dest + B) <> src + B Then
            c = c + 1 'we copy banks in groups, not 1 at a time. a little faster.
            prevBSSrc(dest + B) = src + B
        Else
            If c > 0 Then CopyBanks dest + Aa, src + Aa, c
            Aa = B + 1
            c = 0
        End If
    Next B
    If c > 0 Then CopyBanks dest + Aa, src + Aa, c
End Sub
'resets the info used to decide if a bankswitch is needed.
Public Sub mmc_reset()
    P8 = -1
    PA = -1
    RC = -1
    PE = -1
    Dim i As Long
    For i = 0 To 7
        prevBSSrc(i) = -1
    Next i
End Sub
Public Sub map4_sync()
    If swap Then
        reg8 = &HFE
        regA = PrgSwitch2
        regC = PrgSwitch1
        regE = &HFF
    Else
        reg8 = PrgSwitch1
        regA = PrgSwitch2
        regC = &HFE
        regE = &HFF
    End If
    SetupBanks
End Sub
Public Function MaskBankAddress(Bank As Byte)
    If Bank >= PrgCount * 2 Then
        Dim i As Byte: i = &HFF
        Do While (Bank And i) >= PrgCount * 2
            i = i \ 2
        Loop
        MaskBankAddress = (Bank And i)
    Else
        MaskBankAddress = Bank
    End If
End Function
Public Function MaskVROM(page As Byte, ByVal mask As Long) As Byte
    Dim i As Long
    If mask = 0 Then mask = 256
    If mask And mask - 1 Then 'if mask is not a power of 2
        i = 1
        Do While i < mask 'find smallest power of 2 >= mask
            i = i + i
        Loop
    Else
        i = mask
    End If
    i = (page And (i - 1))
    If i >= mask Then i = mask - 1
    MaskVROM = i
End Function
'Only switches banks when needed
Public Sub SetupBanks()
    reg8 = MaskBankAddress(reg8)
    regA = MaskBankAddress(regA)
    regC = MaskBankAddress(regC)
    regE = MaskBankAddress(regE)
    
    If P8 <> reg8 Then MemCopy Bank8(0), GameImage(reg8 * &H2000&), &H2000&
    If PA <> regA Then MemCopy BankA(0), GameImage(regA * &H2000&), &H2000&
    If RC <> regC Then MemCopy BankC(0), GameImage(regC * &H2000&), &H2000&
    If PE <> regE Then MemCopy BankE(0), GameImage(regE * &H2000&), &H2000&
    P8 = reg8
    PA = regA
    RC = regC
    PE = regE
End Sub
Public Sub Select8KVROM(ByVal Val1 As Byte)
    CurVBank = Val1
    Val1 = MaskVROM(Val1, ChrCount)
    BankSwitch 0, Val1 * 8, 8
End Sub
Public Sub Select4KVROM(ByVal Val1 As Byte, ByVal Bank As Byte)
    CurVBank = Val1 \ 2
    Val1 = MaskVROM(Val1, ChrCount * 2)
    BankSwitch Bank * 4, Val1 * 4, 4
End Sub
Public Sub Select2KVROM(ByVal Val1 As Byte, ByVal Bank As Byte)
    CurVBank = Val1 \ 4
    Val1 = MaskVROM(Val1, ChrCount * 4)
    BankSwitch Bank * 2, Val1 * 2, 2
End Sub
Public Sub Select1KVROM(ByVal Val1 As Byte, ByVal Bank As Byte)
    CurVBank = Val1 \ 8
    Val1 = MaskVROM(Val1, ChrCount * 8)
    BankSwitch Bank, Val1, 1
End Sub
