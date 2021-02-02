Attribute VB_Name = "M6502Core"
Option Explicit

DefLng A-Z

' Declarations for M6502
' Addressing Modes
Public Const ADR_ABS As Long = 0
Public Const ADR_ABSX As Long = 1
Public Const ADR_ABSY  As Long = 2
Public Const ADR_IMM As Long = 3
Public Const ADR_IMP As Long = 4
Public Const ADR_INDABSX As Long = 5
Public Const ADR_IND As Long = 6
Public Const ADR_INDX As Long = 7
Public Const ADR_INDY As Long = 8
Public Const ADR_INDZP As Long = 9
Public Const ADR_REL As Long = 10
Public Const ADR_ZP As Long = 11
Public Const ADR_ZPX As Long = 12
Public Const ADR_ZPY As Long = 13

' Opcodes
Public Const INS_ADC As Long = 0
Public Const INS_AND As Long = 1
Public Const INS_ASL As Long = 2
Public Const INS_ASLA As Long = 3
Public Const INS_BCC As Long = 4
Public Const INS_BCS As Long = 5
Public Const INS_BEQ As Long = 6
Public Const INS_BIT As Long = 7
Public Const INS_BMI As Long = 8
Public Const INS_BNE As Long = 9
Public Const INS_BPL As Long = 10
Public Const INS_BRK As Long = 11
Public Const INS_BVC As Long = 12
Public Const INS_BVS As Long = 13
Public Const INS_CLC As Long = 14
Public Const INS_CLD As Long = 15
Public Const INS_CLI As Long = 16
Public Const INS_CLV As Long = 17
Public Const INS_CMP As Long = 18
Public Const INS_CPX As Long = 19
Public Const INS_CPY As Long = 20
Public Const INS_DEC As Long = 21
Public Const INS_DEA As Long = 22
Public Const INS_DEX As Long = 23
Public Const INS_DEY As Long = 24
Public Const INS_EOR As Long = 25
Public Const INS_INC As Long = 26
Public Const INS_INX As Long = 27
Public Const INS_INY As Long = 28
Public Const INS_JMP As Long = 29
Public Const INS_JSR As Long = 30
Public Const INS_LDA As Long = 31
Public Const INS_LDX As Long = 32
Public Const INS_LDY As Long = 33
Public Const INS_LSR As Long = 34
Public Const INS_LSRA As Long = 35
Public Const INS_NOP As Long = 36
Public Const INS_ORA As Long = 37
Public Const INS_PHA As Long = 38
Public Const INS_PHP As Long = 39
Public Const INS_PLA As Long = 40
Public Const INS_PLP As Long = 41
Public Const INS_ROL As Long = 42
Public Const INS_ROLA As Long = 43
Public Const INS_ROR As Long = 44
Public Const INS_RORA As Long = 45
Public Const INS_RTI As Long = 46
Public Const INS_RTS As Long = 47
Public Const INS_SBC As Long = 48
Public Const INS_SEC As Long = 49
Public Const INS_SED As Long = 50
Public Const INS_SEI As Long = 51
Public Const INS_STA As Long = 52
Public Const INS_STX As Long = 53
Public Const INS_STY As Long = 54
Public Const INS_TAX As Long = 55
Public Const INS_TAY As Long = 56
Public Const INS_TSX As Long = 57
Public Const INS_TXA As Long = 58
Public Const INS_TXS As Long = 59
Public Const INS_TYA As Long = 60
Public Const INS_BRA As Long = 61
Public Const INS_INA As Long = 62
Public Const INS_PHX As Long = 63
Public Const INS_PLX As Long = 64
Public Const INS_PHY As Long = 65
Public Const INS_PLY As Long = 66

Public CurrentLine As Long 'Integer

'Registers and tempregisters
'DF: Be careful. Anything, anywhere that uses a variable of the same name without declaring it will be using these:
Public a As Byte
Public X As Byte
Public Y As Byte
Public s As Byte
Public P As Byte

'32bit instructions are faster in protected mode than 16bit
Public PC As Long
Public SavePC As Long
Public Value As Long 'Integer
Public Value2 As Long 'Integer
Public Sum As Long 'Integer
Public SaveFlags As Long 'Integer
Public Help As Long

Public Opcode As Byte
Public clockticks6502 As Long

' arrays
Public Ticks(0 To &H100&) As Byte
Public AddrMode(0 To &H100&) As Byte
Public Instruction(0 To &H100&) As Byte
Public GameImage() As Byte

Public MaxIdle As Long

Public CPUPaused As Boolean

Public AddrModeBase As Long

Public MaxCycles As Long 'max cycles until next scanline
Public RealFrames As Long 'actual # of frames rendered

Public IdleDetect As Boolean
Public IdleCheck(65535) As Byte

Public AutoSpeed As Boolean

Public KeyCodes(7) As Long

Public rCycles As Long 'real number of cycles executed
Public nCycles As Long 'number that should be executed
Public SaveCPU As Boolean ' Pause CPU when YoshiNES loses focus?

Public Const M6502_INTERNAL_REVISION = "v.18"
Public Function init6502()
      Ticks(&H0) = 7: Instruction(&H0) = INS_BRK: AddrMode(&H0) = ADR_IMP
      Ticks(&H1) = 6: Instruction(&H1) = INS_ORA: AddrMode(&H1) = ADR_INDX
      Ticks(&H2) = 2: Instruction(&H2) = INS_NOP: AddrMode(&H2) = ADR_IMP
      Ticks(&H3) = 2: Instruction(&H3) = INS_NOP: AddrMode(&H3) = ADR_IMP
      Ticks(&H4) = 3: Instruction(&H4) = INS_NOP: AddrMode(&H4) = ADR_ZP
      Ticks(&H5) = 3: Instruction(&H5) = INS_ORA: AddrMode(&H5) = ADR_ZP
      Ticks(&H6) = 5: Instruction(&H6) = INS_ASL: AddrMode(&H6) = ADR_ZP
      Ticks(&H7) = 2: Instruction(&H7) = INS_NOP: AddrMode(&H7) = ADR_IMP
      Ticks(&H8) = 3: Instruction(&H8) = INS_PHP: AddrMode(&H8) = ADR_IMP
      Ticks(&H9) = 3: Instruction(&H9) = INS_ORA: AddrMode(&H9) = ADR_IMM
      Ticks(&HA) = 2: Instruction(&HA) = INS_ASLA: AddrMode(&HA) = ADR_IMP
      Ticks(&HB) = 2: Instruction(&HB) = INS_NOP: AddrMode(&HB) = ADR_IMP
      Ticks(&HC) = 4: Instruction(&HC) = INS_NOP: AddrMode(&HC) = ADR_ABS
      Ticks(&HD) = 4: Instruction(&HD) = INS_ORA: AddrMode(&HD) = ADR_ABS
      Ticks(&HE) = 6: Instruction(&HE) = INS_ASL: AddrMode(&HE) = ADR_ABS
      Ticks(&HF) = 2: Instruction(&HF) = INS_NOP: AddrMode(&HF) = ADR_IMP
      Ticks(&H10) = 2: Instruction(&H10) = INS_BPL: AddrMode(&H10) = ADR_REL
      Ticks(&H11) = 5: Instruction(&H11) = INS_ORA: AddrMode(&H11) = ADR_INDY
      Ticks(&H12) = 3: Instruction(&H12) = INS_ORA: AddrMode(&H12) = ADR_INDZP
      Ticks(&H13) = 2: Instruction(&H13) = INS_NOP: AddrMode(&H13) = ADR_IMP
      Ticks(&H14) = 3: Instruction(&H14) = INS_NOP: AddrMode(&H14) = ADR_ZP
      Ticks(&H15) = 4: Instruction(&H15) = INS_ORA: AddrMode(&H15) = ADR_ZPX
      Ticks(&H16) = 6: Instruction(&H16) = INS_ASL: AddrMode(&H16) = ADR_ZPX
      Ticks(&H17) = 2: Instruction(&H17) = INS_NOP: AddrMode(&H17) = ADR_IMP
      Ticks(&H18) = 2: Instruction(&H18) = INS_CLC: AddrMode(&H18) = ADR_IMP
      Ticks(&H19) = 4: Instruction(&H19) = INS_ORA: AddrMode(&H19) = ADR_ABSY
      Ticks(&H1A) = 2: Instruction(&H1A) = INS_INA: AddrMode(&H1A) = ADR_IMP
      Ticks(&H1B) = 2: Instruction(&H1B) = INS_NOP: AddrMode(&H1B) = ADR_IMP
      Ticks(&H1C) = 4: Instruction(&H1C) = INS_NOP: AddrMode(&H1C) = ADR_ABS
      Ticks(&H1D) = 4: Instruction(&H1D) = INS_ORA: AddrMode(&H1D) = ADR_ABSX
      Ticks(&H1E) = 7: Instruction(&H1E) = INS_ASL: AddrMode(&H1E) = ADR_ABSX
      Ticks(&H1F) = 2: Instruction(&H1F) = INS_NOP: AddrMode(&H1F) = ADR_IMP
      Ticks(&H20) = 6: Instruction(&H20) = INS_JSR: AddrMode(&H20) = ADR_ABS
      Ticks(&H21) = 6: Instruction(&H21) = INS_AND: AddrMode(&H21) = ADR_INDX
      Ticks(&H22) = 2: Instruction(&H22) = INS_NOP: AddrMode(&H22) = ADR_IMP
      Ticks(&H23) = 2: Instruction(&H23) = INS_NOP: AddrMode(&H23) = ADR_IMP
      Ticks(&H24) = 3: Instruction(&H24) = INS_BIT: AddrMode(&H24) = ADR_ZP
      Ticks(&H25) = 3: Instruction(&H25) = INS_AND: AddrMode(&H25) = ADR_ZP
      Ticks(&H26) = 5: Instruction(&H26) = INS_ROL: AddrMode(&H26) = ADR_ZP
      Ticks(&H27) = 2: Instruction(&H27) = INS_NOP: AddrMode(&H27) = ADR_IMP
      Ticks(&H28) = 4: Instruction(&H28) = INS_PLP: AddrMode(&H28) = ADR_IMP
      Ticks(&H29) = 3: Instruction(&H29) = INS_AND: AddrMode(&H29) = ADR_IMM
      Ticks(&H2A) = 2: Instruction(&H2A) = INS_ROLA: AddrMode(&H2A) = ADR_IMP
      Ticks(&H2B) = 2: Instruction(&H2B) = INS_NOP: AddrMode(&H2B) = ADR_IMP
      Ticks(&H2C) = 4: Instruction(&H2C) = INS_BIT: AddrMode(&H2C) = ADR_ABS
      Ticks(&H2D) = 4: Instruction(&H2D) = INS_AND: AddrMode(&H2D) = ADR_ABS
      Ticks(&H2E) = 6: Instruction(&H2E) = INS_ROL: AddrMode(&H2E) = ADR_ABS
      Ticks(&H2F) = 2: Instruction(&H2F) = INS_NOP: AddrMode(&H2F) = ADR_IMP
      Ticks(&H30) = 2: Instruction(&H30) = INS_BMI: AddrMode(&H30) = ADR_REL
      Ticks(&H31) = 5: Instruction(&H31) = INS_AND: AddrMode(&H31) = ADR_INDY
      Ticks(&H32) = 3: Instruction(&H32) = INS_AND: AddrMode(&H32) = ADR_INDZP
      Ticks(&H33) = 2: Instruction(&H33) = INS_NOP: AddrMode(&H33) = ADR_IMP
      Ticks(&H34) = 4: Instruction(&H34) = INS_BIT: AddrMode(&H34) = ADR_ZPX
      Ticks(&H35) = 4: Instruction(&H35) = INS_AND: AddrMode(&H35) = ADR_ZPX
      Ticks(&H36) = 6: Instruction(&H36) = INS_ROL: AddrMode(&H36) = ADR_ZPX
      Ticks(&H37) = 2: Instruction(&H37) = INS_NOP: AddrMode(&H37) = ADR_IMP
      Ticks(&H38) = 2: Instruction(&H38) = INS_SEC: AddrMode(&H38) = ADR_IMP
      Ticks(&H39) = 4: Instruction(&H39) = INS_AND: AddrMode(&H39) = ADR_ABSY
      Ticks(&H3A) = 2: Instruction(&H3A) = INS_DEA: AddrMode(&H3A) = ADR_IMP
      Ticks(&H3B) = 2: Instruction(&H3B) = INS_NOP: AddrMode(&H3B) = ADR_IMP
      Ticks(&H3C) = 4: Instruction(&H3C) = INS_BIT: AddrMode(&H3C) = ADR_ABSX
      Ticks(&H3D) = 4: Instruction(&H3D) = INS_AND: AddrMode(&H3D) = ADR_ABSX
      Ticks(&H3E) = 7: Instruction(&H3E) = INS_ROL: AddrMode(&H3E) = ADR_ABSX
      Ticks(&H3F) = 2: Instruction(&H3F) = INS_NOP: AddrMode(&H3F) = ADR_IMP
      Ticks(&H40) = 6: Instruction(&H40) = INS_RTI: AddrMode(&H40) = ADR_IMP
      Ticks(&H41) = 6: Instruction(&H41) = INS_EOR: AddrMode(&H41) = ADR_INDX
      Ticks(&H42) = 2: Instruction(&H42) = INS_NOP: AddrMode(&H42) = ADR_IMP
      Ticks(&H43) = 2: Instruction(&H43) = INS_NOP: AddrMode(&H43) = ADR_IMP
      Ticks(&H44) = 2: Instruction(&H44) = INS_NOP: AddrMode(&H44) = ADR_IMP
      Ticks(&H45) = 3: Instruction(&H45) = INS_EOR: AddrMode(&H45) = ADR_ZP
      Ticks(&H46) = 5: Instruction(&H46) = INS_LSR: AddrMode(&H46) = ADR_ZP
      Ticks(&H47) = 2: Instruction(&H47) = INS_NOP: AddrMode(&H47) = ADR_IMP
      Ticks(&H48) = 3: Instruction(&H48) = INS_PHA: AddrMode(&H48) = ADR_IMP
      Ticks(&H49) = 3: Instruction(&H49) = INS_EOR: AddrMode(&H49) = ADR_IMM
      Ticks(&H4A) = 2: Instruction(&H4A) = INS_LSRA: AddrMode(&H4A) = ADR_IMP
      Ticks(&H4B) = 2: Instruction(&H4B) = INS_NOP: AddrMode(&H4B) = ADR_IMP
      Ticks(&H4C) = 3: Instruction(&H4C) = INS_JMP: AddrMode(&H4C) = ADR_ABS
      Ticks(&H4D) = 4: Instruction(&H4D) = INS_EOR: AddrMode(&H4D) = ADR_ABS
      Ticks(&H4E) = 6: Instruction(&H4E) = INS_LSR: AddrMode(&H4E) = ADR_ABS
      Ticks(&H4F) = 2: Instruction(&H4F) = INS_NOP: AddrMode(&H4F) = ADR_IMP
      Ticks(&H50) = 2: Instruction(&H50) = INS_BVC: AddrMode(&H50) = ADR_REL
      Ticks(&H51) = 5: Instruction(&H51) = INS_EOR: AddrMode(&H51) = ADR_INDY
      Ticks(&H52) = 3: Instruction(&H52) = INS_EOR: AddrMode(&H52) = ADR_INDZP
      Ticks(&H53) = 2: Instruction(&H53) = INS_NOP: AddrMode(&H53) = ADR_IMP
      Ticks(&H54) = 2: Instruction(&H54) = INS_NOP: AddrMode(&H54) = ADR_IMP
      Ticks(&H55) = 4: Instruction(&H55) = INS_EOR: AddrMode(&H55) = ADR_ZPX
      Ticks(&H56) = 6: Instruction(&H56) = INS_LSR: AddrMode(&H56) = ADR_ZPX
      Ticks(&H57) = 2: Instruction(&H57) = INS_NOP: AddrMode(&H57) = ADR_IMP
      Ticks(&H58) = 2: Instruction(&H58) = INS_CLI: AddrMode(&H58) = ADR_IMP
      Ticks(&H59) = 4: Instruction(&H59) = INS_EOR: AddrMode(&H59) = ADR_ABSY
      Ticks(&H5A) = 3: Instruction(&H5A) = INS_PHY: AddrMode(&H5A) = ADR_IMP
      Ticks(&H5B) = 2: Instruction(&H5B) = INS_NOP: AddrMode(&H5B) = ADR_IMP
      Ticks(&H5C) = 2: Instruction(&H5C) = INS_NOP: AddrMode(&H5C) = ADR_IMP
      Ticks(&H5D) = 4: Instruction(&H5D) = INS_EOR: AddrMode(&H5D) = ADR_ABSX
      Ticks(&H5E) = 7: Instruction(&H5E) = INS_LSR: AddrMode(&H5E) = ADR_ABSX
      Ticks(&H5F) = 2: Instruction(&H5F) = INS_NOP: AddrMode(&H5F) = ADR_IMP
      Ticks(&H60) = 6: Instruction(&H60) = INS_RTS: AddrMode(&H60) = ADR_IMP
      Ticks(&H61) = 6: Instruction(&H61) = INS_ADC: AddrMode(&H61) = ADR_INDX
      Ticks(&H62) = 2: Instruction(&H62) = INS_NOP: AddrMode(&H62) = ADR_IMP
      Ticks(&H63) = 2: Instruction(&H63) = INS_NOP: AddrMode(&H63) = ADR_IMP
      Ticks(&H64) = 3: Instruction(&H64) = INS_NOP: AddrMode(&H64) = ADR_ZP
      Ticks(&H65) = 3: Instruction(&H65) = INS_ADC: AddrMode(&H65) = ADR_ZP
      Ticks(&H66) = 5: Instruction(&H66) = INS_ROR: AddrMode(&H66) = ADR_ZP
      Ticks(&H67) = 2: Instruction(&H67) = INS_NOP: AddrMode(&H67) = ADR_IMP
      Ticks(&H68) = 4: Instruction(&H68) = INS_PLA: AddrMode(&H68) = ADR_IMP
      Ticks(&H69) = 3: Instruction(&H69) = INS_ADC: AddrMode(&H69) = ADR_IMM
      Ticks(&H6A) = 2: Instruction(&H6A) = INS_RORA: AddrMode(&H6A) = ADR_IMP
      Ticks(&H6B) = 2: Instruction(&H6B) = INS_NOP: AddrMode(&H6B) = ADR_IMP
      Ticks(&H6C) = 5: Instruction(&H6C) = INS_JMP: AddrMode(&H6C) = ADR_IND
      Ticks(&H6D) = 4: Instruction(&H6D) = INS_ADC: AddrMode(&H6D) = ADR_ABS
      Ticks(&H6E) = 6: Instruction(&H6E) = INS_ROR: AddrMode(&H6E) = ADR_ABS
      Ticks(&H6F) = 2: Instruction(&H6F) = INS_NOP: AddrMode(&H6F) = ADR_IMP
      Ticks(&H70) = 2: Instruction(&H70) = INS_BVS: AddrMode(&H70) = ADR_REL
      Ticks(&H71) = 5: Instruction(&H71) = INS_ADC: AddrMode(&H71) = ADR_INDY
      Ticks(&H72) = 3: Instruction(&H72) = INS_ADC: AddrMode(&H72) = ADR_INDZP
      Ticks(&H73) = 2: Instruction(&H73) = INS_NOP: AddrMode(&H73) = ADR_IMP
      Ticks(&H74) = 4: Instruction(&H74) = INS_NOP: AddrMode(&H74) = ADR_ZPX
      Ticks(&H75) = 4: Instruction(&H75) = INS_ADC: AddrMode(&H75) = ADR_ZPX
      Ticks(&H76) = 6: Instruction(&H76) = INS_ROR: AddrMode(&H76) = ADR_ZPX
      Ticks(&H77) = 2: Instruction(&H77) = INS_NOP: AddrMode(&H77) = ADR_IMP
      Ticks(&H78) = 2: Instruction(&H78) = INS_SEI: AddrMode(&H78) = ADR_IMP
      Ticks(&H79) = 4: Instruction(&H79) = INS_ADC: AddrMode(&H79) = ADR_ABSY
      Ticks(&H7A) = 4: Instruction(&H7A) = INS_PLY: AddrMode(&H7A) = ADR_IMP
      Ticks(&H7B) = 2: Instruction(&H7B) = INS_NOP: AddrMode(&H7B) = ADR_IMP
      Ticks(&H7C) = 6: Instruction(&H7C) = INS_JMP: AddrMode(&H7C) = ADR_INDABSX
      Ticks(&H7D) = 4: Instruction(&H7D) = INS_ADC: AddrMode(&H7D) = ADR_ABSX
      Ticks(&H7E) = 7: Instruction(&H7E) = INS_ROR: AddrMode(&H7E) = ADR_ABSX
      Ticks(&H7F) = 2: Instruction(&H7F) = INS_NOP: AddrMode(&H7F) = ADR_IMP
      Ticks(&H80) = 2: Instruction(&H80) = INS_BRA: AddrMode(&H80) = ADR_REL
      Ticks(&H81) = 6: Instruction(&H81) = INS_STA: AddrMode(&H81) = ADR_INDX
      Ticks(&H82) = 2: Instruction(&H82) = INS_NOP: AddrMode(&H82) = ADR_IMP
      Ticks(&H83) = 2: Instruction(&H83) = INS_NOP: AddrMode(&H83) = ADR_IMP
      Ticks(&H84) = 2: Instruction(&H84) = INS_STY: AddrMode(&H84) = ADR_ZP
      Ticks(&H85) = 2: Instruction(&H85) = INS_STA: AddrMode(&H85) = ADR_ZP
      Ticks(&H86) = 2: Instruction(&H86) = INS_STX: AddrMode(&H86) = ADR_ZP
      Ticks(&H87) = 2: Instruction(&H87) = INS_NOP: AddrMode(&H87) = ADR_IMP
      Ticks(&H88) = 2: Instruction(&H88) = INS_DEY: AddrMode(&H88) = ADR_IMP
      Ticks(&H89) = 2: Instruction(&H89) = INS_BIT: AddrMode(&H89) = ADR_IMM
      Ticks(&H8A) = 2: Instruction(&H8A) = INS_TXA: AddrMode(&H8A) = ADR_IMP
      Ticks(&H8B) = 2: Instruction(&H8B) = INS_NOP: AddrMode(&H8B) = ADR_IMP
      Ticks(&H8C) = 4: Instruction(&H8C) = INS_STY: AddrMode(&H8C) = ADR_ABS
      Ticks(&H8D) = 4: Instruction(&H8D) = INS_STA: AddrMode(&H8D) = ADR_ABS
      Ticks(&H8E) = 4: Instruction(&H8E) = INS_STX: AddrMode(&H8E) = ADR_ABS
      Ticks(&H8F) = 2: Instruction(&H8F) = INS_NOP: AddrMode(&H8F) = ADR_IMP
      Ticks(&H90) = 2: Instruction(&H90) = INS_BCC: AddrMode(&H90) = ADR_REL
      Ticks(&H91) = 6: Instruction(&H91) = INS_STA: AddrMode(&H91) = ADR_INDY
      Ticks(&H92) = 3: Instruction(&H92) = INS_STA: AddrMode(&H92) = ADR_INDZP
      Ticks(&H93) = 2: Instruction(&H93) = INS_NOP: AddrMode(&H93) = ADR_IMP
      Ticks(&H94) = 4: Instruction(&H94) = INS_STY: AddrMode(&H94) = ADR_ZPX
      Ticks(&H95) = 4: Instruction(&H95) = INS_STA: AddrMode(&H95) = ADR_ZPX
      Ticks(&H96) = 4: Instruction(&H96) = INS_STX: AddrMode(&H96) = ADR_ZPY
      Ticks(&H97) = 2: Instruction(&H97) = INS_NOP: AddrMode(&H97) = ADR_IMP
      Ticks(&H98) = 2: Instruction(&H98) = INS_TYA: AddrMode(&H98) = ADR_IMP
      Ticks(&H99) = 5: Instruction(&H99) = INS_STA: AddrMode(&H99) = ADR_ABSY
      Ticks(&H9A) = 2: Instruction(&H9A) = INS_TXS: AddrMode(&H9A) = ADR_IMP
      Ticks(&H9B) = 2: Instruction(&H9B) = INS_NOP: AddrMode(&H9B) = ADR_IMP
      Ticks(&H9C) = 4: Instruction(&H9C) = INS_NOP: AddrMode(&H9C) = ADR_ABS
      Ticks(&H9D) = 5: Instruction(&H9D) = INS_STA: AddrMode(&H9D) = ADR_ABSX
      Ticks(&H9E) = 5: Instruction(&H9E) = INS_NOP: AddrMode(&H9E) = ADR_ABSX
      Ticks(&H9F) = 2: Instruction(&H9F) = INS_NOP: AddrMode(&H9F) = ADR_IMP
      Ticks(&HA0) = 3: Instruction(&HA0) = INS_LDY: AddrMode(&HA0) = ADR_IMM
      Ticks(&HA1) = 6: Instruction(&HA1) = INS_LDA: AddrMode(&HA1) = ADR_INDX
      Ticks(&HA2) = 3: Instruction(&HA2) = INS_LDX: AddrMode(&HA2) = ADR_IMM
      Ticks(&HA3) = 2: Instruction(&HA3) = INS_NOP: AddrMode(&HA3) = ADR_IMP
      Ticks(&HA4) = 3: Instruction(&HA4) = INS_LDY: AddrMode(&HA4) = ADR_ZP
      Ticks(&HA5) = 3: Instruction(&HA5) = INS_LDA: AddrMode(&HA5) = ADR_ZP
      Ticks(&HA6) = 3: Instruction(&HA6) = INS_LDX: AddrMode(&HA6) = ADR_ZP
      Ticks(&HA7) = 2: Instruction(&HA7) = INS_NOP: AddrMode(&HA7) = ADR_IMP
      Ticks(&HA8) = 2: Instruction(&HA8) = INS_TAY: AddrMode(&HA8) = ADR_IMP
      Ticks(&HA9) = 3: Instruction(&HA9) = INS_LDA: AddrMode(&HA9) = ADR_IMM
      Ticks(&HAA) = 2: Instruction(&HAA) = INS_TAX: AddrMode(&HAA) = ADR_IMP
      Ticks(&HAB) = 2: Instruction(&HAB) = INS_NOP: AddrMode(&HAB) = ADR_IMP
      Ticks(&HAC) = 4: Instruction(&HAC) = INS_LDY: AddrMode(&HAC) = ADR_ABS
      Ticks(&HAD) = 4: Instruction(&HAD) = INS_LDA: AddrMode(&HAD) = ADR_ABS
      Ticks(&HAE) = 4: Instruction(&HAE) = INS_LDX: AddrMode(&HAE) = ADR_ABS
      Ticks(&HAF) = 2: Instruction(&HAF) = INS_NOP: AddrMode(&HAF) = ADR_IMP
      Ticks(&HB0) = 2: Instruction(&HB0) = INS_BCS: AddrMode(&HB0) = ADR_REL
      Ticks(&HB1) = 5: Instruction(&HB1) = INS_LDA: AddrMode(&HB1) = ADR_INDY
      Ticks(&HB2) = 3: Instruction(&HB2) = INS_LDA: AddrMode(&HB2) = ADR_INDZP
      Ticks(&HB3) = 2: Instruction(&HB3) = INS_NOP: AddrMode(&HB3) = ADR_IMP
      Ticks(&HB4) = 4: Instruction(&HB4) = INS_LDY: AddrMode(&HB4) = ADR_ZPX
      Ticks(&HB5) = 4: Instruction(&HB5) = INS_LDA: AddrMode(&HB5) = ADR_ZPX
      Ticks(&HB6) = 4: Instruction(&HB6) = INS_LDX: AddrMode(&HB6) = ADR_ZPY
      Ticks(&HB7) = 2: Instruction(&HB7) = INS_NOP: AddrMode(&HB7) = ADR_IMP
      Ticks(&HB8) = 2: Instruction(&HB8) = INS_CLV: AddrMode(&HB8) = ADR_IMP
      Ticks(&HB9) = 4: Instruction(&HB9) = INS_LDA: AddrMode(&HB9) = ADR_ABSY
      Ticks(&HBA) = 2: Instruction(&HBA) = INS_TSX: AddrMode(&HBA) = ADR_IMP
      Ticks(&HBB) = 2: Instruction(&HBB) = INS_NOP: AddrMode(&HBB) = ADR_IMP
      Ticks(&HBC) = 4: Instruction(&HBC) = INS_LDY: AddrMode(&HBC) = ADR_ABSX
      Ticks(&HBD) = 4: Instruction(&HBD) = INS_LDA: AddrMode(&HBD) = ADR_ABSX
      Ticks(&HBE) = 4: Instruction(&HBE) = INS_LDX: AddrMode(&HBE) = ADR_ABSY
      Ticks(&HBF) = 2: Instruction(&HBF) = INS_NOP: AddrMode(&HBF) = ADR_IMP
      Ticks(&HC0) = 3: Instruction(&HC0) = INS_CPY: AddrMode(&HC0) = ADR_IMM
      Ticks(&HC1) = 6: Instruction(&HC1) = INS_CMP: AddrMode(&HC1) = ADR_INDX
      Ticks(&HC2) = 2: Instruction(&HC2) = INS_NOP: AddrMode(&HC2) = ADR_IMP
      Ticks(&HC3) = 2: Instruction(&HC3) = INS_NOP: AddrMode(&HC3) = ADR_IMP
      Ticks(&HC4) = 3: Instruction(&HC4) = INS_CPY: AddrMode(&HC4) = ADR_ZP
      Ticks(&HC5) = 3: Instruction(&HC5) = INS_CMP: AddrMode(&HC5) = ADR_ZP
      Ticks(&HC6) = 5: Instruction(&HC6) = INS_DEC: AddrMode(&HC6) = ADR_ZP
      Ticks(&HC7) = 2: Instruction(&HC7) = INS_NOP: AddrMode(&HC7) = ADR_IMP
      Ticks(&HC8) = 2: Instruction(&HC8) = INS_INY: AddrMode(&HC8) = ADR_IMP
      Ticks(&HC9) = 3: Instruction(&HC9) = INS_CMP: AddrMode(&HC9) = ADR_IMM
      Ticks(&HCA) = 2: Instruction(&HCA) = INS_DEX: AddrMode(&HCA) = ADR_IMP
      Ticks(&HCB) = 2: Instruction(&HCB) = INS_NOP: AddrMode(&HCB) = ADR_IMP
      Ticks(&HCC) = 4: Instruction(&HCC) = INS_CPY: AddrMode(&HCC) = ADR_ABS
      Ticks(&HCD) = 4: Instruction(&HCD) = INS_CMP: AddrMode(&HCD) = ADR_ABS
      Ticks(&HCE) = 6: Instruction(&HCE) = INS_DEC: AddrMode(&HCE) = ADR_ABS
      Ticks(&HCF) = 2: Instruction(&HCF) = INS_NOP: AddrMode(&HCF) = ADR_IMP
      Ticks(&HD0) = 2: Instruction(&HD0) = INS_BNE: AddrMode(&HD0) = ADR_REL
      Ticks(&HD1) = 5: Instruction(&HD1) = INS_CMP: AddrMode(&HD1) = ADR_INDY
      Ticks(&HD2) = 3: Instruction(&HD2) = INS_CMP: AddrMode(&HD2) = ADR_INDZP
      Ticks(&HD3) = 2: Instruction(&HD3) = INS_NOP: AddrMode(&HD3) = ADR_IMP
      Ticks(&HD4) = 2: Instruction(&HD4) = INS_NOP: AddrMode(&HD4) = ADR_IMP
      Ticks(&HD5) = 4: Instruction(&HD5) = INS_CMP: AddrMode(&HD5) = ADR_ZPX
      Ticks(&HD6) = 6: Instruction(&HD6) = INS_DEC: AddrMode(&HD6) = ADR_ZPX
      Ticks(&HD7) = 2: Instruction(&HD7) = INS_NOP: AddrMode(&HD7) = ADR_IMP
      Ticks(&HD8) = 2: Instruction(&HD8) = INS_CLD: AddrMode(&HD8) = ADR_IMP
      Ticks(&HD9) = 4: Instruction(&HD9) = INS_CMP: AddrMode(&HD9) = ADR_ABSY
      Ticks(&HDA) = 3: Instruction(&HDA) = INS_PHX: AddrMode(&HDA) = ADR_IMP
      Ticks(&HDB) = 2: Instruction(&HDB) = INS_NOP: AddrMode(&HDB) = ADR_IMP
      Ticks(&HDC) = 2: Instruction(&HDC) = INS_NOP: AddrMode(&HDC) = ADR_IMP
      Ticks(&HDD) = 4: Instruction(&HDD) = INS_CMP: AddrMode(&HDD) = ADR_ABSX
      Ticks(&HDE) = 7: Instruction(&HDE) = INS_DEC: AddrMode(&HDE) = ADR_ABSX
      Ticks(&HDF) = 2: Instruction(&HDF) = INS_NOP: AddrMode(&HDF) = ADR_IMP
      Ticks(&HE0) = 3: Instruction(&HE0) = INS_CPX: AddrMode(&HE0) = ADR_IMM
      Ticks(&HE1) = 6: Instruction(&HE1) = INS_SBC: AddrMode(&HE1) = ADR_INDX
      Ticks(&HE2) = 2: Instruction(&HE2) = INS_NOP: AddrMode(&HE2) = ADR_IMP
      Ticks(&HE3) = 2: Instruction(&HE3) = INS_NOP: AddrMode(&HE3) = ADR_IMP
      Ticks(&HE4) = 3: Instruction(&HE4) = INS_CPX: AddrMode(&HE4) = ADR_ZP
      Ticks(&HE5) = 3: Instruction(&HE5) = INS_SBC: AddrMode(&HE5) = ADR_ZP
      Ticks(&HE6) = 5: Instruction(&HE6) = INS_INC: AddrMode(&HE6) = ADR_ZP
      Ticks(&HE7) = 2: Instruction(&HE7) = INS_NOP: AddrMode(&HE7) = ADR_IMP
      Ticks(&HE8) = 2: Instruction(&HE8) = INS_INX: AddrMode(&HE8) = ADR_IMP
      Ticks(&HE9) = 3: Instruction(&HE9) = INS_SBC: AddrMode(&HE9) = ADR_IMM
      Ticks(&HEA) = 2: Instruction(&HEA) = INS_NOP: AddrMode(&HEA) = ADR_IMP
      Ticks(&HEB) = 2: Instruction(&HEB) = INS_NOP: AddrMode(&HEB) = ADR_IMP
      Ticks(&HEC) = 4: Instruction(&HEC) = INS_CPX: AddrMode(&HEC) = ADR_ABS
      Ticks(&HED) = 4: Instruction(&HED) = INS_SBC: AddrMode(&HED) = ADR_ABS
      Ticks(&HEE) = 6: Instruction(&HEE) = INS_INC: AddrMode(&HEE) = ADR_ABS
      Ticks(&HEF) = 2: Instruction(&HEF) = INS_NOP: AddrMode(&HEF) = ADR_IMP
      Ticks(&HF0) = 2: Instruction(&HF0) = INS_BEQ: AddrMode(&HF0) = ADR_REL
      Ticks(&HF1) = 5: Instruction(&HF1) = INS_SBC: AddrMode(&HF1) = ADR_INDY
      Ticks(&HF2) = 3: Instruction(&HF2) = INS_SBC: AddrMode(&HF2) = ADR_INDZP
      Ticks(&HF3) = 2: Instruction(&HF3) = INS_NOP: AddrMode(&HF3) = ADR_IMP
      Ticks(&HF4) = 2: Instruction(&HF4) = INS_NOP: AddrMode(&HF4) = ADR_IMP
      Ticks(&HF5) = 4: Instruction(&HF5) = INS_SBC: AddrMode(&HF5) = ADR_ZPX
      Ticks(&HF6) = 6: Instruction(&HF6) = INS_INC: AddrMode(&HF6) = ADR_ZPX
      Ticks(&HF7) = 2: Instruction(&HF7) = INS_NOP: AddrMode(&HF7) = ADR_IMP
      Ticks(&HF8) = 2: Instruction(&HF8) = INS_SED: AddrMode(&HF8) = ADR_IMP
      Ticks(&HF9) = 4: Instruction(&HF9) = INS_SBC: AddrMode(&HF9) = ADR_ABSY
      Ticks(&HFA) = 4: Instruction(&HFA) = INS_PLX: AddrMode(&HFA) = ADR_IMP
      Ticks(&HFB) = 2: Instruction(&HFB) = INS_NOP: AddrMode(&HFB) = ADR_IMP
      Ticks(&HFC) = 2: Instruction(&HFC) = INS_NOP: AddrMode(&HFC) = ADR_IMP
      Ticks(&HFD) = 4: Instruction(&HFD) = INS_SBC: AddrMode(&HFD) = ADR_ABSX
      Ticks(&HFE) = 7: Instruction(&HFE) = INS_INC: AddrMode(&HFE) = ADR_ABSX
      Ticks(&HFF) = 2: Instruction(&HFF) = INS_NOP: AddrMode(&HFF) = ADR_IMP
End Function
Public Sub indabsx6502()
    Help = Read6502(PC) + (Read6502(PC + 1) * &H100&) + X
    SavePC = Read6502(Help) + (Read6502(Help + 1) * &H100&)
End Sub
Public Sub indx6502()
    'TS: Changed PC++ and removed ' (?)
      Value = Read6502(PC) And &HFF
      Value = (Value + X) And &HFF
      PC = PC + 1
      SavePC = Read6502(Value) + (Read6502(Value + 1) * &H100&)
End Sub
Public Sub indy6502()
    'TS: Changed PC++ and == to != (If then else)
      Value = Read6502(PC)
      PC = PC + 1
          
      SavePC = Read6502(Value) + (Read6502(Value + 1) * &H100&)
      If (Ticks(Opcode) = 5) Then
        If ((SavePC \ &H100&) = ((SavePC + Y) \ &H100&)) Then
        Else
          clockticks6502 = clockticks6502 + 1
        End If
      End If
      SavePC = SavePC + Y
End Sub
Public Sub zpx6502()
    'TS: Rewrote everything!
    'Overflow stupid check
    SavePC = Read6502(PC)
    SavePC = SavePC + X
    PC = PC + 1
    SavePC = SavePC And &HFF
End Sub
Public Sub exec6502()
    Dim f As Long
    f = Frames
    While CPUPaused
        DoEvents
    Wend
    While Frames = f And CPURunning
        Opcode = Read6502(PC)  ' Fetch Next Operation
        PC = PC + 1
        If IdleDetect Then
            If IdleCheck(PC) > 8 Then
                If CurrentLine > 240 Or CurrentLine < MaxIdle Then
                    IdleCheck(PC) = IdleCheck(PC) - 8
                Else
                    clockticks6502 = clockticks6502 + CurrentLine \ 2: rCycles = rCycles - CurrentLine \ 2
                End If
            End If
            If CurrentLine >= 231 And CurrentLine < 238 And IdleCheck(PC) < 240 Then
                IdleCheck(PC) = IdleCheck(PC) + 1
            End If
        End If
    
        clockticks6502 = clockticks6502 + Ticks(Opcode)
    
        Select Case Instruction(Opcode)
                Case INS_JMP: ' jmp6502
                    adrmode Opcode
                    PC = SavePC
                Case INS_LDA: ' lda6502
                    adrmode Opcode
                    a = Read6502(SavePC)
                    SetFlags a
                Case INS_LDX:
                    adrmode (Opcode)
                    X = Read6502(SavePC)
                    SetFlags X
                Case INS_LDY
                    adrmode (Opcode)
                    Y = Read6502(SavePC)
                    SetFlags Y
                Case INS_BNE: bne6502
                Case INS_CMP: cmp6502
                Case INS_STA
                    adrmode (Opcode)
                    Write6502 SavePC, a
                Case INS_BIT: bit6502
                Case INS_BVC: bvc6502
                Case INS_BEQ: beq6502
                Case INS_INY: iny6502
                Case INS_BPL: bpl6502
                Case INS_DEX: dex6502
                Case INS_INC: inc6502
                Case INS_DEC: dec6502
                Case INS_JSR: jsr6502
                Case INS_AND: and6502
                Case INS_NOP:
                
                Case INS_BRK: brk6502
                Case INS_ADC: adc6502
                Case INS_EOR: eor6502
                Case INS_ASL: asl6502
                Case INS_ASLA: asla6502
                Case INS_BCC: bcc6502
                Case INS_BCS: bcs6502
                Case INS_BMI: bmi6502
                Case INS_BVS: bvs6502
                Case INS_CLC: P = P And &HFE
                Case INS_CLD: P = P And &HF7
                Case INS_CLI: P = P And &HFB
                Case INS_CLV: P = P And &HBF
                Case INS_CPX: cpx6502
                Case INS_CPY: cpy6502
                Case INS_DEA: dea6502
                Case INS_DEY: dey6502
                Case INS_INA: ina6502
                Case INS_INX: inx6502
                Case INS_LSR: lsr6502
                Case INS_LSRA: lsra6502
                Case INS_ORA
                    adrmode Opcode
                    a = a Or Read6502(SavePC)
                    SetFlags a
                Case INS_PHA: pha6502
                Case INS_PHX: phx6502
                Case INS_PHP: php6502
                Case INS_PHY: phy6502
                Case INS_PLA: pla6502
                Case INS_PLP: plp6502
                Case INS_PLX: plx6502
                Case INS_PLY: ply6502
                Case INS_ROL: rol6502
                Case INS_ROLA: rola6502
                Case INS_ROR: ror6502
                Case INS_RORA: rora6502
                Case INS_RTI: rti6502
                Case INS_RTS: rts6502
                Case INS_SBC: sbc6502
                Case INS_SEC: P = P Or &H1
                Case INS_SED: P = P Or &H8
                Case INS_SEI: P = P Or &H4
                Case INS_STX
                    adrmode (Opcode)
                    Write6502 SavePC, X
                Case INS_STY
                    adrmode (Opcode)
                    Write6502 SavePC, Y
                Case INS_TAX: tax6502
                Case INS_TAY: tay6502
                Case INS_TXA: txa6502
                Case INS_TYA: tya6502
                Case INS_TXS: txs6502
                Case INS_TSX: tsx6502
                Case INS_BRA: bra6502
                Case Else: MsgBox "Opcode inválido - " & Hex$(Opcode)
        End Select
      
        If clockticks6502 > MaxCycles Then
            nCycles = nCycles + 114
            rCycles = rCycles + MaxCycles
            Select Case Mapper
                Case 4, 12, 47, 64, 74, 95, 115, 118, 158, 245: map4_hblank CurrentLine, PPU_Control2
                Case 6: map6_hblank CurrentLine
                Case 16: map16_irq
                Case 17: map17_doirq
                Case 19, 184, 87, 86, 79, 113: map19_irq
                Case 23: map23_irq
                Case 24, 26: map24_irq
                Case 40, 61, 200: map40_irq
                Case 85: map85_irq
                Case 117: map117_hblank CurrentLine
            End Select
            RenderScanline CurrentLine
            If CurrentLine >= 240 Then
                If CurrentLine = 240 Then
                    If Render Then
                        BlitScreen
                        RealFrames = RealFrames + 1
                    End If
                    Frames = Frames + 1
                                      
                    DoEvents
                    
                    CheckJoy
                    
                    If Gamepad1 = 0 Then
                        Joypad1(0) = Keyboard(nes_ButA)
                        Joypad1(1) = Keyboard(nes_ButB)
                        Joypad1(2) = Keyboard(nes_ButSel)
                        Joypad1(3) = Keyboard(nes_ButSta)
                        Joypad1(4) = Keyboard(nes_ButUp)
                        Joypad1(5) = Keyboard(nes_ButDn)
                        Joypad1(6) = Keyboard(nes_ButLt)
                        Joypad1(7) = Keyboard(nes_ButRt)
                    End If
                    If Gamepad2 = 0 Then
                        Joypad2(0) = Keyboard(nes2_ButA)
                        Joypad2(1) = Keyboard(nes2_ButB)
                        Joypad2(2) = Keyboard(nes2_ButSel)
                        Joypad2(3) = Keyboard(nes2_ButSta)
                        Joypad2(4) = Keyboard(nes2_ButUp)
                        Joypad2(5) = Keyboard(nes2_ButDn)
                        Joypad2(6) = Keyboard(nes2_ButLt)
                        Joypad2(7) = Keyboard(nes2_ButRt)
                    End If
                    
                    If Zapper Then
                        Dim Tmp As Long
                        Tmp = ZapperY
                        Tmp = Tmp * 256
                        Tmp = Tmp + ZapperX
                        If VRAM(vBuffer(Tmp) + &H3F00) = &H30 Then
                            If ZapperTrigger Then ZapperLight = 0
                        Else
                            ZapperLight = &HD1
                        End If
                    End If
                    
                    'Movie recording
                    If Record = True Then
                        Put #1, , Joypad1
                        Put #1, , Joypad2
                    ElseIf Playing = True Then
                        Get #1, , Joypad1
                        Get #1, , Joypad2
                        If EOF(1) Then
                            StopPlaying
                            If Lang = 1 Then
                                frmNES.mnuPlayMovie.Caption = "&Play"
                            Else
                                frmNES.mnuPlayMovie.Caption = "&Reproduzir"
                            End If
                            Playing = False
                        End If
                    End If
                End If
                PPU_Status = &H80
                If CurrentLine = 240 Then
                    If (PPU_Control1 And &H80) Then
                        If IdleDetect Then IdleCheck(PC) = 1
                        nmi6502
                    End If
                End If
            End If
            
            If CurrentLine = 0 Or CurrentLine = 131 Then UpdateSounds
            If CurrentLine = 258 Then PPU_Status = &H0
    
            If CurrentLine = 262 Then
                If MaxIdle < 240 And Not SpritesChanged Then
                    MaxIdle = MaxIdle + 16
                Else
                    If MaxIdle > 8 Then MaxIdle = MaxIdle - 8
                    SpritesChanged = False
                End If
                CurrentLine = 0
                
                'Frame Limit
                If Keyboard(192) And 1 Then
                    Render = (Frames Mod 3) = 0
                Else
                    If AutoSpeed Then
                        Dim Delay As Long
                        Static Tmr As Double
                        Static PTime As Double
                        Static Facc As Double
                        Static PFrame As Long
                        Dim Tme As Double
                        
                        Tme = Timer
                        If Tme - PTime < 0.2 And PTime < Tme Then Tmr = Tmr - Tme + PTime
                        Tmr = Tmr + 0.01667
                        PTime = Tme
                        Render = True
                        
                        If Tmr > 0 Then Delay = 10000 * Tmr * Tmr: Sleep (Delay)
                    End If
                    Render = (Frames Mod FrameSkip = 0)
                End If
                PPU_Status = &H0
            Else
                CurrentLine = CurrentLine + 1
            End If
            clockticks6502 = clockticks6502 - MaxCycles
        End If
    Wend
End Sub
Public Sub SetFlags(ByVal Value As Byte)
    If (Value) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Value And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub indzp6502()
    Value = Read6502(PC)
    PC = PC + 1
    SavePC = Read6502(Value) + (Read6502(Value + 1) * &H100&)
End Sub
Public Sub zpy6502()
    SavePC = Read6502(PC)
    SavePC = SavePC + Y
    PC = PC + 1
End Sub
Public Sub absy6502()
    'TS: Changed to != instead of == (Look at absx for more details)
    SavePC = Read6502(PC) + (Read6502(PC + 1) * &H100&)
    PC = PC + 2
    If (Ticks(Opcode) = 4) Then
        If ((SavePC \ &H100&) = ((SavePC + Y) \ &H100&)) Then
        Else
            clockticks6502 = clockticks6502 + 1
        End If
    End If
    SavePC = SavePC + Y
End Sub
Public Sub immediate6502()
    SavePC = PC
    PC = PC + 1
End Sub
Public Sub indirect6502()
    Help = Read6502(PC) + (Read6502(PC + 1) * &H100&)
    SavePC = Read6502(Help) + (Read6502(Help + 1) * &H100&)
    PC = PC + 2
End Sub
Public Sub absx6502()
    'TS: Changed to if then else instead of if then (!= instead of ==)
    SavePC = Read6502(PC)
    SavePC = SavePC + (Read6502(PC + 1) * &H100&)
    PC = PC + 2
    If (Ticks(Opcode) = 4) Then
        If ((SavePC \ &H100&) = ((SavePC + X) \ &H100&)) Then
        Else
            clockticks6502 = clockticks6502 + 1
        End If
    End If
    SavePC = SavePC + X
End Sub
Public Sub abs6502()
    SavePC = Read6502(PC) + (Read6502(PC + 1) * &H100&)
    PC = PC + 2
End Sub
Public Sub relative6502()
    'Changed to PC++ and == to != (If then else)
  
    SavePC = Read6502(PC)
    PC = PC + 1

    If (SavePC And &H80) Then SavePC = (SavePC - &H100&)
End Sub
Public Sub reset6502()
    Dim i As Long
    For i = 0 To 65535
        IdleCheck(i) = 0
    Next i

    a = 0: X = 0: Y = 0: P = &H20
    s = &HFF
      
    PC = Read6502(&HFFFC&) + (Read6502(&HFFFD&) * &H100&)
    Debug.Print "Reset to $" & Hex$(PC) & "[" & PC & "]"
End Sub
Public Sub zp6502()
    SavePC = Read6502(PC)
    PC = PC + 1
End Sub
Public Sub irq6502()
    ' Maskable interrupt
    If (P And &H4) = 0 Then
        Write6502 &H100& + s, (PC \ &H100&)
        s = (s - 1) And &HFF
        Write6502 &H100& + s, (PC And &HFF)
        s = (s - 1) And &HFF
        Write6502 &H100& + s, P
        s = (s - 1) And &HFF
        P = P Or &H4
        PC = Read6502(&HFFFE&) + (Read6502(&HFFFF&) * &H100&)
        clockticks6502 = clockticks6502 + 7
    End If
End Sub
Public Sub nmi6502()
    Write6502 (s + &H100&), (PC \ &H100&)
    s = (s - 1) And &HFF
    Write6502 (s + &H100&), (PC And &HFF)
    s = (s - 1) And &HFF
    Write6502 (s + &H100&), P
    P = P Or &H4
    s = (s - 1) And &HFF
    PC = Read6502(&HFFFA&) + (Read6502(&HFFFB&) * &H100&)
    clockticks6502 = clockticks6502 + 7
End Sub
' This is where all 6502 instructions are kept.
Public Sub adc6502()
    Dim Tmp As Long ' Integer
      
    adrmode Opcode
    Value = Read6502(SavePC)
     
    SaveFlags = (P And &H1)

    Sum = a
    Sum = (Sum + Value) And &HFF
    Sum = (Sum + SaveFlags) And &HFF
      
    If (Sum > &H7F) Or (Sum < -&H80) Then
        P = P Or &H40
    Else
        P = (P And &HBF)
    End If
      
    Sum = a + (Value + SaveFlags)
    If (Sum > &HFF) Then
        P = P Or &H1
    Else
        P = (P And &HFE)
    End If
      
    a = Sum And &HFF
    If (P And &H8) Then
        P = (P And &HFE)
        If ((a And &HF) > &H9) Then
            a = (a + &H6) And &HFF
        End If
        If ((a And &HF0) > &H90) Then
            a = (a + &H60) And &HFF
            P = P Or &H1
        End If
    Else
        clockticks6502 = clockticks6502 + 1
    End If
    SetFlags a
End Sub
Public Sub adrmode(Opcode As Byte)
Select Case AddrMode(Opcode)
    Case ADR_ABS: SavePC = Read6502(PC) + (Read6502(PC + 1) * &H100&): PC = PC + 2
    Case ADR_ABSX: absx6502
    Case ADR_ABSY: absy6502
    Case ADR_IMP: ' nothing really necessary cause
    Case ADR_IMM: SavePC = PC: PC = PC + 1
    Case ADR_INDABSX: indabsx6502
    Case ADR_IND: indirect6502
    Case ADR_INDX: indx6502
    Case ADR_INDY: indy6502
    Case ADR_INDZP: indzp6502
    Case ADR_REL: SavePC = Read6502(PC): PC = PC + 1: If (SavePC And &H80) Then SavePC = SavePC - &H100&
    Case ADR_ZP: SavePC = Read6502(PC): SavePC = SavePC And &HFF: PC = PC + 1
    Case ADR_ZPX: zpx6502
    Case ADR_ZPY: zpy6502
    Case Else: Debug.Print AddrMode(Opcode)
End Select
End Sub
Public Sub and6502()
    adrmode Opcode
    Value = Read6502(SavePC)
    a = (a And Value)
    SetFlags a
End Sub
Public Sub asl6502()
    adrmode Opcode
    Value = Read6502(SavePC)
    
    P = (P And &HFE) Or ((Value \ 128) And &H1)
    Value = (Value * 2) And &HFF
    
    Write6502 SavePC, (Value And &HFF)
    SetFlags Value
End Sub
Public Sub asla6502()
    P = (P And &HFE) Or ((a \ 128) And &H1)
    a = (a * 2) And &HFF
    SetFlags a
End Sub
Public Sub bcc6502()
    If ((P And &H1) = 0) Then
        adrmode Opcode
        PC = PC + SavePC
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub
Public Sub bcs6502()
    If (P And &H1) Then
        adrmode Opcode
        PC = PC + SavePC
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub
Public Sub beq6502()
    If (P And &H2) Then
        adrmode Opcode
        PC = PC + SavePC
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub
Public Sub bit6502()
    adrmode Opcode
    Value = Read6502(SavePC)
  
    If (Value And a) Then
        P = (P And &HFD)
    Else
        P = P Or &H2
    End If
    P = ((P And &H3F) Or (Value And &HC0))
End Sub
Public Sub bmi6502()
    If (P And &H80) Then
        adrmode Opcode
        PC = PC + SavePC
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub
Public Sub bne6502()
    If ((P And &H2) = 0) Then
        adrmode Opcode
        PC = PC + SavePC
    Else
        PC = PC + 1
    End If
End Sub
Public Sub bpl6502()
    If ((P And &H80) = 0) Then
        adrmode Opcode
        PC = PC + SavePC
    Else
        PC = PC + 1
    End If
End Sub
Public Sub brk6502()
    PC = PC + 1
    Write6502 &H100& + s, (PC \ &H100&) And &HFF
    s = (s - 1) And &HFF
    Write6502 &H100& + s, (PC And &HFF)
    s = (s - 1) And &HFF
    Write6502 &H100& + s, P
    s = (s - 1) And &HFF
    P = P Or &H14
    PC = Read6502(&HFFFE&) + (Read6502(&HFFFF&) * &H100&)
End Sub
Public Sub bvc6502()
    If ((P And &H40) = 0) Then
        adrmode Opcode
        PC = PC + SavePC
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub
Public Sub bvs6502()
    If (P And &H40) Then
        adrmode Opcode
        PC = PC + SavePC
        clockticks6502 = clockticks6502 + 1
    Else
        PC = PC + 1
    End If
End Sub
Public Sub clc6502()
    P = P And &HFE
End Sub
Public Sub cld6502()
    P = P And &HF7
End Sub
Public Sub cli6502()
    P = P And &HFB
End Sub
Public Sub clv6502()
    P = P And &HBF
End Sub
Public Sub cmp6502()
    adrmode Opcode
    Value = Read6502(SavePC)
    
    If (a + &H100 - Value) > &HFF Then
        P = P Or &H1
    Else
        P = (P And &HFE)
    End If
    
    Value = (a + &H100 - Value) And &HFF
    SetFlags Value
End Sub
Public Sub cpx6502()
    adrmode Opcode
    Value = Read6502(SavePC)
        
    If (X + &H100 - Value > &HFF) Then
        P = P Or &H1
    Else
        P = (P And &HFE)
    End If
    
    Value = (X + &H100 - Value) And &HFF
    SetFlags Value
End Sub
Public Sub cpy6502()
    adrmode Opcode
    Value = Read6502(SavePC)
        
    If (Y + &H100 - Value > &HFF) Then
        P = (P Or &H1)
    Else
        P = (P And &HFE)
    End If
    Value = (Y + &H100 - Value) And &HFF
    SetFlags Value
End Sub
Public Sub dec6502()
    adrmode Opcode
    Write6502 (SavePC), (Read6502(SavePC) - 1) And &HFF
    Value = Read6502(SavePC)
    If (Value) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Value And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub dex6502()
    X = (X - 1) And &HFF
    If (X) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (X And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub dey6502()
    Y = (Y - 1) And &HFF
    If (Y) Then
          P = P And &HFD
    Else
          P = P Or &H2
    End If
    If (Y And &H80) Then
          P = P Or &H80
    Else
          P = P And &H7F
    End If
End Sub
Public Sub eor6502()
    adrmode Opcode
    a = a Xor Read6502(SavePC)
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub inc6502()
    adrmode Opcode
    Write6502 (SavePC), (Read6502(SavePC) + 1) And &HFF
    Value = Read6502(SavePC)
    If (Value) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Value And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub inx6502()
    X = (X + 1) And &HFF
    If (X) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (X And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub iny6502()
    Y = (Y + 1) And &HFF
    If (Y) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Y And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub jmp6502()
    adrmode Opcode
    PC = SavePC
End Sub
Public Sub jsr6502()
    PC = PC + 1
    Write6502 s + &H100&, (PC \ &H100&)
    s = (s - 1) And &HFF
    Write6502 s + &H100&, (PC And &HFF)
    s = (s - 1) And &HFF
    PC = PC - 1
    adrmode Opcode
    PC = SavePC
End Sub
Public Sub lda6502()
    adrmode Opcode
    
    a = Read6502(SavePC)
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub ldx6502()
    adrmode Opcode
    X = Read6502(SavePC)
    If (X) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (X And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub ldy6502()
    adrmode Opcode
    Y = Read6502(SavePC)
    If (Y) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Y And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub lsr6502()
    adrmode Opcode
    Value = Read6502(SavePC)
           
    P = ((P And &HFE) Or (Value And &H1))
    
    Value = (Value \ 2) And &HFF
    Write6502 SavePC, (Value And &HFF)
    If (Value) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Value And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub lsra6502()
    P = (P And &HFE) Or (a And &H1)
    a = (a \ 2) And &HFF
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub nop6502()
    ' TS: Implemented complex code structure ;)
End Sub
Public Sub ora6502()
    adrmode Opcode
    a = a Or Read6502(SavePC)
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub pha6502()
    Write6502 &H100& + s, a
    s = (s - 1) And &HFF
End Sub
Public Sub php6502()
    Write6502 &H100& + s, P
    s = (s - 1) And &HFF
End Sub
Public Sub pla6502()
    s = (s + 1) And &HFF
    a = Read6502(s + &H100)
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub plp6502()
    s = (s + 1) And &HFF
    P = Read6502(s + &H100) Or &H20
End Sub
Public Sub rol6502()
    SaveFlags = (P And &H1)
    adrmode Opcode
    Value = Read6502(SavePC)
        
    P = (P And &HFE) Or ((Value \ 128) And &H1)
    
    Value = (Value * 2) And &HFF
    Value = Value Or SaveFlags
    
    Write6502 SavePC, (Value And &HFF)
    
    If (Value) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Value And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub rola6502()
    SaveFlags = (P And &H1)
    P = (P And &HFE) Or ((a \ 128) And &H1)
    a = (a * 2) And &HFF
    a = a Or SaveFlags
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub ror6502()
    SaveFlags = (P And &H1)
    adrmode Opcode
    Value = Read6502(SavePC)
        
    P = (P And &HFE) Or (Value And &H1)
    Value = (Value \ 2) And &HFF
    If (SaveFlags) Then
        Value = Value Or &H80
    End If
    Write6502 (SavePC), Value And &HFF
    If (Value) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Value And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub rora6502()
    SaveFlags = (P And &H1)
    P = (P And &HFE) Or (a And &H1)
    a = (a \ 2) And &HFF
    
    If (SaveFlags) Then
        a = a Or &H80
    End If
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub rti6502()
    s = (s + 1) And &HFF
    P = Read6502(s + &H100&) Or &H20
    s = (s + 1) And &HFF
    PC = Read6502(s + &H100&)
    s = (s + 1) And &HFF
    PC = PC + (Read6502(s + &H100) * &H100&)
End Sub
Public Sub rts6502()
    s = (s + 1) And &HFF
    PC = Read6502(s + &H100)
    s = (s + 1) And &HFF
    PC = PC + (Read6502(s + &H100) * &H100&)
    PC = PC + 1
End Sub
Public Sub sbc6502()
    adrmode Opcode
    Value = Read6502(SavePC) Xor &HFF
    
    SaveFlags = (P And &H1)
    
    Sum = a
    Sum = (Sum + Value) And &HFF
    Sum = (Sum + (SaveFlags * 16)) And &HFF
    
    If ((Sum > &H7F) Or (Sum <= -&H80)) Then
        P = P Or &H40
    Else
        P = P And &HBF
    End If
    
    Sum = a + (Value + SaveFlags)
    
    If (Sum > &HFF) Then
        P = P Or &H1
    Else
        P = P And &HFE
    End If
    
    a = Sum And &HFF
    If (P And &H8) Then
        a = (a - &H66) And &HFF
        P = P And &HFE
        If ((a And &HF) > &H9) Then
            a = (a + &H6) And &HFF
        End If
        If ((a And &HF0) > &H90) Then
            a = (a + &H60) And &HFF
            P = P Or &H1
        End If
    Else
        clockticks6502 = clockticks6502 + 1
    End If

    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub sec6502()
    P = P Or &H1
End Sub
Public Sub sed6502()
    P = P Or &H8
End Sub
Public Sub sei6502()
    P = P Or &H4
End Sub
Public Sub sta6502()
    adrmode Opcode
    Write6502 (SavePC), a
End Sub
Public Sub stx6502()
    adrmode Opcode
    Write6502 (SavePC), X
End Sub
Public Sub sty6502()
    adrmode Opcode
    Write6502 (SavePC), Y
End Sub
Public Sub tax6502()
    X = a
    If (X) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (X And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub tay6502()
    Y = a
    If (Y) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Y And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub tsx6502()
    X = s
    If (X) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (X And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub txa6502()
    a = X
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub txs6502()
  s = X
End Sub
Public Sub tya6502()
    a = Y
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub bra6502()
    adrmode Opcode
    PC = PC + SavePC
    clockticks6502 = clockticks6502 + 1
End Sub
Public Sub dea6502()
    a = (a - 1) And &HFF
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub ina6502()
    a = (a + 1) And &HFF
    If (a) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (a And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub phx6502()
    Write6502 &H100 + s, X
    s = (s - 1) And &HFF
End Sub
Public Sub plx6502()
    s = (s + 1) And &HFF
    X = Read6502(s + &H100)
    If (X) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (X And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
Public Sub phy6502()
    Write6502 &H100 + s, Y
    s = (s - 1) And &HFF
End Sub
Public Sub ply6502()
    s = (s + 1) And &HFF
    
    Y = Read6502(s + &H100)
    If (Y) Then
        P = P And &HFD
    Else
        P = P Or &H2
    End If
    If (Y And &H80) Then
        P = P Or &H80
    Else
        P = P And &H7F
    End If
End Sub
