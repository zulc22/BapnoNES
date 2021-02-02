VERSION 5.00
Begin VB.Form frmNES 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3360
   ClientLeft      =   3495
   ClientTop       =   2580
   ClientWidth     =   3840
   Icon            =   "frmNES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3360
      Left            =   0
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   0
      Width           =   3840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   3240
   End
   Begin VB.Timer MsgTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3000
      Top             =   3240
   End
   Begin VB.Image Splash 
      Height          =   3360
      Left            =   0
      Picture         =   "frmNES.frx":2AFA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3840
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Abrir ROM..."
      End
      Begin VB.Menu mnuFileFree 
         Caption         =   "&Fechar ROM"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_t1 
         Caption         =   "-"
      End
      Begin VB.Menu mSave 
         Caption         =   "Salvar rápido (F5)"
      End
      Begin VB.Menu mRestore 
         Caption         =   "Carregar rápido (F7)"
      End
      Begin VB.Menu mnu_t22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWriteVROM 
         Caption         =   "&Gravar VROM"
      End
      Begin VB.Menu mnuFileRomInfo 
         Caption         =   "&Info. da ROM..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_t26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSs 
         Caption         =   "Screenshot (F12)"
      End
      Begin VB.Menu mnuMovies 
         Caption         =   "&Filme"
         Begin VB.Menu mnuStartRecord 
            Caption         =   "&Gravar"
         End
         Begin VB.Menu mnuPlayMovie 
            Caption         =   "&Reproduzir"
         End
      End
      Begin VB.Menu mnu_t11 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveSlt 
         Caption         =   "&Slot"
         Begin VB.Menu msaveslots 
            Caption         =   "0"
            Index           =   0
         End
         Begin VB.Menu msaveslots 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu msaveslots 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu msaveslots 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu msaveslots 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu msaveslots 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu msaveslots 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu msaveslots 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu msaveslots 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu msaveslots 
            Caption         =   "9"
            Index           =   9
         End
      End
      Begin VB.Menu mnu_t17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentRom 
         Caption         =   "&Recente"
         Begin VB.Menu Rct 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnu_t2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu mnuEmulation 
      Caption         =   "&Opções"
      Begin VB.Menu mnuCPUPause 
         Caption         =   "&Pausar (P)"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reiniciar"
         Begin VB.Menu mnuEmuReset 
            Caption         =   "&Software (R)"
         End
         Begin VB.Menu mnuHardEmuReset 
            Caption         =   "&Hardware"
         End
      End
      Begin VB.Menu mnu_t10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Video"
         Begin VB.Menu mzm 
            Caption         =   "&Zoom"
            Begin VB.Menu mZoom 
               Caption         =   "1 x"
               Index           =   0
            End
            Begin VB.Menu mZoom 
               Caption         =   "2 x"
               Index           =   1
            End
            Begin VB.Menu mZoom 
               Caption         =   "3 x"
               Index           =   2
            End
            Begin VB.Menu mZoom 
               Caption         =   "4 x"
               Index           =   3
            End
            Begin VB.Menu mnu_t7 
               Caption         =   "-"
            End
            Begin VB.Menu mnuFull 
               Caption         =   "Tela Cheia (Alt + Enter)"
            End
         End
         Begin VB.Menu mSmoothTop 
            Caption         =   "&Filtros"
            Begin VB.Menu mSmooth 
               Caption         =   "Nenhum"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mSmooth 
               Caption         =   "Interpolado"
               Index           =   1
            End
            Begin VB.Menu mSmooth 
               Caption         =   "Edge-Finding"
               Index           =   2
            End
            Begin VB.Menu mnu_t8 
               Caption         =   "-"
            End
            Begin VB.Menu mnuScan 
               Caption         =   "Scanlines"
            End
            Begin VB.Menu mMotionBlur 
               Caption         =   "Motion blur"
            End
            Begin VB.Menu mnu_t14 
               Caption         =   "-"
            End
            Begin VB.Menu mnuCut 
               Caption         =   "Cortar borda"
            End
         End
         Begin VB.Menu mLayersTop 
            Caption         =   "&Camadas"
            Begin VB.Menu mLayer1 
               Caption         =   "Fundo"
               Checked         =   -1  'True
            End
            Begin VB.Menu mLayer2 
               Caption         =   "Sprites"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnu_t23 
               Caption         =   "-"
            End
            Begin VB.Menu mnuShowStatus 
               Caption         =   "Exibir FPS"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu mnu_t24 
            Caption         =   "-"
         End
         Begin VB.Menu mPalette 
            Caption         =   "Paleta"
            Begin VB.Menu mSelPalette 
               Caption         =   ""
               Index           =   0
            End
            Begin VB.Menu mnu_t9 
               Caption         =   "-"
            End
            Begin VB.Menu mnuRandomColors 
               Caption         =   "Cores aleatórias"
            End
            Begin VB.Menu mnuInv 
               Caption         =   "Inverter cores"
            End
            Begin VB.Menu mnu_t15 
               Caption         =   "-"
            End
            Begin VB.Menu mnuViewColors 
               Caption         =   "Editar cores..."
            End
         End
      End
      Begin VB.Menu mnuSound 
         Caption         =   "&Áudio"
         Begin VB.Menu mnuCh 
            Caption         =   "Canais"
            Begin VB.Menu mnuCh1 
               Caption         =   "Square 1"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuCh2 
               Caption         =   "Square 2"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuCh3 
               Caption         =   "Triangle"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuCh4 
               Caption         =   "Noise"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu mnuMidiInstr 
            Caption         =   "MIDI..."
         End
         Begin VB.Menu mnu_t6 
            Caption         =   "-"
         End
         Begin VB.Menu mMute 
            Caption         =   "Mudo"
         End
      End
      Begin VB.Menu mnuFrameSkip 
         Caption         =   "&Timing"
         Begin VB.Menu mexec 
            Caption         =   "Velocidade"
            Begin VB.Menu mExecV 
               Caption         =   "Auto ajustar"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu mExecV 
               Caption         =   "200% = Extreme"
               Index           =   1
            End
            Begin VB.Menu mExecV 
               Caption         =   "150% = Overclock"
               Index           =   2
            End
            Begin VB.Menu mExecV 
               Caption         =   "100% = Normal"
               Checked         =   -1  'True
               Index           =   3
            End
            Begin VB.Menu mExecV 
               Caption         =   "75% = Abaixo do normal"
               Index           =   4
            End
            Begin VB.Menu mExecV 
               Caption         =   "50% = Lento"
               Index           =   5
            End
            Begin VB.Menu mExecV 
               Caption         =   "25% = Muito lento"
               Index           =   6
            End
            Begin VB.Menu mnu_t5 
               Caption         =   "-"
            End
            Begin VB.Menu mIdle 
               Caption         =   "Detecção de Ociosidade"
            End
         End
         Begin VB.Menu mnuFrameJumps 
            Caption         =   "Frame Skip"
            Begin VB.Menu mnuFS 
               Caption         =   "0"
               Index           =   0
            End
            Begin VB.Menu mnuFS 
               Caption         =   "1"
               Index           =   1
            End
            Begin VB.Menu mnuFS 
               Caption         =   "2"
               Index           =   2
            End
            Begin VB.Menu mnuFS 
               Caption         =   "3"
               Index           =   3
            End
            Begin VB.Menu mnuFS 
               Caption         =   "4"
               Index           =   4
            End
            Begin VB.Menu mnuFS 
               Caption         =   "5"
               Index           =   5
            End
            Begin VB.Menu mnuFS 
               Caption         =   "6"
               Index           =   6
            End
            Begin VB.Menu mnuFS 
               Caption         =   "7"
               Index           =   7
            End
            Begin VB.Menu mnuFS 
               Caption         =   "8"
               Index           =   8
            End
            Begin VB.Menu mnuFS 
               Caption         =   "9"
               Index           =   9
            End
         End
         Begin VB.Menu mnu_t4 
            Caption         =   "-"
         End
         Begin VB.Menu mAutoSpeed 
            Caption         =   "Limitar a 60 fps"
         End
      End
      Begin VB.Menu mnuEmuConfg 
         Caption         =   "&Controles"
         Begin VB.Menu mnuCSlot1 
            Caption         =   "Entrada 1"
            Begin VB.Menu mnuC1k 
               Caption         =   "Teclado"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuC1j1 
               Caption         =   "Joystick 1"
            End
            Begin VB.Menu mnuC1j2 
               Caption         =   "Joystick 2"
            End
         End
         Begin VB.Menu mnuCSlot2 
            Caption         =   "Entrada 2"
            Begin VB.Menu mnuC2k 
               Caption         =   "Teclado"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuC2j1 
               Caption         =   "Joystick 1"
            End
            Begin VB.Menu mnuC2j2 
               Caption         =   "Joystick 2"
            End
         End
         Begin VB.Menu mnu_t13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuZap 
            Caption         =   "&Zapper"
         End
         Begin VB.Menu mnu_t18 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEmuConfgKeys 
            Caption         =   "Configurar..."
         End
      End
      Begin VB.Menu mnuDipSw 
         Caption         =   "&Dip Switches"
         Begin VB.Menu mnuDipOn 
            Caption         =   "&Lig."
         End
         Begin VB.Menu mnuDipOff 
            Caption         =   "&Desl."
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnu_t16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&Language"
         Begin VB.Menu mnuSelLang 
            Caption         =   "&Português"
            Index           =   0
         End
         Begin VB.Menu mnuSelLang 
            Caption         =   "&English"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuShowDebugger 
         Caption         =   "CPU Regs"
      End
      Begin VB.Menu mnuViewHex 
         Caption         =   "Hex Viewer"
      End
      Begin VB.Menu mnuRamEditor 
         Caption         =   "RAM Editor"
      End
      Begin VB.Menu mnuPT 
         Caption         =   "Pattern Tables Editor"
      End
      Begin VB.Menu mnu_t3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGg 
         Caption         =   "Game Genie"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Sobre"
      End
   End
End
Attribute VB_Name = "frmNES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private FileName As String
Private FPS As String

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Function Open_File() As String
    Dim lReturn As Long
    Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = Me.hwnd
    OpenFile.hInstance = App.hInstance
    If Lang = 1 Then
        sFilter = "NES Roms (*.nes)" & Chr(0) & "*.nes"
    Else
        sFilter = "Roms de NES (*.nes)" & Chr(0) & "*.nes"
    End If
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    If Lang = 1 Then
        OpenFile.lpstrTitle = "Open NES ROM..."
    Else
        OpenFile.lpstrTitle = "Abrir ROM de NES..."
    End If
    OpenFile.flags = 4
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
       Open_File = vbNullString
    Else
       Open_File = OpenFile.lpstrFile
    End If
End Function
Public Function ReturnFileName(ByVal FilePath As String) As String
    Dim fPath As String
    
    fPath = FilePath
    Do While InStr(fPath, "\")
        fPath = Mid(fPath, InStr(fPath, "\") + 1)
    Loop
    If InStr(fPath, ".nes") Then
        fPath = Left(fPath, InStr(fPath, ".nes") - 1)
    End If
    ReturnFileName = fPath
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode >= Asc("0") And KeyCode <= Asc("9") Then
        mSaveSlots_Click KeyCode - Asc("0")
    ElseIf KeyCode = vbKeyF5 Then
        mSave_Click
    ElseIf KeyCode = vbKeyF7 Then
        mRestore_Click
    ElseIf KeyCode = vbKeyP Then
        mnuCPUPause_Click
    ElseIf KeyCode = vbKeyR Then
        mnuEmuReset_Click
    ElseIf KeyCode = vbKeyReturn Then
        If Keyboard(vbKeyMenu) = &H41 Then mnuFull_Click
    End If
    Keyboard(KeyCode) = &H41
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Keyboard(KeyCode) = &H40
End Sub
Private Sub GetPaletteList()
    Dim s As String
    Dim i As Integer
    
    s = Dir(App.Path & "\*.pal")
    Do While s <> vbNullString
        If i Then Load mSelPalette(i)
        mSelPalette(i).Caption = s
        i = i + 1
        s = Dir
    Loop
End Sub
Private Sub GetRecentRomsList()
    Dim i As Integer
    
    For i = 0 To 4
        If Len(Recents(i)) > 0 Then
            If i > 0 Then Load Rct(i)
            Rct(i).Caption = ReturnFileName(Recents(i))
        End If
    Next i
End Sub
Private Sub Form_Initialize()
    XPStyle False
End Sub
Private Sub Form_Terminate()
    MidiClose
End Sub
Private Sub Form_Load()
    Dim i As Integer
    
    VERSION = "YoshiNES v" & App.Major & "." & App.Minor
    Caption = VERSION
    
    DoSound = True
    MidiOpen
    pAPUinit
    
    mExecV_Click 3
    mAutoSpeed_Click
    AutoSpeed = True
    mnuFS_Click 0
    mSaveSlots_Click 0
    mZoom_Click 0
    
    LoadCfg
    
    mnuSelLang_Click Lang
    If Gamepad1 = 1 Then mnuC1j1_Click Else If Gamepad1 = 2 Then mnuC1j2_Click
    If Gamepad2 = 1 Then mnuC2j1_Click Else If Gamepad2 = 2 Then mnuC2j2_Click
    
    GetPaletteList
    GetRecentRomsList
    SetLang
    
    Show
    
    mPalette.Enabled = True
    
    For i = 0 To 30
        Pow2(i) = 2 ^ i
    Next i
    Pow2(31) = -2147483648#
    fillTLook
    
    mnuFileFree_Click
    
    On Error Resume Next
    MkDir (App.Path & "\States\")
    MkDir (App.Path & "\Movies\")
    MkDir (App.Path & "\Srams\")
    MkDir (App.Path & "\Screenshots\")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MidiClose
    mnuFileExit_Click
End Sub
Private Sub mAutoSpeed_Click()
    AutoSpeed = Not AutoSpeed
    mAutoSpeed.Checked = AutoSpeed
End Sub
Private Sub mExecV_Click(Index As Integer)
    Select Case Index
        Case 3: MaxCycles = 115
        Case Else: MaxCycles = (114& * 262& * (6& - (Index - 1)) \ 4& - 114& * 44&) \ 218&
    End Select

    Dim i As Integer
    For i = 0 To mExecV.UBound
        mExecV(i).Checked = False
    Next i
    mExecV(Index).Checked = True
End Sub
Private Sub mIdle_Click()
    mIdle.Checked = Not mIdle.Checked
    IdleDetect = mIdle.Checked
End Sub
Private Sub mLayer1_Click()
    mLayer1.Checked = Not mLayer1.Checked
End Sub
Private Sub mLayer2_Click()
    mLayer2.Checked = Not mLayer2.Checked
End Sub
Private Sub mMotionBlur_Click()
    MotionBlur = Not MotionBlur
    mMotionBlur.Checked = MotionBlur
End Sub
Private Sub mMute_Click()
    mMute.Checked = DoSound
    DoSound = Not DoSound
End Sub
Private Sub mnuC1j1_Click()
    mnuC1k.Checked = False
    mnuC1j1.Checked = True
    mnuC1j2.Checked = False
    Gamepad1 = 1
End Sub
Private Sub mnuC1j2_Click()
    mnuC1k.Checked = False
    mnuC1j1.Checked = False
    mnuC1j2.Checked = True
    Gamepad1 = 2
End Sub
Private Sub mnuC1k_Click()
    mnuC1k.Checked = True
    mnuC1j1.Checked = False
    mnuC1j2.Checked = False
    Gamepad1 = 0
End Sub
Private Sub mnuC2j1_Click()
    mnuC2k.Checked = False
    mnuC2j1.Checked = True
    mnuC2j2.Checked = False
    Gamepad2 = 1
End Sub
Private Sub mnuC2j2_Click()
    mnuC2k.Checked = False
    mnuC2j1.Checked = False
    mnuC2j2.Checked = True
    Gamepad2 = 2
End Sub
Private Sub mnuC2k_Click()
    mnuC2k.Checked = True
    mnuC2j1.Checked = False
    mnuC2j2.Checked = False
    Gamepad2 = 0
End Sub
Private Sub mnuCh1_Click()
    mnuCh1.Checked = Not mnuCh1.Checked
End Sub
Private Sub mnuCh2_Click()
    mnuCh2.Checked = Not mnuCh2.Checked
End Sub
Private Sub mnuCh3_Click()
    mnuCh3.Checked = Not mnuCh3.Checked
End Sub
Private Sub mnuCh4_Click()
    mnuCh4.Checked = Not mnuCh4.Checked
End Sub
Public Sub mnuCPUPause_Click()
    If CPURunning = False Then Exit Sub
    CPUPaused = Not CPUPaused
    mnuCPUPause.Checked = CPUPaused
    If CPUPaused Then StopSound
End Sub
Private Sub mnuDipOff_Click()
    DipSwitch = 0
    mnuDipOn.Checked = False
    mnuDipOff.Checked = True
End Sub
Private Sub mnuDipOn_Click()
    DipSwitch = &HFF
    mnuDipOn.Checked = True
    mnuDipOff.Checked = False
End Sub
Private Sub mnuEmuConfgKeys_Click()
    If CPUPaused = False Then mnuCPUPause_Click
    Load frmConfig
End Sub
Private Sub mnuEmuReset_Click()
    reset6502

    CPURunning = True
    FirstRead = True
    PPU_AddressIsHi = True
    PPUAddress = 0
    SpriteAddress = 0
    PPU_Status = 0
    PPU_Control1 = 0
    PPU_Control2 = 0

    Do Until CPURunning = False
        exec6502
    Loop

    If Mirroring = 1 Then MirrorXor = &H800& Else MirrorXor = &H400&: DoMirror
End Sub
Private Sub mnuFileExit_Click()
    Dim FileNum As Integer
    FileNum = FreeFile
    
    SaveCfg
    If UsesSRAM = True Then ' save the SRAM to a file.
        Open App.Path & "\Srams\" & RomName & ".wrm" For Binary As #FileNum
            Put #FileNum, , Bank6
        Close #FileNum
    End If
    End
End Sub
Private Sub mnuFileFree_Click()
    Timer1.Enabled = False
    MsgTimer.Enabled = False
    Caption = VERSION
    
    ' Stop the sound
    StopSound
    
    ' Erase all known content of rom.
    Erase VROM: Erase GameImage: Erase VRAM: Erase SpriteRAM:
    Erase Bank0: Erase Bank6: Erase Bank8: Erase BankA: Erase BankC: Erase BankE
    Erase Joypad1: Erase Joypad2
    
    PicScreen.Cls
    CPURunning = False
    mSave.Enabled = False
    mRestore.Enabled = False
    mnuStartRecord.Enabled = False
    mnuPlayMovie.Enabled = False
    mnuCPUPause.Enabled = False
    mnuEmuReset.Enabled = False
    mnuHardEmuReset.Enabled = False
    mnuFileRomInfo.Enabled = False
    mnuFileFree.Enabled = False
    mnuSs.Enabled = False
    mnuWriteVROM.Enabled = False
    mnuShowDebugger.Enabled = False
    mnuViewHex.Enabled = False
    mnuRamEditor.Enabled = False
    mnuPT.Enabled = False
    If UsesSRAM = True Then ' save the SRAM to a file.
        Open App.Path & "\Srams\" & RomName & ".wrm" For Binary Access Write As #11
            Put #11, , Bank6
        Close #11
    End If
    UsesSRAM = False
    
    PicScreen.Visible = False
    Splash.Visible = True
End Sub
Private Sub mnuFileLoad_Click()
    StopSound
    
    FileName = Open_File()
        
    If FileName = vbNullString Then Exit Sub
    RomName = ReturnFileName(FileName)
    
    LoadRom FileName
End Sub
Private Sub mnuSelLang_Click(Index As Integer)
    Dim i As Integer
    If Index <> Lang Then
        Lang = Index
        MsgBox "Language has been changed successfully!" & vbCrLf & "You must restart emulator for changes to take effect!", vbInformation, "Done"
    End If
    For i = 0 To mnuSelLang.UBound
        mnuSelLang(i).Checked = False
    Next i
    mnuSelLang(Lang).Checked = True
End Sub
Private Sub mnuTools_Click()
    StopSound
End Sub
Private Sub mnuFile_Click()
    StopSound
End Sub
Private Sub mnuEmulation_Click()
    StopSound
End Sub
Private Sub mnuHardEmuReset_Click()
    StopSound
    mnuFileFree_Click
    LoadRom FileName
End Sub
Private Sub mnuInv_Click()
    Dim n As Long
    Dim mRgb(2) As Byte
    Dim R As Byte, G As Byte, B As Byte
    
    mnuInv.Checked = Not mnuInv.Checked
    For n = 0 To 63
        MemCopy mRgb(0), Pal(n), Len(Pal(n))
        R = (255 - mRgb(2))
        G = (255 - mRgb(1))
        B = (255 - mRgb(0))
        SetPalVal R, G, B, n
    Next n
End Sub
Private Sub mnuRandomColors_Click()
    Dim n As Long
    Dim R As Byte, G As Byte, B As Byte

    For n = 0 To 63
        R = Int(Rnd * 255)
        G = Int(Rnd * 255)
        B = Int(Rnd * 255)
        SetPalVal R, G, B, n
    Next n
End Sub
Private Sub mnuMidiInstr_Click()
    frmMidi.Show
End Sub
Private Sub mnuPT_Click()
    frmPattern.Show
End Sub
Private Sub mnuRamEditor_Click()
    frmRamEditor.Show
End Sub
Private Sub mnuScan_Click()
    mScanlines = Not mScanlines
    mnuScan.Checked = mScanlines
End Sub
Private Sub mnuView_Click()
    StopSound
End Sub
Private Sub mnuHelp_Click()
    StopSound
End Sub
Private Sub mnuFileRomInfo_Click()
    Load frmROMInfo
End Sub
Public Sub mnuFS_Click(Index As Integer)
    Dim i As Long
    FrameSkip = Index + 1
    For i = 0 To mnuFS.UBound
        mnuFS(i).Checked = False
    Next i
    On Error Resume Next
    mnuFS(Index).Checked = True
End Sub
Private Sub mnuFull_Click()
    FScreen = Not FScreen
    If FScreen Then frmRender.Show
End Sub
Private Sub mnuGg_Click()
    If CPUPaused = False Then mnuCPUPause_Click
    frmGameGenie.Show
End Sub
Private Sub mnuHelpAbout_Click()
    If CPUPaused = False Then mnuCPUPause_Click
    Load frmAbout
End Sub
Private Sub mnuPlayMovie_Click()
    Dim mTocar, mParar As String
    
    If Lang = 1 Then
        mTocar = "&Play"
        mParar = "&Stop"
    Else
        mTocar = "&Reproduzir"
        mParar = "&Parar"
    End If
    
    If mnuPlayMovie.Caption = mTocar Then
        PlayMovie CLng(SlotIndex)
        Record = False
        Playing = True
        mnuPlayMovie.Caption = mParar
    Else
        StopPlaying
        mnuPlayMovie.Caption = mTocar
    End If
End Sub
Private Sub mnuShowDebugger_Click()
    frmDebug.Show
End Sub
Private Sub mnuShowStatus_Click()
    mnuShowStatus.Checked = Not mnuShowStatus.Checked
End Sub
Private Sub mnuSs_Click()
    On Error GoTo ErrH
    Dim a As Long
    
    Do
        a = a + 1
        If Dir(App.Path & "\Screenshots\Screenshot " & a & ".jpg") = vbNullString Then
            Call SavePicture(ExportBmp(PicScreen), App.Path & "\Screenshots\Screenshot " & a & ".jpg")
            ExibeMsg "Imagem salva!"
            Exit Sub
        End If
    Loop
    Exit Sub
ErrH:
    ExibeMsg "Erro ao salvar imagem!" 'In case of some overflow
End Sub
Private Sub mnuStartRecord_Click()
    Dim mRec As String
    Dim mParar As String
    
    If Lang = 1 Then
        mRec = "&Record"
        mParar = "&Stop"
    Else
        mRec = "&Gravar"
        mParar = "&Parar"
    End If
    
    If mnuStartRecord.Caption = mRec Then
        RecordMovie CLng(SlotIndex)
        Playing = False
        mnuStartRecord.Caption = mParar
    Else
        StopRecording
        mnuStartRecord.Caption = mRec
    End If
End Sub
Private Sub mnuViewColors_Click()
    frmVPal.Show
End Sub
Private Sub mnuViewHex_Click()
    frmHexView.Show
End Sub
Private Sub mnuWriteVROM_Click()
    On Error GoTo ErrH
    If ChrCount Then
        Open FileName For Binary As #1
            Put #1, PrgMark, VROM
        Close #1
    Else
        ExibeMsg "Sem VROM!"
        Exit Sub
    End If
    ExibeMsg "VROM gravada com sucesso!"
    Exit Sub
ErrH:
    ExibeMsg "ERRO ao gravar VROM!!!"
End Sub
Private Sub mnuZap_Click()
    Zapper = Not Zapper
    mnuZap.Checked = Zapper
End Sub
Private Sub mRestore_Click()
    LoadState SlotIndex
End Sub
Private Sub mSave_Click()
    SaveState SlotIndex
End Sub
Private Sub mSaveSlots_Click(Index As Integer)
    Dim i As Integer
    
    SlotIndex = Index
    For i = 0 To 9
        msaveslots(i).Checked = False
    Next i
    msaveslots(Index).Checked = True
End Sub
Private Sub mSelPalette_Click(Index As Integer)
    PalName = mSelPalette(Index).Caption
    LoadPal PalName
End Sub
Private Sub MsgTimer_Timer()
    StrMsg = vbNullString
    MsgTimer.Enabled = False
End Sub
Private Sub mSmooth_Click(Index As Integer)
    mSmooth(Smooth2x).Checked = False
    Smooth2x = Index
    mSmooth(Smooth2x).Checked = True
End Sub
Private Sub mZoom_Click(Index As Integer)
    PicScreen.Move PicScreen.Left, PicScreen.Top, 256 * (Index + 1) * Screen.TwipsPerPixelX, 240 * (Index + 1) * Screen.TwipsPerPixelY
    Move Left, Top, PicScreen.Width * Screen.TwipsPerPixelX, PicScreen.Height * Screen.TwipsPerPixelY
       
    ResScreen
End Sub
Public Function ResScreen()
    frmNES.Width = PicScreen.Width + 90
    frmNES.Height = PicScreen.Height + 720
        
    Me.Top = (Screen.Height) / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    Splash.Width = PicScreen.Width
    Splash.Height = PicScreen.Height
End Function
Private Sub PicScreen_Paint()
    If CPUPaused Then BlitScreen
End Sub
Private Sub Rct_Click(Index As Integer)
    StopSound
    mnuFileFree_Click
    
    FileName = Recents(Index)
    
    RomName = ReturnFileName(FileName)
    If LoadNES(FileName) = 0 Then Exit Sub
    
    LoadRom FileName
End Sub
Private Sub Timer1_Timer()
    Static P As Long, Pr As Long
    Static PTime As Double
    Dim CTime As Double
    
    CTime = Timer
    FPS = Format((RealFrames - Pr) / (CTime - PTime), "0.0") & " fps (" + CStr(CLng((Frames - P) / (CTime - PTime))) & " virtual)"
    PTime = CTime
    P = Frames
    Pr = RealFrames
    
    UpdateFPS
End Sub
Private Function LoadRom(ByVal FileName As String)
    If LoadNES(FileName) = 0 Then Exit Function
    
    Timer1.Enabled = True
    MsgTimer.Enabled = True
    
    CPUPaused = False
    
    Recents(4) = Recents(3)
    Recents(3) = Recents(2)
    Recents(2) = Recents(1)
    Recents(1) = Recents(0)
    Recents(0) = FileName
    
    Me.Caption = VERSION & " - " & RomName
    
    PicScreen.Visible = True
    Splash.Visible = False
        
    CPURunning = True
    FirstRead = True
    PPU_AddressIsHi = True
    PPUAddress = 0: SpriteAddress = 0: PPU_Status = 0: PPU_Control1 = 0: PPU_Control2 = 0
    TileBased = False
    init6502
    mSave.Enabled = True
    mRestore.Enabled = True
    mnuStartRecord.Enabled = True
    mnuPlayMovie.Enabled = True
    mnuCPUPause.Enabled = True
    mnuEmuReset.Enabled = True
    mnuHardEmuReset.Enabled = True
    mnuSs.Enabled = True
    mnuWriteVROM.Enabled = True
    mnuShowDebugger.Enabled = True
    mnuViewHex.Enabled = True
    mnuRamEditor.Enabled = True
    mnuPT.Enabled = True
    
    If mnuCPUPause.Checked = True Then mnuCPUPause.Checked = False
    
    Do Until CPURunning = False
        exec6502
    Loop
End Function
Private Sub PicScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ZapperX = X \ ((PicScreen.Width / Screen.TwipsPerPixelX) / 256)
    ZapperY = Y \ ((PicScreen.Height / Screen.TwipsPerPixelY) / 240)
End Sub
Private Sub PicScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then ZapperTrigger = &HD1
End Sub
Private Sub PicScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then ZapperTrigger = 0
End Sub
Private Sub SetLang()
    Select Case Lang
        Case 1
            'File
            mnuFile.Caption = "&File"
            mnuFileLoad.Caption = "&Load ROM..."
            mnuFileFree.Caption = "&Close ROM"
            mSave.Caption = "Quick save (F5)"
            mRestore.Caption = "Quick load (F7)"
            mnuWriteVROM.Caption = "&Write VROM"
            mnuFileRomInfo.Caption = "&ROM Info..."
            mnuMovies.Caption = "&Movie"
            mnuStartRecord.Caption = "&Record"
            mnuPlayMovie.Caption = "&Play"
            mnuRecentRom.Caption = "&Recent"
            mnuFileExit.Caption = "Exit"
            'Options
            mnuEmulation.Caption = "&Options"
            mnuCPUPause.Caption = "Pause (P)"
            mnuReset.Caption = "Reset"
            mnuFull.Caption = "Full Screen (Alt + Return)"
            mSmoothTop.Caption = "Filters"
            mSmooth(0).Caption = "None"
            mLayersTop.Caption = "Layers"
            mLayer1.Caption = "Background"
            mnuShowStatus.Caption = "Show FPS"
            mPalette.Caption = "Palette"
            mnuRandomColors.Caption = "Random colors"
            mnuInv.Caption = "Invert colors"
            mnuViewColors.Caption = "Edit Palette..."
            mnuSound.Caption = "Sound"
            mnuCh.Caption = "Channels"
            mMute.Caption = "Mute"
            mexec.Caption = "CPU Speed"
            mIdle.Caption = "Idle detection"
            mExecV(4).Caption = "75% = Under normal"
            mExecV(5).Caption = "50% = Slow"
            mExecV(6).Caption = "25% = Very slow"
            mAutoSpeed.Caption = "Limit to 60 fps"
            mnuEmuConfg.Caption = "Input"
            mnuCSlot1.Caption = "Input 1"
            mnuCSlot2.Caption = "Input 2"
            mnuC1k.Caption = "Keyboard"
            mnuC2k.Caption = "Keyboard"
            mnuEmuConfgKeys.Caption = "Config..."
            mnuDipOn.Caption = "On"
            mnuDipOff.Caption = "Off"
            'Tools
            mnuTools.Caption = "&Tools"
            'Help
            mnuHelp.Caption = "&Help"
            mnuHelpAbout.Caption = "&About"
    End Select
End Sub
Public Sub UpdateFPS()
    'Draw status and FPS
    If StrMsg = vbNullString Then
        If frmNES.mnuShowStatus.Checked Then
            Caption = VERSION & " - " & FPS
        ElseIf Caption <> VERSION & " - " & RomName Then
            Caption = VERSION & " - " & RomName
        End If
    Else
        Caption = VERSION & " - " & StrMsg
    End If
End Sub
