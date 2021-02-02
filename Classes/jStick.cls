VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "jStick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type JOYCAPS
        wMid As Integer
        wPid As Integer
        wXmin As Integer
        wXmax As Integer
        wYmin As Integer
        wYmax As Integer
        wZmin As Integer
        wZmax As Integer
        wNumButtons As Integer
        wPeriodMin As Integer
        wPeriodMax As Integer
End Type
Private Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type
Private Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type

Private Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Private Declare Function joyGetNumDevs Lib "winmm.dll" Alias "joyGetNumDev" () As Long
Private Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Private Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Private Declare Function joyGetThreshold Lib "winmm.dll" (ByVal id As Long, lpuThreshold As Long) As Long
Private Declare Function joyReleaseCapture Lib "winmm.dll" (ByVal id As Long) As Long
Private Declare Function joySetCapture Lib "winmm.dll" (ByVal hwnd As Long, ByVal uID As Long, ByVal uPeriod As Long, ByVal bChanged As Long) As Long
Private Declare Function joySetThreshold Lib "winmm.dll" (ByVal id As Long, ByVal uThreshold As Long) As Long
Private Const JOY_RETURNBUTTONS = &H80&
Private Const JOY_RETURNCENTERED = &H400&
Private Const JOY_RETURNPOV = &H40&
Private Const JOY_RETURNPOVCTS = &H200&
Private Const JOY_RETURNR = &H8&
Private Const JOY_RETURNRAWDATA = &H100&
Private Const JOY_RETURNU = &H10                             ' Axis 5
Private Const JOY_RETURNV = &H20                             ' Axis 6
Private Const JOY_RETURNX = &H1&
Private Const JOY_RETURNY = &H2&
Private Const JOY_RETURNZ = &H4&
Private Const JOY_RETURNALL = (JOY_RETURNX Or JOY_RETURNY Or JOY_RETURNZ Or JOY_RETURNR Or JOY_RETURNU Or JOY_RETURNV Or JOY_RETURNPOV Or JOY_RETURNBUTTONS)

Private Const JOYSTICKID1 = 0
Private Const JOYSTICKID2 = 1 ' This is for a option to play head 2 head with 2 gamepads
Private Const MAXPNAMELEN = 32  ' Max product name length (including NULL)
Private ji As JOYINFOEX     ' Joystick state buffer
Private jiC As JOYINFOEX    ' Joystick blank buffer (so dwSize doesn't get messed up)

Private mVar As Variant
Public Property Get CurButton()
    ji.dwFlags = JOY_RETURNALL
    ji.dwSize = Len(ji)
    joyGetPosEx JOYSTICKID1, ji: mVar = ji.dwButtons
    CurButton = mVar
End Property
Public Property Get JoyX() 'Get X Pos
    ji.dwFlags = JOY_RETURNALL
    ji.dwSize = Len(ji)
    joyGetPosEx JOYSTICKID1, ji: mVar = ji.dwXpos
    JoyX = mVar
End Property
Public Property Get JoyY() 'Get Y Pos
    ji.dwFlags = JOY_RETURNALL
    ji.dwSize = Len(ji)
    joyGetPosEx JOYSTICKID1, ji: mVar = ji.dwYpos
    JoyY = mVar
End Property
Public Property Get Joy2CurButton()
    ji.dwFlags = JOY_RETURNALL
    ji.dwSize = Len(ji)
    joyGetPosEx JOYSTICKID2, ji: mVar = ji.dwButtons
    Joy2CurButton = mVar
End Property
Public Property Get Joy2X() 'Get X Pos (Joy 2)
    ji.dwFlags = JOY_RETURNALL
    ji.dwSize = Len(ji)
    joyGetPosEx JOYSTICKID2, ji: mVar = ji.dwXpos
    Joy2X = mVar
End Property
Public Property Get Joy2Y() 'Get Y Pos (Joy 2)
    ji.dwFlags = JOY_RETURNALL
    ji.dwSize = Len(ji)
    joyGetPosEx JOYSTICKID2, ji: mVar = ji.dwYpos
    Joy2Y = mVar
End Property
