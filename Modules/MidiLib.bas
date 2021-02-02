Attribute VB_Name = "MidiLib"
'Written by David Finch
Option Explicit
DefLng A-Z

Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Public Instrumental(3) As Long

Private mdh As Long
Private midiOpened As Boolean
Public Sub MidiOpen()
    If midiOutOpen(mdh, 0, 0, 0, 0) Then
        Exit Sub
    End If
    midiOpened = True
End Sub
Public Sub MidiClose()
    If midiOpened Then
        midiOutClose mdh
        midiOpened = False
    End If
End Sub
Public Sub SelectInstrument(ByVal Channel As Long, ByVal Patch As Long)
    If midiOpened Then midiOutShortMsg mdh, &HC0 Or Patch * 256 Or Channel
    Instrumental(Channel) = Patch
End Sub
Public Sub ToneOn(ByVal Channel As Long, ByVal Tone As Long, ByVal Volume As Long)
    If midiOpened Then
        If Tone < 0 Then Tone = 0
        If Tone > 127 Then Tone = 127
        midiOutShortMsg mdh, &H90 Or Tone * 256 Or Channel Or Volume * 65536
    End If
End Sub
Public Sub ToneOff(ByVal Channel As Long, ByVal Tone As Long)
    If midiOpened Then
        If Tone < 0 Then Tone = 0
        If Tone > 127 Then Tone = 127
        midiOutShortMsg mdh, &H80 Or Tone * 256 Or Channel
    End If
End Sub
