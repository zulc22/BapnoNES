Attribute VB_Name = "pAPUMidi"
'Partially emulates the Audio Processing Unit (APU).

Option Explicit
DefLng A-Z

Private Tones(3)
Private Volume(3)
Private LastFrame(3)
Private StopTones(3)

Public DoSound As Boolean
Private vLengths(31) As Long
Public ChannelWrite(3) As Boolean

Const MASTER_VOLUME As Integer = 10
Private Sub FillArray(a() As Long, ParamArray B() As Variant)
    Dim i
    For i = 0 To UBound(a)
        a(i) = B(i)
    Next i
End Sub
Public Sub pAPUinit()
    'Lookup table used by nester.
    FillArray vLengths, 5, 127, 10, 1, 19, 2, 40, 3, 80, 4, 30, 5, 7, 6, 13, 7, 6, 8, 12, 9, 24, 10, 48, 11, 96, 12, 36, 13, 8, 14, 16, 15
    
    ' known problems:
    '   no instantaneous volume with midi. Really short notes can't be generated and heard.
    '   couldn't find an adequate noise generator.
    '   can't control the shape of the square wave.
    
    SelectInstrument 0, 80 'Square wave
    SelectInstrument 1, 80 'Square wave
    SelectInstrument 2, 74 'Triangle wave. Used recorder (like a flute)
    SelectInstrument 3, 127 'Noise. Used gunshot. Sometimes inadequate.
End Sub
Public Sub PlayTone(Channel, Tone, v)
    If Tone <> Tones(Channel) Or v < Volume(Channel) - 3 Or v > Volume(Channel) Or v = 0 Then
        If Tones(Channel) <> 0 Then
            ToneOff Channel, Tones(Channel)
            Tones(Channel) = 0
            Volume(Channel) = 0
        End If
        If DoSound And Tone > 0 And Tone <= 127 And v > 0 Then
            Volume(Channel) = v
            Tones(Channel) = Tone
            ToneOn Channel, Tone, v * MASTER_VOLUME
        End If
    End If
End Sub
Public Sub StopTone(Channel)
    If Tones(Channel) <> 0 Then
        StopTones(Channel) = Tones(Channel)
        Tones(Channel) = 0
        Volume(Channel) = 0
    End If
End Sub
Public Sub ReallyStopTones()
    Dim i
    For i = 0 To 3
        If StopTones(i) <> 0 And StopTones(i) <> Tones(i) Then
            ToneOff i, StopTones(i)
            StopTones(i) = 0
        End If
    Next i
End Sub
'Calculates a midi tone given an nes frequency.
'Frequency passed is actual interval in 1/65536's of a second (I hope). nope.
Public Function GetTone(ByVal Freq) As Long
    If Freq <= 0 Then Exit Function
    
    Dim t As Long
    
    ' Hopefully this is correct. Convert period to frequency.
    ' freq = 65536 / freq
    ' wow. I was way off. Almost an entire octave. -DF
    Freq = 111861 / (Freq + 1)
    
    'convert to frequency to closest note
    t = CLng(Log(Freq / 8.176) * 17.31234)
    
    GetTone = t
End Function
Public Sub PlayRect(Ch)
    Dim f, l, v
    If SoundCtrl And Pow2(Ch) Then
        v = (Sound(Ch * 4 + 0) And 15) 'Get volume
        l = vLengths(Sound(Ch * 4 + 3) \ 8) 'Get length
        If v > 0 Then
            f = Sound(Ch * 4 + 2) + (Sound(Ch * 4 + 3) And 7) * 256 'Get frequency
            If f > 1 Then
                If ChannelWrite(Ch) Then 'Ensures that a note doesn't replay unless memory written
                    ChannelWrite(Ch) = False
                    LastFrame(Ch) = Frames + l
                    PlayTone Ch, GetTone(f), v
                End If
            Else
                StopTone Ch
            End If
        Else
            StopTone Ch
        End If
    Else
        ChannelWrite(Ch) = True
        StopTone Ch
    End If
    If Frames >= LastFrame(Ch) Then
        StopTone Ch
    End If
End Sub
Public Sub PlayTriangle(Ch)
    Dim f, l, v
    If SoundCtrl And Pow2(Ch) Then
        v = 6 'triangle
        l = vLengths(Sound(Ch * 4 + 3) \ 8)
        If v > 0 Then
            f = Sound(Ch * 4 + 2) + (Sound(Ch * 4 + 3) And 7) * 256
            If f > 1 Then
                If ChannelWrite(Ch) Then
                    ChannelWrite(Ch) = False
                    LastFrame(Ch) = Frames + l
                    PlayTone Ch, GetTone(f), v
                End If
            Else
                StopTone Ch
            End If
        Else
            StopTone Ch
        End If
    Else
        ChannelWrite(Ch) = True
        StopTone Ch
    End If
    If Frames >= LastFrame(Ch) Then
        StopTone Ch
    End If
End Sub
Public Sub PlayNoise(Ch)
    Dim f, l, v
    If SoundCtrl And Pow2(Ch) Then
        v = 6
        l = vLengths(Sound(Ch * 4 + 3) \ 8)
        If v > 0 Then
            f = (Sound(Ch * 4 + 2) And 15) * 128
            If f > 1 Then
                If ChannelWrite(Ch) Then
                    ChannelWrite(Ch) = False
                    LastFrame(Ch) = Frames + l
                    PlayTone Ch, GetTone(f), v
                End If
            Else
                StopTone Ch
            End If
        Else
            StopTone Ch
        End If
    Else
        ChannelWrite(Ch) = True
        StopTone Ch
    End If
    If Frames >= LastFrame(Ch) Then
        StopTone Ch
    End If
End Sub
Public Sub StopSound()
    StopTone 0
    StopTone 1
    StopTone 2
    StopTone 3
    ReallyStopTones
End Sub
Public Sub UpdateSounds()
    If DoSound Then
        ReallyStopTones
        If frmNES.mnuCh1.Checked = True Then PlayRect 0
        If frmNES.mnuCh2.Checked = True Then PlayRect 1
        If frmNES.mnuCh3.Checked = True Then PlayTriangle 2
        If frmNES.mnuCh4.Checked = True Then PlayNoise 3
    Else
        StopSound
    End If
End Sub
