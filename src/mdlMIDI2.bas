Attribute VB_Name = "mdlMIDI"
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Enum midiformat
 midiSingleTrack = 0
 midiMultiTrack = 1
 midiAsyncMultiTrack = 2
End Enum
Public lastrange(0 To 15) As Long
Public tracksizeoff As Long
Public datasize As Long
Public eventpass As Long
Public lastcmd As Long
Public Type MIDIHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        lpNext As Long
        Reserved As Long
End Type



Global MidiEventOut, MidiNoteOut, MidiVelOut As Long
Global hMidiOut As Long
Global hMidiOutCopy As Long
Global MidiOpenError As String

Global Const MODAL = 1
Global Const ShiftKey = 1
Public Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Public Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * 32
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type
Public Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
   Public Declare Function MidiOutClose Lib "winmm.dll" Alias "midiOutClose" (ByVal hMidiOut As Long) As Long
   Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
   Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Public Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

Public Sub SetControl(ByVal channel As Byte, ControlNum, ByVal ControlVal As Byte)
SendMidiOut &HB0 + channel, ControlNum, ControlVal
End Sub

Public Sub SetPitchBendRange(ByVal channel As Byte, ByVal NewRange As Byte)
SetControl channel, &H64, 0
SetControl channel, &H65, 0
SetControl channel, &H6, NewRange
SetControl channel, &H26, 0
End Sub
Public Sub SetVibrato(ByVal channel As Byte, ByVal Depth As Byte, ByVal Rate As Byte)
SetControl 0, &H62, 8
SetControl 0, &H63, 1
SetControl 0, &H66, (Rate / 2) + 64
SetControl 0, &H62, 9
SetControl 0, &H63, 1
SetControl 0, &H66, (Depth / 2) + 64
SetControl 0, &H62, &HA
SetControl 0, &H63, 1
SetControl 0, &H66, 0
If Depth = 0 Then
SetControl 0, 1, 0
Else
SetControl 0, 1, 127
End If
End Sub
Public Sub SetFineTune(ByVal channel As Byte, ByVal FineTune As Byte)
Dim rangeo As Long
finetune1 = (FineTune / 100) * (&H7F + (&H7F * 128))
'MsgBox rangeo
finetuneH = finetune1 \ 128
'MsgBox newrangeh
finetuneL = finetune1 Mod 128
'MsgBox newrangel

SetControl channel, &H64, 0         'RPN LSB
SetControl channel, &H65, 0         'RPN MSB
SetControl channel, &H6, finetuneH  'DATA MSB
SetControl channel, &H26, finetuneL 'DATA LSB
End Sub



Public Sub SetPanning(ByVal channel As Byte, ByVal pan As Byte)
SetControl channel, &HA, pan
End Sub

Public Sub SetBank(ByVal channel As Byte, ByVal MSB As Byte, ByVal LSB As Byte)
SetControl channel, 0, MSB
SetControl channel, 32, LSB
End Sub
Public Sub SetRPN(ByVal channel As Byte, ByVal rlsb As Byte, ByVal rmsb As Byte, ByVal dmsb As Byte)

End Sub
Public Sub SetPitchBend(ByVal channel As Byte, ByVal range As Long)
Dim rangeo As Long
'rangeo = (range / 100) * (&H7F + (&H7F * 128))
'x = ((lastrange(channel) - range) + 8192)
'rangeo = x
rangeo = range
'rangeo = IIf(rangeo > 8191, rangeo - 8192, rangeo + 8192)
lastrange(channel) = range
'MsgBox rangeo
newrangeh = (rangeo \ 128) Mod 128
'MsgBox newrangeh
newrangel = rangeo Mod 128
'MsgBox newrangel
SendMidiOut &HE0 + channel, newrangel, newrangeh
End Sub

Public Sub SetVolume(ByVal channel As Byte, ByVal newvol As Byte)
SendMidiOut &HB0 + channel, 7, newvol
End Sub

Public Sub SetPatch(ByVal channel As Byte, ByVal newpatch As Byte)
SendMidiOut2 &HC0 + channel, newpatch
End Sub

Public Sub PlayNote(ByVal channel As Byte, ByVal note As Byte, ByVal velo As Byte)
SendMidiOut &H90 + channel, note, velo
End Sub
Public Sub NoteOff(ByVal channel As Byte, ByVal note As Byte)
SendMidiOut &H80 + channel, note
End Sub



Public Sub MidiOutOpenPort(ByVal portid As Integer)
   MidiOpenError = Str$(midiOutOpen(hMidiOut, portid, 0, 0, 0))
   hMidiOutCopy = hMidiOut
End Sub

Public Sub MidiClose()
If hMidiOutCopy <> 0 Then
For i = 0 To 15
SetControl i, &H78, 0
SetControl i, &H79, 0
SetControl i, &H7B, 0
Next i
   MidiOutClose hMidiOutCopy
hMidiOutCopy = 0
End If
End Sub
Function LongToVar(ByVal longinfo As Long)
Dim a() As Byte
Dim lb(1 To 4) As Byte
lb(4) = longinfo Mod 128
lb(3) = (longinfo \ 128) Mod 128
lb(2) = (longinfo \ 128 \ 128) Mod 128
lb(1) = (longinfo \ 128 \ 128 \ 128) Mod 128
x = 4
'For i = 1 To 3
'If lb(i) = 0 Then x = x - 1
'Next i
x = IIf(longinfo < 2 ^ 7, 1, IIf(longinfo < 2 ^ 14, 2, IIf(longinfo < 2 ^ 21, 3, IIf(longinfo < 2 ^ 28, 4, 4))))
ReDim Preserve a(x)
y = 1
For i = 4 To 4 - (x - 1) Step -1
a(x - (y - 1)) = lb(i) + IIf(i < 4, 128, 0)
y = y + 1
Next i
LongToVar = a
End Function

Sub SendMidiOut(ByVal MidiEvent As Long, ByVal MidiNote As Long, Optional ByVal MidiVel As Long = 0)
Dim MidiMessage As Long
Dim lowbyte As Byte
Dim midbyte As Byte
Dim highbyte As Byte
a = LongToVar(eventpass - lastcmd)
lastcmd = eventpass
   lowbyte = MidiEvent
   midbyte = MidiNote
   highbyte = MidiVel
   For i = 1 To UBound(a)
Put #256, , CByte(a(i))
   Next i
Put #256, , lowbyte
Put #256, , midbyte
Put #256, , highbyte
datasize = datasize + UBound(a) + 3
   
'   MidiMessage = lowint + highint
'   x% = midiOutShortMsg(hMidiOutCopy, MidiMessage)

End Sub

Sub SendMidiOut2(ByVal MidiEvent As Long, ByVal MidiNote As Long)
Dim MidiMessage As Long
Dim lowbyte As Byte
Dim midbyte As Byte
a = LongToVar(eventpass - lastcmd)
lastcmd = eventpass
   lowbyte = MidiEvent
   midbyte = MidiNote
   For i = 1 To UBound(a)
Put #256, , CByte(a(i))
   Next i
Put #256, , lowbyte
Put #256, , midbyte
datasize = datasize + UBound(a) + 2
   
'   MidiMessage = lowint + highint
'   x% = midiOutShortMsg(hMidiOutCopy, MidiMessage)

End Sub

Sub midiNewFile(ByVal filename As String, ByVal format As midiformat, ByVal ppqn As Integer, Optional ByVal numtracks As Integer = 1)
Open filename For Output As #1
Print #1, "x"
Close #1
Open filename For Binary As #256
Put #256, , CByte(Asc("M"))
Put #256, , CByte(Asc("T"))
Put #256, , CByte(Asc("h"))
Put #256, , CByte(Asc("d"))
Put #256, , CByte(0)
Put #256, , CByte(0)
Put #256, , CByte(0)
Put #256, , CByte(6)
Put #256, , CByte(format)
Put #256, , CByte(0)
Put #256, , CByte(numtracks \ 256)
Put #256, , CByte(numtracks Mod 256)
Put #256, , CByte(ppqn \ 256)
Put #256, , CByte(ppqn Mod 256)
End Sub
Sub newMidiTrack(ByVal tracktitle As String)
Put #256, , CByte(Asc("M"))
Put #256, , CByte(Asc("T"))
Put #256, , CByte(Asc("r"))
Put #256, , CByte(Asc("k"))
tracksizeoff = Loc(256) + 1
Put #256, , CLng(0)
Put #256, , CByte(0)
Put #256, , CByte(&HFF)
Put #256, , CByte(&H3)
Put #256, , CByte(Len(tracktitle))
datasize = datasize + 4
For i = 1 To Len(tracktitle)
Put #256, , CByte(Asc(Mid(tracktitle, i, 1)))
datasize = datasize + 1
Next i
End Sub
Sub endMidiTrack()
Dim lowbyte As Byte
Dim midbyte As Byte
Dim highbyte As Byte
a = LongToVar(eventpass - lastcmd)
lastcmd = eventpass
   lowbyte = &HFF
   midbyte = &H2F
   highbyte = &H0
   For i = 1 To UBound(a)
Put #256, , CByte(a(i))
   Next i
Put #256, , lowbyte
Put #256, , midbyte
Put #256, , highbyte
datasize = datasize + UBound(a) + 3
Put #256, tracksizeoff, CByte((datasize \ (256 ^ 3)))
Put #256, , CByte((datasize \ (256 ^ 2)) Mod 256)
Put #256, , CByte((datasize \ 256) Mod 256)
Put #256, , CByte((datasize) Mod 256)
End Sub
Public Sub settempo(ByVal tempo As Byte)
Dim txt As Long
a = LongToVar(eventpass - lastcmd)
lastcmd = eventpass
   For i = 1 To UBound(a)
Put #256, , CByte(a(i))
   Next i

txt = (48 / (tempo * (2048 / 5120))) * 100000
Put #256, , CByte(&HFF)
Put #256, , CByte(&H51)
Put #256, , CByte(3)
Put #256, , CByte((txt \ (256 ^ 2)) Mod 256)
Put #256, , CByte((txt \ 256) Mod 256)
Put #256, , CByte(txt Mod 256)
datasize = datasize + UBound(a) + 6

End Sub

Sub endmidifile()
Close #256
End Sub
