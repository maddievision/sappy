Attribute VB_Name = "mdlSapp"
Public timestart As Long
Public lastgbbank As Long
Public debugpath As String
Public Type minstmap
 MapTo As Byte
 Transpose As Integer
 SecondNote As Integer
 ThirdNote As Integer '
 VolumeEnvelopeID As Byte
 PitchEnvelopeID As Byte
 Sustain As Boolean
End Type
Public Type s2d
 x As Long
 y As Long
End Type
Public exiting As Boolean
Public fullstop As Boolean
Public pitchenv(0 To &HFF, 0 To &H3F) As s2d
Public volenv(0 To &HFF, 0 To &H3F) As s2d
Public pitchrange(0 To &HFF) As Byte
Public songlayer As Byte
Public gbc_fpn(0 To &H10) As Byte
Public gbc_octave(0 To &H10) As Byte
Public notedot(0 To &H10) As Byte
Public custommap As String
Public snote(0 To &HFF, 0 To &H10) As Byte
Public snotes(0 To &H10) As Long
Public intable As Long
Public lastdrumkit As Byte
Public looplimit As Byte
Public lastvelo(0 To &H10) As Byte
Public songpointer As Long
Public pc(0 To &H10) As Long
Public timepass(0 To &H10) As Boolean
'ublic tracks(0 To &HF) As Long
Public numtracks As Byte
Public disabledtracks(0 To &H10) As Boolean
Public instmap(0 To &H7F) As minstmap
Public drummap(0 To &H7F) As Byte
Public drumkits(0 To &H7F) As Byte
Public noisemap(0 To &H7F) As Byte
Public noisekits(0 To &H7F) As Byte
Public vex(0 To &H10) As Boolean
Public notelen(0 To &H3F) As Single
Public xnote(0 To &H10) As Boolean
Public wait(0 To &H10) As Byte
Public beat As Long
Public iteration As Long
Public notewait(0 To &H10) As Byte
Public songplay As Boolean
Public instrument(0 To &H10) As Byte
Public spd2 As Byte
Public volcmd(0 To &H10) As Byte
Public volume(0 To &H10) As Byte
Public notedelay(0 To &H10) As Byte
Public velo(0 To &H7F, 0 To &H10) As Byte
Public pitch(0 To &H7F, 0 To &H10) As Byte
Public Transpose(0 To &H10) As Integer
Public lastnote(0 To &H10) As Byte
Public lastchan(0 To &H10) As Byte
Public lastarg3(0 To &H10) As Byte
Public arg3(0 To &H7F, 0 To &H10) As Byte
Public exnote(0 To &H7F, 0 To &H10) As Byte
Public exnotes(0 To &H10) As Byte
Public note(0 To &H10) As Boolean
Public inloop(0 To &H10) As Boolean
Public loopreturn(0 To &H10) As Long
Public speed As Integer
Public loops As Byte
Public xstop As Boolean
Public FileName As String
Public com As Byte
Public argx As Byte
Public argy As Byte
Public Sub lmf(xind, xarr)
notelen(xind - &H80) = xarr
End Sub

Public Function getgbapointer(ByVal offset As Long)
Dim a(0 To 3) As Byte
Get #256, offset + 1, a(0)
Get #256, offset + 2, a(1)
Get #256, offset + 3, a(2)
Get #256, offset + 4, a(3)
If a(3) = 8 Then
getgbapointer = (CLng(a(2)) * CLng(&H10000)) + (CLng(a(1)) * CLng(&H100)) + a(0)
Else
getgbapointer = -1
End If
End Function
Public Function GetPkmnGbcPointerLong(ByVal offset As Long)
Dim a(0 To 2) As Byte
Get #256, offset + 1, a(2)
Get #256, offset + 2, a(0)
Get #256, offset + 3, a(1)
If a(1) < &H40 Or a(1) > &H7F Then
GetPkmnGbcPointerLong = -1
Else
b = a(2) * CLng(&H4000)
c = (a(1) - &H40) * CLng(&H100) + a(0)
GetPkmnGbcPointerLong = b + c
lastgbbank = a(2)
End If

End Function
Public Function GetPkmnGbcPointer(ByVal offset As Long)
Dim a(0 To 2) As Byte
Get #256, offset + 1, a(0)
Get #256, offset + 2, a(1)
If a(1) < &H40 Or a(1) > &H7F Then
GetPkmnGbcPointer = -1
Else
b = lastgbbank * &H4000
c = (a(1) - &H40) * CLng(&H100) + a(0)
GetPkmnGbcPointer = b + c
End If

End Function
Public Sub inc(cv, Optional ByVal incd = 1)
cv = cv + incd
End Sub
Public Sub dec(cv, Optional ByVal decd = 1)
cv = cv - decd
End Sub

'Public Function notelen(ByVal nle) As Byte
'x = nle
''x = nle - &H80
''If x > &H30 Then
''x = x * 3
''ElseIf x > &H20 Then
''x = x * 1.5
''End If
'Select Case x
' Case &H12: x = &H22
' Case &H1C: x = &H2C
' Case &H20: x = &H30
' Case &H28: x = &H38
' Case &H30: x = &H80
'End Select
'notelen = x
'End Function
Public Function notedel(ByVal nle) As Byte
x = nle - &HD0
If x > &H20 Then inc x, &H10
notedel = x
End Function
Public Sub clearinstmap()
For i = 0 To &H7F
With instmap(i)
.MapTo = i
.SecondNote = 0
.ThirdNote = 0
.Transpose = 0
.Sustain = False
.VolumeEnvelopeID = 0
.PitchEnvelopeID = 0
End With
pitchenv(0, 0).x = 0
pitchenv(0, 0).y = 0
pitchenv(0, 1).x = -1
pitchenv(0, 1).y = 999
volenv(0, 0).x = 0
volenv(0, 0).y = 0
volenv(0, 1).x = -1
volenv(0, 1).y = 999
pitchrange(0) = 0
drummap(i) = i
noisemap(i) = i
drumkits(i) = 128
noisekits(i) = 128
Next i
End Sub
Public Function midi2note(ByVal midinum As Byte) As String
If Form1.cM.value = vbChecked Then
midi2note = midinum
Exit Function
End If
xoct = (midinum \ 12)
xnot = midinum Mod 12
Select Case xnot
 Case 0: notx = "C"
 Case 1: notx = "C#"
 Case 2: notx = "D"
 Case 3: notx = "D#"
 Case 4: notx = "E"
 Case 5: notx = "F"
 Case 6: notx = "F#"
 Case 7: notx = "G"
 Case 8: notx = "G#"
 Case 9: notx = "A"
 Case 10: notx = "A#"
 Case 11: notx = "B"
End Select
midi2note = notx & "" & Trim(Str(xoct))
End Function
Public Function note2midi(ByVal MidiNote As String) As Byte
Select Case UCase(Left(MidiNote, 2))
 Case "C#": x = 1: y = 3
 Case "D#": x = 3: y = 3
 Case "F#": x = 6: y = 3
 Case "G#": x = 8: y = 3
 Case "A#": x = 10: y = 3
 Case Else
  Select Case UCase(Left(MidiNote, 1))
   Case "C": x = 0: y = 2
   Case "D": x = 2: y = 2
   Case "E": x = 4: y = 2
   Case "F": x = 5: y = 2
   Case "G": x = 7: y = 2
   Case "A": x = 9: y = 2
   Case "B": x = 11: y = 2
  End Select
 End Select
o = Mid(MidiNote, y)
note2midi = (o * 12) + x
End Function
Public Function panx(ByVal panxy As Byte) As String
If panxy = 64 Then
panx = "C"
ElseIf panxy < 64 Then
panx = "L" & (64 - panxy)
Else
panx = "R" & (panxy - 64)
End If
End Function
Public Sub loadinstmap(ByVal FileName As String)
Dim x As String
Dim mf As Long
Dim mt As Long
Dim t As Integer
Dim sn As Long
Dim tn As Long
Dim dm As String

Open FileName For Input As #5
Do
Line Input #5, x
Select Case x
Case "sust"
 Line Input #5, dm
 mf = CByte(dm)
 instmap(mf).Sustain = True
Case "inst"
 Line Input #5, dm
 mf = CByte(dm)
 Line Input #5, dm
 mt = CByte(dm)
 Line Input #5, dm
 t = CInt(dm)
 Line Input #5, dm
 sn = CByte(dm)
 Line Input #5, dm
 tn = CByte(dm)
 With instmap(mf)
   .MapTo = mt
   .SecondNote = sn
   .ThirdNote = tn
   .Transpose = t
 Line Input #5, dm
 sn = CByte(dm)
 Line Input #5, dm
 tn = CByte(dm)
   .VolumeEnvelopeID = sn
   .PitchEnvelopeID = tn
  End With
 Case "drum"
  Line Input #5, dm
  mf = note2midi(dm)
  Line Input #5, dm
  mt = note2midi(dm)
  drummap(mf) = mt
  Line Input #5, dm
  mt = CByte(dm)
  drumkits(mf) = mt
Case "noise"
  Line Input #5, dm
  mf = note2midi(dm)
  Line Input #5, dm
  mt = note2midi(dm)
  noisemap(mf) = mt
  Line Input #5, dm
  mt = CByte(dm)
  noisekits(mf) = mt
 Case "envelope_pitch"
  Line Input #5, dm
  mf = CByte(dm)
  Line Input #5, dm
  mt = CByte(dm)
  pitchrange(mf) = mt
  it = 0
   Do
   Line Input #5, dm
   If dm = "end_envelope" Then Exit Do
   sn = CLng(dm)
   Line Input #5, dm
   tn = CLng(dm)
   pitchenv(mf, it).x = sn
   pitchenv(mf, it).y = tn
   it = it + 1
   Loop
 Case "envelope_vol"
  Line Input #5, dm
  mf = CByte(dm)
  it = 0
   Do
   Line Input #5, dm
   If dm = "end_envelope" Then Exit Do
   sn = CLng(dm)
   Line Input #5, dm
   tn = CLng(dm)
   volenv(mf, it).x = sn
   volenv(mf, it).y = tn
   it = it + 1
   Loop
 Case "ENDFILE"
  Exit Do
 Case Else
  Exit Do
 End Select
Loop
Close #5
End Sub

Public Function Hex2(ByVal decnumber As Long, Optional ByVal pad As Byte = 2) As String
m = Hex(decnumber)
Do While Len(m) < pad
m = "0" & m
Loop
Hex2 = m
End Function

Public Function GetFilePath(ByVal FileName As String) As String
f = InStrRev(FileName, "\")
GetFilePath = Left(FileName, f)
End Function
Public Function GetFilename(ByVal FileName As String) As String
f = InStrRev(FileName, "\")
GetFilename = Mid(FileName, f + 1)
End Function

