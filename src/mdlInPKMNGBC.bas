Attribute VB_Name = "mdlInPKMNGBC"
Public Sub callback_pkmngbc()
songpointer = (Form1.tstart * 3) + Form1.ttable
Form1.Label3 = Hex(songpointer)

clearinstmap
'xstop = True
'Do
'DoEvents
'Loop Until songplay = False
MidiOutOpenPort -1
loops = 0
xstop = False
songplay = True
iteration = 0

Open filename For Binary As #256
songlayer = 0
Form1.Label22 = "---"
songpointer = GetPkmnGbcPointerLong(songpointer)

If songpointer = -1 Then GoTo hell
Form1.Label4 = Hex(songpointer)
numtracks = 4
If numtracks = 0 Then GoTo hell
For i = 0 To numtracks
z = IIf(i > 8, i + 1, i)
volcmd(z) = 0
wait(z) = 0

lastvelo(z) = &H7F
notewait(z) = 0
instrument(z) = 0
vex(z) = False
volume(z) = &H7F
lastdrumkit = 255

If fileout = False Then
SetVolume z, &H7F
SetPitchBend z, &H1FFF
SetPitchBendRange z, 2
SetPanning z, 64
SetControl z, 1, 0
SetControl z, 64, 0
End If
notedot(z) = 2
Transpose(z) = 0
beat = 0
notedelay(z) = &H80
exnotes(z) = 0
note(z) = False
inloop(z) = False
xnote(z) = False
Next i
If fileout = False Then
SetVolume 9, &H7F
SetPitchBend 9, &H1FFF
SetPitchBendRange 9, 2
SetControl 9, 1, 0
End If


intable = 0
Form1.Label15 = "---"
Open IIf(Form1.cD.value = vbChecked, debugpath, App.Path & "\") & "sappy.xtt" For Output As #4
Print #4, numtracks & "|" & songlayer & "|"; Hex(songpointer) & "|" & Hex(intable)
Close #4
Dim r As Byte
If Form1.cIMAP.value = vbUnchecked Then
loadinstmap custommap
End If

For i = 0 To numtracks - 1

Form1.Label2(j) = Hex(pc(j))
z = IIf(i > 8, i + 1, i)
pc(z) = GetPkmnGbcPointer(songpointer + 1 + (i * 3))
MsgBox Hex(pc(z))
Next i
spd2 = 64
speed = 70
'settempo 70
timestart = timeGetTime
MidiStart

Do
'
If iteration Mod 2 = 0 Then MidiClock
If iteration Mod 48 = 0 Then
Form1.Label24 = iteration
inc beat
Form1.Label14 = beat
End If
If fileout = False Then
xx = timeGetTime
yy = (xx - timestart) \ 1000
mm = yy \ 60
ss = yy Mod 60

xyz = mm & ":" & IIf(Len(Trim(Str(ss))) < 2, "0" & Trim(Str(ss)), Trim(Str(ss)))
If Form1.Label25 <> xyz Then Form1.Label25 = xyz
End If
If fileout = False Then
Do
DoEvents
Loop While (timeGetTime - xx) < ((((255 - 64) / 60) * 60) / (speed / 3)) And xstop = False
Else
DoEvents
End If
For j = 0 To numtracks - 1
z = IIf(j > 8, j + 1, j)

Form1.Label2(z) = Hex(pc(z))
parsecom_pkmngbc z
Next j
Form1.Label12 = loops
Form1.Label11 = speed
inc iteration
eventpass = iteration
If looplimit > 0 And loops > (looplimit - 1) Then Exit Do
Loop Until xstop = True 'Or loops > 1
'Loop Until xstop = True
hell:
Close #256
MidiStop
   MidiClose
songplay = False
Form1.tLoops.Enabled = True
'Form1.Command2.Enabled = True
Form1.Label13 = "Stopped"
'MsgBox "stopped"
If fullstop = False And fileout = False Then
If Form1.cAuto.value = vbChecked Then
If Form1.cRand.value = vbChecked Then
'Form1.tstart = "&H" & Hex(Fix(Rnd * (Form1.tend - Form1.tSt2)) + Form1.tSt2)
'If Form1.tstart > Form1.tend Then Form1.tstart = "&H" & Hex(Form1.tend)
Randomize Timer
If Form1.auto1.value = True Then
Form1.tstart = "&H" & Hex(Fix(Rnd * &H200))
Else
With Form1.listSongs
If .ListCount > 1 Then
 .ListIndex = Fix(Rnd * (.ListCount))
End If
End With
End If
Else
If Form1.auto1.value = True Then
Form1.tstart = "&H" & Hex(Val(Form1.tstart))
  If Form1.tstart < &H200 Then Form1.tstart = "&H" & Hex(Form1.tstart + 1)

Else
With Form1
If .listSongs.ListIndex < (.listSongs.ListCount - 1) And .listSongs.ListCount > 0 Then .listSongs.ListIndex = .listSongs.ListIndex + 1
End With
End If
End If
 
 Form1.Timer1.Enabled = True
End If
Else
fullstop = False
If fileout = True Then
endMidiTrack
endmidifile
fileout = False
MsgBox "MIDI File Successfully Written"
End If
End If
Form1.Command10.Enabled = True
Form1.Command5.Enabled = True
Form1.Command3.Enabled = False
Form1.mStop.Enabled = False
If exiting = True Then End
End Sub









Public Sub parsecom_pkmngbc(ByVal tracknum As Byte)
timepass(tracknum) = True

If notewait(tracknum) > 1 Then
  dec notewait(tracknum)
'  Form1.Label8(tracknum) = notewait(tracknum) \ 2
  'Form1.Label8(tracknum).ForeColor = RGB(255, 255, 0)
  Form1.Label5(tracknum).BackStyle = 1
  Form1.Label20(tracknum).BackStyle = 1
  Form1.Label21(tracknum).BackStyle = 1
    
  ElseIf note(tracknum) = True Then

If vex(tracknum) = False Then
note(tracknum) = False
'Form1.Label8(tracknum).ForeColor = RGB(255, 0, 0)
  Form1.Label5(tracknum).BackStyle = 0
  Form1.Label20(tracknum).BackStyle = 0
  Form1.Label21(tracknum).BackStyle = 0
'  Form1.Label5(tracknum) = "-"
'NoteOff lastchan(tracknum), lastnote(tracknum)
If exnotes(tracknum) > 0 Then
     For i = 1 To exnotes(tracknum)
        NoteOff lastchan(tracknum), exnote(i - 1, tracknum)
     
     Next i
For i = 0 To snotes(tracknum) - 1
If snotes(tracknum) = 0 Then Exit For
NoteOff lastchan(tracknum), snote(i, tracknum)

Next i
End If
'exnotes(tracknum) = 0
End If
End If

If wait(tracknum) > 1 Then
 dec wait(tracknum)
 ' Form1.Label1(tracknum) = wait(tracknum) \ 2
Else
'''''''''''''''''''''''''
Do Until wait(tracknum) > 1

 Get #256, pc(tracknum) + 1, com
'Form1.Label1(tracknum) = Hex(com)
'Form1.Refresh
 Select Case com
 Case &HFF
   If inloop(tracknum) = True Then
   pc(tracknum) = loopreturn(tracknum)
   Else
   xstop = True
   Exit Do
   End If
 Case &HFD
   If tracknum = 0 Then inc loops
   pc(tracknum) = GetPkmnGbcPointer(pc(tracknum) + 2)
 Case &HFE
   loopreturn(tracknum) = pc(tracknum) + 3
   inloop(tracknum) = True
   pc(tracknum) = GetPkmnGbcPointer(pc(tracknum) + 1)
 Case &HDA
   Get #256, pc(tracknum) + 2, argx
   Get #256, pc(tracknum) + 3, argy
   speed = (argx * CLng(&H100)) + argy
   settempo argy
'   volcmd(tracknum) = 6
   inc pc(tracknum), 3
'  Case &HBC
'   Get #256, pc(tracknum) + 2, argx
' Transpose(tracknum) = IIf(argx < &H81, argx, CInt(argx) - CInt(&H100))
'
'   inc pc(tracknum), 2
 Case &HDB
   Select Case tracknum
    Case 0: ib = 0: tl = "SQ1"
    Case 1: ib = 0: tl = "SQ2"
    Case 2: ib = 32: tl = "TRI"
    Case 3: ib = 64: tl = "NOI"
    Case Else: ib = 0: tl = "???"
   End Select
   Get #256, pc(tracknum) + 2, argx
   argx = argx + ib
   instrument(tracknum) = argx
'   If instrument(tracknum) = 1 Then instrument(tracknum) = 0
'     If instrument(tracknum) > 119 Then
'   instrument(tracknum) = 0
'   If instrument(tracknum) = 119 Then transpose(tracknum) = -12
'   End If
'   If instrument(tracknum) > 118 Then instrument(tracknum) = 80
   Form1.Label6(tracknum) = (argx)
   Form1.LabelC(tracknum) = tl
   If instmap(argx).MapTo < 128 Then
   SetPatch tracknum, instmap(argx).MapTo
   Else
   SetPatch 9, 0
   End If
   inc pc(tracknum), 2
 Case &HE0
   Select Case tracknum
    Case 0: ib = 0: tl = "SQ1"
    Case 1: ib = 0: tl = "SQ2"
    Case 2: ib = 32: tl = "TRI"
    Case 3: ib = 64: tl = "NOI"
    Case Else: ib = 0: tl = "???"
   End Select
   Get #256, pc(tracknum) + 2, argx
   argx = argx + ib
   instrument(tracknum) = argx
'   If instrument(tracknum) = 1 Then instrument(tracknum) = 0
'     If instrument(tracknum) > 119 Then
'   instrument(tracknum) = 0
'   If instrument(tracknum) = 119 Then transpose(tracknum) = -12
'   End If
'   If instrument(tracknum) > 118 Then instrument(tracknum) = 80
   Form1.Label6(tracknum) = (argx)
   Form1.LabelC(tracknum) = tl
   If instmap(argx).MapTo < 128 Then
   SetPatch tracknum, instmap(argx).MapTo
   Else
   SetPatch 9, 0
   End If
   inc pc(tracknum), 2
 Case &HD8
   Get #256, pc(tracknum) + 2, argx
   Get #256, pc(tracknum) + 3, argy
   gbc_fpn(tracknum) = argx
   lastvelo(tracknum) = (argy \ &H10) \ 2
   notedelay(tracknum) = argy Mod &H10
   If notedelay(tracknum) = 0 Then
   vex(tracknum) = True
   Else
   vex(tracknum) = False
   End If
   inc pc(tracknum), 3
 Case &HD8
   Get #256, pc(tracknum) + 2, argy
   lastvelo(tracknum) = (argy \ &H10) \ 2
   notedelay(tracknum) = argy Mod &H10
   If notedelay(tracknum) = 0 Then
   vex(tracknum) = True
   Else
   vex(tracknum) = False
   End If
   inc pc(tracknum), 2
 Case &HE5
   inc pc(tracknum), 2
Case &HEF
   Get #256, pc(tracknum) + 2, argx
   spd2 = argx
'volcmd(tracknum) = 6
'SetPanning tracknum, spd2
'Form1.Label17(tracknum) = panx(spd2)
'Form1.Label17(tracknum).BackStyle = IIf(spd2 = &H40, 0, 1)
   inc pc(tracknum), 2
 Case &HE1
'   Get #256, pc(tracknum) + 2, argx
''volcmd(tracknum) = argx
'   SetNRPN tracknum, 1, 9, (argx * 3)
'SetControl numtrack, 1, IIf(argx = 0, 0, 127)
' Form1.Label9(tracknum) = argx
'Form1.Label9(tracknum).BackStyle = IIf(argx = 0, 0, 1)
'volcmd(tracknum) = &HC4
   inc pc(tracknum), 3
  
 Case &HE6 'tune
   inc pc(tracknum), 3
 
 Case Is > &HD7
   MsgBox "unknown code: " & Hex(com)
   xstop = True
   Exit Do
 Case Is > &HCF
   gbc_octave(tracknum) = 7 - (com - &HD0)
 Case Is < &HC1
   notenum = com \ &H10
    If notenum = 0 Then
    notenum = 12
    Else
    notenum = notenum - 1
    End If
   tnlen = com Mod &H10
   wait(tracknum) = gbc_fpn(tracknum) * tnlen
   notewait(tracknum) = gbc_fpn(tracknum) * notedelay(tracknum)
If timepass(tracknum) = True Then
If note(tracknum) = True Then
note(tracknum) = False
'NoteOff lastchan(tracknum), lastnote(tracknum)
If exnotes(tracknum) > 0 Then
     For i = 1 To exnotes(tracknum)
        NoteOff lastchan(tracknum), exnote(i - 1, tracknum)
     Next i
'exnotes(tracknum) = 0
End If
End If
exnotes(tracknum) = 0
Form1.Label5(tracknum) = ""
Form1.Label20(tracknum) = ""
Form1.Label21(tracknum) = ""
timepass(tracknum) = False
End If
'If xnote(tracknum) = False Then
Form1.Label5(tracknum) = Form1.Label5(tracknum) & IIf(exnotes(tracknum) > 0, "/", "") & midi2note(com)
If notenum < 11 Then
n2p = (notenum * gbc_octave(tracknum)) + Transpose(tracknum) + instmap(instrument(tracknum)).Transpose
pitch(exnotes(tracknum), tracknum) = n2p
If instmap(instrument(tracknum)).MapTo > 127 Then
 exnote(exnotes(tracknum), tracknum) = drummap(n2p)
Else
exnote(exnotes(tracknum), tracknum) = n2p
End If
velo(exnotes(tracknum), tracknum) = lastvelo(tracknum)
xnote(tracknum) = True
End If
'Else
'     xnote(tracknum) = True
'End If
inc pc(tracknum)
'If instrument(tracknum) > 0 Then
inc exnotes(tracknum)
Form1.Label20(tracknum) = Form1.Label20(tracknum) & IIf(exnotes(tracknum) > 1, "/", "") & velo(exnotes(tracknum) - 1, tracknum)
Form1.Label21(tracknum) = "---"

 'Case Is < &H80
 '    inc pc(tracknum)
'     pitch(tracknum) = (com - 70) + pitch(tracknum)
'     xnote(tracknum) = True
'        Get #256, pc(tracknum) + 2, argx
'   Get #256, pc(tracknum) + 3, argy
' If argx < &H80 Then
' velo(tracknum) = argx
' Form1.Label9(tracknum) = Hex(argx)
'' notedelay(tracknum) = argx - &H5C
' inc pc(tracknum)
'' If argy < &H80 Then
'' inc pc(tracknum)
''
'' End If
' End If
Case Else
   MsgBox "unknown code: " & Hex(com)
   xstop = True
   Exit Do
End Select

If wait(tracknum) > 1 Then
     Form1.Label1(tracknum) = wait(tracknum)
 If Form1.Check1(tracknum).value = vbChecked Then
If xnote(tracknum) = True Then
dontplay = 0
   If note(tracknum) = True Then
     note(tracknum) = False
     notewait(tracknum) = 0
'     NoteOff lastchan(tracknum), lastnote(tracknum)
If exnotes(tracknum) > 0 Then
     For i = 1 To exnotes(tracknum)
        NoteOff lastchan(tracknum), exnote(i - 1, tracknum)
     
     Next i
For i = 0 To snotes(tracknum) - 1
If snotes(tracknum) = 0 Then Exit For
NoteOff lastchan(tracknum), snote(i, tracknum)
Next i
vex(tracknum) = False
'exnotes(tracknum) = 0
End If
End If
snotes(tracknum) = 0
     If instmap(instrument(tracknum)).MapTo > 127 Then
     lastchan(tracknum) = 9
     Else
     lastchan(tracknum) = tracknum
     End If
'     lastnote(tracknum) = pitch(tracknum) + IIf(instrument(tracknum) > 0, transpose(tracknum), 0)
'     lastnote(tracknum) = pitch(tracknum) + transpose(tracknum)

'  If lastnote(tracknum) >= &H54 Then dontplay = 1
'  If lastnote(tracknum) <= &H1C Then dontplay = 1
'  If lastnote(tracknum) = &H54
'  And instrument(tracknum) = 0 Then dontplay = 1
'  If lastnote(tracknum) = &H54 And instrument(tracknum) = 0 Then dontplay = 1
If dontplay = 0 Then
     note(tracknum) = True
'     PlayNote lastchan(tracknum), lastnote(tracknum), IIf(instrument(tracknum) = 0, volume(tracknum), &H7F)
If exnotes(tracknum) > 0 Then
     For i = 0 To exnotes(tracknum) - 1

If instmap(instrument(tracknum)).MapTo > 127 Then
 sx = drumkits(pitch(i, tracknum))
 If sx < 128 And sx <> lastdrumkit Then
  lastdrumkit = sx
  SetPatch 9, sx
 End If
End If
PlayNote lastchan(tracknum), exnote(i, tracknum), IIf(instmap(instrument(tracknum)).MapTo > 127, (volume(tracknum) / &H7F) * cvelo, cvelo)
If instmap(instrument(tracknum)).SecondNote > 0 Then
snote(snotes(tracknum), tracknum) = exnote(i, tracknum) + instmap(instrument(tracknum)).SecondNote
inc snotes(tracknum)
 PlayNote lastchan(tracknum), exnote(i, tracknum) + instmap(instrument(tracknum)).SecondNote, IIf(instmap(instrument(tracknum)).MapTo > 127, (volume(tracknum) / &H7F) * cvelo, cvelo)
End If
If instmap(instrument(tracknum)).ThirdNote > 0 Then
snote(snotes(tracknum), tracknum) = exnote(i, tracknum) + instmap(instrument(tracknum)).ThirdNote
inc snotes(tracknum)
 PlayNote lastchan(tracknum), exnote(i, tracknum) + instmap(instrument(tracknum)).ThirdNote, IIf(instmap(instrument(tracknum)).MapTo > 127, (volume(tracknum) / &H7F) * cvelo, cvelo)
End If

'        PlayNote lastchan(tracknum), exnote(i - 1, tracknum), IIf(instmap(instrument(tracknum)).MapTo > 127, (volume(tracknum) / &H7F) * velo(i - 1, tracknum), velo(i - 1, tracknum))
        
     Next i
End If
     notewait(tracknum) = notedelay(tracknum)
 
 xnote(tracknum) = False
End If
End If
End If
End If
 
 
 
 
 

 
 If wait(tracknum) > 1 Then Exit Do
Loop
End If
End Sub

