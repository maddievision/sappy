Attribute VB_Name = "mdlInSapphire"
Public Sub parsecom(ByVal tracknum As Byte)
If disabledtracks(tracknum) = True Then Exit Sub
gb1 = 1.5
gb2 = 1.5
timepass(tracknum) = True

If notewait(tracknum) > 1 Then
  dec notewait(tracknum)
'  Form1.Label8(tracknum) = notewait(tracknum) \ 2
  'Form1.Label8(tracknum).ForeColor = RGB(255, 255, 0)
  Form1.Label5(tracknum).BackStyle = 1
  Form1.Label20(tracknum).BackStyle = 1
  Form1.Label21(tracknum).BackStyle = 1
    
  ElseIf note(tracknum) = True Then

If vex(tracknum) = False And instmap(instrument(tracknum)).Sustain = False Then
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
 Case &HB1
 disabledtracks(tracknum) = True
 xbm = True
For lll = 0 To numtracks - 1
   If lll = 9 Then lll = lll + 1
   If disabledtracks(lll) = False Then xbm = False
Next lll
 
 With Form1
 .Label2(tracknum) = ""
 .Label8(tracknum) = ""
 .Label1(tracknum) = ""
  .Label16(tracknum) = ""
   .Label20(tracknum) = ""
.Label20(tracknum).BackStyle = 0
.Label21(tracknum) = ""
.Label21(tracknum).BackStyle = 0
 .Label17(tracknum) = ""
 .Label5(tracknum) = ""
.Label5(tracknum).BackStyle = 0
 .Label9(tracknum) = ""
  .LabelC(tracknum) = ""

 .Label9(tracknum).BackStyle = 0
 .Label6(tracknum) = ""
 .Label16(tracknum).BackStyle = 0
 .Label7(tracknum) = ""
 .Label17(tracknum).BackStyle = 0
End With
 
 xstop = xbm
   
   
   
   Exit Do
   inc pc(tracknum), 2
  
  Case &HB2
   If tracknum = 0 Then inc loops
   pc(tracknum) = getgbapointer(pc(tracknum) + 1)
 Case &HB3
   loopreturn(tracknum) = pc(tracknum) + 5
   inloop(tracknum) = True
   pc(tracknum) = getgbapointer(pc(tracknum) + 1)
 Case &HB4
   If inloop(tracknum) = True Then
   pc(tracknum) = loopreturn(tracknum)
   inloop(tracknum) = False
   Else
   inc pc(tracknum)
   End If
 Case &HB5
   inc pc(tracknum), 2
 Case &HB6
   inc pc(tracknum), 2
 Case &HB7
   inc pc(tracknum), 2
 Case &HB8
   inc pc(tracknum), 2
 Case &HB9
   Get #256, pc(tracknum) + 2, argx
   Get #256, pc(tracknum) + 3, argy
sp = Hex(argy) & "." & Hex(argx) & "."
   Get #256, pc(tracknum) + 4, argx
sp = sp & Hex(argx)
Form1.Label19 = sp
   inc pc(tracknum), 4
    
 Case &HBA
   inc pc(tracknum), 2
  Case &HBB
   Get #256, pc(tracknum) + 2, argx
   speed = argx
   settempo argx
'   volcmd(tracknum) = 6
   inc pc(tracknum), 2
  Case &HBC
   Get #256, pc(tracknum) + 2, argx
 Transpose(tracknum) = IIf(argx < &H81, argx, CInt(argx) - CInt(&H100))
 
   inc pc(tracknum), 2
 Case &HBD
   Get #256, pc(tracknum) + 2, argx
   instrument(tracknum) = argx
'   If instrument(tracknum) = 1 Then instrument(tracknum) = 0
'     If instrument(tracknum) > 119 Then
'   instrument(tracknum) = 0
'   If instrument(tracknum) = 119 Then transpose(tracknum) = -12
'   End If
'   If instrument(tracknum) > 118 Then instrument(tracknum) = 80
   Form1.Label6(tracknum) = (argx)
   If instmap(argx).MapTo < 128 Then
   If Form1.cGBMode.value = vbChecked Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 80
   Else
   SetPatch tracknum, instmap(argx).MapTo
   End If
   Form1.LabelC(tracknum) = "DIR"
   ElseIf instmap(argx).MapTo = 252 Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 80
   Form1.LabelC(tracknum) = "SQ1"
   ElseIf instmap(argx).MapTo = 253 Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 80
   Form1.LabelC(tracknum) = "SQ2"
   ElseIf instmap(argx).MapTo = 254 Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 81
   Form1.LabelC(tracknum) = "TRI"
   ElseIf instmap(argx).MapTo = 129 Then
   SetPatch 9, 0
   Form1.LabelC(tracknum) = "NOI"
   Else
   SetPatch 9, 0
   Form1.LabelC(tracknum) = "DIR"
   End If
   inc pc(tracknum), 2
    volcmd(tracknum) = &HBD
 Case &HBE
   Get #256, pc(tracknum) + 2, argx
   volume(tracknum) = argx
   volcmd(tracknum) = 4
    Form1.Label7(tracknum) = argx
If instmap(instrument(tracknum)).MapTo > 251 Then
 mv = argx * gb1
 If mv > 127 Then mv = 127
 SetVolume tracknum, mv
Else
 SetVolume tracknum, argx
End If
   inc pc(tracknum), 2
 Case &HBF
   Get #256, pc(tracknum) + 2, argx
   spd2 = argx
volcmd(tracknum) = 6
SetPanning tracknum, spd2
Form1.Label17(tracknum) = panx(spd2)
Form1.Label17(tracknum).BackStyle = IIf(spd2 = &H40, 0, 1)
   inc pc(tracknum), 2
 Case &HC0
   Get #256, pc(tracknum) + 2, argx
SetPitchBend tracknum, ((argx / 2) * &H100) '(argx - &H40) + &H1FFF
Form1.Label16(tracknum) = argx - 64
Form1.Label16(tracknum).BackStyle = IIf(argx = &H40, 0, 1)
     volcmd(tracknum) = 5
   inc pc(tracknum), 2

 Case &HC1
    Get #256, pc(tracknum) + 2, argx
   SetPitchBendRange tracknum, argx
   inc pc(tracknum), 2
'inc pc(tracknum), 1
 Case &HC2
   inc pc(tracknum), 2
   SetNRPN tracknum, 1, 8, argx
 Case &HC3
   inc pc(tracknum), 2
 Case &HC4
   Get #256, pc(tracknum) + 2, argx
'volcmd(tracknum) = argx
   SetNRPN tracknum, 1, 9, (argx * 3)
SetControl numtrack, 1, IIf(argx = 0, 0, 127)
 Form1.Label9(tracknum) = argx
Form1.Label9(tracknum).BackStyle = IIf(argx = 0, 0, 1)
volcmd(tracknum) = &HC4
   inc pc(tracknum), 2
 Case &HC5
   inc pc(tracknum), 1
 Case &HC6
   inc pc(tracknum), 1
 Case &HC7
   inc pc(tracknum), 1
 Case &HC8
   inc pc(tracknum), 1
 Case &HC9
   inc pc(tracknum), 1
 Case &HCA
   inc pc(tracknum), 1
 Case &HCB
   inc pc(tracknum), 1
 Case &HCC
   inc pc(tracknum), 1
 Case &HCD
    Get #256, pc(tracknum) + 2, argx
   notedot(tracknum) = argx
   inc pc(tracknum), 2
   volcmd(tracknum) = &HCD
 Case &HCE
'SetControl tracknum, 64, 0
If note(tracknum) = True Then
  If exnotes(tracknum) > 0 Then
     For i = 1 To exnotes(tracknum)
        NoteOff lastchan(tracknum), exnote(i - 1, tracknum)
     
     Next i
For i = 0 To snotes(tracknum) - 1
If snotes(tracknum) = 0 Then Exit For
NoteOff lastchan(tracknum), snote(i, tracknum)

Next i
End If
End If
   inc pc(tracknum), 1
   vex(tracknum) = False
 Case &HCF
'SetControl tracknum, 64, 127
xnote(tracknum) = True
      volcmd(tracknum) = 0
      inc pc(tracknum), 1
   vex(tracknum) = True
 Case Is < &H80
If volcmd(tracknum) = 0 Then
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

n2p = com + Transpose(tracknum) + instmap(instrument(tracknum)).Transpose
pitch(exnotes(tracknum), tracknum) = n2p
If instmap(instrument(tracknum)).MapTo = 129 Then
 exnote(exnotes(tracknum), tracknum) = noisemap(n2p)
ElseIf instmap(instrument(tracknum)).MapTo > 127 And instmap(instrument(tracknum)).MapTo < 252 Then
 exnote(exnotes(tracknum), tracknum) = drummap(n2p)
Else
exnote(exnotes(tracknum), tracknum) = n2p
End If
velo(exnotes(tracknum), tracknum) = lastvelo(tracknum)
arg3(exnotes(tracknum), tracknum) = lastarg3(tracknum)
   
     xnote(tracknum) = True
'Else
'     xnote(tracknum) = True
'End If
inc pc(tracknum)
'If instrument(tracknum) > 0 Then
 Get #256, pc(tracknum) + 1, argx
Get #256, pc(tracknum) + 2, argy
 
 If argx < &H80 Then
' And instrument(tracknum) > 0 Then
'If instrument(tracknum) = 0 Then
'exnote(0, tracknum) = argx + transpose(tracknum)
'exnotes(tracknum) = 1
'Else
velo(exnotes(tracknum), tracknum) = argx
lastvelo(tracknum) = velo(exnotes(tracknum), tracknum)
' SetVolume tracknum, argx
''''''''''''''''' Form1.Label7(tracknum) = argx
' notedelay(tracknum) = argx - &H5C
'End If
 inc pc(tracknum)
   If argy < &H80 Then
arg3(exnotes(tracknum), tracknum) = argy
lastarg3(tracknum) = argy
      
   inc pc(tracknum)
   
   End If
' If argy < &H80 Then
' inc pc(tracknum)
'
' End If
 'End If
End If
inc exnotes(tracknum)
Form1.Label20(tracknum) = Form1.Label20(tracknum) & IIf(exnotes(tracknum) > 1, "/", "") & velo(exnotes(tracknum) - 1, tracknum)
Form1.Label21(tracknum) = Form1.Label21(tracknum) & IIf(exnotes(tracknum) > 1, "/", "") & arg3(exnotes(tracknum) - 1, tracknum)

ElseIf volcmd(tracknum) = 4 Then
Get #256, pc(tracknum) + 1, argx
' Get #256, pc(tracknum) + 2, argy
 
' If argx < &H80 Then
 volume(tracknum) = argx
If instmap(instrument(tracknum)).MapTo > 251 Then
 mv = argx * gb1
 If mv > 127 Then mv = 127
 SetVolume tracknum, mv
Else
 SetVolume tracknum, argx
End If
 Form1.Label7(tracknum) = argx
' notedelay(tracknum) = argx - &H5C
'   If argy < &H80 Then
 '  inc pc(tracknum)
  ' End If
' inc pc(tracknum)
' If argy < &H80 Then
 inc pc(tracknum)
'
' End If
 'End If
ElseIf volcmd(tracknum) = 5 Then
Get #256, pc(tracknum) + 1, argx

SetPitchBend tracknum, ((argx / 2) * &H100) '(argx - &H40) + &H1FFF 'IIf(argx > &H80, argx - &H80, argx) + &H3FFF
Form1.Label16(tracknum) = argx - 64
Form1.Label16(tracknum).BackStyle = IIf(argx = &H40, 0, 1)
inc pc(tracknum)
ElseIf volcmd(tracknum) = 6 Then
Get #256, pc(tracknum) + 1, argx
spd2 = argx
SetPanning tracknum, spd2
Form1.Label17(tracknum) = panx(spd2)
Form1.Label17(tracknum).BackStyle = IIf(spd2 = &H40, 0, 1)
inc pc(tracknum)
ElseIf volcmd(tracknum) = &HBD Then
   Get #256, pc(tracknum) + 1, argx
   instrument(tracknum) = argx
'   If instrument(tracknum) = 1 Then instrument(tracknum) = 0
'     If instrument(tracknum) > 119 Then
'   instrument(tracknum) = 0
'   If instrument(tracknum) = 119 Then transpose(tracknum) = -12
'   End If
'   If instrument(tracknum) > 118 Then instrument(tracknum) = 80
   Form1.Label6(tracknum) = (argx)
   If instmap(argx).MapTo < 128 Then
   If Form1.cGBMode.value = vbChecked Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 80
   Else
   SetPatch tracknum, instmap(argx).MapTo
   End If
   Form1.LabelC(tracknum) = "DIR"
   ElseIf instmap(argx).MapTo = 252 Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 80
   Form1.LabelC(tracknum) = "SQ1"
   ElseIf instmap(argx).MapTo = 253 Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 80
   Form1.LabelC(tracknum) = "SQ2"
   ElseIf instmap(argx).MapTo = 254 Then
   SetControl tracknum, 32, 6
   SetPatch tracknum, 81
   Form1.LabelC(tracknum) = "TRI"
   ElseIf instmap(argx).MapTo = 129 Then
   SetPatch 9, 0
   Form1.LabelC(tracknum) = "NOI"
   Else
   SetPatch 9, 0
   Form1.LabelC(tracknum) = "DIR"
   End If
   inc pc(tracknum), 1

ElseIf volcmd(tracknum) = &HC4 Then
Get #256, pc(tracknum) + 1, argx
   SetNRPN tracknum, 1, 9, (argx * 2)
SetControl numtrack, 1, IIf(argx = 0, 0, 127)
 Form1.Label9(tracknum) = argx
Form1.Label9(tracknum).BackStyle = IIf(argx = 0, 0, 1)
inc pc(tracknum)
Else
inc pc(tracknum)
End If
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
 Case Is < &HB1
     Form1.Label1(tracknum) = (notelen(com - &H80))
    wait(tracknum) = (notelen(com - &H80) * 2) * IIf((notedot(tracknum) Mod &HC) = 1, 0.75, 1)
'If vex(tracknum) = True Then xnote(tracknum) = True
     inc pc(tracknum)
 
 
 
 
 
 
 
 
 
 
 
 
'  Get #256, pc(tracknum) + 1, argx
' If argx = &HC0 Then inc pc(tracknum)
 
 
 
 
 
 
 
 
 
 
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
     If instmap(instrument(tracknum)).MapTo > 127 And instmap(instrument(tracknum)).MapTo < 252 Then
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

If instmap(instrument(tracknum)).MapTo = 129 Then
 sx = noisekits(pitch(i, tracknum))
 If sx < 128 And sx <> lastdrumkit Then
  lastdrumkit = sx
  SetPatch 9, sx
 End If
ElseIf instmap(instrument(tracknum)).MapTo > 127 And instmap(instrument(tracknum)).MapTo < 252 Then
 sx = drumkits(pitch(i, tracknum))
 If sx < 128 And sx <> lastdrumkit Then
  lastdrumkit = sx
  SetPatch 9, sx
 End If
End If
If instmap(instrument(tracknum)).MapTo > 251 Then
cvelo = velo(i, tracknum) * gb2
Else
cvelo = velo(i, tracknum)
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
        
        MidiAfterTouch lastchan(tracknum), arg3(i, tracknum)
     Next i
End If
     notewait(tracknum) = notedelay(tracknum)
 
 xnote(tracknum) = False
End If
End If
End If

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 Case Is > &HCF
     notedelay(tracknum) = (notelen(com - &HD0) * 2) + 1
     Form1.Label8(tracknum) = notelen(com - &HD0) + 1
'     notedelay(tracknum) = (com - &HD0) * 2
'     Form1.Label8(tracknum) = com - &HD0
     inc pc(tracknum)
     xnote(tracknum) = True
   volcmd(tracknum) = 0
 Case Else
   MsgBox "unknown code: " & Hex(com)
   xstop = True
   Exit Do
 End Select
 If wait(tracknum) > 1 Then Exit Do
Loop
End If
End Sub
Public Sub callback()

songpointer = (Form1.tstart * 8) + Form1.ttable
Form1.Label3 = Hex(songpointer)
lmf &H80, &H0
lmf &H81, &H1
lmf &H82, &H2
lmf &H83, &H3
lmf &H84, &H4
lmf &H85, &H5
lmf &H86, &H6
lmf &H87, &H7
lmf &H88, &H8
lmf &H89, &H9
lmf &H8A, &HA
lmf &H8B, &HB
lmf &H8C, &HC
lmf &H8D, &HD
lmf &H8E, &HE
lmf &H8F, &HF
lmf &H90, &H10
lmf &H91, &H11
lmf &H92, &H12
lmf &H93, &H13
lmf &H94, &H14
lmf &H95, &H15
lmf &H96, &H16
lmf &H97, &H17
lmf &H98, &H18
lmf &H99, &H1C
lmf &H9A, &H1E
lmf &H9B, &H20
lmf &H9C, &H24
lmf &H9D, &H28
lmf &H9E, &H2A
lmf &H9F, &H2C
lmf &HA0, &H30
lmf &HA1, &H34
lmf &HA2, &H36
lmf &HA3, &H38
lmf &HA4, &H3C
lmf &HA5, &H40
lmf &HA6, &H42
lmf &HA7, &H44
lmf &HA8, &H48
lmf &HA9, &H4C
lmf &HAA, &H4E
lmf &HAB, &H50
lmf &HAC, &H54
lmf &HAD, &H58
lmf &HAE, &H5A
lmf &HAF, &H5C
lmf &HB0, &H60
lmf &HB1, &H71
lmf &HB2, &H72
lmf &HB3, &H73
lmf &HB4, &H74
lmf &HB5, &H75
lmf &HB6, &H76
lmf &HB7, &H77
lmf &HB8, &H78
lmf &HB9, &H99
lmf &HBA, &H9A
lmf &HBB, &H9B
lmf &HBC, &H9C
lmf &HBD, &H9D
lmf &HBE, &H9E
lmf &HBF, &H9F

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

Open FileName For Binary As #256
Get #256, songpointer + 5, songlayer
Form1.Label22 = songlayer
songpointer = getgbapointer(songpointer)

If songpointer = -1 Then GoTo hell
Form1.Label4 = Hex(songpointer)
Get #256, songpointer + 1, numtracks
If numtracks = 0 Then GoTo hell
For i = 0 To numtracks
z = IIf(i > 8, i + 1, i)
volcmd(z) = 4
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



intable = getgbapointer(songpointer + 4)
Form1.Label15 = Hex(intable)
Open IIf(Form1.cD.value = vbChecked, debugpath, App.Path & "\") & "sappy.xtt" For Output As #4
Print #4, numtracks & "|" & songlayer & "|"; Hex(songpointer) & "|" & Hex(intable)
Close #4
Dim r As Byte
If Form1.cIMAP.value = vbUnchecked Then
loadinstmap custommap
End If
For i = 0 To ((128 * &HC) - 1) Step &HC
Get #256, i + intable + 1, r
If r = &H80 Then
instmap(i \ 12).MapTo = 128
ElseIf r Mod 8 = &H4 Then
instmap(i \ 12).MapTo = 129
ElseIf r Mod 4 = 1 Then
instmap(i \ 12).MapTo = 252
ElseIf r Mod 4 = 2 Then
instmap(i \ 12).MapTo = 253
ElseIf r Mod 4 = 3 Then
instmap(i \ 12).MapTo = 254
End If
Next i

For i = 0 To numtracks - 1
disabledtracks(i) = False
Form1.Label2(j) = Hex(pc(j))
z = IIf(i > 8, i + 1, i)
pc(z) = getgbapointer(songpointer + 8 + (i * 4))
Next i
disabledtracks(9) = True
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
parsecom z
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


