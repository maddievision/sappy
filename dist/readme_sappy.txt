--Sappy---------------------------------------
-----The GBA Music Player---------------------
-------version 1.6 [July 23, 2003]------------
-----------by Bouche (Maddie Lim)-------------
-----------------[REDACTED EMAIL]-------------

What's New:
1.6 [Wednesday, July 23, 2003]

MIDI EXPORT FUNCTION!! :)
 [Important Note: although now you don't have to sequence your own MIDIs, I would ask
  all of you please to NOT submit any MIDIs converted by this program to 
  VGmusic.com, as 1. they do not accept converted MIDIs, 2. it really isn't your
  talent that is being submitted, it's a machine/program doing it for you.
  I'm sure you and I don't want to get into trouble for this.  You may go ahead
  and host these on other sites or use them in your projects, as long as you give
  credit to who made the converter (me! ^_^). I'd really appreciate your help here guys!]

Added support for new games (in no particular order):
 -ESPN Final Round Golf 2002 (U)
 -Wario Land 4 (J)
 -Golden Sun (U)
 -Super Robot Wars (J)
 -Sonic Advance (J)
 -Sonic Advance 2 (U)
 -Advance Wars 2: Black Hole Rising (U)
 -Fire Emblem (J) [thank you Chiklit!]
 -Castlevania - Aria of Sorrow (U) (and partial MAP file)
 -Castlevania: Aria of Sorrow 
 -Tom Clancy's Splinter Cell (U)
 -F-Zero Advance (J)
Added two new .lst files from someone who wishes to remain anonymous.. (have
 yet to be finalised by me)
Added Playlist Groups (see LST file section to see how to set these up, old LST
 format not supported :P)!
You can now change the speed using the <, >, - and +, keys, (with or without shift)
Changed from Buttons to Menu interface ^_^!
Now the Time/Frame counters only update every beat, or when necessary, this
 clears up major speed problems on some computers.
The Play button no longer switches to the next track on Auto Advance mode,
 unless you press any of the "next" buttons while the song is playing (bug)
GBC channels are louder now compared to PCM channels
GBC channels now use the appropiate closest instrument
GBC channels on XG use LSB bank 6
Removed "GB Mode"
Added new status field "Chan" which shows which channel sound is being
 output through (Square 1, Square 2, Triangle, Noise or Direct PCM)

1.4 [Tuesday April 22, 2003]
First Official Public Release
14 Games Supported
Pok�mon Ruby/Sapphire and Mario Kart Super Circuit sound almost perfect.
Advance Wars and Super Mario Advance instrument mapping taken to attention.

0.01 [Saturday April 12, 2003]
Project Start

----------------------------------------------

File Checklist:
readme_sappy.txt	Sappy Readme, this file
sappy.exe		Sappy Executable
sappy.mrc		Sample Sappy mIRCscript
data\sappy.lst		Sappy ROM list
data\blank.lst		Blank Song List
data\advw.lst		Advance Wars Song List
data\cv2.lst		Castlevania AoS Song List
data\advw.lst		Advance Wars Song List
data\marokart.lst	Mario Kart Super Circuit Song List
data\sapphire.lst	Pokemon Sapphire/Ruby Song List
data\kirby.lst	Kirby Song List
data\marioadv.lst	Super Mario Advance Song List
data\blank.map		Blank Instrument Map
data\advw.map		Advance Wars Map
data\cv2.map		Castlevania AoS Map
data\mksc.map		Mario Kart Super Circuit Map
data\sapp.map		Pokemon Sapphire/Ruby Map
data\marioadv.map	Super Mario Advance Map

----------------------------------------------
--Table of Contents---------------------------
1. Introduction
2. What need to be done?
3. Usage
4. Status Dumps
5. Instrument Maps and Song Lists
6. Author & Credits
----------------------------------------------

--1. Introduction-----------------------------
So what in the blazes is Sappy?


Sappy is a program that plays a certain GBA [Gameboy Advance] music 
engine's format through MIDI.

How perfect is it?

It depends on the game and how the game organizes it's instruments,
if it's near MIDI standards (eg Pokemon Sapphire, Mario Kart) then
it's "instrument mapping" needs only to be small (see section 4 for
details on instrument mapping).  For other games, such as Super Mario
Advance, every single instrument may need to be mapped.

For games that haven't had their map done, they've been assigned the
"blank" map.

For games that haven't had their song name list done, they've been
assigned the "blank" list.

--2. What needs to be done?-------------------
As you might gather (or will gather), Sappy is not perfect, it could
(and will) be improved in these ways:
[in no particular order]

* Supporting ALL the commands of the Sappy engine. (There's like 2 or 3 left...)
* Using DirectMusic and DLS banks to output the music in the same way
  it would sound in the original game :)
* Adding support for other GBA engines, such as the one that Super
  Mario Advance 2-4 and Zelda: A Link to the Past 4 Swords use.
* Adding support for some GBC engines. (more details at a later date :p)
* Standalone File Playing, and Ripping
* bunch more stuff, will add to this list later :)

--3. Usage------------------------------------
Usage should be quite straight forward.
Load the ROM with the Load ROM function.
(To see a list of supported ROMs, click the "Supported ROMs" button)
Then press Play for instant action. :)
For games that don't have any song names you'll have to navigate
using the "Song Number" +/- buttons.  For games that do, you can
use the List Box and </> to navigate.  You can choose what navigation
method Shuffle and AutoAdvance use by setting the Radio button
to the respective mode.
Note that there might be some songs that are not on the list box but
are in the song, these are probably unimportant, or sound effects.

With the status viewer (the big black bit with all the track info),
you can change the speed of the music:
< - Halves the current speed
- - Decreases speed by 1
+ - Increases speed by 1
> - Doubles the current speed
(You can also press these keys, with/without shift)

You can mute tracks by unchecking the box next to their number.
The box next to "Trk" controls all tracks.  You can "solo" a single
track by right clicking it, this will enable only this channel.

Loop Limit:
If you want songs to loop forever set loop limit to 0 (infinite).
Otherwise it will loop how many times you set it minus 1.
(1 will make the song play once without looping, 2 will play the 
song and loop once, 3 will loop twice, etc)

--4. Status Dumps------------------------------
Whenever you play a song Sappy dumps statistics info into a file
"sappy.stt" and "sappy.xtt".  The information can be used in
places such as an IRC script to show what song you are playing.
The fields are separated by "|". A sample irc script 'sappy.mrc'
is included with this package.

SAPPY.STT
sappyversion|romfilename|romheader|songnumber|gamename|songname
SAPPY.XTT
numberoftracks|songlayer|songaddress|songinstaddress

--5. Instrument Maps and Song Lists------------
In order for a game's instruments to sound correct, you need
to define an instrument map for it. The best examples are the
ones included with this package.  The games which don't have
song list and map info, are directed to the blank templates
"blank.lst" and "blank.map".

Here's how a song list works:
xxxx songname
Where xxxx is the song number (it must be 4 digits hexadecimal),
and songame is the song's name (obviously).
For example:
0002 Title Screen
0195 Littleroot Town
005E Dying Sound Effect

It should be noted that I haven't added any game's sound effects
to the song list as of now.

Instrument maps work like this:
For each element of the instrument map is a block, which can be
one of the following: inst, drum, or noise (you may see
envelope_pitch/envelope_volume/end_envelope in some maps, these 
are not documented yet as they are not yet supported by Sappy itself).

Block arguments:

inst
a - MIDI instrument to map FROM (0-127)
b - MIDI instrument to map TO (0-127)
    if for some reason sappy does not detect a drum/noise channel, you
    can map it yourself.  drum is instrument 128, and noise is instrument
    129.
c - transposition, as you know, samples can be at the wrong pitch,
    this is here to allow transposing of the sample to match the
    correct MIDI pitch (signed value)
d - secondary note, some instruments are recorded like a single chord,
    second/third notes are here to accompany for those,  they work
    like the tranposition argument, but they apply to a second and third
    note, to disable any of these two, set them to 0. (signed value)
e - see argument d (signed value)
f - ignored (will be used for volume envelope ID)
g - ignored (will be used for pitch envelope ID)

drum
a - MIDI key to map FROM
b - MIDI key to map TO

c - Drum kit patch, some drums may sound better on a particular drum kit
    or some games might have more than one set of drums that may sound
    better on their respective kits.  This is to accomodate for those.
    For notes that drum kit patch doesn't really matter (eg Hi-hats),
    you can set them to drum patch 128, which tells Sappy not to change
    drum kits upon this note.

noise
---noise works exactly the same as drum, except it is for noise instruments

Finally if you want to add a game to the list, or add a map for one of the
existing games that use blank, you'll need to know how sappy.lst works.

The first line is for unsupported games, which is not currently used yet,
just leave it.  Also the 'error' games with rom type as 'error' should be
left at the very end of the list.

romheader,
 romname, romtype, songlist, instmap, tableoffset, songstart, 
   songend


romheader - the four character ROM header in the game.  you can wildcard 
            the last character (language) with �. You should only do this
            if all versions are identical in where the music is located

romname - the name of the game (including the (x) country spec.)

romtype - must be 'sapphire', or it will not play in Sappy.
songlist - filename of the songlist (without the ".lst" extension)
instmap - filename of the instrument map (without the ".map" extension)
tableoffset - address of the song table in the ROM
songstart - song number to start on (used only if a blank songlist is used)
songend - song number to end on, not used at all

--6. Author and Credits------------------------

This program is by Bouche, Maddie Lim, who is 15 and has been ROM Hacking
for more than 3 years.  She has been music hacking for more than a year.

You can catch me at the following:
'BoucheanBouche' or 'DJ Bouche' on AOL Instant Messenger
'69427310' on ICQ
'[REDACTED EMAIL]' on E-mail AND Windows/MSN Messenger
'[REDACTED USERNAME]' on Yahoo! Messenger
'DJ Bouche' on Acmlm's Board (http://acmlm.overclocked.org/board)
'#acmlm', '#romhacking' and "#rom-hacking" on EmuNET 
 (irc server: Akron.OH.US.irc.acmlm.org, port 6667)

Special Thanks:

I'd like to thank the community of AcmlmBoard for supporting me in
ROM Hacking and for sharing their useful information.  I'd like to also
thank Pikachu14/Kyoufu Kawa for giving me much inspiration into hacking
Pokemon Sapphire/Ruby, in which lead me into cracking this music format.
I'd also like to thank the people who have sent in feedback and support
for the program, in the forms of email, good words, new game entries,
instrument maps and song lists. Your help is greatly appreciated ^_^.

I'll put this into more detail later.

-----------------------------------------------
Enjoy the program! Send any bug reports or requests to any of my 
contacts above. (ie, such as a game not loading right, an interface
bug, a non-working map or song list, et al)
And MOST of all, do NOT ask me for ROMs, find them
the way everyone else does, or dump your own :).
