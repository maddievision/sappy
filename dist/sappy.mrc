;Sappy mIRCScript, set sappydir to your Sappy directory
;Use Shift+F5 or /sappyinfo to output to channel :)
;Customize all you want!

alias sappydir {
  return C:\sappy
  ;set this to your Sappy directory
}


alias sappy {
  %sappyx = $read(" $+ $sappydir $+ \sappy.stt",1)
  %sappy.version = $gettok(%sappyx,1,124)
  %sappy.filename = $gettok(%sappyx,2,124)
  %sappy.header = $gettok(%sappyx,3,124)
  %sappy.songnumber = $gettok(%sappyx,4,124)
  %sappy.game = $gettok(%sappyx,5,124)
  %sappy.songname = $gettok(%sappyx,6,124)
  %sappyx = $read(" $+ %sappydir $+ \sappy.xtt",1)
  %sappy.tracks = $gettok(%sappyx,1,124)
  %sappy.layer = $gettok(%sappyx,2,124)
  %sappy.address = $gettok(%sappyx,3,124)
  %sappy.instaddress = $gettok(%sappyx,4,124)
}

alias sF5 {
  sappyinfo
}

alias sappyinfo {
  sappy
  me - [Sappy v $+ %sappy.version $+ ]  GBA Music - %sappy.game - ( $+ %sappy.songnumber $+ ) %sappy.songname < $+ %sappy.tracks tracks/MIDI/Stereo>
}

alias F5 {
  sappy
  clipboard [Sappy v $+ %sappy.version $+ ]  GBA Music - %sappy.game - ( $+ %sappy.songnumber $+ ) %sappy.songname < $+ %sappy.tracks tracks/MIDI/Stereo>
}
