# Sappy

This is the original source for the 1.6 release for archival purposes.

1.6 was the original live MIDI player + MIDI file dumper before it was rewritten into [Sappy 2005](https://github.com/Touched/Sappy) which was when I had handed off the project to Kawa. **Sappy 2005** had additional features such as simulate the mixing of the audio driver, among other fancy things.

I have been looking for the source for this original project for at least 20 years now and had finally come across it.

Beware, the code is *rough*, I was 15 when this I released this and good programming learning resources on the internet were much harder to come by in 2003.

## Directories:

- `src` - contains the Visual Basic 6.0 source code. `prjsapp.vbp` is the project file.
- `dist` - the compiled `sappy.exe` should go into this directory, which contains the other distribution files such as the [readme](./dist/readme_sappy.txt), a [mIRC](https://www.mirc.com/) Now Playing script and data files for specific supported games.

## Source notes

This contains the original source as-is with a few updates:

- Updated the relative paths in the prjsapp.vbp to work with this current directory structure.
- Included the **CDS_Cini** dependency. It does not appear to have any license and looks like it originally comes from this post: https://www.codeguru.com/visual-basic/a-class-for-easy-ini-file-handling/
- Updated some names in the original readme.

And one final caveat, I don't have a Windows setup with VB6 installed to test building the project. I assume that all the necessary files are present and that I have made the correct changes to `.vbp`, but it may still need adjustments. I will provide an update here if/when I get the chance to confirm it all works.
