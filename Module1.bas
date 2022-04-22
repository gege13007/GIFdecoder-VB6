Attribute VB_Name = "Module1"
'https://www.matthewflickinger.com/lab/whatsinagif/
'https://commandlinefanatic.com/cgi-bin/showarticle.cgi?article=art011
'https://stackoverflow.com/questions/67912876/what-exactly-to-do-at-max-code-table-size-followed-by-clear-code-read-in-a-gif-l
'https://www.eecis.udel.edu/~amer/CISC651/lzw.and.gif.explained.html

Type Entry
  long As Long
  code As Integer
  prev As Long
End Type

Global fso As New FileSystemObject
Global logfilobj, logfile

Global gif() As Byte
Global gifptr!

'pour Gif Anims
Global frame_num%
Global display_times%
Global display_ms%
Global display_transpa%
Global display_trans_col%
Global display_disposal As Byte

'dimensions image
Global pixwidth!, pixheight!
Global maxpixels!
'dim & pos du clip
Global clipwidth!, clipheight!
Global clipleft!, cliptop!
'point en cours dans clip
Global clipx!

'global color table x256 xRGB
Global color(256, 2) As Byte

'infos image
Global bgcolindex As Byte
Global field As Byte, pixratio As Byte, interlace As Byte
Global glob_col_tab As Byte
Global glob_coltab_size!, loc_coltab_size!

'pour transform en 9 bits
Global newb9 As Long
Global packleft!
Global newx As Byte
Global mask As Byte
Global loopbits As Byte
Global MSBbit!
Global tablesize!, initablesize!

Global clearcode!, endcode!, oldcode!

'dico
Global dico2(4096) As Entry

'sortie pix
Global pix() As Byte
Global outptr As Long
Global dicoptr As Long

'pour la pile de pixels (ordre à inverser)
'(premiers entrés -> derniers setpix)
Global pixstackptr%
Global pixstack(255) As Byte
