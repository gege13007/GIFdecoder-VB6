VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   662
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox loopchk 
      Caption         =   "loop"
      Height          =   330
      Left            =   7920
      TabIndex        =   4
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox logtxt 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4440
      Width           =   7260
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GIF parse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7800
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDlg 
      Left            =   4080
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pict 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   4800
      ScaleHeight     =   176
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label gifptrlbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Noto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub Command2_Click()
Dim x!, nn!, nxx As Long
Dim maxdicoptr!
Dim pref$, ver$, f$, t$, f2$
Dim nbitcode As Byte
Dim col As Byte, x2%, y2%
Dim a%, b%, ext%
'Chars pour test extension blocs type
Dim c1 As Byte, c2 As Byte, c3 As Byte

'Ouvre le fichier gpx entrée en lecture
CommonDlg.ShowOpen
f$ = CommonDlg.FileName
If InStr(UCase(f$), ".GIF") < 1 Then
  MsgBox "Pas de fichier GIF !"
  Unload Me
End If

pict.Picture = LoadPicture(f$)
Me.Cls
Me.Refresh

RESTARTIMAGE:

'Ouvre le fichier histo en lecture
Open (f$) For Binary As #1
ReDim gif(LOF(1))
Get #1, , gif
gifptr = LOF(1)
Close #1
gifptr = 0

'---------------------------------------------
'         H E A D E R   G I F  (6)
'---------------------------------------------
'Prefixe GIF
pref$ = Chr$(getgif()) + Chr$(getgif()) + Chr$(getgif())
'Version GIF
ver$ = Chr$(getgif()) + Chr$(getgif()) + Chr$(getgif())

'---------------------------------------------
'        LOGICAL  SCREEN  DESCRIPTOR (7)
'---------------------------------------------
pixwidth = getgif() + 256 * getgif()
pixheight = getgif() + 256 * getgif()

a = pixwidth: If a < 220 Then a = 250
Form1.Width = (2 * a + 200) * Screen.TwipsPerPixelX
pict.Left = (150 + a)
Form1.Height = (pixheight + 300) * Screen.TwipsPerPixelY
logtxt.Top = pixheight + 50
Command2.Top = logtxt.Top
loopchk.Top = logtxt.Top + 40
logtxt.Text = ""

'redim le canvas
maxpixels = pixwidth * pixheight
ReDim pix(maxpixels) As Byte

field = getgif()
'Test Global color table ?
glob_coltab_size = 2 ^ ((field And 7) + 1)
glob_col_tab = (field And 128)
bgcolindex = getgif()
pixratio = getgif()

f2$ = Mid$(f$, InStr(f$, "\"))
f2$ = Mid$(f2$, 1 + InStr(f2$, "\"))
f2$ = Mid$(f2$, 1 + InStr(f2$, "\"))
Me.Caption = f2$
log ("> " + f2$ + "  " + CStr(pixwidth) + "x" + CStr(pixheight) + " pixels")

frame_num = 0

'Ouvre le fichier .h en écriture
'Set logfilobj = CreateObject("Scripting.FileSystemObject")
'Set logfile = logfilobj.CreateTextFile(App.Path + "\log_b9.txt", True)
'logfile.WriteLine (f2$)
'logfile.WriteLine ("")

'--------------------------------------------
'         Global Color Table (si oui)
'--------------------------------------------
If glob_col_tab > 0 Then
  For nxx = 0 To glob_coltab_size - 1
    color(nxx, 0) = getgif()
    color(nxx, 1) = getgif()
    color(nxx, 2) = getgif()
 Next nxx
End If

'---------------------------------------------
'      Re-bouclage éventuel si gif-anim !
'---------------------------------------------
Do
 Me.MousePointer = vbHourglass

'passe eventuelles extensions
Do
 x = getgif()
 While x = 0: x = getgif(): Wend
 'securité fin fichier ?
 If x = &H3B Then Exit Do
 
 'Test si Extension ?
 If x <> &H21 Then Exit Do
   'Type extension ?
   ext = getgif()
   Select Case ext
   Case &HFF
   '----------------------------------------------
   '            Application Extension
   '----------------------------------------------
     b = getgif()
     'capte les 3 premiers chars pour test
     'soit 'NETSCAPE' (gifanim) soit 'XMP' (???)
     c1 = getgif()
     c2 = getgif()
     c3 = getgif()
     For a = 1 To b - 3: x = getgif(): Next a
     
     'Si NETSCAPE (gifanim)
     If c1 = Asc("N") And c2 = Asc("E") And c3 = Asc("T") Then
       a = getgif()  ' must be = 3
       b = getgif()  ' must be = 1
       display_times = getgif()                       ' \
       display_times = display_times + 256 * getgif() ' / n fois loops (0 = infini)
       If display_times = 0 Then loopchk.Value = Checked
     Else
     'Si XMP vide ???
      While x <> 0: x = getgif(): Wend
     End If
     
   '----------------------------------------------
   ' Graphics Control extension (Anim & Transpar)
   '----------------------------------------------
   Case &HF9
     'len of data sub-block
     a = getgif()       ' must be = 4 (byte size)
     b = getgif()       ' packed fields - graphics disposal
     '>0 si color transparence
     display_transpa = b And 1
     '01 (dont dispose / graphic), 02 (overwrite graph / background color), 04 (overwrite graphic with previous graphic)
     display_disposal = (b And &H1C) / 4
     display_ms = getgif()                     ' \
     display_ms = display_ms + 256 * getgif()  ' / ms / frames
     'color transparence index
     display_trans_col = getgif()           ' Transpar Color Index
     
   Case Else
   '--------------------------------------------
   '            'Comment' Extension
   '--------------------------------------------
     b = getgif()
     t$ = ""
     For a = 1 To b
       x = getgif(): If x > 31 Then t$ = t$ + Chr$(x)
     Next a
     log ">Comment: " + t$
   End Select
   
   'doit être = 0 !
   x = getgif()
'sort de la série 'Extensions'
Loop

'securité fin fichier TRAILER ?
 If x = &H3B Then Exit Do
 
'--------------------------------------------
'              Image  Descriptor
'--------------------------------------------
If x <> &H2C Then MsgBox "Problème 2C"

clipleft = getgif() + 256 * getgif()   ' clip left pos
cliptop = getgif() + 256 * getgif()    ' clip top pos
clipwidth = getgif() + 256 * getgif()
clipheight = getgif() + 256 * getgif()
field = getgif()
'Lobal color table ?
loc_coltab_size = 2 ^ ((field And 7) + 1)
interlace = field And 64

'--------------------------------------------
'            Local Color Table ?
'--------------------------------------------
If loc_coltab_size > 4 Then         ' !!!!!!!!!!!! ICI !!!!!!!!!!!!
  For nxx = 0 To loc_coltab_size - 1
    color(nxx, 0) = getgif()
    color(nxx, 1) = getgif()
    color(nxx, 2) = getgif()
 Next nxx
End If

'---------------------------------------------
'           I M A G E     D A T A S
'---------------------------------------------
initablesize = (1 + getgif())  ' 9 bits au debut
tablesize = initablesize
MSBbit = 2 ^ (initablesize - 1)
maxdicoptr = (2 ^ tablesize) - 1
clearcode = 2 ^ (tablesize - 1)
endcode = 1 + clearcode
 
'Init ptr pix
outptr = clipleft + cliptop * pixwidth
clipx = 0

'Lecture par paquets de 'byteslong' <255
'si packleft=0 -> la première lecture est une taille de bloc !
packleft = 0
mask = initablesize

' A VOIR ! A VOIR !!!!
If display_disposal = 2 Then
  For nxx = 0 To UBound(pix) - 1
    pix(nxx) = bgcolindex
  Next nxx
End If
  
'--------------------------------------------------
'                    DECODE  L Z W
'--------------------------------------------------
Raz_dico
 
Do
  newb9 = getCode()
       
  If newb9 = clearcode Then
    tablesize = initablesize
    MSBbit = 2 ^ (tablesize - 1)
    maxdicoptr = (2 ^ tablesize) - 1
    Raz_dico
    oldcode = newb9
    newb9 = getCode()
    'Ajouté après ? !
    Setpix (dico(oldcode).code)
  End If
  
  ' Increase Size -> GIF89a mandates that this stops at 12 bits
  If tablesize < 12 Then
    If dicoptr = maxdicoptr Then
      tablesize = tablesize + 1
      MSBbit = 2 ^ (tablesize - 1)
      maxdicoptr = (2 ^ tablesize) - 1
    End If
  End If
  
  If newb9 = endcode Then Exit Do
    
  'YES code is in the code table
  If dico(newb9).code > -1 Then
    a = dico(newb9).long
    b = newb9
    Do
      pushpix ((dico(b).code))
      b = dico(b).prev
      If b > 0 And b < clearcode Then
        pushpix (b)
        Exit Do
      End If
      If a = 0 Then Exit Do
      a = a - 1
     Loop
     'affiche les pixels dans l'ordre inverse
     poppix

  Else
    'Not in DICO : out {code-1}+K to stream
    a = dico(oldcode).long
    b = oldcode
    pushpix ((dico(b).code))
    Do
      pushpix ((dico(b).code))
      b = dico(b).prev
      If b > 0 And b < clearcode Then
        pushpix (b)
        Exit Do
      End If
      If a = 0 Then Exit Do
      a = a - 1
     Loop
     'affiche les pixels dans l'ordre inverse
     poppix
  
  End If
  
  'ADD to DICO S+first symbol / {code-1}+K
  dico(dicoptr).prev = oldcode
  dico(dicoptr).long = dico(oldcode).long + 1
  'cherche le premier char(code) en récursif si long>1 avec .prev
  b = newb9
  If b < clearcode Then
    dico(dicoptr).code = dico(b).code
  Else
  Do
    a = dico(b).long
    If a > 1 Then
      b = dico(b).prev
    Else
      dico(dicoptr).code = dico(b).prev
      Exit Do
    End If
  Loop
  End If
  If dicoptr < 4096 Then dicoptr = dicoptr + 1
      
  oldcode = newb9
Loop

frame_num = frame_num + 1
log ("> " + CStr(frame_num) + ". disposal=" + CStr(display_disposal) + ", transp:" + CStr(display_transpa) + " ,transcol=" + colstr(display_trans_col) + " ,bgcol=" + colstr(bgcolindex) _
 + "  Clip:" + CStr(clipleft) + "/" + CStr(cliptop) + " " + CStr(clipwidth) + "/" + CStr(clipheight))
  
Me.MousePointer = vbDefault
Call drawpix
Sleep (2 * display_ms)   '  PLUS LENT !!!

Loop

finprog:
'logfile.Close
Me.MousePointer = vbDefault

 If loopchk.Value = Checked Then GoTo RESTARTIMAGE

End Sub

'---------------------------------------------------
' Lit les octets du .gif et passe en code (newb9) de
' longueur variable ...de 8 à 9 bits (ou plus...)
' Gère les paquets d'octets (packleft)
'---------------------------------------------------
'Sort si packleft=0 : plus de paquets à lire (getCode = endcode)
'Sort si fin du fichier dépassée (getCode = endcode)
Function getCode()
Dim nn%
 'enquille les 9 bits - rentre sur le 8 - shift vers droite
 For nn = 1 To tablesize
   'shift droit mot 9 bits destination
   newb9 = (newb9 And &HFFE) / 2
   'met MSB dans b9 précédent
   If ((newx And 1) <> 0) Then newb9 = newb9 + MSBbit
   'shift droit l'octet en cours de lecture
   newx = (newx And &HFE) / 2
   
   mask = mask - 1
   'un code (newb9) est-il prêt ?
   If mask = 0 Then
     'nouveau pack de 8 bits (ou moins ?)
     mask = 8
     'Test si charger nouveau paquet de gif bytes ?
     If packleft = 0 Then
       packleft = getgif()
       
       'POUR TEST
       gifptrlbl.Caption = CStr(gifptr) + " bytes"
       'POUR SECURITE
       If gifptr > UBound(gif) Then
         getCode = endcode
         Exit Function
       End If
       
       'logfile.WriteLine ("paquet > " + str4$(packleft) + " dicoptr:" + CStr(dicoptr))
       'test si plus de paquets picture datas
       If packleft = 0 Then
         getCode = endcode
         Exit Function
       End If
     End If
     newx = getgif()
     packleft = packleft - 1
   End If
    
 Next nn
 
 getCode = newb9
End Function

'--------------------------------------------------------
'             Dessine de l'Image en Clair
'--------------------------------------------------------
Sub drawpix()
Dim pt!, a!, b!, col As Byte

pt = 0
 
If interlace <> 0 Then
 'Interlacé - passe 1
  For b = 0 To pixheight - 1 Step 8
    For a = 0 To pixwidth - 1
      col = pix(pt)
      pt = pt + 1
      Call pset2(a, b, col)
    Next a
  Next b
  'passe 2
  For b = 4 To pixheight - 1 Step 8
    For a = 0 To pixwidth - 1
      col = pix(pt)
      pt = pt + 1
      Call pset2(a, b, col)
    Next a
  Next b
  'passe 3
  For b = 2 To pixheight - 1 Step 4
    For a = 0 To pixwidth - 1
      col = pix(pt)
      pt = pt + 1
      Call pset2(a, b, col)
    Next a
  Next b
  'passe 4
  For b = 1 To pixheight - 1 Step 2
    For a = 0 To pixwidth - 1
      col = pix(pt)
      pt = pt + 1
      Call pset2(a, b, col)
    Next a
  Next b
  
Else
'Non Interlacé
  For b = 0 To pixheight - 1
    For a = 0 To pixwidth - 1
      col = pix(pt)
      pt = pt + 1
      Call pset2(a, b, col)
    Next a
  Next b
 End If
 
 DoEvents
 Me.Refresh
End Sub


Sub pset2(x!, y!, coul As Byte)

 Select Case display_disposal
' Case 0
' Me.PSet (x + 20, y + 20), (RGB(color(coul, 0), color(coul, 1), color(coul, 2)))
 
 'test si transparence
 Case 0, 1
  'Si dans le nouveau clip ?
  If y < cliptop Then Exit Sub
  If y > cliptop + clipheight Then Exit Sub
  'test si draw
  If ((coul <> display_trans_col) Or (display_transpa = 0)) Then
    Me.PSet (x + 20, y + 20), (RGB(color(coul, 0), color(coul, 1), color(coul, 2)))
  End If
 
 Case 2 To 7
 'si disposal to bground / met fond gris forcé sur transparence
 'englobe aussi le disposal mode 3 (?)
  If ((coul = display_trans_col) And (display_transpa > 0)) Then
    Me.PSet (x + 20, y + 20), (RGB(239, 239, 239))
  Else
    Me.PSet (x + 20, y + 20), (RGB(color(coul, 0), color(coul, 1), color(coul, 2)))
  End If
 
 End Select
End Sub


Sub Setpix(z As Byte)
 'Store un Pixel / sauf si pixel en transparence
 If outptr > maxpixels Then Exit Sub
  
 pix(outptr) = z
 
 'Si le Clip + petit quePix -> test si retour ligne ?
 clipx = clipx + 1
 If clipx >= clipwidth Then
   clipx = 0
   outptr = outptr + pixwidth - clipwidth
 End If
 
 outptr = outptr + 1
End Sub

'Système de pile pour inverser le sens d'envoi des pixels
' (les pixels sont décodés et arrivent (PUSH) comme 3.2.1 et doivent
' être affichés (POP) comme 1.2.3 !)
Sub pushpix(z As Byte)
  pixstack(pixstackptr) = z
  pixstackptr = pixstackptr + 1
End Sub

Sub poppix()
 Do
   pixstackptr = pixstackptr - 1
   Setpix (pixstack(pixstackptr))
   If pixstackptr < 1 Then Exit Do
 Loop
End Sub


'Raz le dico / clearcode doit etre réglé
Sub Raz_dico()
Dim a%, b%

tablesize = initablesize
 
 For a = 0 To UBound(dico, 1)
   dico(a).code = -1
   dico(a).long = 0
   dico(a).prev = 0
 Next a
 
 For a = 0 To clearcode - 1
   dico(a).code = a
 Next a
 
 dicoptr = 1 + endcode
 
 newb9 = getCode()
 
' logfile.WriteLine ("> RAZ DICO ")
End Sub

'Renvoie la couleur index en RVB hexa
Function colstr$(z)
  colstr$ = hex2$(color(z, 0)) + "/" + hex2$(color(z, 1)) + "/" + hex2$(color(z, 2))
End Function

Function str4$(z)
  str4$ = Right$("000" + CStr(z), 4)
End Function

Function hex2$(z)
  hex2$ = Right$("0" + Hex$(z), 2)
End Function

Sub log(s$)
 logtxt = logtxt + s$ + Chr$(13) + Chr$(10)
End Sub

Function bin8$(z As Byte)
Dim q!, s$, puis!
puis = 128
s$ = ""
For q = 1 To 8
  If (puis And z) > 0 Then s$ = s$ + "1" Else s$ = s$ + "0"
  puis = puis / 2
Next q
bin8$ = s$
End Function

Function bin16$(z As Long)
Dim q!, s$, puis!

 puis = 32768
 s$ = ""
 For q = 1 To 16
   If (puis And z) > 0 Then s$ = s$ + "1" Else s$ = s$ + "0"
   puis = puis / 2
 Next q
 bin16$ = s$
End Function

Function getgif() As Byte
  getgif = gif(gifptr)
  gifptr = gifptr + 1
End Function

