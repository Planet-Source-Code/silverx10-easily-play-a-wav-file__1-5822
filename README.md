<div align="center">

## Easily Play a WAV File


</div>

### Description

This source code will easily allow you to play a WAV file from within your Visual Basic application WITHOUT using an ActiveX Control. The two most amazing facts about this source code that amazed me were:

&#8226; The value for the "flag." I had absolutely no idea what to do for the flag. But I took a good guess at it, and I got it!

&#8226; It will actually play the WAV file from the Visual Basic environment (you don't need to compile it to hear the sound).
 
### More Info
 
1) Create one (1) CommandButton Control on the Form, and set the Name property to "cmdSound" (minus the quotes).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[silverx10](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/silverx10.md)
**Level**          |Beginner
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/silverx10-easily-play-a-wav-file__1-5822/archive/master.zip)





### Source Code

```
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Sub PlayWav(Filename As String)
  sndPlaySound (Filename), &H80
End Sub
Private Sub cmdSound_Click()
  PlayWav "C:\WINDOWS\Media\Chord.wav"
  'Chord.wav is a file that comes along with both
  'Windows 95 and 98 Operating Systems. If your
  'system is missing this file, specify a different WAV.
End Sub
'Now, press F5, or the Run button in the Visual Basic
'Environment, and then click the button. If you enjoy
'this source code, please let me know by posting feedback.
'Thanks!
```

