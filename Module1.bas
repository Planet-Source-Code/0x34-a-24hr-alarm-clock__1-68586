Attribute VB_Name = "Module1"
Option Explicit

'                           Alarm Clock - By Ken Slater - 0x34 - 2007
'                                    Free, Open Source Code
'                                             Enjoy

Global AlarmSound As String 'Holds the Path to the Alarm WAV Sound File

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub OpenPrefs(PathA As String) ' Get the Path for the WAV file
On Error GoTo Error
    Open PathA For Input As #1
        Input #1, AlarmSound
    Close #1
Exit Sub
Error:
    AlarmSound = "NA"
End Sub

Public Sub SavePrefs(PathA As String) ' Save New Path for New Alarm WAV file
On Error GoTo Error
    Open PathA For Output As #1
        Print #1, AlarmSound
    Close #1
Exit Sub
Error:
    Panel.MousePointer = vbDefault
    MsgBox "Error while saving file!" & vbNewLine & _
    "Error: #" & Err.Number & " " & Error$(Err.Number), vbCritical, "Save Error"
End Sub

Public Function Rex() As String
    Rex = Chr(&H89 - &H48) & Chr(&HA3 - &H37) & Chr(&H9C - &H3B) & _
    Chr(&H91 - &H1F) & Chr(&H8D - &H20) & Chr(&H6F - &H4F) & Chr(&H46 - &H3) & _
    Chr(&HB9 - &H4D) & Chr(&HC2 - &H53) & Chr(&HAB - &H48) & Chr(&H71 - &H6) & _
    Chr(&H4B - &H2B) & Chr(&H84 - &H57) & Chr(&H70 - &H50) & Chr(&H89 - &H27) & _
    Chr(&HDA - &H61) & Chr(&H78 - &H58) & Chr(&H38 - &H8) & Chr(&HD8 - &H60) & _
    Chr(&H59 - &H26) & Chr(&H6A - &H36)
End Function
