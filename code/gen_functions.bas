Attribute VB_Name = "gen_functions"
Option Explicit

Global sSettingsForm As String

Sub Pause(Seconds)
    Dim PauseTime, Start, finish
    PauseTime = Seconds   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + Seconds
        DoEvents    ' Yield to other processes.
    Loop
    finish = Timer  ' Set end time.
End Sub
