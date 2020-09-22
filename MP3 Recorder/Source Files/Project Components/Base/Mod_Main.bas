Attribute VB_Name = "Mod_Main"
Option Explicit
Rem -> **************************************************************************************************************************************************************
Rem -> Sub Main Function To Intialize Application
Public Sub Main()
    If InitCommonControlsVB Then
        Frm_Main.Show
    Else
        Beep
        MsgBox "Environment Error [Int.001]" & vbCrLf & vbCrLf & "Error Initializing Windows XP Skin Manifest!", vbCritical + vbSystemModal, "Environment Error"
        End
    End If
End Sub
Rem -> **************************************************************************************************************************************************************
Rem -> Function To Check If The Common Controls Library Loaded Or Not
Public Function InitCommonControlsVB() As Boolean
    On Error Resume Next
    Dim iccex As TagInitCommonControlsEx
    With iccex
        .LngSize = LenB(iccex)
        .LngICC = &H200
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)
    On Error GoTo 0
End Function
Rem -> **************************************************************************************************************************************************************
Public Function SecToTime(ByVal Ticks As Long) As String
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Converts Ticks (Milliseconds) To A Usable Time
    Dim IntHours        As Integer
    Dim IntMinutes      As Integer
    Dim IntSeconds      As Integer
    Rem -> ~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Rem -> Split Seconds Into HH:MM:SS
    IntSeconds = Ticks Mod 60
    Ticks = (Ticks - IntSeconds) / 60
    IntMinutes = Ticks Mod 60
    Ticks = (Ticks - IntMinutes) / 60
    IntHours = Ticks
    Rem -> ~~~~~~~~~~~~~~~
    Rem -> Format The Time
    SecToTime = Format(IntHours, "00") & ":" & _
                  Format(IntMinutes, "00") & ":" & _
                  Format(IntSeconds, "00")
End Function
Rem -> **************************************************************************************************************************************************************
Public Function GetFileSize(ByVal Length As Long) As String
    Dim B, KB, MB, GB As Long
    Select Case (Length)
        Case Is < 1024
            GetFileSize = Length & " Bytes"
            Exit Function
        Case Is > 1024
            B = Length
            While B >= 1024
                KB = KB + 1
                If KB >= 1024 Then
                    KB = 0
                    MB = MB + 1
                End If
                If MB >= 1024 Then
                    MB = 0
                    GB = GB + 1
                End If
                B = B - 1024
            Wend
            If KB > 0 Then GetFileSize = KB & "." & B & "  KBytes"
            If MB > 0 Then GetFileSize = MB & "." & KB & "  MBytes"
            If GB > 0 Then GetFileSize = GB & "." & MB & "  GBytes"
    End Select
End Function
