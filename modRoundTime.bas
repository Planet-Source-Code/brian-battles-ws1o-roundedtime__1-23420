Attribute VB_Name = "modRoundTime"
Option Explicit

Public gstrRoundedTime As String  ' time rounded  to specified interval
Public gstrRealTime    As String  ' actual unrounded time
Public glngInterval    As Long    ' rounding interval, in minutes

Public Const gcintMin  As Integer = 1  ' constant that defines minimum interval in minutes
Public Const gcintMax  As Integer = 60 ' constant that defines maximum interval in minutes
Public Function dhRoundTime(dtmTime As Date, ByVal lngInterval As Long) As Date

    ' Round the time value in varTime to the
    ' nearest minute interval in lngInterval
    
    Dim intTime   As Integer
    Dim sglTime   As Single
    Dim intHour   As Integer
    Dim intMinute As Integer
    Dim lngdate   As Long
    
    On Error GoTo Err_dhRoundTime

    ' Get the date portion of the date/time value
    lngdate = DateValue(dtmTime)
    ' Get the time portion as a number like 11.5 for 11:30.
    sglTime = TimeValue(dtmTime) * 24
    ' Get the hour and store it away. Int truncates, CInt rounds, so use Int
    intHour = Int(sglTime)
    ' Get the number of minutes, and then round to the nearest
    ' occurrence of the interval specified.
    intMinute = CInt((sglTime - intHour) * 60)
    ' let's force it to give us the PREVIOUS 10-minute interval
    If intMinute >= 5 Then
        intMinute = intMinute - 5
    End If
    intMinute = CInt(intMinute / lngInterval) * lngInterval
    ' Build back up the original date/time value, rounded to the nearest interval
    dhRoundTime = CDate(lngdate + ((intHour + intMinute / 60) / 24))

Exit_dhRoundTime:

    On Error GoTo 0
    Exit Function

Err_dhRoundTime:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & Err.Source & vbCrLf & vbCrLf & "In modCallQueue" & " during " & "dhRoundTime", vbInformation, App.Title & " ADVISORY"
            Resume Exit_dhRoundTime
    End Select

End Function
Sub Main()
   
    '---------------------------------------------------------------
    ' Purpose   :
    ' Modified  : 5/24/2001 By BB
    '---------------------------------------------------------------

    On Error GoTo Err_Main

    ' to test this out, generate the calls on the frmTest's Timer event
    frmTest.Show

Exit_Main:
    
    On Error GoTo 0
    Exit Sub

Err_Main:
     
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In Module1" & " during " & "Main" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Main
    End Select

End Sub
