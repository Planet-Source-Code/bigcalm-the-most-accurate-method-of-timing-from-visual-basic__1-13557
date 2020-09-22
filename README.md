<div align="center">

## The most accurate method of Timing from Visual Basic


</div>

### Description

By using the "Performance Timer" in all modern PC's it is possible to achieve timing accuracy of greater than one microsecond (yes, 1 millionth of a second). This code shows you how to use API calls to access and use it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BigCalm](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bigcalm.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bigcalm-the-most-accurate-method-of-timing-from-visual-basic__1-13557/archive/master.zip)

### API Declarations

```
' Unsigned 64-bit long
Public Type LongLong
  LowPart As Long
  HighPart As Long
End Type
Declare Function QueryPerformanceCounter Lib "kernel32" _
        (lpPerformanceCount As LongLong) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" _
        (lpFrequency As LongLong) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long
```


### Source Code

```
' Unsigned 64-bit long
Public Type LongLong
  LowPart As Long
  HighPart As Long
End Type
Declare Function QueryPerformanceCounter Lib "kernel32" _
        (lpPerformanceCount As LongLong) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" _
        (lpFrequency As LongLong) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Function TimerElapsed(Optional µS As Long = 0) As Boolean
Static StartTime As Variant ' Decimal
Static PerformanceFrequency As LongLong
Static EndTime As Variant ' Decimal
Dim CurrentTime As LongLong
Dim Dec As Variant
  If µS > 0 Then
    ' Initialize
    If QueryPerformanceFrequency(PerformanceFrequency) Then
      ' Performance Timer available
      Debug.Print PerformanceFrequency.HighPart & " " & PerformanceFrequency.LowPart
      If QueryPerformanceCounter(CurrentTime) Then
      Else
        ' Performance timer is available, but is not responding
        CurrentTime.HighPart = 0
        CurrentTime.LowPart = timeGetTime
        PerformanceFrequency.HighPart = 0
        PerformanceFrequency.LowPart = 1000
      End If
    Else
      ' Performance timer is not available.
      CurrentTime.HighPart = 0
      CurrentTime.LowPart = timeGetTime
      PerformanceFrequency.HighPart = 0
      PerformanceFrequency.LowPart = 1000
    End If
    ' Work out start time...
    ' Convert to DECIMAL
    Dec = CDec(CurrentTime.LowPart)
    ' make this UNSIGNED
    If Dec < 0 Then
      Dec = CDec(Dec + (2147483648# * 2))
    End If
    ' Add higher value
    StartTime = CDec(Dec + (CurrentTime.HighPart * 2147483648# * 2))
    ' Put performance frequency into Dec variable
    Dec = CDec(PerformanceFrequency.LowPart)
    ' Convert to unsigned
    If Dec < 0 Then
      Dec = CDec(Dec + (2147483648# * 2))
    End If
    ' Add higher value
    Dec = CDec(Dec + (PerformanceFrequency.HighPart * 2147483648# * 2))
    ' Work out end time from this
    EndTime = CDec(StartTime + µS * Dec / 1000000)
    TimerElapsed = False
  Else
    If PerformanceFrequency.LowPart = 1000 And PerformanceFrequency.HighPart = 0 Then
      ' Using standard windows timer
      Dec = CDec(timeGetTime)
      If Dec < 0 Then
        Dec = CDec(Dec + (2147483648# * 2))
      End If
      If Dec > EndTime Then
        TimerElapsed = True
      Else
        TimerElapsed = False
      End If
    Else
      If QueryPerformanceCounter(CurrentTime) Then
        Dec = CDec(CurrentTime.LowPart)
        ' make this UNSIGNED
        If Dec < 0 Then
          Dec = CDec(Dec + (2147483648# * 2))
        End If
        Dec = CDec(Dec + (CurrentTime.HighPart * 2147483648# * 2))
        If Dec > EndTime Then
          TimerElapsed = True
        Else
          TimerElapsed = False
        End If
      Else
        ' Should never happen in theory
        Err.Raise vbObjectError + 2, "Timer Elapsed", "Your performance timer has stopped functioning!!!"
        TimerElapsed = True
      End If
    End If
  End If
End Function
' Example use
Public Sub DummySub()
Dim i As Long
  ' count for 5 seconds and then display result
  TimerElapsed (5000000)
  i = 0
  Do While TimerElapsed = False
    i = i + 1
    DoEvents
  Loop
  MsgBox i
End Sub
```

