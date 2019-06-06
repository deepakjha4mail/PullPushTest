Option Explicit
Dim st, i

'st = Timer()
st= Inputbox("Please Enter the Time Value in Seconds: ")
Msgbox ("The Exact Value of time in HR:MIN:SEC Format is: "& FormatTime(st))

'For i = 1 to 2
'  Wait 10
'  Msgbox FormatTime(Timer() - st)
'Next
'Msgbox

Function FormatTime(secs)
  Dim t, a
  secs = Int(secs)
  a = Array(CStr(Right("00" & Int(secs / 3600) Mod 24, 2)), CStr(Right("00" & Int(secs / 60) Mod 60, 2)), CStr(Right("00" & secs Mod 60, 2)))
  FormatTime = Join(a, ":")
End Function