Attribute VB_Name = "basTimeFunctions"
Public Function NumberOfMinutes(ByVal Time1 As String, ByVal Time2 As String) As Integer
'Returns the number of minutes of difference between Time1 and Time2'
'Time1 must be later than Time2'
'Time1 and Time2 must be strings in the form of "14:26:43" OR "14.26.43"'

Dim HoursDiff As Integer
Dim MinutesDiff As Integer
HoursDiff = Hour(Time1) - Hour(Time2)
MinutesDiff = Minute(Time1) - Minute(Time2)
NumberOfMinutes = (HoursDiff * 60) + MinutesDiff


End Function

Public Function ShortenTime(ByVal FullTime As String) As String
'This removes the seconds from a time value and returns the hours and minutes'
'The string FullTime is in the format "14:26:43"'


ShortenTime = Left(FullTime, 5)

End Function
