Attribute VB_Name = "SZAfunctions"
'****************************************************************************
' These SZA functions were written by Roland Leigh (RL40@le.ac.uk)
' Please use them freely, but leave in this comment as an acknowledgement
'****************************************************************************


' This module contains functions for the calculation of the elevation angle of the sun
' at a particular time, on a particular date, at a particular location.
' The main procedure returns what is known as the Solar Zenith Angle (SZA).
' This is the angle between a line directly up from the earth's surface, and a line
' from the same point on the surface to the sun. So SZA at dawn and dusk is
' approximately 90 degrees, and the SZA at noon is its lowest value during the day.

' The SZA function requires latitude and longitude in degrees, a time (in the standard
' time format of (HH:MM:SS) and a date (in the standard date format of DD/MM/YYYY)
' An optional timediff argument can also be used, which is the time difference between
' the time entered and GMT (because the algorithm works exclusively in GMT)
' I personally deal with local time differences outside of the main function, but
' I've left the option there for you if you want it.
' SZA is returned in degrees as a single.
' To get SZA in radians, just divide by 57.296

Public Function SZA(lat, lon, szatime, szadate, Optional timediff) As Single
If IsMissing(timediff) Then timediff = 0


Pi = 3.14128

mytime = szatime
myhour = Hour(mytime) + timediff
myminute = Minute(mytime)
mysecond = Second(mytime)
GMT = myhour + (myminute / 60) + (mysecond / 3600)


longitude = lon
latitude = lat
latitude = latitude / 57.296


mydate = szadate
bits = Split(mydate, "/")
darray = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 303, 334)
dn = darray(bits(1) - 1) + bits(0) - 1
'This is the actual SZA calculation

N = (2 * Pi * dn) / 365

EQT = 0.000075 + (0.001868 * Cos(N)) - (0.032077 * Sin(N)) - (0.014615 * Cos(2 * N)) - (0.040849 * Sin(2 * N))

th = Pi * ((GMT / 12) - 1 + (longitude / 180)) + EQT

delta = 0.006918 - (0.399912 * Cos(N)) + (0.070257 * Sin(N)) - (0.006758 * Cos(2 * N)) + (0.000907 * Sin(2 * N)) - (0.002697 * Cos(3 * N)) + (0.00148 * Sin(3 * N))
cossza = (Sin(delta) * Sin(latitude)) + (Cos(delta) * Cos(latitude) * Cos(th))
SZA = (Atn(-cossza / Sqr(-cossza * cossza + 1)) + 2 * Atn(1)) * 57.296 ' the 57.296 converts this SZA to degrees


End Function
'This function calculates solar Noon to the nearest minute, it's a really clumsy method
' that I use, but it's pretty quick anyway, if you need to minimise processing then
' this proceedure can be significantly improved. But it works fine as it is for most purposes.
Public Function SolarNoon(lat, lon, noondate) As String


best = 1000

For a = 0 To 1440
mytime = TimeSerial(0, a, 0)
asza = SZA(lat, lon, mytime, noondate)
If asza < best Then
best = asza
besta = mytime
End If
Next a
SolarNoon = besta
End Function


'This function calculates Dawn to the nearest minute, it's a really clumsy method
' that I use, but it's pretty quick anyway, if you need to minimise processing then
' this proceedure can be significantly improved. But it works fine as it is for most purposes.
Public Function Dawn(lat, lon, dawndate) As String


best = 1000
Last = 0
For a = 0 To 1440
mytime = TimeSerial(0, a, 0)
asza = SZA(lat, lon, mytime, dawndate)
If asza <= 90 And Last >= 90 Then
best = asza
besta = mytime
End If
Last = asza

Next a
Dawn = besta

End Function

'This function calculates Dusk to the nearest minute, it's a really clumsy method
' that I use, but it's pretty quick anyway, if you need to minimise processing then
' this proceedure can be significantly improved. But it works fine as it is for most purposes.
Public Function Dusk(lat, lon, duskdate) As String


best = 1000
Last = 0
Dim mytime As String

For a = 0 To 1440
mytime = ""
dusktime = TimeSerial(0, a, 0)

asza = SZA(lat, lon, dusktime, duskdate)



If asza >= 90 And Last <= 90 Then
best = asza
besta = dusktime
End If
Last = asza

Next a
Dusk = besta

End Function

' This function calculates the Julian Day, which is basically the day of the year.
' It assumes that it's not a leap year I'm afraid, perhaps an improvement to version 2?
Public Function GetJulianDay(JDdate) As Integer
mydate = JDdate
bits = Split(mydate, "/")
darray = Array(0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 303, 334)
GetJulianDay = darray(bits(1) - 1) + bits(0) - 1
End Function
'This is just the Quick Error handler code which I use as default in all my projects
Public Function EHandler() As Boolean
response = msgbox("Error Description = " & Err.Description & vbCrLf & "Error Source = " & Err.Source & vbCrLf & " Error Number = " & Err.Number & vbCrLf & "Operation Cancelled", , "Error Encountered")
Close
End Function
