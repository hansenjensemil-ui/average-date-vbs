' VBScript to calculate the average date from a list of dates
' Considers day, month, year and handles leap years automatically via DateSerial

Option Explicit

Dim dates(4) ' Example dates, can be modified or input
dates(0) = "2023-01-15"
dates(1) = "2023-03-20"
dates(2) = "2024-02-29" ' Leap year
dates(3) = "2023-12-31"
dates(4) = "2024-06-15"

Function AverageDate(dateArray)
    Dim totalDays, count, i, avgDays, avgDate
    totalDays = 0
    count = 0
    
    For i = 0 To UBound(dateArray)
        If IsDate(dateArray(i)) Then
            totalDays = totalDays + CDbl(CDate(dateArray(i)))
            count = count + 1
        End If
    Next
    
    If count = 0 Then
        AverageDate = "No valid dates"
        Exit Function
    End If
    
    avgDays = totalDays / count
    avgDate = CDate(Round(avgDays)) ' Round to nearest day
    
    AverageDate = FormatDateTime(avgDate, vbShortDate)
End Function

WScript.Echo "Average Date: " & AverageDate(dates)