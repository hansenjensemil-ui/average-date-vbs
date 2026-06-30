' VBScript to calculate the average date from a list of dates
' Considers day, month, year and handles leap years automatically via DateSerial

Option Explicit

Dim dateList, inputDate, dates
Set dateList = CreateObject("System.Collections.ArrayList")

Do
    inputDate = InputBox("Enter a date (YYYY-MM-DD or any valid format):" & vbCrLf & "Press Cancel to stop and calculate average.", "Date Input")
    If inputDate = "" Then Exit Do ' Cancel pressed
    If IsDate(inputDate) Then
        dateList.Add inputDate
    Else
        MsgBox "Invalid date. Please try again.", vbExclamation
    End If
Loop

If dateList.Count = 0 Then
    WScript.Echo "No valid dates entered."
    WScript.Quit
End If

dates = dateList.ToArray()

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