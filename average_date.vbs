' VBScript to calculate the average date from a list of dates
' Considers day, month, year and handles leap years automatically via DateSerial

Option Explicit

Dim dateList, inputDate, dates
Set dateList = CreateObject("System.Collections.ArrayList")

' Default dates (suggested - enter them or your own)
Dim defaultDates
defaultDates = Array("27. oct. 1985", "14. jun. 1986", "31. jan. 2017", "4. may 2019", "22. sept. 2022")

WScript.Echo "Suggested default dates (copy/paste into input or enter new ones):" & vbCrLf & Join(defaultDates, vbCrLf)

Do
    inputDate = InputBox("Enter a date (e.g. 27. oct. 1985, YYYY-MM-DD or any valid):" & vbCrLf & "Press Cancel to finish and calculate average.", "Date Input", "")
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