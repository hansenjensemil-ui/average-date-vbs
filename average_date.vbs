' VBScript to calculate the average date from a list of dates
' Considers day, month, year and handles leap years automatically via DateSerial

Option Explicit

Dim dateList, inputDate, dates
Set dateList = CreateObject("System.Collections.ArrayList")

' Default dates pre-filled one-by-one (press Enter to accept each as-is)
Dim defaultDates, i
defaultDates = Array("27. oct. 1985", "14. jun. 1986", "31. jan. 2017", "4. may 2019", "22. sept. 2022")

For i = 0 To UBound(defaultDates)
    inputDate = InputBox("Date #" & (dateList.Count + 1) & " (press Enter to accept default, edit if needed):" & vbCrLf & "Cancel to stop adding.", "Date Input", defaultDates(i))
    If inputDate = "" Then Exit For ' Cancel
    If IsDate(inputDate) Then
        dateList.Add inputDate
    Else
        MsgBox "Invalid date '" & inputDate & "'. Skipping.", vbExclamation
    End If
Next

' Optional: add more dates after defaults
Do While True
    inputDate = InputBox("Enter additional date (or Cancel to calculate average):", "Additional Dates", "")
    If inputDate = "" Then Exit Do
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