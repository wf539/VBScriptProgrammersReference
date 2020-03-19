Dim intAge As Integer
    intAge = DateDiff(DateInterval.Year, dtBirthDate, Today())
    If Today() < DateSerial(Year(Today()), Month(dtBirthDate), Day(dtBirthDate)) Then
        intAge = intAge - 1
    End If
    Return intAge
