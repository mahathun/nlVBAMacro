Sub ClaculateKPI()

    Dim person As Integer
    Dim currentMonth, week, weeklyTotal, workingDaysPerWeek, weklyViewTotalCellLetter As String
    
    Sheets("Weekly").Select
    currentMonth = Range("B1")
    week = Range("B2")
        
    For person = 6 To 13
        Sheets("Weekly").Select
        
        weeklyTotal = Range("W" + CStr(person))
        workingDaysPerWeek = Range("Y" + CStr(person))
        
        Sheets("WeekView").Select
        Select Case week
            Case Is = 1
                weeklyViewTotalCellLEtter = "B"
            Case Is = 2
                weeklyViewTotalCellLEtter = "C"
            Case Is = 3
                weeklyViewTotalCellLEtter = "D"
            Case Is = 4
                weeklyViewTotalCellLEtter = "E"
            Case Is = 5
                weeklyViewTotalCellLEtter = "F"
            Case Is = 6
                weeklyViewTotalCellLEtter = "G"
            Case Is = 7
                weeklyViewTotalCellLEtter = "H"
            Case Is = 8
                weeklyViewTotalCellLEtter = "I"
            Case Is = 9
                weeklyViewTotalCellLEtter = "J"
            Case Is = 10
                weeklyViewTotalCellLEtter = "K"
            Case Is = 11
                weeklyViewTotalCellLEtter = "L"
            Case Is = 12
                weeklyViewTotalCellLEtter = "M"
            Case Is = 13
                weeklyViewTotalCellLEtter = "N"
            Case Is = 14
                weeklyViewTotalCellLEtter = "O"
            Case Is = 15
                weeklyViewTotalCellLEtter = "P"
            Case Is = 16
                weeklyViewTotalCellLEtter = "Q"
            Case Is = 17
                weeklyViewTotalCellLEtter = "R"
            Case Is = 18
                weeklyViewTotalCellLEtter = "S"
            Case Is = 19
                weeklyViewTotalCellLEtter = "T"
            Case Is = 20
                weeklyViewTotalCellLEtter = "U"
            Case Is = 21
                weeklyViewTotalCellLEtter = "V"
            Case Is = 22
                weeklyViewTotalCellLEtter = "W"
            Case Is = 23
                weeklyViewTotalCellLEtter = "X"
            Case Is = 24
                weeklyViewTotalCellLEtter = "Y"
            Case Is = 25
                weeklyViewTotalCellLEtter = "Z"
            Case Is = 26
                weeklyViewTotalCellLEtter = "AA"
            Case Is = 27
                weeklyViewTotalCellLEtter = "AB"
            Case Is = 28
                weeklyViewTotalCellLEtter = "AC"
            Case Is = 29
                weeklyViewTotalCellLEtter = "AD"
            Case Is = 30
                weeklyViewTotalCellLEtter = "AE"
            Case Is = 31
                weeklyViewTotalCellLEtter = "AF"
            Case Is = 32
                weeklyViewTotalCellLEtter = "AG"
            Case Is = 33
                weeklyViewTotalCellLEtter = "AH"
            Case Is = 34
                weeklyViewTotalCellLEtter = "AI"
            Case Is = 35
                weeklyViewTotalCellLEtter = "AJ"
            Case Is = 36
                weeklyViewTotalCellLEtter = "AK"
            Case Is = 37
                weeklyViewTotalCellLEtter = "AL"
            Case Is = 38
                weeklyViewTotalCellLEtter = "AM"
            Case Is = 39
                weeklyViewTotalCellLEtter = "AN"
            Case Is = 40
                weeklyViewTotalCellLEtter = "AO"
            Case Is = 41
                weeklyViewTotalCellLEtter = "AP"
            Case Is = 42
                weeklyViewTotalCellLEtter = "AQ"
            Case Is = 43
                weeklyViewTotalCellLEtter = "AR"
            Case Is = 44
                weeklyViewTotalCellLEtter = "AS"
            Case Is = 45
                weeklyViewTotalCellLEtter = "AT"
            Case Is = 46
                weeklyViewTotalCellLEtter = "AU"
            Case Is = 47
                weeklyViewTotalCellLEtter = "AV"
            Case Is = 48
                weeklyViewTotalCellLEtter = "AW"
            Case Is = 49
                weeklyViewTotalCellLEtter = "AX"
            Case Is = 50
                weeklyViewTotalCellLEtter = "AY"
            Case Is = 51
                weeklyViewTotalCellLEtter = "AZ"
            Case Is = 52
                weeklyViewTotalCellLEtter = "BA"
            Case Is = 53
                weeklyViewTotalCellLEtter = "BB"
            Case Is = 54
                weeklyViewTotalCellLEtter = "BC"
            Case Is = 55
                weeklyViewTotalCellLEtter = "BD"


        End Select
        
        Range(weeklyViewTotalCellLEtter + CStr(person)).Value = weeklyTotal
        Range(weeklyViewTotalCellLEtter + CStr(person + 20)).Value = workingDaysPerWeek
        Range(weeklyViewTotalCellLEtter + "4").Value = currentMonth
    
    Next person
    
    
    
    Dim weeksPerYear, month As Integer
    Dim monthTotal(1 To 12) As Integer
    Dim countTotal(1 To 12) As Integer
    
    
    For person = 6 To 13
        For month = 1 To 12
            For weeksPerYear = 1 To 52
                    Select Case weeksPerYear
                        Case Is = 1
                            weeklyViewTotalCellLEtter = "B"
                        Case Is = 2
                            weeklyViewTotalCellLEtter = "C"
                        Case Is = 3
                            weeklyViewTotalCellLEtter = "D"
                        Case Is = 4
                            weeklyViewTotalCellLEtter = "E"
                        Case Is = 5
                            weeklyViewTotalCellLEtter = "F"
                        Case Is = 6
                            weeklyViewTotalCellLEtter = "G"
                        Case Is = 7
                            weeklyViewTotalCellLEtter = "H"
                        Case Is = 8
                            weeklyViewTotalCellLEtter = "I"
                        Case Is = 9
                            weeklyViewTotalCellLEtter = "J"
                        Case Is = 10
                            weeklyViewTotalCellLEtter = "K"
                        Case Is = 11
                            weeklyViewTotalCellLEtter = "L"
                        Case Is = 12
                            weeklyViewTotalCellLEtter = "M"
                        Case Is = 13
                            weeklyViewTotalCellLEtter = "N"
                        Case Is = 14
                            weeklyViewTotalCellLEtter = "O"
                        Case Is = 15
                            weeklyViewTotalCellLEtter = "P"
                        Case Is = 16
                            weeklyViewTotalCellLEtter = "Q"
                        Case Is = 17
                            weeklyViewTotalCellLEtter = "R"
                        Case Is = 18
                            weeklyViewTotalCellLEtter = "S"
                        Case Is = 19
                            weeklyViewTotalCellLEtter = "T"
                        Case Is = 20
                            weeklyViewTotalCellLEtter = "U"
                        Case Is = 21
                            weeklyViewTotalCellLEtter = "V"
                        Case Is = 22
                            weeklyViewTotalCellLEtter = "W"
                        Case Is = 23
                            weeklyViewTotalCellLEtter = "X"
                        Case Is = 24
                            weeklyViewTotalCellLEtter = "Y"
                        Case Is = 25
                            weeklyViewTotalCellLEtter = "Z"
                        Case Is = 26
                            weeklyViewTotalCellLEtter = "AA"
                        Case Is = 27
                            weeklyViewTotalCellLEtter = "AB"
                        Case Is = 28
                            weeklyViewTotalCellLEtter = "AC"
                        Case Is = 29
                            weeklyViewTotalCellLEtter = "AD"
                        Case Is = 30
                            weeklyViewTotalCellLEtter = "AE"
                        Case Is = 31
                            weeklyViewTotalCellLEtter = "AF"
                        Case Is = 32
                            weeklyViewTotalCellLEtter = "AG"
                        Case Is = 33
                            weeklyViewTotalCellLEtter = "AH"
                        Case Is = 34
                            weeklyViewTotalCellLEtter = "AI"
                        Case Is = 35
                            weeklyViewTotalCellLEtter = "AJ"
                        Case Is = 36
                            weeklyViewTotalCellLEtter = "AK"
                        Case Is = 37
                            weeklyViewTotalCellLEtter = "AL"
                        Case Is = 38
                            weeklyViewTotalCellLEtter = "AM"
                        Case Is = 39
                            weeklyViewTotalCellLEtter = "AN"
                        Case Is = 40
                            weeklyViewTotalCellLEtter = "AO"
                        Case Is = 41
                            weeklyViewTotalCellLEtter = "AP"
                        Case Is = 42
                            weeklyViewTotalCellLEtter = "AQ"
                        Case Is = 43
                            weeklyViewTotalCellLEtter = "AR"
                        Case Is = 44
                            weeklyViewTotalCellLEtter = "AS"
                        Case Is = 45
                            weeklyViewTotalCellLEtter = "AT"
                        Case Is = 46
                            weeklyViewTotalCellLEtter = "AU"
                        Case Is = 47
                            weeklyViewTotalCellLEtter = "AV"
                        Case Is = 48
                            weeklyViewTotalCellLEtter = "AW"
                        Case Is = 49
                            weeklyViewTotalCellLEtter = "AX"
                        Case Is = 50
                            weeklyViewTotalCellLEtter = "AY"
                        Case Is = 51
                            weeklyViewTotalCellLEtter = "AZ"
                        Case Is = 52
                            weeklyViewTotalCellLEtter = "BA"
                        Case Is = 53
                            weeklyViewTotalCellLEtter = "BB"
                        Case Is = 54
                            weeklyViewTotalCellLEtter = "BC"
                        Case Is = 55
                            weeklyViewTotalCellLEtter = "BD"
            
            
                    End Select
            
                currentMonth = Range(weeklyViewTotalCellLEtter + "4")
                
                
                    If (currentMonth = month) Then
                        countTotal(month) = countTotal(month) + Range(weeklyViewTotalCellLEtter + CStr(person + 20)).Value
                        monthTotal(month) = monthTotal(month) + Range(weeklyViewTotalCellLEtter + CStr(person)).Value
                    End If

                
                
               
            Next weeksPerYear
            Next month
            
            Sheets("MonthView").Select
            For month = 1 To 12
                Select Case month
                    Case Is = 1
                        weeklyViewTotalCellLEtter = "B"
                    Case Is = 2
                        weeklyViewTotalCellLEtter = "C"
                    Case Is = 3
                        weeklyViewTotalCellLEtter = "D"
                    Case Is = 4
                        weeklyViewTotalCellLEtter = "E"
                    Case Is = 5
                        weeklyViewTotalCellLEtter = "F"
                    Case Is = 6
                        weeklyViewTotalCellLEtter = "G"
                    Case Is = 7
                        weeklyViewTotalCellLEtter = "H"
                    Case Is = 8
                        weeklyViewTotalCellLEtter = "I"
                    Case Is = 9
                        weeklyViewTotalCellLEtter = "J"
                    Case Is = 10
                        weeklyViewTotalCellLEtter = "K"
                    Case Is = 11
                        weeklyViewTotalCellLEtter = "L"
                    Case Is = 12
                        weeklyViewTotalCellLEtter = "M"
                End Select
                
                If (countTotal(month) <> 0) Then
                    Range(weeklyViewTotalCellLEtter + CStr(person)).Value = monthTotal(month) / countTotal(month)
                End If
                

                
            Next month
            Sheets("WeekView").Select
            
             
            monthTotal(1) = 0
            monthTotal(2) = 0
            monthTotal(3) = 0
            monthTotal(4) = 0
            monthTotal(5) = 0
            monthTotal(6) = 0
            monthTotal(7) = 0
            monthTotal(8) = 0
            monthTotal(9) = 0
            monthTotal(10) = 0
            monthTotal(11) = 0
            monthTotal(12) = 0
            
            
            countTotal(1) = 0
            countTotal(2) = 0
            countTotal(3) = 0
            countTotal(4) = 0
            countTotal(5) = 0
            countTotal(6) = 0
            countTotal(7) = 0
            countTotal(8) = 0
            countTotal(9) = 0
            countTotal(10) = 0
            countTotal(11) = 0
            countTotal(12) = 0
    
    Next person
    
    Dim mbResult As Integer
    mbResult = MsgBox("These changes cannot be undone. Would you like to save a copy before proceeding?", vbYesNo + vbQuestion)
    Select Case mbResult
    Case vbYes
        'Modify as needed, this is a simple example with no error handling:
        With ActiveWorkbook
            If Not .Saved Then .SaveAs Application.GetSaveAsFilename()
        End With
    Case vbNo
        ' Do nothing and allow the macro to run
        Sheets("Weekly").Select
        Range("B6:C13").Value = "" 'Monday
        Range("E6:F13").Value = "" 'Tuesday
        Range("H6:I13").Value = "" 'Wednsday
        Range("K6:L13").Value = "" 'Thursday
        Range("N6:O13").Value = "" 'Friday
        Range("Q6:R13").Value = "" 'Saturday
        Range("T6:U13").Value = "" 'Sunday
        
        Range("B1:B2").Value = "" 'Clearing the month and week
        

    Case vbCancel
        ' Do NOT allow the macro to run
        Exit Sub

End Select

    
    
End Sub
