'Version: 1.00
'CCI-Comcast
'For Echobroadband
'By Farhat Abbas

Sub Main()
    'Call WorkloadCCI
    'Call SurroundingCCI
    'Call WorkloadCommscope
    'Call SurroundingCommscope
End Sub
Sub WorkloadCCI()
    'Gethering Data
        filedate = Format(Date, "ddmmyyyy")
        OutputFileName = "CCI-Comcast Workload " & filedate & ".xlsx"
        Call Check_if_workbook_is_open(OutputFileName)
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\" & OutputFileName
        Application.DisplayAlerts = True
        DayName = Format(Date, "dddd")
        If DayName = "Monday" Then
            x = 3
        End If
        If DayName = "Tuesday" Or DayName = "Wednesday" Or DayName = "Thursday" Or DayName = "Friday" Then
            x = 1
        End If
        EDate = Date - x
        PreviousDate = Format(EDate, "mm/dd/yyyy")
        MsgBox "Starting Date: " & PreviousDate
        Filename = ThisWorkbook.name
        
        
    'Input and Output
        'Row 6  'Copy the formate and call the function below another example of 2 status is at row 10, Copy and paste it at right before "'Calling surrounding and saving the workbook"
            Sheetname = "Node Split" 'Data require form tracking sheet, Sheet name
            JobType = "NS Design"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL" 'Dilvery date
            DateColumNum = 64 'Dilvery date Column number
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK" 'Final Status
            StatusFColumNum = 63 'Final Status number
            Out1 = "E6" 'Output where it should be prinited on output sheet
            Out2 = "F6"
            Out3 = "G6"
            Out4 = "H6"
            Out5 = "I6"
            Out6 = "L6"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR" 'Fiber Drsigner status
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 7
            Sheetname = "Asbuilt Coax & Fiber"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AV:AV"
            DateColumNum = 48
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AA:AA"
            Status2ColumNum = 27
            Status3Colum = "AL:AL"
            Status3ColumNum = 38
            StatusFColum = "AU:AU"
            StatusFColumNum = 47
            Out1 = "E7"
            Out2 = "F7"
            Out3 = "G7"
            Out4 = "H7"
            Out5 = "I7"
            Out6 = "L7"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 8
            Sheetname = "SFU&MDU Design"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BC:BC"
            DateColumNum = 55
            Status1Colum = "AB:AB"
            Status1ColumNum = 28
            Status2Colum = "AH:AH"
            Status2ColumNum = 34
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BB:BB"
            StatusFColumNum = 54
            Out1 = "E8"
            Out2 = "F8"
            Out3 = "G8"
            Out4 = "H8"
            Out5 = "I8"
            Out6 = "L8"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AQ:AQ"
            DesginerStatus3ColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 9
            Sheetname = "Node Split"
            JobType = "NS Asbuilt"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E9"
            Out2 = "F9"
            Out3 = "G9"
            Out4 = "H9"
            Out5 = "I9"
            Out6 = "L9"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 10
            Sheetname = "Commercial_ Expense Design"
            JobType = "Hyperbuild Design"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E10"
            Out2 = "F10"
            Out3 = "G10"
            Out4 = "H10"
            Out5 = "I10"
            Out6 = "L10"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 11
            Sheetname = "Asbuilt Coax & Fiber"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Beltway"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AV:AV"
            DateColumNum = 48
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AA:AA"
            Status2ColumNum = 27
            Status3Colum = "AL:AL"
            Status3ColumNum = 38
            StatusFColum = "AU:AU"
            StatusFColumNum = 47
            Out1 = "E11"
            Out2 = "F11"
            Out3 = "G11"
            Out4 = "H11"
            Out5 = "I11"
            Out6 = "L11"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 12
            Sheetname = "Commercial_ Expense Design"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Beltway"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E12"
            Out2 = "F12"
            Out3 = "G12"
            Out4 = "H12"
            Out5 = "I12"
            Out6 = "L12"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 13
            Sheetname = "SFU&MDU Design"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Beltway"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BC:BC"
            DateColumNum = 55
            Status1Colum = "AB:AB"
            Status1ColumNum = 28
            Status2Colum = "AH:AH"
            Status2ColumNum = 34
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BB:BB"
            StatusFColumNum = 54
            Out1 = "E13"
            Out2 = "F13"
            Out3 = "G13"
            Out4 = "H13"
            Out5 = "I13"
            Out6 = "L13"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AQ:AQ"
            DesginerStatus3ColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 14
            Sheetname = "Node Split"
            JobType = "NS Design"
            Region = "Twin City"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E14"
            Out2 = "F14"
            Out3 = "G14"
            Out4 = "H14"
            Out5 = "I14"
            Out6 = "L14"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 15
            Sheetname = "ME Design,Asbuit&Desktop Srvy"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Seattle"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT"
            DateColumNum = 46
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            StatusFColum = "AS:AS"
            StatusFColumNum = 45
            Out1 = "E15"
            Out2 = "F15"
            Out3 = "G15"
            Out4 = "H15"
            Out5 = "I15"
            Out6 = "L15"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 16
            Sheetname = "Node Split"
            JobType = "NS Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E16"
            Out2 = "F16"
            Out3 = "G16"
            Out4 = "H16"
            Out5 = "I16"
            Out6 = "L16"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 17
            Sheetname = "Node Split"
            JobType = "NS Asbuilt"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E17"
            Out2 = "F17"
            Out3 = "G17"
            Out4 = "H17"
            Out5 = "I17"
            Out6 = "L17"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 18
            Sheetname = "Asbuilt Coax & Fiber"
            JobType = "Span Replacement Asbuilt"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AV:AV"
            DateColumNum = 48
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AA:AA"
            Status2ColumNum = 27
            Status3Colum = "AL:AL"
            Status3ColumNum = 38
            StatusFColum = "AU:AU"
            StatusFColumNum = 47
            Out1 = "E18"
            Out2 = "F18"
            Out3 = "G18"
            Out4 = "H18"
            Out5 = "I18"
            Out6 = "L18"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 19
            Sheetname = "Commercial_ Expense Design"
            JobType = "Forced Relocation Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E19"
            Out2 = "F19"
            Out3 = "G19"
            Out4 = "H19"
            Out5 = "I19"
            Out6 = "L19"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 20  Different Job Type
            Sheetname = "Asbuilt Coax & Fiber"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AV:AV"
            DateColumNum = 48
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AA:AA"
            Status2ColumNum = 27
            Status3Colum = "AL:AL"
            Status3ColumNum = 38
            StatusFColum = "AU:AU"
            StatusFColumNum = 47
            Out1 = "E20"
            Out2 = "F20"
            Out3 = "G20"
            Out4 = "H20"
            Out5 = "I20"
            Out6 = "L20"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 21
            Sheetname = "ME Design,Asbuit&Desktop Srvy"
            JobType = "Metro-E Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT"
            DateColumNum = 46
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            StatusFColum = "AS:AS"
            StatusFColumNum = 45
            Out1 = "E21"
            Out2 = "F21"
            Out3 = "G21"
            Out4 = "H21"
            Out5 = "I21"
            Out6 = "L21"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 22
            Sheetname = "ME Design,Asbuit&Desktop Srvy"
            JobType = "Metro-E Asbuilt"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT"
            DateColumNum = 46
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            StatusFColum = "AS:AS"
            StatusFColumNum = 45
            Out1 = "E22"
            Out2 = "F22"
            Out3 = "G22"
            Out4 = "H22"
            Out5 = "I22"
            Out6 = "L22"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 23
            Sheetname = "Commercial_ Expense Design"
            JobType = "Forced Relocation Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E23"
            Out2 = "F23"
            Out3 = "G23"
            Out4 = "H23"
            Out5 = "I23"
            Out6 = "L23"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 24
            Sheetname = "Commercial_ Expense Design"
            JobType = "Hyperbuild Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E24"
            Out2 = "F24"
            Out3 = "G24"
            Out4 = "H24"
            Out5 = "I24"
            Out6 = "L24"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 25
            Sheetname = "SFU&MDU Design"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BC:BC"
            DateColumNum = 55
            Status1Colum = "AB:AB"
            Status1ColumNum = 28
            Status2Colum = "AH:AH"
            Status2ColumNum = 34
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BB:BB"
            StatusFColumNum = 54
            Out1 = "E25"
            Out2 = "F25"
            Out3 = "G25"
            Out4 = "H25"
            Out5 = "I25"
            Out6 = "L25"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AQ:AQ"
            DesginerStatus3ColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 26 NS
            Sheetname = "Node Split"
            JobType = "Legacy Doc Updates"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E26"
            Out2 = "F26"
            Out3 = "G26"
            Out4 = "H26"
            Out5 = "I26"
            Out6 = "L26"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 27 MDADS
            Sheetname = "ME Design,Asbuit&Desktop Srvy"
            JobType = "EDP& Wavelength"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT"
            DateColumNum = 46
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            StatusFColum = "AS:AS"
            StatusFColumNum = 45
            Out1 = "E27"
            Out2 = "F27"
            Out3 = "G27"
            Out4 = "H27"
            Out5 = "I27"
            Out6 = "L27"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 28 CED
            Sheetname = "Commercial_ Expense Design"
            JobType = "Coax Design"
            Region = "Atlanta"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E28"
            Out2 = "F28"
            Out3 = "G28"
            Out4 = "H28"
            Out5 = "I28"
            Out6 = "L28"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 29 ACF
            Sheetname = "Asbuilt Coax & Fiber"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Atlanta"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AV:AV"
            DateColumNum = 48
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AA:AA"
            Status2ColumNum = 27
            Status3Colum = "AL:AL"
            Status3ColumNum = 38
            StatusFColum = "AU:AU"
            StatusFColumNum = 47
            Out1 = "E29"
            Out2 = "F29"
            Out3 = "G29"
            Out4 = "H29"
            Out5 = "I29"
            Out6 = "L29"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 30 MDADS
            Sheetname = "ME Design,Asbuit&Desktop Srvy"
            JobType = "Metro-E Design"
            Region = "Atlanta"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT"
            DateColumNum = 46
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            StatusFColum = "AS:AS"
            StatusFColumNum = 45
            Out1 = "E30"
            Out2 = "F30"
            Out3 = "G30"
            Out4 = "H30"
            Out5 = "I30"
            Out6 = "L30"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 31 MDADS
            Sheetname = "ME Design,Asbuit&Desktop Srvy"
            JobType = "Metro-E Asbuilt"
            Region = "Atlanta"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT"
            DateColumNum = 46
            Status1Colum = "X:X"
            Status1ColumNum = 24
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            StatusFColum = "AS:AS"
            StatusFColumNum = 45
            Out1 = "E31"
            Out2 = "F31"
            Out3 = "G31"
            Out4 = "H31"
            Out5 = "I31"
            Out6 = "L31"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 32 SMD
            Sheetname = "SFU&MDU Design"
            JobType = Array("Coax/Fiber Asbuilt", "Coax Design", "Forced Relocation Asbuilt", "NS Asbuilt", "NS Design", "Metro-E Design", "Metro-E Asbuilt", "SDU Design", "MDU Design", "Span Replacement Asbuilt", "SDU Asbuilt", "MDU Asbuilt", "Hyperbuild Design", "Expense Design", "Forced Relocation Design", "Hyperbuild Asbuilt", "Legacy Doc Updates", "Coax & Fiber Design", "EDP& Wavelength")
            Region = "Atlanta"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BC:BC"
            DateColumNum = 55
            Status1Colum = "AB:AB"
            Status1ColumNum = 28
            Status2Colum = "AH:AH"
            Status2ColumNum = 34
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BB:BB"
            StatusFColumNum = 54
            Out1 = "E32"
            Out2 = "F32"
            Out3 = "G32"
            Out4 = "H32"
            Out5 = "I32"
            Out6 = "L32"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AQ:AQ"
            DesginerStatus3ColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 33 CED
            Sheetname = "Commercial_ Expense Design"
            JobType = "Hyperbuild Design"
            Region = "Atlanta"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AY:AY"
            DateColumNum = 51
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            StatusFColum = "AX:AX"
            StatusFColumNum = 50
            Out1 = "E33"
            Out2 = "F33"
            Out3 = "G33"
            Out4 = "H33"
            Out5 = "I33"
            Out6 = "L33"
            numofstatus = 2
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AI:AI"
            DesginerStatus3ColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 34 NS
            Sheetname = "Node Split"
            JobType = "NS Asbuilt"
            Region = "Florida"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E34"
            Out2 = "F34"
            Out3 = "G34"
            Out4 = "H34"
            Out5 = "I34"
            Out6 = "L34"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        'Row 35 NS
            Sheetname = "Node Split"
            JobType = "NS Design"
            Region = "Florida"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BL:BL"
            DateColumNum = 64
            Status1Colum = "AC:AC"
            Status1ColumNum = 29
            Status2Colum = "AI:AI"
            Status2ColumNum = 35
            Status3Colum = "AS:AS"
            Status3ColumNum = 45
            StatusFColum = "BK:BK"
            StatusFColumNum = 63
            Out1 = "E35"
            Out2 = "F35"
            Out3 = "G35"
            Out4 = "H35"
            Out5 = "I35"
            Out6 = "L35"
            numofstatus = 3
            DateColumA = "A:A"
            DateColumANum = 1
            DesginerStatus3Colum = "AR:AR"
            DesginerStatus3ColumNum = 44
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatus3ColumNum, DesginerStatus3Colum)
        
    'Clean up
        Call SurroundingCCI
        Workbooks(OutputFileName).Save
End Sub
Sub SurroundingCCI()
        filedate = Format(Date, "ddmmyyyy")
        OutputFileName = "CCI-Comcast Workload " & filedate & ".xlsx"
        TodayDate = Format(Date, "dd/m/yyyy")
    'Surrounding 'Copy from here to other part of macro for Surrounding Workload macro
        Dim wb As Workbook
        Dim ws As Worksheet
        Set wb = Workbooks(OutputFileName)
        Set ws = wb.Sheets("Sheet1")
        With ws
        ThickBoarder = "B6:M10, B11:M13, B14:M14, B15:M15, B16:M27, B28:M33, B34:M35" 'Thick black Boarder
        .Range("B4").FormulaR1C1 = "CCI"
        .Range("C4").FormulaR1C1 = TodayDate
        .Range("B5").FormulaR1C1 = "Region"
        .Range("C5").FormulaR1C1 = "Basic Rate Per Job"
        .Range("D5").FormulaR1C1 = "Scopes"
        .Range("E5").FormulaR1C1 = "New Received"
        .Range("F5").FormulaR1C1 = "Spill Over"
        .Range("G5").FormulaR1C1 = "IP/RFI Reply"
        .Range("H5").FormulaR1C1 = "QC Pending"
        .Range("I5").FormulaR1C1 = "RFI & Pess Lock Pending"
        .Range("J5").FormulaR1C1 = "Remaining Jobs"
        .Range("K5").FormulaR1C1 = "Target to deliver"
        .Range("L5").FormulaR1C1 = "Actual Delivery Previous Day"
        .Range("M5").FormulaR1C1 = "Actual Billable Hours" ' Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Range("B6").FormulaR1C1 = "California"
        .Range("B11").FormulaR1C1 = "Beltway"
        .Range("B14").FormulaR1C1 = "Twin City"
        .Range("B15").FormulaR1C1 = "Seattle"
        .Range("B16").FormulaR1C1 = "Chicago"
        .Range("B28").FormulaR1C1 = "Atlanta"
        .Range("B34").FormulaR1C1 = "Florida"  'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Range("C6").FormulaR1C1 = "10.90"
        .Range("C7").FormulaR1C1 = "2.00"
        .Range("C8").FormulaR1C1 = "3.86"
        .Range("C9").FormulaR1C1 = "6.10"
        .Range("C10").FormulaR1C1 = "6.86"
        .Range("C11").FormulaR1C1 = "2.00"
        .Range("C12").FormulaR1C1 = "3.86"
        .Range("C13").FormulaR1C1 = "6.00"
        .Range("C14").FormulaR1C1 = "3.70"
        .Range("C15").FormulaR1C1 = "3.65"
        .Range("C16").FormulaR1C1 = "5.96"
        .Range("C17").FormulaR1C1 = "6.90"
        .Range("C18").FormulaR1C1 = "0.62"
        .Range("C19").FormulaR1C1 = "3.31"
        .Range("C20").FormulaR1C1 = "2.00"
        .Range("C21").FormulaR1C1 = "5.43"
        .Range("C22").FormulaR1C1 = "2.79"
        .Range("C23:C25").FormulaR1C1 = "7.65"
        .Range("C26:C27").FormulaR1C1 = "2.00"
        .Range("C28").FormulaR1C1 = "3.86"
        .Range("C29").FormulaR1C1 = "2.00"
        .Range("C30").FormulaR1C1 = "4.64"
        .Range("C31").FormulaR1C1 = "2.00"
        .Range("C32,C33").FormulaR1C1 = "6.86"
        .Range("C34").FormulaR1C1 = "6.10"
        .Range("C35").FormulaR1C1 = "10.90" 'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Range("D6").FormulaR1C1 = "Node Split Design"
        .Range("D7").FormulaR1C1 = "Asbuilt"
        .Range("D8").FormulaR1C1 = "SDU & MDU Design"
        .Range("D9").FormulaR1C1 = "Node Split Asbuilt"
        .Range("D10").FormulaR1C1 = "Hyperbuild Design"
        .Range("D11").FormulaR1C1 = "Coax Asbuilt"
        .Range("D12").FormulaR1C1 = "Coax Design"
        .Range("D13").FormulaR1C1 = "SDU & MDU Design"
        .Range("D14").FormulaR1C1 = "Node Split Design"
        .Range("D15").FormulaR1C1 = "Metro-E Asbuilt"
        .Range("D16").FormulaR1C1 = "Node Split Design"
        .Range("D17").FormulaR1C1 = "Node Split Asbuilt"
        .Range("D18").FormulaR1C1 = "Span Replacement Asbuilt"
        .Range("D19").FormulaR1C1 = "Force Relocation Design"
        .Range("D20").FormulaR1C1 = "Asbuilt"
        .Range("D21").FormulaR1C1 = "Metro-E Design"
        .Range("D22").FormulaR1C1 = "Metro-E Asbuilt"
        .Range("D23").FormulaR1C1 = "Forced Relocation"
        .Range("D24").FormulaR1C1 = "Hyperbuilt Design"
        .Range("D25").FormulaR1C1 = "SDU & MDU Design"
        .Range("D26").FormulaR1C1 = "Legacy Doc Update"
        .Range("D27").FormulaR1C1 = "EDP & Wavelength"
        .Range("D28").FormulaR1C1 = "Coax Design"
        .Range("D29").FormulaR1C1 = "Asbuilt"
        .Range("D30").FormulaR1C1 = "Metro-E Design"
        .Range("D31").FormulaR1C1 = "Metro-E Asbuilt"
        .Range("D32").FormulaR1C1 = "SDU/MDU Design"
        .Range("D33").FormulaR1C1 = "Hyperbuilt Design"
        .Range("D34").FormulaR1C1 = "Node Split Asbuilt"
        .Range("D35").FormulaR1C1 = "Tier 2 Node Split Design" 'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
    'Count
        FindlastemptyCell = WorksheetFunction.CountA(.Range("D:D")) + 4
        TotalJob = FindlastemptyCell + 1
        Manday = FindlastemptyCell + 2
        .Range("D" & TotalJob).FormulaR1C1 = "Total Jobs"
        .Range("D" & Manday).FormulaR1C1 = "Manday"
        .Range("E" & Manday & ":L" & Manday).NumberFormat = "0"
        For x = 6 To FindlastemptyCell
        .Range("J" & x).FormulaR1C1 = WorksheetFunction.Sum(.Range("E" & x & ":H" & x))
        Next

        
        .Range("E" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("E6:E" & FindlastemptyCell))
        .Range("F" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("F6:F" & FindlastemptyCell))
        .Range("G" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("G6:G" & FindlastemptyCell))
        .Range("H" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("H6:H" & FindlastemptyCell))
        .Range("I" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("I6:I" & FindlastemptyCell))
        .Range("J" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("J6:J" & FindlastemptyCell))
        .Range("K" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("K6:K" & FindlastemptyCell))
        .Range("L" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("L6:L" & FindlastemptyCell))
        .Range("M" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("M6:M" & FindlastemptyCell))

        For x = 6 To FindlastemptyCell
        TotalSumNR = .Range("E" & x).Value * .Range("C" & x).Value
        NRSum = TotalSumNR + NRSum
        TotalSumSO = .Range("F" & x).Value * .Range("C" & x).Value
        SOSum = TotalSumSO + SOSum
        TotalSumIP = .Range("G" & x).Value * .Range("C" & x).Value
        IPSum = TotalSumIP + IPSum
        TotalSumQC = .Range("H" & x).Value * .Range("C" & x).Value
        QCSum = TotalSumQC + QCSum
        TotalSumRFI = .Range("I" & x).Value * .Range("C" & x).Value
        RFISum = TotalSumRFI + RFISum
        TotalSumRJ = .Range("J" & x).Value * .Range("C" & x).Value
        RJSum = TotalSumRJ + RJSum
        TotalSumTD = .Range("K" & x).Value * .Range("C" & x).Value
        TDSum = TotalSumTD + TDSum
        TotalSumADPD = .Range("L" & x).Value * .Range("C" & x).Value
        ADPDSum = TotalSumADPD + ADPDSum
        Next

        .Range("E" & Manday).FormulaR1C1 = NRSum / 8
        .Range("F" & Manday).FormulaR1C1 = SOSum / 8
        .Range("G" & Manday).FormulaR1C1 = IPSum / 8
        .Range("H" & Manday).FormulaR1C1 = QCSum / 8
        .Range("I" & Manday).FormulaR1C1 = RFISum / 8
        .Range("J" & Manday).FormulaR1C1 = RJSum / 8
        .Range("K" & Manday).FormulaR1C1 = TDSum / 8
        .Range("L" & Manday).FormulaR1C1 = ADPDSum / 8

        
    'Merging & Size etc
        .Range("B6:B10").Merge  'Add here this is to merge the region new region with multi job scope merge here
        .Range("B11:B13").Merge
        .Range("B16:B27").Merge
        .Range("B28:B33").Merge
        .Range("B34:B35").Merge
        .Range("B6:B10").Merge 'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Columns("B").ColumnWidth = 11.53 'Column width
        .Columns("C").ColumnWidth = 12.29
        .Columns("D").ColumnWidth = 19.43
        .Columns("E").ColumnWidth = 19.86
        .Columns("F").ColumnWidth = 24.86
        .Columns("G").ColumnWidth = 8.43
        .Columns("H").ColumnWidth = 8.43
        .Columns("I").ColumnWidth = 11.29
        .Columns("J").ColumnWidth = 12.14
        .Columns("K").ColumnWidth = 8.43
        .Columns("L").ColumnWidth = 8.43
        .Columns("M").ColumnWidth = 8.43
        .Rows("4:4").RowHeight = 45.75
        .Rows("5:5").RowHeight = 60.75
        .Range("B4:M" & Manday).Font.Size = 12 'Size of font
        .Range("B6:B" & FindlastemptyCell).Font.Size = 14
        .Range(" C4, D5, I5, J5, B4:B" & FindlastemptyCell & ", E5:H" & Manday & ", D" & TotalJob & ", D" & Manday & ", J" & Manday & ", I" & Manday & ", J5:M" & Manday).Font.FontStyle = "Bold" 'bolding the data

    'Color Interiar and Font
        With .Range("B4:M" & Manday)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = True
            .ReadingOrder = xlContext
        End With
        With .Range("E4:E" & Manday & ", J5:J" & Manday & " , D" & TotalJob & ",F" & Manday & ", D" & Manday & ", J" & Manday & ":G" & Manday).Font 'Changing the font color any addition
            .TintAndShade = 0
        End With
        With .Range("J" & Manday & ":L" & Manday).Font 'Changing the font color Red Color
            .Color = -16776961
            .TintAndShade = 0
        End With
        With .Range("L5:M5").Interior 'Changing the backgroup color
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 6299648
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Range("L5:M5").Font 'Changing the font color
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With .Range("K6:K" & FindlastemptyCell & ",B4,C4,D" & TotalJob & ",D" & Manday).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        With .Range("B6:B" & FindlastemptyCell).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
        With .Range("B6:B" & FindlastemptyCell).Font
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0
        End With
        With .Range("I5:I" & Manday).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark2
            .TintAndShade = -9.99786370433668E-02
            .PatternTintAndShade = 0
        End With
    'Table
        'Normal boarder
            NormalBoarder = "B5:M" & FindlastemptyCell & ",B4:C4,D" & TotalJob & ":L" & Manday & ",M" & TotalJob 'Adding square boarder for each celll add inside the "NormalBoarder" if new data is at M36 then B5:M36 Can be done or just add seperatly with comma but inside the bracket
            With .Range(NormalBoarder).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        'Think boarder for Index
            ThickIndexBoarder = "B5,J5" 'Adding square Thick boarder for each celll add inside the "ThickIndexBoarder" if New title B7 is at M36 then B5,J5,B7 Can be done or just add seperatly with comma but inside the bracket
            With .Range(ThickIndexBoarder).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickIndexBoarder).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickIndexBoarder).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickIndexBoarder).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        'Think boarder Outside
            With .Range(ThickBoarder).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(ThickBoarder).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
End Sub
Sub WorkloadCommscope()
    'Gethering Data
        filedate = Format(Date, "ddmmyyyy")
        OutputFileName = "Workload " & filedate & ".xlsx"
        Call Check_if_workbook_is_open(OutputFileName)
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\" & OutputFileName
        Application.DisplayAlerts = True
        DayName = Format(Date, "dddd")
        If DayName = "Monday" Then
            x = 3
        End If
        If DayName = "Tuesday" Or DayName = "Wednesday" Or DayName = "Thursday" Or DayName = "Friday" Then
            x = 1
        End If
        EDate = Date - x
        PreviousDate = Format(EDate, "mm/dd/yyyy")
        MsgBox "Starting Date: " & PreviousDate
        Filename = ThisWorkbook.name

        
    'Input and Output
        'Row 4
            Sheetname = "Comm+Res+Other Design" 'Data require form tracking sheet, Sheet name
            JobType = "Comm Design"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AW:AW" 'Dilvery date
            DateColumNum = 49 'Dilvery date Column number
            Status1Colum = "AF:AF"
            Status1ColumNum = 32
            Status2Colum = "AN:AN"
            Status2ColumNum = 40
            StatusFColum = "AV:AV" 'Final Status
            StatusFColumNum = 48 'Final Status number
            Out1 = "E4" 'Output where it should be prinited on output sheet
            Out2 = "F4"
            Out3 = "G4"
            Out4 = "H4"
            Out5 = "I4"
            Out6 = "L4"
            numofstatus = 2 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AM:AM" 'Fiber Drsigner status
            DesginerStatusColumNum = 39
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        'Row 5
            Sheetname = "Comm+Res+Other Design" 'Data require form tracking sheet, Sheet name
            JobType = "<>*Comm Design*"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AW:AW" 'Dilvery date
            DateColumNum = 49 'Dilvery date Column number
            Status1Colum = "AF:AF"
            Status1ColumNum = 32
            Status2Colum = "AN:AN"
            Status2ColumNum = 40
            StatusFColum = "AV:AV" 'Final Status
            StatusFColumNum = 48 'Final Status number
            Out1 = "E5" 'Output where it should be prinited on output sheet
            Out2 = "F5"
            Out3 = "G5"
            Out4 = "H5"
            Out5 = "I5"
            Out6 = "L5"
            numofstatus = 2 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AM:AM" 'Fiber Drsigner status
            DesginerStatusColumNum = 39
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        'Row 6
            Sheetname = "Comm+Res+Other Asbuilt" 'Data require form tracking sheet, Sheet name
            JobType = "Comm Asbuilt"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT" 'Dilvery date
            DateColumNum = 46 'Dilvery date Column number
            Status1Colum = "Y:Y"
            Status1ColumNum = 25
            Status2Colum = "AC:AC"
            Status2ColumNum = 29
            Status3Colum = "AJ:AJ"
            Status3ColumNum = 36
            StatusFColum = "AS:AS" 'Final Status
            StatusFColumNum = 45 'Final Status number
            Out1 = "E6" 'Output where it should be prinited on output sheet
            Out2 = "F6"
            Out3 = "G6"
            Out4 = "H6"
            Out5 = "I6"
            Out6 = "L6"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AI:AI" 'Fiber Drsigner status
            DesginerStatusColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    
        'Row 7
            Sheetname = "Comm+Res+Other Asbuilt" 'Data require form tracking sheet, Sheet name
            JobType = "<>*Comm Asbuilt*"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT" 'Dilvery date
            DateColumNum = 46 'Dilvery date Column number
            Status1Colum = "Y:Y"
            Status1ColumNum = 25
            Status2Colum = "AC:AC"
            Status2ColumNum = 29
            Status3Colum = "AJ:AJ"
            Status3ColumNum = 36
            StatusFColum = "AS:AS" 'Final Status
            StatusFColumNum = 45 'Final Status number
            Out1 = "E7" 'Output where it should be prinited on output sheet
            Out2 = "F7"
            Out3 = "G7"
            Out4 = "H7"
            Out5 = "I7"
            Out6 = "L7"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AI:AI" 'Fiber Drsigner status
            DesginerStatusColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    
        'Row 8
            Sheetname = "Node Split" 'Data require form tracking sheet, Sheet name
            JobType = "NS Design"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BG:BG" 'Dilvery date
            DateColumNum = 59 'Dilvery date Column number
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BF:BF" 'Final Status
            StatusFColumNum = 58 'Final Status number
            Out1 = "E8" 'Output where it should be prinited on output sheet
            Out2 = "F8"
            Out3 = "G8"
            Out4 = "H8"
            Out5 = "I8"
            Out6 = "L8"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AQ:AQ" 'Fiber Drsigner status
            DesginerStatusColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    
        'Row 9
            Sheetname = "Node Split" 'Data require form tracking sheet, Sheet name
            JobType = "NS Asbuilt"
            Region = "California"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BG:BG" 'Dilvery date
            DateColumNum = 59 'Dilvery date Column number
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BF:BF" 'Final Status
            StatusFColumNum = 58 'Final Status number
            Out1 = "E9" 'Output where it should be prinited on output sheet
            Out2 = "F9"
            Out3 = "G9"
            Out4 = "H9"
            Out5 = "I9"
            Out6 = "L9"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AQ:AQ" 'Fiber Drsigner status
            DesginerStatusColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        
        'Row 10
            Sheetname = "Metro E" 'Data require form tracking sheet, Sheet name
            JobType = "Metro E Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AP:AP" 'Dilvery date
            DateColumNum = 42 'Dilvery date Column number
            Status1Colum = "X:X"
            Status1ColumNum = 24
            StatusFColum = "AO:AO" 'Final Status
            StatusFColumNum = 41 'Final Status number
            Out1 = "E10" 'Output where it should be prinited on output sheet
            Out2 = "F10"
            Out3 = "G10"
            Out4 = "H10"
            Out5 = "I10"
            Out6 = "L10"
            numofstatus = 1 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        'Row 11
            Sheetname = "Metro E" 'Data require form tracking sheet, Sheet name
            JobType = "Metro E Asbuilt"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AP:AP" 'Dilvery date
            DateColumNum = 42 'Dilvery date Column number
            Status1Colum = "X:X"
            Status1ColumNum = 24
            StatusFColum = "AO:AO" 'Final Status
            StatusFColumNum = 41 'Final Status number
            Out1 = "E11" 'Output where it should be prinited on output sheet
            Out2 = "F11"
            Out3 = "G11"
            Out4 = "H11"
            Out5 = "I11"
            Out6 = "L11"
            numofstatus = 1 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)

        'Row 12
            Sheetname = "Metro E" 'Data require form tracking sheet, Sheet name
            JobType = "Fiber Maintenance"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AP:AP" 'Dilvery date
            DateColumNum = 42 'Dilvery date Column number
            Status1Colum = "X:X"
            Status1ColumNum = 24
            StatusFColum = "AO:AO" 'Final Status
            StatusFColumNum = 41 'Final Status number
            Out1 = "E12" 'Output where it should be prinited on output sheet
            Out2 = "F12"
            Out3 = "G12"
            Out4 = "H12"
            Out5 = "I12"
            Out6 = "L12"
            numofstatus = 1 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        
        'Row 13
            Sheetname = "Comm+Res+Other Design" 'Data require form tracking sheet, Sheet name
            JobType = "<>*=*"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AW:AW" 'Dilvery date
            DateColumNum = 49 'Dilvery date Column number
            Status1Colum = "AF:AF"
            Status1ColumNum = 32
            Status2Colum = "AN:AN"
            Status2ColumNum = 40
            StatusFColum = "AV:AV" 'Final Status
            StatusFColumNum = 48 'Final Status number
            Out1 = "E13" 'Output where it should be prinited on output sheet
            Out2 = "F13"
            Out3 = "G13"
            Out4 = "H13"
            Out5 = "I13"
            Out6 = "L13"
            numofstatus = 2 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AM:AM" 'Fiber Drsigner status
            DesginerStatusColumNum = 39
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
 
        'Row 14
            Sheetname = "Comm+Res+Other Asbuilt" 'Data require form tracking sheet, Sheet name
            JobType = "<>*=*"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT" 'Dilvery date
            DateColumNum = 46 'Dilvery date Column number
            Status1Colum = "Y:Y"
            Status1ColumNum = 25
            Status2Colum = "AC:AC"
            Status2ColumNum = 29
            Status3Colum = "AJ:AJ"
            Status3ColumNum = 36
            StatusFColum = "AS:AS" 'Final Status
            StatusFColumNum = 45 'Final Status number
            Out1 = "E14" 'Output where it should be prinited on output sheet
            Out2 = "F14"
            Out3 = "G14"
            Out4 = "H14"
            Out5 = "I14"
            Out6 = "L14"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AI:AI" 'Fiber Drsigner status
            DesginerStatusColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)

        'Row 15
            Sheetname = "Node Split" 'Data require form tracking sheet, Sheet name
            JobType = "NS Design"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BG:BG" 'Dilvery date
            DateColumNum = 59 'Dilvery date Column number
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BF:BF" 'Final Status
            StatusFColumNum = 58 'Final Status number
            Out1 = "E15" 'Output where it should be prinited on output sheet
            Out2 = "F15"
            Out3 = "G15"
            Out4 = "H15"
            Out5 = "I15"
            Out6 = "L15"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AQ:AQ" 'Fiber Drsigner status
            DesginerStatusColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
     
        'Row 16
            Sheetname = "Node Split" 'Data require form tracking sheet, Sheet name
            JobType = "NS Asbuilt"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BG:BG" 'Dilvery date
            DateColumNum = 59 'Dilvery date Column number
            Status1Colum = "AD:AD"
            Status1ColumNum = 30
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            Status3Colum = "AR:AR"
            Status3ColumNum = 44
            StatusFColum = "BF:BF" 'Final Status
            StatusFColumNum = 58 'Final Status number
            Out1 = "E16" 'Output where it should be prinited on output sheet
            Out2 = "F16"
            Out3 = "G16"
            Out4 = "H16"
            Out5 = "I16"
            Out6 = "L16"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AQ:AQ" 'Fiber Drsigner status
            DesginerStatusColumNum = 43
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    
        'Row 17
            Sheetname = "Comm+Res+Other Asbuilt" 'Data require form tracking sheet, Sheet name
            JobType = "Spatial Asbuilt"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT" 'Dilvery date
            DateColumNum = 46 'Dilvery date Column number
            Status1Colum = "Y:Y"
            Status1ColumNum = 25
            Status2Colum = "AC:AC"
            Status2ColumNum = 29
            Status3Colum = "AJ:AJ"
            Status3ColumNum = 36
            StatusFColum = "AS:AS" 'Final Status
            StatusFColumNum = 45 'Final Status number
            Out1 = "E17" 'Output where it should be prinited on output sheet
            Out2 = "F17"
            Out3 = "G17"
            Out4 = "H17"
            Out5 = "I17"
            Out6 = "L17"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AI:AI" 'Fiber Drsigner status
            DesginerStatusColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    
        'Row 18
            Sheetname = "FD Design" 'Data require form tracking sheet, Sheet name
            JobType = "Res design (FD)"
            Region = "Chicago"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "BM:BM" 'Dilvery date
            DateColumNum = 65 'Dilvery date Column number
            Status1Colum = "AB:AB"
            Status1ColumNum = 28
            Status2Colum = "AJ:AJ"
            Status2ColumNum = 36
            Status3Colum = "AT:AT"
            Status3ColumNum = 46
            StatusFColum = "BL:BL" 'Final Status
            StatusFColumNum = 64 'Final Status number
            Out1 = "E18" 'Output where it should be prinited on output sheet
            Out2 = "F18"
            Out3 = "G18"
            Out4 = "H18"
            Out5 = "I18"
            Out6 = "L18"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AS:AS" 'Fiber Drsigner status
            DesginerStatusColumNum = 45
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)

        'Row 19
            Sheetname = "Comm+Res+Other Design" 'Data require form tracking sheet, Sheet name
            JobType = "<>*=*"
            Region = "Houston"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AW:AW" 'Dilvery date
            DateColumNum = 49 'Dilvery date Column number
            Status1Colum = "AF:AF"
            Status1ColumNum = 32
            Status2Colum = "AN:AN"
            Status2ColumNum = 40
            StatusFColum = "AV:AV" 'Final Status
            StatusFColumNum = 48 'Final Status number
            Out1 = "E19" 'Output where it should be prinited on output sheet
            Out2 = "F19"
            Out3 = "G19"
            Out4 = "H19"
            Out5 = "I19"
            Out6 = "L19"
            numofstatus = 2 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AM:AM" 'Fiber Drsigner status
            DesginerStatusColumNum = 39
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        'Row 20
            Sheetname = "Comm+Res+Other Asbuilt" 'Data require form tracking sheet, Sheet name
            JobType = "<>*=*"
            Region = "Houston"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT" 'Dilvery date
            DateColumNum = 46 'Dilvery date Column number
            Status1Colum = "Y:Y"
            Status1ColumNum = 25
            Status2Colum = "AC:AC"
            Status2ColumNum = 29
            Status3Colum = "AJ:AJ"
            Status3ColumNum = 36
            StatusFColum = "AS:AS" 'Final Status
            StatusFColumNum = 45 'Final Status number
            Out1 = "E20" 'Output where it should be prinited on output sheet
            Out2 = "F20"
            Out3 = "G20"
            Out4 = "H20"
            Out5 = "I20"
            Out6 = "L20"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AI:AI" 'Fiber Drsigner status
            DesginerStatusColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    
        'Row 21
            Sheetname = "Metro E" 'Data require form tracking sheet, Sheet name
            JobType = "Metro E Asbuilt"
            Region = "Houston"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AP:AP" 'Dilvery date
            DateColumNum = 42 'Dilvery date Column number
            Status1Colum = "X:X"
            Status1ColumNum = 24
            StatusFColum = "AO:AO" 'Final Status
            StatusFColumNum = 41 'Final Status number
            Out1 = "E21" 'Output where it should be prinited on output sheet
            Out2 = "F21"
            Out3 = "G21"
            Out4 = "H21"
            Out5 = "I21"
            Out6 = "L21"
            numofstatus = 1 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        'Row 22
            Sheetname = "Comm+Res+Other Design" 'Data require form tracking sheet, Sheet name
            JobType = "<>*=*"
            Region = "Seattle"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AW:AW" 'Dilvery date
            DateColumNum = 49 'Dilvery date Column number
            Status1Colum = "AF:AF"
            Status1ColumNum = 32
            Status2Colum = "AN:AN"
            Status2ColumNum = 40
            StatusFColum = "AV:AV" 'Final Status
            StatusFColumNum = 48 'Final Status number
            Out1 = "E22" 'Output where it should be prinited on output sheet
            Out2 = "F22"
            Out3 = "G22"
            Out4 = "H22"
            Out5 = "I22"
            Out6 = "L22"
            numofstatus = 2 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AM:AM" 'Fiber Drsigner status
            DesginerStatusColumNum = 39
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        'Row 23
            Sheetname = "Comm+Res+Other Asbuilt" 'Data require form tracking sheet, Sheet name
            JobType = "<>*=*"
            Region = "Seattle"
            RegionColum = "D:D"
            RegionColumNum = 4
            ScopeColum = "E:E"
            ScopeColumNum = 5
            DateColum = "AT:AT" 'Dilvery date
            DateColumNum = 46 'Dilvery date Column number
            Status1Colum = "Y:Y"
            Status1ColumNum = 25
            Status2Colum = "AC:AC"
            Status2ColumNum = 29
            Status3Colum = "AJ:AJ"
            Status3ColumNum = 36
            StatusFColum = "AS:AS" 'Final Status
            StatusFColumNum = 45 'Final Status number
            Out1 = "E23" 'Output where it should be prinited on output sheet
            Out2 = "F23"
            Out3 = "G23"
            Out4 = "H23"
            Out5 = "I23"
            Out6 = "L23"
            numofstatus = 3 ' Number of status, Status 1, 2 and 3 are counted, Status 4 is main status shouldnt be counted here
            DateColumA = "A:A" 'Date New Received
            DateColumANum = 1
            DesginerStatusColum = "AI:AI" 'Fiber Drsigner status
            DesginerStatusColumNum = 35
            Call RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
        
    'Clean Up
        Call SurroundingCommscope
        Workbooks(OutputFileName).Save
        
End Sub
Sub SurroundingCommscope()
        filedate = Format(Date, "ddmmyyyy")
        OutputFileName = "Workload " & filedate & ".xlsx"
        TodayDate = Format(Date, "dd/m/yyyy")
    'Surrounding 'Copy from here to other part of macro for Surrounding Workload macro
        Dim wb As Workbook
        Dim ws As Worksheet
        Set wb = Workbooks(OutputFileName)
        Set ws = wb.Sheets("Sheet1")
        With ws
        .Range("B2").FormulaR1C1 = "Commscope"
        .Range("C2").FormulaR1C1 = TodayDate
        .Range("B3").FormulaR1C1 = "Region"
        .Range("C3").FormulaR1C1 = "Basic Rate Per Job"
        .Range("D3").FormulaR1C1 = "Scopes"
        .Range("E3").FormulaR1C1 = "New Received"
        .Range("F3").FormulaR1C1 = "Spill Over"
        .Range("G3").FormulaR1C1 = "IP/RFI Reply"
        .Range("H3").FormulaR1C1 = "QC Pending"
        .Range("I3").FormulaR1C1 = "RFI & Pess Lock Pending"
        .Range("J3").FormulaR1C1 = "Remaining Jobs"
        .Range("K3").FormulaR1C1 = "Target to deliver"
        .Range("L3").FormulaR1C1 = "Planned to Deliver"
        .Range("M3").FormulaR1C1 = "Actual Delivery Previous Day"
        .Range("N3").FormulaR1C1 = "Actual Billable Hours" ' Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Range("B4").FormulaR1C1 = "California"
        .Range("B10").FormulaR1C1 = "Chicago"
        .Range("B18").FormulaR1C1 = "Houston"
        .Range("B22").FormulaR1C1 = "Seattle"  'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Range("C4").FormulaR1C1 = "3.86"
        .Range("C5").FormulaR1C1 = "5.00"
        .Range("C6").FormulaR1C1 = "2.86"
        .Range("C7").FormulaR1C1 = "4.00"
        .Range("C8").FormulaR1C1 = "17.00"
        .Range("C9").FormulaR1C1 = "5.56"
        .Range("C10").FormulaR1C1 = "6.20"
        .Range("C11").FormulaR1C1 = "5.46"
        .Range("C12").FormulaR1C1 = ""
        .Range("C13").FormulaR1C1 = "3.86"
        .Range("C14").FormulaR1C1 = "2.00"
        .Range("C15").FormulaR1C1 = "10.90"
        .Range("C16").FormulaR1C1 = "6.10"
        .Range("C17").FormulaR1C1 = "2.50"
        .Range("C18").FormulaR1C1 = "10.75"
        .Range("C19").FormulaR1C1 = "6.86"
        .Range("C20").FormulaR1C1 = "2.00"
        .Range("C21").FormulaR1C1 = "2.00"
        .Range("C22").FormulaR1C1 = "3.86"
        .Range("C23").FormulaR1C1 = "2.00" 'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Range("D4").FormulaR1C1 = "Commercial RF Design"
        .Range("D5").FormulaR1C1 = "Residential/ Forced Relocation/Hyperbuild Design"
        .Range("D6").FormulaR1C1 = "Commercial RF Asbuilt"
        .Range("D7").FormulaR1C1 = "Residential/ Forced Relocation/Hyperbuild Asbuilt"
        .Range("D8").FormulaR1C1 = "Node Split Design"
        .Range("D9").FormulaR1C1 = "Node Split Asbuilt"
        .Range("D10").FormulaR1C1 = "Metro E Design"
        .Range("D11").FormulaR1C1 = "Metro E Asbuilt"
        .Range("D12").FormulaR1C1 = "Fiber Maintenance"
        .Range("D13").FormulaR1C1 = "Comm/Res Design"
        .Range("D14").FormulaR1C1 = "Comm/Res As-built"
        .Range("D15").FormulaR1C1 = "Nodes Split Design"
        .Range("D16").FormulaR1C1 = "Nodes Split AsBuilt"
        .Range("D17").FormulaR1C1 = "Spatial Asbuild New node"
        .Range("D18").FormulaR1C1 = "Fiber Deep Design"
        .Range("D19").FormulaR1C1 = "RF Design "
        .Range("D20").FormulaR1C1 = "RF Asbuilt"
        .Range("D21").FormulaR1C1 = "Metro E As-built"
        .Range("D22").FormulaR1C1 = "RF Design "
        .Range("D23").FormulaR1C1 = "RF As-built" 'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
    'Count
        FindlastemptyCell = WorksheetFunction.CountA(.Range("C:C")) + 2
        TotalJob = FindlastemptyCell + 1
        Manday = FindlastemptyCell + 2
        .Range("D" & TotalJob).FormulaR1C1 = "Total Jobs"
        .Range("D" & Manday).FormulaR1C1 = "Manday"
        .Range("E" & Manday & ":L" & Manday).NumberFormat = "0"
        For x = 4 To FindlastemptyCell
        .Range("J" & x).FormulaR1C1 = WorksheetFunction.Sum(.Range("E" & x & ":H" & x))
        Next

        
        .Range("E" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("E6:E" & FindlastemptyCell))
        .Range("F" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("F6:F" & FindlastemptyCell))
        .Range("G" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("G6:G" & FindlastemptyCell))
        .Range("H" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("H6:H" & FindlastemptyCell))
        .Range("I" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("I6:I" & FindlastemptyCell))
        .Range("J" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("J6:J" & FindlastemptyCell))
        .Range("K" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("K6:K" & FindlastemptyCell))
        .Range("L" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("L6:L" & FindlastemptyCell))
        .Range("M" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("M6:M" & FindlastemptyCell))
        .Range("N" & TotalJob).FormulaR1C1 = WorksheetFunction.Sum(.Range("N6:N" & FindlastemptyCell)) 'Add here New column
        '----------------------------------------------------------
        
        
        '----------------------------------------------------------
        For x = 4 To FindlastemptyCell
        TotalSumNR = .Range("E" & x).Value * .Range("C" & x).Value
        NRSum = TotalSumNR + NRSum
        TotalSumSO = .Range("F" & x).Value * .Range("C" & x).Value
        SOSum = TotalSumSO + SOSum
        TotalSumIP = .Range("G" & x).Value * .Range("C" & x).Value
        IPSum = TotalSumIP + IPSum
        TotalSumQC = .Range("H" & x).Value * .Range("C" & x).Value
        QCSum = TotalSumQC + QCSum
        TotalSumRFI = .Range("I" & x).Value * .Range("C" & x).Value
        RFISum = TotalSumRFI + RFISum
        TotalSumRJ = .Range("J" & x).Value * .Range("C" & x).Value
        RJSum = TotalSumRJ + RJSum
        TotalSumTD = .Range("K" & x).Value * .Range("C" & x).Value
        TDSum = TotalSumTD + TDSum
        TotalSumADPD = .Range("L" & x).Value * .Range("C" & x).Value
        ADPDSum = TotalSumADPD + ADPDSum
        TotalSumADPD = .Range("M" & x).Value * .Range("C" & x).Value
        ADPDSum = TotalSumADPD + ADPDSum 'Add here New column
        '----------------------------------------------------------
        
        
        '----------------------------------------------------------
        Next

        .Range("E" & Manday).FormulaR1C1 = NRSum / 8
        .Range("F" & Manday).FormulaR1C1 = SOSum / 8
        .Range("G" & Manday).FormulaR1C1 = IPSum / 8
        .Range("H" & Manday).FormulaR1C1 = QCSum / 8
        .Range("I" & Manday).FormulaR1C1 = RFISum / 8
        .Range("J" & Manday).FormulaR1C1 = RJSum / 8
        .Range("K" & Manday).FormulaR1C1 = TDSum / 8
        .Range("L" & Manday).FormulaR1C1 = ADPDSum / 8
        .Range("M" & Manday).FormulaR1C1 = NSum / 8 'Add here New column
        '----------------------------------------------------------
        

        
    'Merging & Size etc
        .Range("B4:B9").Merge  'Add here this is to merge the region new region with multi job scope merge here
        .Range("B10:B17").Merge
        .Range("B18:B21").Merge
        .Range("B22:B23").Merge 'Add here
        '----------------------------------------------------------


        '----------------------------------------------------------
        .Columns("B").ColumnWidth = 9.75 'Column width
        .Columns("C").ColumnWidth = 8.38
        .Columns("D").ColumnWidth = 21.88
        .Columns("E").ColumnWidth = 8.38
        .Columns("F").ColumnWidth = 8.38
        .Columns("G").ColumnWidth = 8.38
        .Columns("H").ColumnWidth = 8.38
        .Columns("I").ColumnWidth = 8.38
        .Columns("J").ColumnWidth = 8.38
        .Columns("K").ColumnWidth = 8.38
        .Columns("L").ColumnWidth = 8.38
        .Columns("M").ColumnWidth = 8.38
        .Columns("N").ColumnWidth = 8.38
        '.Rows("5:5").RowHeight = 10
        .Range("B2:N" & Manday).Font.Size = 11 'Size of font
        .Range("C2:C" & FindlastemptyCell & ",D3:D" & FindlastemptyCell).Font.Size = 12
        .Range("B2:B" & FindlastemptyCell & ",D3:N3, E4:H" & Manday & ",J4:N" & Manday & ",C2").Font.FontStyle = "Bold"  'bolding the data

    'Color Interiar and Font
        With .Range("B3:N" & Manday)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = True
            .ReadingOrder = xlContext
        End With
        With .Range("B4:B23").Interior 'backgrounds
            .PatternColorIndex = xlAutomatic
            .Color = 682978
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Range("B2:C2,D24:D25").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 15523812
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Range("J3:J25,C12:H12").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 11851260
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Range("L3:N3").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 411543
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

        With .Range("I3:I25").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12566463
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Range("E3:E23,K3:K23").Font
            .Color = -16094238
            .TintAndShade = 0
        End With
        With Range("J3").Font
            .Color = -16365673
            .TintAndShade = 0
        End With
        With .Range("C8,M25").Font
            .Color = -16776961
            .TintAndShade = 0
        End With
        With .Range("B2:C2,D25:L25").Font
            .Color = -4165632
            .TintAndShade = 0
        End With
        With .Range("B2:C3").Font
            .Color = -16365673
            .TintAndShade = 0
        End With
        .Range("L3:N3").Font.Color = vbWhite

    'Table
        'Normal boarder  ADD HERE
            NormalBoarder = "B3:N" & FindlastemptyCell & ",B2:C2,D" & TotalJob & ":M" & Manday & ",N" & TotalJob 'Adding square boarder for each celll add inside the "NormalBoarder" if new data is at M36 then B5:M36 Can be done or just add seperatly with comma but inside the bracket
            
            With .Range(NormalBoarder).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(NormalBoarder).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        'Think boarder for Index
            ThickIndexBoarder = "B2,C2,J3,B3" 'ADD HERE Adding square Thick boarder for each celll add inside the "ThickIndexBoarder" if New title B7 is at M36 then B5,J5,B7 Can be done or just add seperatly with comma but inside the bracket
            With .Range(ThickIndexBoarder).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickIndexBoarder).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickIndexBoarder).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickIndexBoarder).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        'Think boarder Outside
            ThickBoarder = "B4:N9, B10:N17, B18:N21, B22:N23, J4:J9,J10:J17,J18:J21,J22:J23 " 'Thick black Boarder ADD HERE

            With .Range(ThickBoarder).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With .Range(ThickBoarder).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Range(ThickBoarder).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
End Sub
Sub Check_if_workbook_is_open(OutputFileName)
    Dim wb As Workbook 'to test if workbook is open. No change here
        For Each wb In Workbooks
            If wb.name = OutputFileName Then
                Workbooks(OutputFileName).Save
                Workbooks(OutputFileName).Close
            End If
        Next
End Sub
Sub CheckDataSheet(Filename)
    For Each Sheet In Workbooks(Filename).Worksheets ' Checking if VBA Sheet exist
        If Sheet.name = "VBA" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
    Workbooks(Filename).Sheets.Add.name = "VBA"
End Sub
Sub RegionTitle(Filename, OutputFileName, Sheetname, PreviousDate, JobType, Region, RegionColum, RegionColumNum, ScopeColum, ScopeColumNum, DateColumA, DateColumANum, DateColum, DateColumNum, Status1Colum, Status1ColumNum, Status2Colum, Status2ColumNum, Status3Colum, Status3ColumNum, StatusFColum, StatusFColumNum, Out1, Out2, Out3, Out4, Out5, Out6, numofstatus, DesginerStatusColumNum, DesginerStatusColum)
    'NOTE: IF u want to add new colum Make a new function like this And call it after region title. Dont touch this one its the main part of program, This one works like fill in the blank. Everything is variable except for Sheet name. U can add after "' Complete Actual de" Finish and make sure to declare those varible abbove
    'Data
        TodayDate = Format(Date, "mm/dd/yyyy")
        Call CheckDataSheet(Filename) 'Delete the existing VBA sheet if exist and create a new one for data recording
        Workbooks(Filename).Sheets(Sheetname).Range(DateColumA).NumberFormat = "mm/dd/yyyy" 'Changing the formate of the date
        Workbooks(Filename).Sheets(Sheetname).Range(DateColum).NumberFormat = "mm/dd/yyyy"

    'New Recieved
        Workbooks(Filename).Sheets(Sheetname).Range(DateColumA).AutoFilter Field:=DateColumANum, Operator:=xlFilterValues, _
                                                 Criteria1:=TodayDate
        Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:="=", Operator:=xlFilterValues
        Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
        Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
        CountND = Workbooks(Filename).Sheets(Sheetname).AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        Workbooks(Filename).Sheets(Sheetname).ShowAllData
        If CountND <> 0 Then
            Workbooks(OutputFileName).Worksheets("Sheet1").Range(Out1).FormulaR1C1 = CountND
        End If

    'Spilover
        Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
        Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
        If numofstatus = 1 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            
        End If
        If numofstatus = 2 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
        End If
        If numofstatus = 3 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
        End If
        If numofstatus = 4 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(Status4Colum).AutoFilter Field:=Status4ColumNum, Criteria1:="="
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
        End If
        CountSO = Workbooks(Filename).Sheets(Sheetname).AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        Workbooks(Filename).Sheets(Sheetname).ShowAllData
        If CountSO <> 0 Then
        Workbooks(OutputFileName).Worksheets("Sheet1").Range(Out2).FormulaR1C1 = CountSO
        End If
        
    'IP
        If numofstatus = 1 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="IP" 'Checking for IP
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData




        End If
        If numofstatus = 2 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="IP" 'Checking for IP
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:="IP"
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets("VBA").Range("A1:B1").Delete
            CountIPA = Workbooks(Filename).Sheets(Sheetname).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
            CountIPB = Workbooks(Filename).Sheets(Sheetname).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
            CountIPB1 = "B1:B" & CountIPB + 1
            CountCopy = "A" & CountIPA + 1 & ":A" & CountIPA + CountIPB + 1
            Workbooks(Filename).Worksheets("VBA").Range(CountIPB1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy).PasteSpecial Paste:=xlPasteValues


        End If
        If numofstatus = 3 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="IP"
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:="IP"
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:="IP"
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets("VBA").Range("A1:C1").Delete

            CountIPA = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("A:A"))
            CountIPB = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("B:B"))
            CountIPC = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("C:C"))

            CountIPD1 = "B1:B" & CountIPB + 1
            CountIPE1 = "C1:C" & CountIPC + 1
            CountPaste = "A" & CountIPA + 1 & ":A" & CountIPA + CountIPB + 1
            CountPaste1 = "A" & CountIPA + CountIPB + 1 & ":A" & CountIPA + CountIPB + CountIPC + 1

            Workbooks(Filename).Worksheets("VBA").Range(CountIPD1).Copy
            Workbooks(Filename).Worksheets("VBA").Range(CountPaste).PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Worksheets("VBA").Range(CountIPE1).Copy
            Workbooks(Filename).Worksheets("VBA").Range(CountPaste1).PasteSpecial Paste:=xlPasteValues

        End If
        If numofstatus = 4 Then
            'Checking for IP
                Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:="IP"
                Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
                Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
                Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValues
                Workbooks(Filename).Sheets(Sheetname).ShowAllData
                Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:="IP"
                Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
                Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
                Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValues
                Workbooks(Filename).Sheets(Sheetname).ShowAllData
                Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:="IP"
                Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
                Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
                Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValues
                Workbooks(Filename).Sheets(Sheetname).ShowAllData
                Workbooks(Filename).Sheets(Sheetname).Range(Status4Colum).AutoFilter Field:=Status4ColumNum, Criteria1:="IP"
                Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
                Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
                Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValues
                Workbooks(Filename).Sheets(Sheetname).ShowAllData
                Workbooks(Filename).Sheets("VBA").Range("A1:D1").Delete

                CountIPA = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("A:A"))
                CountIPB = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("B:B"))
                CountIPC = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("C:C"))
                CountIPD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("D:D"))

                CountIPB1 = "B1:B" & CountIPB + 1
                CountIPC1 = "C1:C" & CountIPC + 1
                CountIPD1 = "D1:D" & CountIPD + 1
                
                CountPaste = "A" & CountIPA + 1 & ":A" & CountIPA + CountIPB + 1
                CountPaste1 = "A" & CountIPA + CountIPB + 1 & ":A" & CountIPA + CountIPB + CountIPC + 1
                CountPaste2 = "A" & CountIPA + CountIPB + CountIPC + 1 & ":A" & CountIPA + CountIPB + CountIPC + CountIPD + 1
                
                
                Workbooks(Filename).Worksheets("VBA").Range(CountIPB1).Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range(CountPaste).PasteSpecial Paste:=xlPasteValues
                Workbooks(Filename).Worksheets("VBA").Range(CountIPC1).Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range(CountPaste1).PasteSpecial Paste:=xlPasteValues
                Workbooks(Filename).Worksheets("VBA").Range(CountIPD1).Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range(CountPaste2).PasteSpecial Paste:=xlPasteValues
        End If
        Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:="IP"
        Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
        Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
        Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
        Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValues
        Workbooks(Filename).Sheets(Sheetname).ShowAllData
        Workbooks(Filename).Sheets("VBA").Range("E1").Delete
        CountIPE = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("E:E"))
        CountIPE1 = "E1:E" & CountIPE + 1
        CountPaste = "A" & CountIPA + CountIPB + CountIPC + CountIPD + 1 & ":A" & CountIPA + CountIPB + CountIPC + CountIPD + CountIPE + 1
        Workbooks(Filename).Worksheets("VBA").Range(CountIPE1).Copy 'Actual Delivery
        Workbooks(Filename).Worksheets("VBA").Range(CountPaste).PasteSpecial Paste:=xlPasteValues
        Workbooks(Filename).Worksheets("VBA").Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
        CountS = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("A:A"))
        If CountS <> 0 Then
            Workbooks(OutputFileName).Worksheets("Sheet1").Range(Out3).FormulaR1C1 = CountS
        End If

    'QC
        If numofstatus = 1 Then
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region 'QC, QC/blank
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            End If
        If numofstatus = 2 Then
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region 'QC, QC/blank
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            If Sheetname = "Commercial_ Expense Design" Or Sheetname = "Comm+Res+Other Design" Then
            Workbooks(Filename).Sheets(Sheetname).Range(DesginerStatusColum).AutoFilter Field:=DesginerStatusColumNum, Criteria1:="<>*=*", Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL", "="), Operator:=xlFilterValues
            End If

        End If
        If numofstatus = 3 Then
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region 'QC, QC, QC/blank
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            If Sheetname = "Node Split" Or Sheetname = "SFU&MDU Design" Or Sheetname = "FD Design" Then
                Workbooks(Filename).Sheets(Sheetname).Range(DesginerStatusColum).AutoFilter Field:=DesginerStatusColumNum, Criteria1:="<>*=*", Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL", "="), Operator:=xlFilterValues
            End If
        End If

        If numofstatus = 4 Then
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region 'QC, QC, QC, QC/blank
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(Status4Colum).AutoFilter Field:=Status4ColumNum, Criteria1:=Array("QC 1", "QC 2", "2nd QC", "FIXING", "EMAIL"), Operator:=xlFilterValues

        End If
        CountQC = Workbooks(Filename).Sheets(Sheetname).AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        Workbooks(Filename).Sheets(Sheetname).ShowAllData
        If CountQC <> 0 Then
        Workbooks(OutputFileName).Worksheets("Sheet1").Range(Out4).FormulaR1C1 = CountQC
        End If

        
    'RFI
        If numofstatus = 1 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData

        End If
        If numofstatus = 2 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets("VBA").Range("F1:G1").Delete
            CountRFIG = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("G:G"))
            CountRFIG1 = "G1:G" & CountRFIG + 1
            CountCopy = "F" & CountRFIF + 1 & ":F" & CountRFIF + CountRFIG + 1
            Workbooks(Filename).Worksheets("VBA").Range(CountRFIG1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy).PasteSpecial Paste:=xlPasteValues
        End If
        If numofstatus = 3 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets("VBA").Range("F1:H1").Delete
            CountRFIF = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("F:F"))
            CountRFIG = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("G:G"))
            CountRFIH = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("H:H"))
            CountRFIG1 = "G1:G" & CountRFIG + 1
            CountRFIH1 = "H1:H" & CountRFIH + 1
            CountCopy = "F" & CountRFIF + 1 & ":F" & CountRFIF + CountRFIG + 1
            CountCopy1 = "F" & CountRFIF + CountRFIG + 1 & ":F" & CountRFIF + CountRFIG + CountRFIH + 1
            
            Workbooks(Filename).Worksheets("VBA").Range(CountRFIG1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy).PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Worksheets("VBA").Range(CountRFIH1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy1).PasteSpecial Paste:=xlPasteValues
        End If
        If numofstatus = 4 Then
            Workbooks(Filename).Sheets(Sheetname).Range(Status1Colum).AutoFilter Field:=Status1ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status2Colum).AutoFilter Field:=Status2ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status3Colum).AutoFilter Field:=Status3ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets(Sheetname).Range(Status4Colum).AutoFilter Field:=Status4ColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
            Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
            Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
            Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Sheets(Sheetname).ShowAllData
            Workbooks(Filename).Sheets("VBA").Range("F1:I1").Delete
            CountRFIF = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("F:F"))
            CountRFIG = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("G:G"))
            CountRFIH = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("H:H"))
            CountRFII = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("I:I"))
            CountRFIG1 = "G1:G" & CountRFIG + 1
            CountRFIH1 = "H1:H" & CountRFIH + 1
            CountRFII1 = "I1:I" & CountRFII + 1
            CountCopy = "F" & CountRFIF + 1 & ":F" & CountRFIF + CountRFIG + 1
            CountCopy1 = "F" & CountRFIF + CountRFIG + 1 + ":F" & CountRFIF + CountRFIG + CountRFIH + 1
            CountCopy2 = "F" & CountRFIF + CountRFIG + CountRFIH + 1 & ":F" & CountRFIF + CountRFIG + CountRFIH + CountRFII + 1
            Workbooks(Filename).Worksheets("VBA").Range(CountRFIG1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy).PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Worksheets("VBA").Range(CountRFIH1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy1).PasteSpecial Paste:=xlPasteValues
            Workbooks(Filename).Worksheets("VBA").Range(CountRFII1).Copy 'Actual Delivery
            Workbooks(Filename).Worksheets("VBA").Range(CountCopy2).PasteSpecial Paste:=xlPasteValues
        End If
        Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("Pess Lock", "RFI"), Operator:=xlFilterValues
        Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
        Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Criteria1:=Array("="), Operator:=xlFilterValues
        Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
        Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
        Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValues
        Workbooks(Filename).Sheets(Sheetname).ShowAllData
        Workbooks(Filename).Sheets("VBA").Range("J1").Delete
        CountRFIJ = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("J:J"))
        CountRFIJ1 = "J1:J" & CountRFIJ + 1
        CountFCopy3 = "F" & CountRFIF + CountRFIG + CountRFIH + CountRFII + 1 & ":F" & CountRFIF + CountRFIG + CountRFIH + CountRFII + CountRFIJ + 1
        Workbooks(Filename).Worksheets("VBA").Range(CountRFIJ1).Copy 'Actual Delivery
        Workbooks(Filename).Worksheets("VBA").Range(CountFCopy3).PasteSpecial Paste:=xlPasteValues
        Workbooks(Filename).Worksheets("VBA").Range("$F$1:$F$1048492").RemoveDuplicates Columns:=1, Header:=xlNo
        CountRFI = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("F:F"))
        If CountRFI <> 0 Then
            Workbooks(OutputFileName).Worksheets("Sheet1").Range(Out5).FormulaR1C1 = CountRFI
        End If

    'Actual Delivery Previous Day
        Workbooks(Filename).Sheets(Sheetname).Range(DateColum).AutoFilter Field:=DateColumNum, Operator:=xlFilterValues, _
                                                 Criteria1:=PreviousDate
        Workbooks(Filename).Sheets(Sheetname).Range(RegionColum).AutoFilter Field:=RegionColumNum, Criteria1:=Region
        Workbooks(Filename).Sheets(Sheetname).Range(ScopeColum).AutoFilter Field:=ScopeColumNum, Criteria1:=JobType, Operator:=xlFilterValues
        Workbooks(Filename).Sheets(Sheetname).Range(StatusFColum).AutoFilter Field:=StatusFColumNum, Criteria1:=Array("Completed", "Completed  2 (Valid)", "Completed 2"), _
                                                    Operator:=xlFilterValues
        CountAD = Workbooks(Filename).Sheets(Sheetname).AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells.Count - 1
        Workbooks(Filename).Sheets(Sheetname).ShowAllData
        If CountAD <> 0 Then
            Workbooks(OutputFileName).Worksheets("Sheet1").Range(Out6).FormulaR1C1 = CountAD
        End If
        Workbooks(Filename).Sheets(Sheetname).Range(DateColumA).NumberFormat = "d-mmm"
        Workbooks(Filename).Sheets(Sheetname).Range(DateColum).NumberFormat = "d-mmm"

End Sub




