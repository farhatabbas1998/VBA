'Version: 1.2
'Weekly Task
'For Echobroadband
'User: Saiful
'By Farhat Abbas

Sub Reports()
    Application.ScreenUpdating = False
    'Collecting data
        x = InputBox("Enter Number Of Days")
        eDate = Date - x
        lastweek = eDate - 7
        Nextweek = eDate + 7
        lastweekdate = Format(lastweek, "d/mmmm/yyyy")
        Nextweekdate = Format(Nextweek, "d/mmmm/yyyy")
        StartingDateValue = 1 * (Date - DatePart("y", Date - 1))
        StartingDateValueold = 1 * (Date - 365 - DatePart("y", Date))
        StartingDateold = Format(StartingDateValueold, "d/mmmm/yyyy")
        StartingDate = Format(StartingDateValue, "d/mmmm/yyyy")
        TodayDate = Format(Date, "dmyyyy")

        MsgBox "Starting year old: " & StartingDateold
        MsgBox "Starting year: " & StartingDate
        MsgBox "Last week date: " & lastweekdate
        MsgBox "Next week date: " & Nextweekdate
        'MsgBox "Todays date " & TodayDate
        'C:\Users\saiful\Documents\ALTICE WEEKLY\WEDNESDAY REPORT\Weekly Report\Output\                 'location in saiful's pc
    'Input And Output
        OutputFileName = "WeeklyReport" & TodayDate & ".xlsx"               'Please Change "WeeklyReport.xlsx"                                         FILE     SHEET   Report
        Filename0 = "ECHO-NETWIN DESIGN Tracking List 2021.xlsx"             'Please Change "ECHO-NETWIN DESIGN Tracking List 2021.xlsx"                A,B,L    1,2,12  1,2
        Filename1 = "ECHO-Netwin-Feeder Tracking 2021.xlsx"                 'Please Change "ECHO-Netwin-Feeder Tracking 2021.xlsx"                     C        3       3
        Filename2 = "ECHO-Asbuild Netwin Tracking 2020.xlsx"                'Please Change "ECHO-Asbuild Netwin Tracking 2020.xlsx"                    D        4       4
        Filename3 = "Altice NAME CHANGES Tracking.xlsx"                     'Please Change "Altice NAME CHANGES Tracking.xlsx"                         E        5       5
        Filename4 = "ECHO Daily Tracking GPON test Sheet.xlsx"              'Please Change "ECHO Daily Tracking GPON test Sheet.xlsx"                  F        6       6
        Filename5 = "Tracking NODE Data Update Log PNI NODE Migration.xlsx" 'Please Change "Tracking NODE Data Update Log PNI NODE Migration.xlsx"     G        7       7
        Filename6 = "Altice Node Split Tracking.xlsx"                       'Please Change "Altice Node Split Tracking.xlsx"                           H        8       8
        Filename7 = "Altice Coax HFC Design Tracking.xlsx"                  'Please Change "Altice Coax HFC Design Tracking.xlsx"                      I,J      9,10    9,10
        Filename8 = "DOT Daily Tracking 2021.xlsx"                          'Please Change "DOT Daily Tracking 2021.xlsx"                              K        11
        Filename9 = "SECO RF Migration Daily Tracking.xlsx"                 'Please Change "SECO RF Migration Daily Tracking.xlsx"                     M        13

        Sheetname_1 = "DESIGN"
        Sheetname_2 = "REDESIGN"
        Sheetname_3 = "Sheet1"
        Sheetname_4 = "Sheet1"
        Sheetname_5 = "NY40CL"
        Sheetname_6 = "GPON"
        Sheetname_7 = "Updates needed"
        Sheetname_8 = "Sheet1"
        Sheetname_9 = "DESIGN"
        Sheetname_10 = "ASBUILT"
        Sheetname_11 = "Job Tracking"
        Sheetname_12 = "TEXAS DESIGN"
        Sheetname_13 = "RF Migration"



    'Working on data
        Workbooks.Open "D:\Downloads\" + Filename0 'Open Data file location 'Please Change "C:\Users\farhanah\Documents\ALTICE WEEKLY\WEDNESDAY REPORT\Weekly Report\Input\" Input Files
        Workbooks.Open "D:\Downloads\" + Filename1
        Workbooks.Open "D:\Downloads\" + Filename2
        Workbooks.Open "D:\Downloads\" + Filename3
        Workbooks.Open "D:\Downloads\" + Filename4
        Workbooks.Open "D:\Downloads\" + Filename5
        Workbooks.Open "D:\Downloads\" + Filename6
        Workbooks.Open "D:\Downloads\" + Filename7
        Workbooks.Open "D:\Downloads\" + Filename8
        Workbooks.Open "D:\Downloads\" + Filename9

        Call Weeklyreportpart1(eDate, StartingDate, Filename0, lastweekdate, Nextweekdate, Sheetname_1, Sheetname_2, OutputFileName, StartingDateold) 'Super old
        Call Weeklyreportpart2(eDate, StartingDate, Filename1, lastweekdate, Nextweekdate, Sheetname_3, OutputFileName, StartingDateold)
        Call Weeklyreportpart3(eDate, StartingDate, Filename2, lastweekdate, Nextweekdate, Sheetname_4, OutputFileName, StartingDateold)
        Call Weeklyreportpart4(eDate, StartingDate, Filename3, lastweekdate, Nextweekdate, Sheetname_5, OutputFileName, StartingDateold)
        Call Weeklyreportpart5(eDate, StartingDate, Filename4, lastweekdate, Nextweekdate, Sheetname_6, OutputFileName, StartingDateold)
        Call Weeklyreportpart6(eDate, StartingDate, Filename5, lastweekdate, Nextweekdate, Sheetname_7, OutputFileName)
        Call Weeklyreportpart7(eDate, StartingDate, Filename6, lastweekdate, Nextweekdate, Sheetname_8, OutputFileName)
        Call Weeklyreportpart8(eDate, StartingDate, Filename7, lastweekdate, Nextweekdate, Sheetname_9, OutputFileName)
        Call Weeklyreportpart9(eDate, StartingDate, Filename7, lastweekdate, Nextweekdate, Sheetname_10, OutputFileName)
        Call Weeklyreportpart11(Filename9, Sheetname_13, OutputFileName)
        Call Weeklyreportpart12(OutputFileName)
        Call Weeklyreportpart13(OutputFileName)
        Call Weeklyreportpart10(eDate, StartingDate, Filename8, lastweekdate, Nextweekdate, Sheetname_11, OutputFileName, StartingDateold)
        
    'Table surrounding
        Workbooks(OutputFileName).Worksheets("Sheet1").Columns("A:W").EntireColumn.AutoFit
        Workbooks(OutputFileName).Worksheets("Sheet1").Columns("B:H").ColumnWidth = 32
        Workbooks(OutputFileName).Save
        Application.ScreenUpdating = True
End Sub

Sub Weeklyreportpart1(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname_1, Sheetname_2, OutputFileName, StartingDateold)
    'Creating files
        Call Check_if_workbook_is_open(OutputFileName)
        Application.DisplayAlerts = False
        Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputFileName 'Please Change "C:\Users\saiful\Documents\ALTICE WEEKLY\WEDNESDAY REPORT\Weekly Report\Output\"
        Application.DisplayAlerts = True
        Dim Wb1In As Workbook
        Dim Ws1In As Worksheet
        Set Wb1In = Workbooks(Filename)
        Set Ws1In = Wb1In.Sheets(Sheetname_1)
        Dim Wb2In As Workbook
        Dim Ws2In As Worksheet
        Set Wb2In = Workbooks(Filename)
        Set Ws2In = Wb2In.Sheets(Sheetname_2)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With Ws1In
        
    'For table A. FTTH Design (Netwin) B. FTTH Redesign (Netwin)
        'Finding Value for total FTTH 1 Design
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & StartingDate, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                Countdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                Countdnc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week FTTH 2 Design
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                CountdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value last week

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Total Value last week

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value last week

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Total Value last week

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountdncL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1   'last week
                .ShowAllData

        'Finding RFI Value 3 Design
            'Finding value for nj
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value RFI

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Total Value RFI

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value RFI

            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Total Value RFI

            'Finding value for Nc
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountdNCRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'RFI
                .ShowAllData

        'Total cells planned to be delivered this week FTTH 4 Design
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                CountdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total cells planned to be delivered this week

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Nc
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountdncTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells FTTH 5 Design
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Remaining cells
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Remaining cells

            'Finding value for nyc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Remaining cells
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Remaining cells

            'Finding value for li
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Remaining cells
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Remaining cells

            'Finding value for Cwn, CT, WC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Remaining cells
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Remaining cells

            'Finding value for Nc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                Countdncrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countdncbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
        End With
        With Ws2In
        'Finding Value for total FTTH 1 redesign
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & StartingDateold, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                Countrnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value this year

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countrnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1 'Total Value this year

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countrli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn,CT,WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countrcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for TX
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("TX"), Operator:=xlFilterValues
                Countrtx = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week FTTH 2 redesign
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                CountrnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value last week

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountrnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountrliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountrcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Tx
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("TX"), Operator:=xlFilterValues
                CountrtxL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value 3 ReDesign
            'Finding value for njs
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountrnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total Value RFI

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountrnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountrliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountrcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Tx
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("TX"), Operator:=xlFilterValues
                CountdTxRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week FTTH 4 redesign
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountrnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Total cells planned to be delivered this week

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountrnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountrliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn,CT,WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountrcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Tx
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("TX"), Operator:=xlFilterValues
                CountrtxTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells FTTH 5 redesign
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countrnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1  'Remaining cells
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countrnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countrnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countrnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countrlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countrlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn,CT,WC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countrcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countrcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Tx
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(Nextweekdate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("TX"), Operator:=xlFilterValues
                Countrtxrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countrtxbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        End With
        With WsOut
        'Creating Surrounding
            'FTTH Design
                .Range("A5").FormulaR1C1 = "A. FTTH Design (Netwin)"
                .Range("A6").FormulaR1C1 = "Status Description"
                .Range("B6").FormulaR1C1 = "NJ"
                .Range("C6").FormulaR1C1 = "CWN"
                .Range("D6").FormulaR1C1 = "NYC"
                .Range("E6").FormulaR1C1 = "LI"
                .Range("F6").FormulaR1C1 = "NC"
                .Range("G6").FormulaR1C1 = "Total"
                .Range("H6").FormulaR1C1 = "Remark"
                .Range("A8").FormulaR1C1 = "Total cells delivered (" & Currentyear & ")"
                .Range("A9").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A10").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A11").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A12").FormulaR1C1 = "Remaining cells"
                .Range("A5:H7").Font.FontStyle = "Bold"
                .Range("A5").Font.Underline = xlUnderlineStyleSingle
                With .Range("A5").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A6:A7").Merge
                .Range("B6:B7").Merge
                .Range("C6:C7").Merge
                .Range("D6:D7").Merge
                .Range("E6:E7").Merge
                .Range("F6:F7").Merge
                .Range("G6:G7").Merge
                .Range("H6:H7").Merge

        'Creating Surrounding for FTTH Redesign
            'FTTH Redesign
                .Range("A17").FormulaR1C1 = "B. FTTH Redesign (Netwin)"
                .Range("A18").FormulaR1C1 = "Status Description"
                .Range("B18").FormulaR1C1 = "NJ"
                .Range("C18").FormulaR1C1 = "CWN"
                .Range("D18").FormulaR1C1 = "NYC"
                .Range("E18").FormulaR1C1 = "LI"
                .Range("F18").FormulaR1C1 = "TX"
                .Range("G18").FormulaR1C1 = "Total"
                .Range("H18").FormulaR1C1 = "Remark"
                .Range("A20").FormulaR1C1 = "Total cells delivered"
                .Range("A21").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A22").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A23").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A24").FormulaR1C1 = "Remaining cells"
                .Range("A17:H19").Font.FontStyle = "Bold"
                .Range("A17").Font.Underline = xlUnderlineStyleSingle
                With .Range("A17").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A18:A19").Merge
                .Range("B18:B19").Merge
                .Range("C18:C19").Merge
                .Range("D18:D19").Merge
                .Range("E18:E19").Merge
                .Range("F18:F19").Merge
                .Range("G18:G19").Merge
                .Range("H18:H19").Merge
        'Calculating the total
            'Design this year
                .Range("B8").FormulaR1C1 = Countdnj
                .Range("C8").FormulaR1C1 = Countdcwn
                .Range("D8").FormulaR1C1 = Countdnyc
                .Range("E8").FormulaR1C1 = Countdli
                .Range("B9").FormulaR1C1 = CountdnjL
                .Range("C9").FormulaR1C1 = CountdcwnL
                .Range("D9").FormulaR1C1 = CountdnycL
                .Range("E9").FormulaR1C1 = CountdliL
                .Range("B11").FormulaR1C1 = CountdnjRFI
                .Range("C11").FormulaR1C1 = CountdcwnRFI
                .Range("D11").FormulaR1C1 = CountdnycRFI
                .Range("E11").FormulaR1C1 = CountdliRFI
                .Range("B10").FormulaR1C1 = CountdnjTW
                .Range("C10").FormulaR1C1 = CountdcwnTW
                .Range("D10").FormulaR1C1 = CountdnycTW
                .Range("E10").FormulaR1C1 = CountdliTW
                .Range("B12").FormulaR1C1 = Countdnjrc + Countdnjbrc
                .Range("C12").FormulaR1C1 = Countdcwnbrc + Countdcwnrc
                .Range("D12").FormulaR1C1 = Countdnycbrc + Countdnycrc
                .Range("E12").FormulaR1C1 = Countdlibrc + Countdlirc
                'design F
                    .Range("F8").FormulaR1C1 = Countdnc  'Total Dilvery Desgin
                    .Range("F9").FormulaR1C1 = CountdncL  'last week Desgin
                    .Range("F10").FormulaR1C1 = CountdncTW 'Total cells planned to be delivered this week
                    .Range("F11").FormulaR1C1 = CountrtxRFI  'rfi Desgin
                    .Range("F12").FormulaR1C1 = Countdncrc + Countdncbrc 'remaing cells Desgin

                Count1TD = WorksheetFunction.Sum(.Range("B8:F8"))
                Count2TD = WorksheetFunction.Sum(.Range("B9:F9"))
                Count3TD = WorksheetFunction.Sum(.Range("B10:F10"))
                Count4TD = WorksheetFunction.Sum(.Range("B11:F11"))
                Count5TD = WorksheetFunction.Sum(.Range("B12:F12"))
                .Range("G8").FormulaR1C1 = Count1TD
                .Range("G9").FormulaR1C1 = Count2TD
                .Range("G10").FormulaR1C1 = Count3TD
                .Range("G11").FormulaR1C1 = Count4TD
                .Range("G12").FormulaR1C1 = Count5TD


            'Redesign
                .Range("B20").FormulaR1C1 = Countrnj
                .Range("C20").FormulaR1C1 = Countrcwn 'Total Dilvery
                .Range("D20").FormulaR1C1 = Countrnyc
                .Range("E20").FormulaR1C1 = Countrli

                .Range("B21").FormulaR1C1 = CountrnjL
                .Range("C21").FormulaR1C1 = CountrcwnL 'last week
                .Range("D21").FormulaR1C1 = CountrnycL
                .Range("E21").FormulaR1C1 = CountrliL

                .Range("B23").FormulaR1C1 = CountrnjRFI
                .Range("C23").FormulaR1C1 = CountrcwnRFI 'rfi
                .Range("D23").FormulaR1C1 = CountrnycRFI
                .Range("E23").FormulaR1C1 = CountrliRFI

                .Range("B22").FormulaR1C1 = CountrnjTW
                .Range("C22").FormulaR1C1 = CountrcwnTW
                .Range("D22").FormulaR1C1 = CountrnycTW
                .Range("E22").FormulaR1C1 = CountrliTW 'this week
                .Range("B24").FormulaR1C1 = Countrnjrc + Countrnjbrc 'remaing cells
                .Range("C24").FormulaR1C1 = Countrcwnbrc + Countrcwnrc
                .Range("D24").FormulaR1C1 = Countrnycbrc + Countrnycrc
                .Range("E24").FormulaR1C1 = Countrlibrc + Countrlirc
                'F  'Redesign

                        .Range("F20").FormulaR1C1 = Countrtx 'Total Dilvery Redesgin
                        .Range("F21").FormulaR1C1 = CountrtxL 'last week Redesgin
                        .Range("F22").FormulaR1C1 = CountrtxTW
                        .Range("F23").FormulaR1C1 = CountrtxRFI 'rfi Redesgin
                        .Range("F24").FormulaR1C1 = Countrtxrc + Countrtxbrc 'remaing cells Redesgin

                Count1TR = WorksheetFunction.Sum(.Range("B20:F20"))
                Count2TR = WorksheetFunction.Sum(.Range("B21:F21"))
                Count3TR = WorksheetFunction.Sum(.Range("B22:F22"))
                Count4TR = WorksheetFunction.Sum(.Range("B23:F23"))
                Count5TR = WorksheetFunction.Sum(.Range("B24:F24"))
                .Range("G20").FormulaR1C1 = Count1TR
                .Range("G21").FormulaR1C1 = Count2TR
                .Range("G22").FormulaR1C1 = Count3TR
                .Range("G23").FormulaR1C1 = Count4TR
                .Range("G24").FormulaR1C1 = Count5TR
            
        End With
        'Table surrounding
            tablesize = "A18:H24"
            tableheader = "A18:H19"
            tabledata = "B20:H24"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            tablesize = "A6:H12"
            tableheader = "A6:H7"
            tabledata = "B8:H12"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)

        'Rfi Data
            Call rfidata01(CountdnjRFI, CountdcwnRFI, CountdnycRFI, CountdliRFI, CountdNCRFI, Count4TD, Sheetname_1, Filename)
            Call rfidata02(CountrnjRFI, CountrcwnRFI, CountrnycRFI, CountrliRFI, CountdTxRFI, Count4TR, Sheetname_2, Filename)
                       
End Sub

Sub Weeklyreportpart2(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName, StartingDateold)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn
    'For table C. FTTH Feeder design (Netwin)
        'Finding Value for total
            'Finding value for nj
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & StartingDateold, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("I:I").AutoFilter Field:=9, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                Countfdnc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week FTTH Feeder design (Netwin)
            'Finding value for nj
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("I:I").AutoFilter Field:=9, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountfdncL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week FTTH Feeder design (Netwin)
            'Finding value for nj
                .Range("J:J").AutoFilter Field:=10, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("I:I").AutoFilter Field:=9, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "="), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountfdncTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value FTTH Feeder design (Netwin)
            'Finding value for njs, NJN
                .Range("I:I").AutoFilter Field:=9, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for nc
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountfdncRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells FTTH Feeder design (Netwin)
            'Finding value for nj
                .Range("J:J").AutoFilter Field:=10, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("I:I").AutoFilter Field:=9, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("J:J").AutoFilter Field:=10, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("J:J").AutoFilter Field:=10, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("J:J").AutoFilter Field:=10, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("J:J").AutoFilter Field:=10, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("J:J").AutoFilter Field:=10, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("J:J").AutoFilter Field:=10, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("J:J").AutoFilter Field:=10, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nc
                .Range("J:J").AutoFilter Field:=10, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdncrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("J:J").AutoFilter Field:=10, Criteria1:="="
                Countfdncbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
        End With
        With WsOut
        'Creating Surrounding for FTTH Feeder design (Netwin)
            
                .Range("A29").FormulaR1C1 = "C. FTTH Feeder design (Netwin)"
                .Range("A30").FormulaR1C1 = "Status Description"
                .Range("B30").FormulaR1C1 = "NJ"
                .Range("C30").FormulaR1C1 = "CWN"
                .Range("D30").FormulaR1C1 = "NYC"
                .Range("E30").FormulaR1C1 = "LI"
                .Range("F30").FormulaR1C1 = "NC"
                .Range("G30").FormulaR1C1 = "Total"
                .Range("H30").FormulaR1C1 = "Remark"
                .Range("A32").FormulaR1C1 = "Total cells delivered"
                .Range("A33").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A34").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A35").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A36").FormulaR1C1 = "Remaining cells"
                .Range("A29:H30").Font.FontStyle = "Bold"
                .Range("A29").Font.Underline = xlUnderlineStyleSingle
                With .Range("A29").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A30:A31").Merge
                .Range("B30:B31").Merge
                .Range("C30:C31").Merge
                .Range("D30:D31").Merge
                .Range("E30:E31").Merge
                .Range("F30:F31").Merge
                .Range("G30:G31").Merge
                .Range("H30:H31").Merge

        'Calculating the total for FTTH Feeder design (Netwin)
                .Range("B32").FormulaR1C1 = Countfdnj
                .Range("C32").FormulaR1C1 = Countfdcwn
                .Range("D32").FormulaR1C1 = Countfdnyc
                .Range("E32").FormulaR1C1 = Countfdli
                .Range("F32").FormulaR1C1 = Countfdnc
                .Range("B33").FormulaR1C1 = CountfdnjL
                .Range("C33").FormulaR1C1 = CountfdcwnL
                .Range("D33").FormulaR1C1 = CountfdnycL
                .Range("E33").FormulaR1C1 = CountfdliL
                .Range("F33").FormulaR1C1 = CountfdncL
                .Range("B34").FormulaR1C1 = CountfdnjTW
                .Range("C34").FormulaR1C1 = CountfdcwnTW
                .Range("D34").FormulaR1C1 = CountfdnycTW
                .Range("E34").FormulaR1C1 = CountfdliTW
                .Range("F34").FormulaR1C1 = CountfdncTW
                .Range("B35").FormulaR1C1 = CountfdnjRFI
                .Range("C35").FormulaR1C1 = CountfdcwnRFI
                .Range("D35").FormulaR1C1 = CountfdnycRFI
                .Range("E35").FormulaR1C1 = CountfdliRFI
                .Range("F35").FormulaR1C1 = CountfdncRFI
                .Range("B36").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C36").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D36").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E36").FormulaR1C1 = Countfdlibrc + Countfdlirc
                .Range("F36").FormulaR1C1 = Countfdncrc + Countfdncbrc
                Count1fd = WorksheetFunction.Sum(.Range("B32:F32"))
                Count2fd = WorksheetFunction.Sum(.Range("B33:F33"))
                Count3fd = WorksheetFunction.Sum(.Range("B34:F34"))
                Count4fd = WorksheetFunction.Sum(.Range("B35:F35"))
                Count5fd = WorksheetFunction.Sum(.Range("B36:F36"))
                .Range("G32").FormulaR1C1 = Count1fd
                .Range("G33").FormulaR1C1 = Count2fd
                .Range("G34").FormulaR1C1 = Count3fd
                .Range("G35").FormulaR1C1 = Count4fd
                .Range("G36").FormulaR1C1 = Count5fd
        End With
        'Table surrounding
            tablesize = "A30:H36"
            tableheader = "A30:H31"
            tabledata = "B32:H36"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)

        'RFI Data
            Call rfidata03(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, CountfdncRFI, Count4fd, Sheetname, Filename)
End Sub

Sub Weeklyreportpart3(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName, StartingDateold)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
    'For table D. FTTH Asbuilt (Netwin)
        With WsIn
        'Finding Value for total
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & StartingDateold, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                Countfdnc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc

                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountfdncL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountfdncTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                CountfdncRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for NC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NC"), Operator:=xlFilterValues
                Countfdncrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdncbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
        End With
        With WsOut
        'Creating Surrounding
                .Range("A41").FormulaR1C1 = "D. FTTH Asbuilt (Netwin)"
                .Range("A42").FormulaR1C1 = "Status Description"
                .Range("B42").FormulaR1C1 = "NJ"
                .Range("C42").FormulaR1C1 = "CWN"
                .Range("D42").FormulaR1C1 = "NYC"
                .Range("E42").FormulaR1C1 = "LI"
                .Range("F42").FormulaR1C1 = "NC"
                .Range("G42").FormulaR1C1 = "Total"
                .Range("H42").FormulaR1C1 = "Remark"
                .Range("A44").FormulaR1C1 = "Total cells delivered"
                .Range("A45").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A46").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A47").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A48").FormulaR1C1 = "Remaining cells"
                .Range("A41:H43").Font.FontStyle = "Bold"
                .Range("A41").Font.Underline = xlUnderlineStyleSingle
                With .Range("A41").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A42:A43").Merge
                .Range("B42:B43").Merge
                .Range("C42:C43").Merge
                .Range("D42:D43").Merge
                .Range("E42:E43").Merge
                .Range("F42:F43").Merge
                .Range("G42:G43").Merge
                .Range("H42:H43").Merge

        'Calculating the total
                .Range("B44").FormulaR1C1 = Countfdnj
                .Range("C44").FormulaR1C1 = Countfdcwn
                .Range("D44").FormulaR1C1 = Countfdnyc
                .Range("E44").FormulaR1C1 = Countfdli
                .Range("F44").FormulaR1C1 = Countfdnc
                .Range("B45").FormulaR1C1 = CountfdnjL
                .Range("C45").FormulaR1C1 = CountfdcwnL
                .Range("D45").FormulaR1C1 = CountfdnycL
                .Range("E45").FormulaR1C1 = CountfdliL
                .Range("F45").FormulaR1C1 = CountfdncL
                .Range("B46").FormulaR1C1 = CountfdnjTW
                .Range("C46").FormulaR1C1 = CountfdcwnTW
                .Range("D46").FormulaR1C1 = CountfdnycTW
                .Range("E46").FormulaR1C1 = CountfdliTW
                .Range("F46").FormulaR1C1 = CountfdncTW
                .Range("B47").FormulaR1C1 = CountfdnjRFI
                .Range("C47").FormulaR1C1 = CountfdcwnRFI
                .Range("D47").FormulaR1C1 = CountfdnycRFI
                .Range("E47").FormulaR1C1 = CountfdliRFI
                .Range("F47").FormulaR1C1 = CountfdncRFI
                .Range("B48").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C48").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D48").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E48").FormulaR1C1 = Countfdlibrc + Countfdlirc
                .Range("F48").FormulaR1C1 = Countfdncrc + Countfdncbrc
                Count1fd = WorksheetFunction.Sum(.Range("B44:F44"))
                Count2fd = WorksheetFunction.Sum(.Range("B45:F45"))
                Count3fd = WorksheetFunction.Sum(.Range("B46:F46"))
                Count4fd = WorksheetFunction.Sum(.Range("B47:F47"))
                Count5fd = WorksheetFunction.Sum(.Range("B48:F48"))
                .Range("G44").FormulaR1C1 = Count1fd
                .Range("G45").FormulaR1C1 = Count2fd
                .Range("G46").FormulaR1C1 = Count3fd
                .Range("G47").FormulaR1C1 = Count4fd
                .Range("G48").FormulaR1C1 = Count5fd
            End With

        'Table surrounding
            tablesize = "A42:H48"
            tableheader = "A42:H43"
            tabledata = "B44:H48"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)

        'RFI Data
            Call rfidata04(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, CountfdncRFI, Count4fd, Sheetname, Filename)
End Sub

Sub Weeklyreportpart4(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName, StartingDateold)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn
    'For table E. FTTH Cell Rename (Netwin)
        'Finding Value for total
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & StartingDateold, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc

                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        End With
        With WsOut
        'Creating Surrounding
                .Range("A53").FormulaR1C1 = "E. FTTH Cell Rename (Netwin)"
                .Range("A54").FormulaR1C1 = "Status Description"
                .Range("B54").FormulaR1C1 = "NJ"
                .Range("C54").FormulaR1C1 = "CWN"
                .Range("D54").FormulaR1C1 = "NYC"
                .Range("E54").FormulaR1C1 = "LI"
                .Range("F54").FormulaR1C1 = "Total"
                .Range("G54").FormulaR1C1 = "Remark"
                .Range("A56").FormulaR1C1 = "Total cells delivered"
                .Range("A57").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A58").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A59").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A60").FormulaR1C1 = "Remaining cells"
                .Range("A53:G55").Font.FontStyle = "Bold"
                .Range("A53").Font.Underline = xlUnderlineStyleSingle
                With .Range("A53").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A54:A55").Merge
                .Range("B54:B55").Merge
                .Range("C54:C55").Merge
                .Range("D54:D55").Merge
                .Range("E54:E55").Merge
                .Range("F54:F55").Merge
                .Range("G54:G55").Merge

        'Calculating the total
                .Range("B56").FormulaR1C1 = Countfdnj
                .Range("C56").FormulaR1C1 = Countfdcwn
                .Range("D56").FormulaR1C1 = Countfdnyc
                .Range("E56").FormulaR1C1 = Countfdli
                .Range("B57").FormulaR1C1 = CountfdnjL
                .Range("C57").FormulaR1C1 = CountfdcwnL
                .Range("D57").FormulaR1C1 = CountfdnycL
                .Range("E57").FormulaR1C1 = CountfdliL
                .Range("B58").FormulaR1C1 = CountfdnjTW
                .Range("C58").FormulaR1C1 = CountfdcwnTW
                .Range("D58").FormulaR1C1 = CountfdnycTW
                .Range("E58").FormulaR1C1 = CountfdliTW
                .Range("B59").FormulaR1C1 = CountfdnjRFI
                .Range("C59").FormulaR1C1 = CountfdcwnRFI
                .Range("D59").FormulaR1C1 = CountfdnycRFI
                .Range("E59").FormulaR1C1 = CountfdliRFI
                .Range("B60").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C60").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D60").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E60").FormulaR1C1 = Countfdlibrc + Countfdlirc
                Count1fd = WorksheetFunction.Sum(.Range("B56:E56"))
                Count2fd = WorksheetFunction.Sum(.Range("B57:E57"))
                Count3fd = WorksheetFunction.Sum(.Range("B58:E58"))
                Count4fd = WorksheetFunction.Sum(.Range("B59:E59"))
                Count5fd = WorksheetFunction.Sum(.Range("B60:E60"))
                .Range("F56").FormulaR1C1 = Count1fd
                .Range("F57").FormulaR1C1 = Count2fd
                .Range("F58").FormulaR1C1 = Count3fd
                .Range("F59").FormulaR1C1 = Count4fd
                .Range("F60").FormulaR1C1 = Count5fd
        End With
        'Table surrounding
            tablesize = "A54:G60"
            tableheader = "A54:G55"
            tabledata = "B56:G60"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            Call rfidata05(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, Count4fd, Sheetname, Filename)
End Sub

Sub Weeklyreportpart5(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName, StartingDateold)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn
    'For table F. EOL test sheet (Netwin)
        'Finding Value for total
            'Finding value for nj
                .Range("R:R").AutoFilter Field:=18, Operator:=xlFilterValues, Criteria1:=">=" & StartingDateold, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("B:B").AutoFilter Field:=2, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
                
        'Finding Value for Last week
            'Finding value for nj
                .Range("R:R").AutoFilter Field:=18, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("B:B").AutoFilter Field:=2, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("Q:Q").AutoFilter Field:=17, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("B:B").AutoFilter Field:=2, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("B:B").AutoFilter Field:=2, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("Q:Q").AutoFilter Field:=17, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("N:N").AutoFilter Field:=14, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Q:Q").AutoFilter Field:=17, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("Q:Q").AutoFilter Field:=17, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("B:B").AutoFilter Field:=2, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Q:Q").AutoFilter Field:=17, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("Q:Q").AutoFilter Field:=17, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Q:Q").AutoFilter Field:=17, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("Q:Q").AutoFilter Field:=17, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("B:B").AutoFilter Field:=2, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Q:Q").AutoFilter Field:=17, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
        End With
        With WsOut
        'Creating Surrounding
                .Range("A65").FormulaR1C1 = "F. EOL test sheet (Netwin)"
                .Range("A66").FormulaR1C1 = "Status Description"
                .Range("B66").FormulaR1C1 = "NJ"
                .Range("C66").FormulaR1C1 = "CWN"
                .Range("D66").FormulaR1C1 = "NYC"
                .Range("E66").FormulaR1C1 = "LI"
                .Range("F66").FormulaR1C1 = "Total"
                .Range("G66").FormulaR1C1 = "Remark"
                .Range("A68").FormulaR1C1 = "Total cells delivered"
                .Range("A69").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A70").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A71").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A72").FormulaR1C1 = "Remaining cells"
                .Range("A65:G67").Font.FontStyle = "Bold"
                .Range("A65").Font.Underline = xlUnderlineStyleSingle
                With .Range("A65").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A66:A67").Merge
                .Range("B66:B67").Merge
                .Range("C66:C67").Merge
                .Range("D66:D67").Merge
                .Range("E66:E67").Merge
                .Range("F66:F67").Merge
                .Range("G66:G67").Merge

        'Calculating the total
            .Range("B68").FormulaR1C1 = Countfdnj
            .Range("C68").FormulaR1C1 = Countfdcwn
            .Range("D68").FormulaR1C1 = Countfdnyc
            .Range("E68").FormulaR1C1 = Countfdli
            .Range("B69").FormulaR1C1 = CountfdnjL
            .Range("C69").FormulaR1C1 = CountfdcwnL
            .Range("D69").FormulaR1C1 = CountfdnycL
            .Range("E69").FormulaR1C1 = CountfdliL
            .Range("B70").FormulaR1C1 = CountfdnjTW
            .Range("C70").FormulaR1C1 = CountfdcwnTW
            .Range("D70").FormulaR1C1 = CountfdnycTW
            .Range("E70").FormulaR1C1 = CountfdliTW
            .Range("B71").FormulaR1C1 = CountfdnjRFI
            .Range("C71").FormulaR1C1 = CountfdcwnRFI
            .Range("D71").FormulaR1C1 = CountfdnycRFI
            .Range("E71").FormulaR1C1 = CountfdliRFI
            .Range("B72").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
            .Range("C72").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
            .Range("D72").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
            .Range("E72").FormulaR1C1 = Countfdlibrc + Countfdlirc
            Count1fd = WorksheetFunction.Sum(.Range("B68:E68"))
            Count2fd = WorksheetFunction.Sum(.Range("B69:E69"))
            Count3fd = WorksheetFunction.Sum(.Range("B70:E70"))
            Count4fd = WorksheetFunction.Sum(.Range("B71:E71"))
            Count5fd = WorksheetFunction.Sum(.Range("B72:E72"))
            .Range("F68").FormulaR1C1 = Count1fd
            .Range("F69").FormulaR1C1 = Count2fd
            .Range("F70").FormulaR1C1 = Count3fd
            .Range("F71").FormulaR1C1 = Count4fd
            .Range("F72").FormulaR1C1 = Count5fd
        End With
        
        'Table surrounding
            tablesize = "A66:G72"
            tableheader = "A66:G67"
            tabledata = "B68:G72"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            Call rfidata06(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, Count4fd, Sheetname, Filename)
End Sub

Sub Weeklyreportpart6(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn

    'For table G. HFC Fiber Node (Fiber)
        'Finding Value for total
            'Finding value for nj
                .Range("AA:AA").AutoFilter Field:=27, Operator:=xlFilterValues, Criteria1:=">=" & StartingDate, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("AI:AI").AutoFilter Field:=35, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("AA:AA").AutoFilter Field:=27, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("AI:AI").AutoFilter Field:=35, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("Z:Z").AutoFilter Field:=26, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("AI:AI").AutoFilter Field:=35, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for nj
                .Range("AI:AI").AutoFilter Field:=35, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("Z:Z").AutoFilter Field:=26, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("AI:AI").AutoFilter Field:=35, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Z:Z").AutoFilter Field:=26, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                
            'Finding value for nyc
                .Range("Z:Z").AutoFilter Field:=26, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Z:Z").AutoFilter Field:=26, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                
            'Finding value for li
                .Range("Z:Z").AutoFilter Field:=26, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Z:Z").AutoFilter Field:=26, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("Z:Z").AutoFilter Field:=26, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("Z:Z").AutoFilter Field:=26, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        End With
        With WsOut
        'Creating Surrounding
                .Range("A77").FormulaR1C1 = "G. HFC Fiber Node (Fiber)"
                .Range("A78").FormulaR1C1 = "Status Description"
                .Range("B78").FormulaR1C1 = "NJ"
                .Range("C78").FormulaR1C1 = "CWN"
                .Range("D78").FormulaR1C1 = "NYC"
                .Range("E78").FormulaR1C1 = "LI"
                .Range("F78").FormulaR1C1 = "Total"
                .Range("G78").FormulaR1C1 = "Remark"
                .Range("A80").FormulaR1C1 = "Total cells delivered (" & Currentyear & ")"
                .Range("A81").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A82").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A83").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A84").FormulaR1C1 = "Remaining cells"
                .Range("A77:G79").Font.FontStyle = "Bold"
                .Range("A77").Font.Underline = xlUnderlineStyleSingle
                With .Range("A77").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A78:A79").Merge
                .Range("B78:B79").Merge
                .Range("C78:C79").Merge
                .Range("D78:D79").Merge
                .Range("E78:E79").Merge
                .Range("F78:F79").Merge
                .Range("G78:G79").Merge

        'Calculating the total
                .Range("B80").FormulaR1C1 = Countfdnj
                .Range("C80").FormulaR1C1 = Countfdcwn
                .Range("D80").FormulaR1C1 = Countfdnyc
                .Range("E80").FormulaR1C1 = Countfdli
                .Range("B81").FormulaR1C1 = CountfdnjL
                .Range("C81").FormulaR1C1 = CountfdcwnL
                .Range("D81").FormulaR1C1 = CountfdnycL
                .Range("E81").FormulaR1C1 = CountfdliL
                .Range("B82").FormulaR1C1 = CountfdnjTW
                .Range("C82").FormulaR1C1 = CountfdcwnTW
                .Range("D82").FormulaR1C1 = CountfdnycTW
                .Range("E82").FormulaR1C1 = CountfdliTW
                .Range("B83").FormulaR1C1 = CountfdnjRFI
                .Range("C83").FormulaR1C1 = CountfdcwnRFI
                .Range("D83").FormulaR1C1 = CountfdnycRFI
                .Range("E83").FormulaR1C1 = CountfdliRFI
                .Range("B84").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C84").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D84").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E84").FormulaR1C1 = Countfdlibrc + Countfdlirc
                Count1fd = WorksheetFunction.Sum(.Range("B80:E80"))
                Count2fd = WorksheetFunction.Sum(.Range("B81:E81"))
                Count3fd = WorksheetFunction.Sum(.Range("B82:E82"))
                Count4fd = WorksheetFunction.Sum(.Range("B83:E83"))
                Count5fd = WorksheetFunction.Sum(.Range("B84:E84"))
                .Range("F80").FormulaR1C1 = Count1fd
                .Range("F81").FormulaR1C1 = Count2fd
                .Range("F82").FormulaR1C1 = Count3fd
                .Range("F83").FormulaR1C1 = Count4fd
                .Range("F84").FormulaR1C1 = Count5fd

        End With
        'Table surrounding
            tablesize = "A78:G84"
            tableheader = "A78:G79"
            tabledata = "B80:G84"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            Call rfidata07(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, Count4fd, Sheetname, Filename)
End Sub

Sub Weeklyreportpart7(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn

    'For table H. NJ Node Split Design (Coax)
        'Finding Value for total
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & StartingDate, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                
            'Finding value for nyc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
                
        End With
        With WsOut
        'Creating Surrounding
            
                .Range("A89").FormulaR1C1 = "H. NJ Node Split Design (Coax)"
                .Range("A90").FormulaR1C1 = "Status Description"
                .Range("B90").FormulaR1C1 = "NJ"
                .Range("C90").FormulaR1C1 = "CWN"
                .Range("D90").FormulaR1C1 = "NYC"
                .Range("E90").FormulaR1C1 = "LI"
                .Range("F90").FormulaR1C1 = "Total"
                .Range("G90").FormulaR1C1 = "Remark"
                .Range("A92").FormulaR1C1 = "Total cells delivered (" & Currentyear & ")"
                .Range("A93").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A94").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A95").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A96").FormulaR1C1 = "Remaining cells"
                .Range("A89:G90").Font.FontStyle = "Bold"
                .Range("A89").Font.Underline = xlUnderlineStyleSingle
                With .Range("A89").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A90:A91").Merge
                .Range("B90:B91").Merge
                .Range("C90:C91").Merge
                .Range("D90:D91").Merge
                .Range("E90:E91").Merge
                .Range("F90:F91").Merge
                .Range("G90:G91").Merge

        'Calculating the total
                .Range("B92").FormulaR1C1 = Countfdnj
                .Range("C92").FormulaR1C1 = Countfdcwn
                .Range("D92").FormulaR1C1 = Countfdnyc
                .Range("E92").FormulaR1C1 = Countfdli
                .Range("B93").FormulaR1C1 = CountfdnjL
                .Range("C93").FormulaR1C1 = CountfdcwnL
                .Range("D93").FormulaR1C1 = CountfdnycL
                .Range("E93").FormulaR1C1 = CountfdliL
                .Range("B94").FormulaR1C1 = CountfdnjTW
                .Range("C94").FormulaR1C1 = CountfdcwnTW
                .Range("D94").FormulaR1C1 = CountfdnycTW
                .Range("E94").FormulaR1C1 = CountfdliTW
                .Range("B95").FormulaR1C1 = CountfdnjRFI
                .Range("C95").FormulaR1C1 = CountfdcwnRFI
                .Range("D95").FormulaR1C1 = CountfdnycRFI
                .Range("E95").FormulaR1C1 = CountfdliRFI
                .Range("B96").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C96").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D96").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E96").FormulaR1C1 = Countfdlibrc + Countfdlirc
                Count1fd = WorksheetFunction.Sum(.Range("B92:E92"))
                Count2fd = WorksheetFunction.Sum(.Range("B93:E93"))
                Count3fd = WorksheetFunction.Sum(.Range("B94:E94"))
                Count4fd = WorksheetFunction.Sum(.Range("B95:E95"))
                Count5fd = WorksheetFunction.Sum(.Range("B96:E96"))
                .Range("F92").FormulaR1C1 = Count1fd
                .Range("F93").FormulaR1C1 = Count2fd
                .Range("F94").FormulaR1C1 = Count3fd
                .Range("F95").FormulaR1C1 = Count4fd
                .Range("F96").FormulaR1C1 = Count5fd
        End With
        'Table surrounding
            tablesize = "A90:G96"
            tableheader = "A90:G91"
            tabledata = "B92:G96"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            Call rfidata08(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, Count4fd, Sheetname, Filename)

End Sub

Sub Weeklyreportpart8(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn
    'For table I. Coax Design (Coax)
        'Finding Value for total
            'Finding value for nj
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:=">=" & StartingDate, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("K:K").AutoFilter Field:=11, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("K:K").AutoFilter Field:=11, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("K:K").AutoFilter Field:=11, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NYC"), Operator:=xlFilterValues
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1


            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                
            'Finding value for lie
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("K:K").AutoFilter Field:=11, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("M:M").AutoFilter Field:=13, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                
            'Finding value for nyc
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("M:M").AutoFilter Field:=13, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("M:M").AutoFilter Field:=13, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("M:M").AutoFilter Field:=13, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData
        End With
        With WsOut
        'Creating Surrounding
                .Range("A101").FormulaR1C1 = "I. Coax Design (Coax)"
                .Range("A102").FormulaR1C1 = "Status Description"
                .Range("B102").FormulaR1C1 = "NJ"
                .Range("C102").FormulaR1C1 = "CWN"
                .Range("D102").FormulaR1C1 = "NYC"
                .Range("E102").FormulaR1C1 = "LI"
                .Range("F102").FormulaR1C1 = "Total"
                .Range("G102").FormulaR1C1 = "Remark"
                .Range("A104").FormulaR1C1 = "Total cells delivered (" & Currentyear & ")"
                .Range("A105").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A106").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A107").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A108").FormulaR1C1 = "Remaining cells"
                .Range("A101:G102").Font.FontStyle = "Bold"
                .Range("A101").Font.Underline = xlUnderlineStyleSingle
                With .Range("A101").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A102:A103").Merge
                .Range("B102:B103").Merge
                .Range("C102:C103").Merge
                .Range("D102:D103").Merge
                .Range("E102:E103").Merge
                .Range("F102:F103").Merge
                .Range("G102:G103").Merge

        'Calculating the total
                .Range("B104").FormulaR1C1 = Countfdnj
                .Range("C104").FormulaR1C1 = Countfdcwn
                .Range("D104").FormulaR1C1 = Countfdnyc
                .Range("E104").FormulaR1C1 = Countfdli
                .Range("B105").FormulaR1C1 = CountfdnjL
                .Range("C105").FormulaR1C1 = CountfdcwnL
                .Range("D105").FormulaR1C1 = CountfdnycL
                .Range("E105").FormulaR1C1 = CountfdliL
                .Range("B106").FormulaR1C1 = CountfdnjTW
                .Range("C106").FormulaR1C1 = CountfdcwnTW
                .Range("D106").FormulaR1C1 = CountfdnycTW
                .Range("E106").FormulaR1C1 = CountfdliTW
                .Range("B107").FormulaR1C1 = CountfdnjRFI
                .Range("C107").FormulaR1C1 = CountfdcwnRFI
                .Range("D107").FormulaR1C1 = CountfdnycRFI
                .Range("E107").FormulaR1C1 = CountfdliRFI
                .Range("B108").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C108").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D108").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E108").FormulaR1C1 = Countfdlibrc + Countfdlirc
                Count1fd = WorksheetFunction.Sum(.Range("B104:E104"))
                Count2fd = WorksheetFunction.Sum(.Range("B105:E105"))
                Count3fd = WorksheetFunction.Sum(.Range("B106:E106"))
                Count4fd = WorksheetFunction.Sum(.Range("B107:E107"))
                Count5fd = WorksheetFunction.Sum(.Range("B108:E108"))
                .Range("F104").FormulaR1C1 = Count1fd
                .Range("F105").FormulaR1C1 = Count2fd
                .Range("F106").FormulaR1C1 = Count3fd
                .Range("F107").FormulaR1C1 = Count4fd
                .Range("F108").FormulaR1C1 = Count5fd
        End With
        'Table surrounding
            tablesize = "A102:G108"
            tableheader = "A102:G103"
            tabledata = "B104:G108"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            Call rfidata09(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, Count4TD, Sheetname, Filename)
End Sub

Sub Weeklyreportpart9(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn
    'For table J. Coax Asbuilt (Coax)
        'Finding Value for total
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & StartingDate, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("M:M").AutoFilter Field:=13, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("O:O").AutoFilter Field:=15, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("A:A").AutoFilter Field:=1, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("O:O").AutoFilter Field:=15, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        End With
        With WsOut
        'Creating Surrounding
                .Range("A113").FormulaR1C1 = "J. Coax Asbuilt (Coax)"
                .Range("A114").FormulaR1C1 = "Status Description"
                .Range("B114").FormulaR1C1 = "NJ"
                .Range("C114").FormulaR1C1 = "CWN"
                .Range("D114").FormulaR1C1 = "NYC"
                .Range("E114").FormulaR1C1 = "LI"
                .Range("F114").FormulaR1C1 = "Total"
                .Range("G114").FormulaR1C1 = "Remark"
                .Range("A116").FormulaR1C1 = "Total cells delivered (" & Currentyear & ")"
                .Range("A117").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A118").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A119").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A120").FormulaR1C1 = "Remaining cells"
                .Range("A113:G114").Font.FontStyle = "Bold"
                .Range("A113").Font.Underline = xlUnderlineStyleSingle
                With .Range("A113").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A114:A115").Merge
                .Range("B114:B115").Merge
                .Range("C114:C115").Merge
                .Range("D114:D115").Merge
                .Range("E114:E115").Merge
                .Range("F114:F115").Merge
                .Range("G114:G115").Merge

        'Calculating the total
                .Range("B116").FormulaR1C1 = Countfdnj
                .Range("C116").FormulaR1C1 = Countfdcwn
                .Range("D116").FormulaR1C1 = Countfdnyc
                .Range("E116").FormulaR1C1 = Countfdli
                .Range("B117").FormulaR1C1 = CountfdnjL
                .Range("C117").FormulaR1C1 = CountfdcwnL
                .Range("D117").FormulaR1C1 = CountfdnycL
                .Range("E117").FormulaR1C1 = CountfdliL
                .Range("B118").FormulaR1C1 = CountfdnjTW
                .Range("C118").FormulaR1C1 = CountfdcwnTW
                .Range("D118").FormulaR1C1 = CountfdnycTW
                .Range("E118").FormulaR1C1 = CountfdliTW
                .Range("B119").FormulaR1C1 = CountfdnjRFI
                .Range("C119").FormulaR1C1 = CountfdcwnRFI
                .Range("D119").FormulaR1C1 = CountfdnycRFI
                .Range("E119").FormulaR1C1 = CountfdliRFI
                .Range("B120").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C120").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D120").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E120").FormulaR1C1 = Countfdlibrc + Countfdlirc
                Count1fd = WorksheetFunction.Sum(.Range("B116:E116"))
                Count2fd = WorksheetFunction.Sum(.Range("B117:E117"))
                Count3fd = WorksheetFunction.Sum(.Range("B118:E118"))
                Count4fd = WorksheetFunction.Sum(.Range("B119:E119"))
                Count5fd = WorksheetFunction.Sum(.Range("B120:E120"))
                .Range("F116").FormulaR1C1 = Count1fd
                .Range("F117").FormulaR1C1 = Count2fd
                .Range("F118").FormulaR1C1 = Count3fd
                .Range("F119").FormulaR1C1 = Count4fd
                .Range("F120").FormulaR1C1 = Count5fd
        End With
        'Table surrounding
            tablesize = "A114:G120"
            tableheader = "A114:G115"
            tabledata = "B116:G120"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
            Call rfidata10(CountfdnjRFI, CountfdcwnRFI, CountfdnycRFI, CountfdliRFI, Count4fd, Sheetname, Filename)

End Sub

Sub Weeklyreportpart10(eDate, StartingDate, Filename, lastweekdate, Nextweekdate, Sheetname, OutputFileName, StartingDateold)
    'Creating files
        Dim WbIn As Workbook
        Dim WsIn As Worksheet
        Set WbIn = Workbooks(Filename)
        Set WsIn = WbIn.Sheets(Sheetname)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        With WsIn
        Currentyear = Year(Date)
    'For table K. New York DOT permit (PNI + Netwin)
        'Finding Value for total
            'Finding value for nj
                .Range("L:L").AutoFilter Field:=12, Operator:=xlFilterValues, Criteria1:=">=" & StartingDateold, Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("J:J").AutoFilter Field:=10, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnj = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("E:E").AutoFilter Field:=5, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnyc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdli = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwn = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Last week
            'Finding value for nj
                .Range("L:L").AutoFilter Field:=12, Operator:=xlFilterValues, Criteria1:=">=" & Format(lastweekdate, "d/mmmm/yyyy"), Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("J:J").AutoFilter Field:=10, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("E:E").AutoFilter Field:=5, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, WC, CT
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnL = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Total cells planned to be delivered this week
            'Finding value for nj
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & Format(eDate, "d/mmmm/yyyy"), Criteria2:="<=" & Nextweekdate
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("J:J").AutoFilter Field:=10, Criteria1:=Array("Completed", "Delivered", "Delivery", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                CountfdnjTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                
            'Finding value for nyc
                .Range("E:E").AutoFilter Field:=5, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnTW = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding RFI Value
            'Finding value for njs, NJN
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:="=" & "RFI"
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                CountfdnjRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("E:E").AutoFilter Field:=5, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                CountfdnycRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for lie
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                CountfdliRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for cwn, WC, CT
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                CountfdcwnRFI = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        'Finding Value for Remaining cells
            'Finding value for nj
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("NJN", "NJS", "NJ"), Operator:=xlFilterValues
                .Range("J:J").AutoFilter Field:=10, Criteria1:=Array("Completed", "Delivered", "Delivery", "HOLD", "DONE", "FIXING", "In Progress", "IP", "=", "Complete"), Operator:=xlFilterValues
                Countfdnjrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("K:K").AutoFilter Field:=11, Criteria1:="="
                Countfdnjbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for nyc
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("E:E").AutoFilter Field:=5, Operator:=xlFilterValues, Criteria1:="=" & "NYC"
                Countfdnycrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("K:K").AutoFilter Field:=11, Criteria1:="="
                Countfdnycbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for li
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("LIE", "LIW", "LI"), Operator:=xlFilterValues
                Countfdlirc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("K:K").AutoFilter Field:=11, Criteria1:="="
                Countfdlibrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1

            'Finding value for Cwn, CT, WC
                .Range("K:K").AutoFilter Field:=11, Operator:=xlFilterValues, Criteria1:=">=" & Nextweekdate
                .Range("E:E").AutoFilter Field:=5, Criteria1:=Array("CT", "CWN", "WC"), Operator:=xlFilterValues
                Countfdcwnrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .Range("K:K").AutoFilter Field:=11, Criteria1:="="
                Countfdcwnbrc = .AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
                .ShowAllData

        End With
        With WsOut
        'Creating Surrounding

                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A123").EntireRow.Insert
                .Range("A125").FormulaR1C1 = "K. New York DOT permit (PNI + Netwin)"
                .Range("A126").FormulaR1C1 = "Status Description"
                .Range("B126").FormulaR1C1 = "NJ"
                .Range("C126").FormulaR1C1 = "CWN"
                .Range("D126").FormulaR1C1 = "NYC"
                .Range("E126").FormulaR1C1 = "LI"
                .Range("F126").FormulaR1C1 = "Total"
                .Range("G126").FormulaR1C1 = "Remark"
                .Range("A128").FormulaR1C1 = "Total cells delivered"
                .Range("A129").FormulaR1C1 = "Total cells delivered (Last week)"
                .Range("A130").FormulaR1C1 = "Total cells planned to be delivered this week"
                .Range("A131").FormulaR1C1 = "Number of cells Pending RFI"
                .Range("A132").FormulaR1C1 = "Remaining cells"
                .Range("A125:G126").Font.FontStyle = "Bold"
                .Range("A125").Font.Underline = xlUnderlineStyleSingle
                With .Range("A125").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A126:A127").Merge
                .Range("B126:B127").Merge
                .Range("C126:C127").Merge
                .Range("D126:D127").Merge
                .Range("E126:E127").Merge
                .Range("F126:F127").Merge
                .Range("G126:G127").Merge

        'Calculating the total
                .Range("B128").FormulaR1C1 = Countfdnj
                .Range("C128").FormulaR1C1 = Countfdcwn
                .Range("D128").FormulaR1C1 = Countfdnyc
                .Range("E128").FormulaR1C1 = Countfdli
                .Range("B129").FormulaR1C1 = CountfdnjL
                .Range("C129").FormulaR1C1 = CountfdcwnL
                .Range("D129").FormulaR1C1 = CountfdnycL
                .Range("E129").FormulaR1C1 = CountfdliL
                .Range("B130").FormulaR1C1 = CountfdnjTW
                .Range("C130").FormulaR1C1 = CountfdcwnTW
                .Range("D130").FormulaR1C1 = CountfdnycTW
                .Range("E130").FormulaR1C1 = CountfdliTW
                .Range("B131").FormulaR1C1 = CountfdnjRFI
                .Range("C131").FormulaR1C1 = CountfdcwnRFI
                .Range("D131").FormulaR1C1 = CountfdnycRFI
                .Range("E131").FormulaR1C1 = CountfdliRFI
                .Range("B132").FormulaR1C1 = Countfdnjrc + Countfdnjbrc
                .Range("C132").FormulaR1C1 = Countfdcwnbrc + Countfdcwnrc
                .Range("D132").FormulaR1C1 = Countfdnycbrc + Countfdnycrc
                .Range("E132").FormulaR1C1 = Countfdlibrc + Countfdlirc
                Count1fd = WorksheetFunction.Sum(.Range("B128:E128"))
                Count2fd = WorksheetFunction.Sum(.Range("B129:E129"))
                Count3fd = WorksheetFunction.Sum(.Range("B130:E130"))
                Count4fd = WorksheetFunction.Sum(.Range("B131:E131"))
                Count5fd = WorksheetFunction.Sum(.Range("B132:E132"))
                .Range("F128").FormulaR1C1 = Count1fd
                .Range("F129").FormulaR1C1 = Count2fd
                .Range("F130").FormulaR1C1 = Count3fd
                .Range("F131").FormulaR1C1 = Count4fd
                .Range("F132").FormulaR1C1 = Count5fd
        End With
        'Table surrounding
            tablesize = "A126:G132"
            tableheader = "A126:G127"
            tabledata = "B128:G132"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)
End Sub

Sub Weeklyreportpart11(Filename, Sheetname, OutputFileName)
    'Creating files
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")
        Currentyear = Year(Date)
        Call CheckDataSheet(Filename, "VBA")
    'For Seco
        'Finding Total Miles & Total cell
            'Finding Fairgrounds
                Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                            "Fairgrounds"), _
                                                            Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Criteria1:="<>"
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).ShowAllData
            'Finding value for Frelinghuysen
                Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                            "Frelinghuysen"), _
                                                            Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Criteria1:="<>"
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).ShowAllData

            'Finding value for Guthry Corner
                Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                            "Guthry Corner"), _
                                                            Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Criteria1:="<>"
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).ShowAllData

            'Finding value for Hardyston
                Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                            "Hardyston"), _
                                                            Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Criteria1:="<>"
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("L1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).ShowAllData

            'Finding value for Sparta
                Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                            "Sparta"), _
                                                            Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("M1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("N1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Criteria1:="<>"
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("O1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).ShowAllData

            'Finding value for Frelinghuysen
                Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                            "Tall Timbers"), _
                                                            Operator:=xlFilterValues
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("P1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("Q1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Criteria1:="<>"
                Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                Workbooks(Filename).Worksheets("VBA").Range("R1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                Workbooks(Filename).Sheets(Sheetname).ShowAllData


        'Calculation Value
            'Calculate Fairgrounds
                Workbooks(Filename).Sheets("VBA").Range("A1:AV1").Delete
                CountFGT = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("A:A"))
                CountFGU = WorksheetFunction.Sum(Workbooks(Filename).Sheets("VBA").Range("B:B"))
                CountFGD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("C:C"))


            'Calculate Frelinghuysen
                CountFRT = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("P:P"))
                CountFRU = WorksheetFunction.Sum(Workbooks(Filename).Sheets("VBA").Range("Q:Q"))
                CountFRD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("R:R"))

            'Calculate Guthry Corner
                CountGCT = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("G:G"))
                CountGCU = WorksheetFunction.Sum(Workbooks(Filename).Sheets("VBA").Range("H:H"))
                CountGCD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("I:I"))

            'Calculate Hardyston
                CountHST = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("J:J"))
                CountHSU = WorksheetFunction.Sum(Workbooks(Filename).Sheets("VBA").Range("K:K"))
                CountHSD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("L:L"))

            'Calculate Sparta
                CountSPT = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("M:M"))
                CountSPU = WorksheetFunction.Sum(Workbooks(Filename).Sheets("VBA").Range("N:N"))
                CountSPD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("O:O"))

            'Calculate Frelinghuysen
                CountFHT = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("D:D"))
                CountFHU = WorksheetFunction.Sum(Workbooks(Filename).Sheets("VBA").Range("E:E"))
                CountFHD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("VBA").Range("F:F"))

        With WsOut
        'Creating Surrounding
                .Range("A137").FormulaR1C1 = "M. SECO coax migration"
                .Range("A138").FormulaR1C1 = "Status Description"
                .Range("B138").FormulaR1C1 = "Fair Grounds Hub"
                .Range("C138").FormulaR1C1 = "Tall Timbers Hub"
                .Range("D138").FormulaR1C1 = "Guthrie Corners Hub"
                .Range("E138").FormulaR1C1 = "Hardyston Hub"
                .Range("F138").FormulaR1C1 = "Sparta Hub"
                .Range("G138").FormulaR1C1 = "Frelinghuysen Hub"
                .Range("A140").FormulaR1C1 = "Approximate miles"
                .Range("A141").FormulaR1C1 = "Miles completed"
                .Range("A142").FormulaR1C1 = "Completion percentage"
                .Range("A143").FormulaR1C1 = "Nodes plan to deliver this week"
                .Range("A144").FormulaR1C1 = "Completed Nodes/Total nodes"
                .Range("A145").FormulaR1C1 = "Pending RFI"
                .Range("A137:G138").Font.FontStyle = "Bold"
                .Range("A137").Font.Underline = xlUnderlineStyleSingle
                With .Range("A137").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A138:A139").Merge
                .Range("B138:B139").Merge
                .Range("C138:C139").Merge
                .Range("D138:D139").Merge
                .Range("E138:E139").Merge
                .Range("F138:F139").Merge
                .Range("G138:G139").Merge

        'Calculating the total
                .Range("B140").FormulaR1C1 = 0
                .Range("C140").FormulaR1C1 = 0
                .Range("D140").FormulaR1C1 = 0
                .Range("E140").FormulaR1C1 = 0
                .Range("F140").FormulaR1C1 = 0
                .Range("G140").FormulaR1C1 = 0
                .Range("B141").FormulaR1C1 = CountFGU
                .Range("C141").FormulaR1C1 = CountFRU
                .Range("D141").FormulaR1C1 = CountGCU
                .Range("E141").FormulaR1C1 = CountHSU
                .Range("F141").FormulaR1C1 = CountSPU
                .Range("G141").FormulaR1C1 = CountFHU
                .Range("B142").FormulaR1C1 = "=B140/B141"
                .Range("C142").FormulaR1C1 = "=C140/C141"
                .Range("D142").FormulaR1C1 = "=D140/D141"
                .Range("E142").FormulaR1C1 = "=E140/E141"
                .Range("F142").FormulaR1C1 = "=F140/F141"
                .Range("G142").FormulaR1C1 = "=G140/G141"
                .Range("B143").FormulaR1C1 = 0
                .Range("C143").FormulaR1C1 = 0
                .Range("D143").FormulaR1C1 = 0
                .Range("E143").FormulaR1C1 = 0
                .Range("F143").FormulaR1C1 = 0
                .Range("G143").FormulaR1C1 = 0
                .Range("B144").FormulaR1C1 = CountFGD & "/" & CountFGT
                .Range("C144").FormulaR1C1 = CountFRD & "/" & CountFRT
                .Range("D144").FormulaR1C1 = CountGCD & "/" & CountGCT
                .Range("E144").FormulaR1C1 = CountHSD & "/" & CountHST
                .Range("F144").FormulaR1C1 = CountSPD & "/" & CountSPT
                .Range("G144").FormulaR1C1 = CountFHD & "/" & CountFHT
                .Range("B145").FormulaR1C1 = 0
                .Range("C145").FormulaR1C1 = 0
                .Range("D145").FormulaR1C1 = 0
                .Range("E145").FormulaR1C1 = 0
                .Range("F145").FormulaR1C1 = 0
                .Range("G145").FormulaR1C1 = 0
                .Range("B142:G142").Font.Bold = True
        
                With .Range("B142:G142").Font
                    .Color = -11489280
                    .TintAndShade = 0
                End With
                .Range("B142:G142").NumberFormat = "0%"
            End With
        'Table surrounding
            tablesize = "A138:G145"
            tableheader = "A138:G139"
            tabledata = "B140:G145"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)

End Sub

Sub Weeklyreportpart12(OutputFileName)
    'Creating files
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")

    'For L. Texas Landbase & FTTH
        'Creating Surrounding
            With WsOut
                .Range("A124").FormulaR1C1 = "L. Texas Landbase & FTTH"
                .Range("A125").FormulaR1C1 = "Status Description"
                .Range("B125").FormulaR1C1 = "Hereford"
                .Range("D125").FormulaR1C1 = "Seminole"
                .Range("F125").FormulaR1C1 = "Brownfield"
                .Range("B126,D126,F126").FormulaR1C1 = "Stage 1: Civil + Addressing"
                .Range("C126,E126,G126").FormulaR1C1 = "Stage 2: FTTH design"
                .Range("A128").FormulaR1C1 = "Total map/cell"
                .Range("A129").FormulaR1C1 = "Total map/cell completed"
                .Range("A130").FormulaR1C1 = "Completetion Percentage"
                .Range("A131").FormulaR1C1 = "Pending RFI"
                .Range("A132").FormulaR1C1 = "Remark"
                .Range("A124:G126").Font.FontStyle = "Bold"
                .Range("A124").Font.Underline = xlUnderlineStyleSingle
                With .Range("A124").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 16
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With
                .Range("A126:A127").Merge
                .Range("B126:B127").Merge
                .Range("C126:C127").Merge
                .Range("D126:D127").Merge
                .Range("E126:E127").Merge
                .Range("F126:F127").Merge
                .Range("G126:G127").Merge

                .Range("B125:C125").Merge
                .Range("D125:E125").Merge
                .Range("F125:G125").Merge

        'Calculating the total
                .Range("B128").FormulaR1C1 = "94.55"
                .Range("C128").FormulaR1C1 = "62"
                .Range("D128").FormulaR1C1 = "105.76"
                .Range("E128").FormulaR1C1 = "10"
                .Range("F128").FormulaR1C1 = "79.19"
                .Range("G128").FormulaR1C1 = "0"
                .Range("B129").FormulaR1C1 = "94.55"
                .Range("C129").FormulaR1C1 = "62"
                .Range("D129").FormulaR1C1 = "105.76"
                .Range("E129").FormulaR1C1 = "10"
                .Range("F129").FormulaR1C1 = "79.19"
                .Range("G129").FormulaR1C1 = "0"
                .Range("B130").FormulaR1C1 = "100.00%"
                .Range("C130").FormulaR1C1 = "100.00%"
                .Range("D130").FormulaR1C1 = "100.00%"
                .Range("E130").FormulaR1C1 = "100.00%"
                .Range("F130").FormulaR1C1 = "100.00%"
                .Range("G130").FormulaR1C1 = "100.00%"
                .Range("B131").FormulaR1C1 = "0"
                .Range("C131").FormulaR1C1 = "0"
                .Range("D131").FormulaR1C1 = "0"
                .Range("E131").FormulaR1C1 = "0"
                .Range("F131").FormulaR1C1 = "0"
                .Range("G131").FormulaR1C1 = "0"

                .Range("B130:G130").NumberFormat = "0%"
            End With

        'Table surrounding
            tablesize = "A125:G132"
            tableheader = "A125:G127"
            tabledata = "B128:G132"
            Call TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)

End Sub

Sub Weeklyreportpart13(OutputFileName)
        Dim WbOut As Workbook
        Dim WsOut As Worksheet
        Set WbOut = Workbooks(OutputFileName)
        Set WsOut = WbOut.Sheets("Sheet1")

    'For Data
        'Creating Surrounding
            With WsOut
                .Range("B150").FormulaR1C1 = "This week's plan"
                .Range("F151").FormulaR1C1 = "Total Cells"
                .Range("B152:D152").FormulaR1C1 = "Areas"
                .Range("C152:E152").FormulaR1C1 = "Qty"
                .Range("B153:B156,D153:D156").FormulaR1C1 = "-"
                .Range("C153:C156,E153:E156,F153:F156").FormulaR1C1 = "0"
                .Range("A153").FormulaR1C1 = "NJ"
                .Range("A154").FormulaR1C1 = "CWN"
                .Range("A155").FormulaR1C1 = "NYC"
                .Range("A156").FormulaR1C1 = "LI"
                .Range("A150:F152,A153:A156").Font.FontStyle = "Bold"
                .Range("A150").Font.Underline = xlUnderlineStyleSingle
                With .Range("A150").Font
                                                                                .Name = "Calibri"
                                                                                .Size = 11
                                                                                .Strikethrough = False
                                                                                .Superscript = False
                                                                                .Subscript = False
                                                                                .OutlineFont = False
                                                                                .Shadow = False
                                                                                .Underline = xlUnderlineStyleSingle
                                                                                .ThemeColor = xlThemeColorLight1
                                                                                .TintAndShade = 0
                                                                                .ThemeFont = xlThemeFontMinor
                End With

                .Range("B151:C151").Merge
                .Range("D151:E151").Merge

                .Range("B151:F156").Borders(xlDiagonalDown).LineStyle = xlNone
                .Range("B151:F156").Borders(xlDiagonalUp).LineStyle = xlNone
                With .Range("B151:F156").Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156").Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156").Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156").Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156").Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                .Range("B151:F156").Borders(xlDiagonalDown).LineStyle = xlNone
                .Range("B151:F156").Borders(xlDiagonalUp).LineStyle = xlNone
                With .Range("B151:F156").Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With .Range("B151:F156").Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With .Range("B151:F156").Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With .Range("B151:F156").Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With .Range("B151:F156").Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156").Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Range("B151:F156")
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                End With

                With .Range("A153:A156")
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With

                With .Range("B151:F151").Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent3
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End With
End Sub

Sub TableArrangment1(tablesize, tableheader, tabledata, OutputFileName)

    Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlDiagonalDown).LineStyle = xlNone
    Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlDiagonalUp).LineStyle = xlNone
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tablesize).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tableheader)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tableheader)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tableheader).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10147522
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Workbooks(OutputFileName).Worksheets("Sheet1").Range(tabledata)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub rfidata01(nj, cwn, nyc, li, nc, total, Sheetname, Filename)
    'Creating file

        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - FTTH Netwin Design " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change C:\Users\farhanah\Documents\ALTICE WEEKLY\WEDNESDAY REPORT\Weekly Report\Output\
            Application.DisplayAlerts = True
            
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNJS
                    RFI = "I3:I" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:K" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountCWN
                    RFI = "I3:I" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:K" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNYC
                    RFI = "I3:I" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:K" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(Filename).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountLI
                    RFI = "I3:I" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NC"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(Filename).Worksheets("NC").Range("A1").FormulaR1C1 = "NC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNC
                    RFI = "I3:I" & CountNC
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountNC
                    Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
    
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
            
        End If
End Sub

Sub rfidata02(nj, cwn, nyc, li, tx, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - FTTH Netwin Redesign " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNJS
                    RFI = "I3:I" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:K" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountCWN
                    RFI = "I3:I" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:K" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNYC
                    RFI = "I3:I" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:K" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountLI
                    RFI = "I3:I" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If tx <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "TX"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "TX"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("TX").Range("A1").FormulaR1C1 = "TX"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("TX").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("TX").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("TX").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountTX = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("TX").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountTX
                    RFI = "I3:I" & CountTX
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountTX
                    Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("TX").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("TX").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("TX").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("TX").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("TX").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata03(nj, cwn, nyc, li, nc, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - FTTH Feeder Design " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("I:I").AutoFilter Field:=9, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNJS
                    RFI = "I3:I" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:K" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("I:I").AutoFilter Field:=9, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountCWN
                    RFI = "I3:I" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:K" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("I:I").AutoFilter Field:=9, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNYC
                    RFI = "I3:I" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:K" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("I:I").AutoFilter Field:=9, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountLI
                    RFI = "I3:I" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NC"
                    Workbooks(Filename).Sheets(Sheetname).Range("I:I").AutoFilter Field:=9, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NC"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("A1").FormulaR1C1 = "NC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("A2:J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNC
                    RFI = "I3:I" & CountNC
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountNC
                    Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata04(nj, cwn, nyc, li, nc, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - FTTH Netwin Asbuilt " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNJS
                    RFI = "I3:I" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:K" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountCWN
                    RFI = "I3:I" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:K" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNYC
                    RFI = "I3:I" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:K" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountLI
                    RFI = "I3:I" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:K" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("A1").FormulaR1C1 = "NC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:K" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NC").Range("K:K").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("A2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range("I2:K2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:K" & CountNC
                    RFI = "I3:I" & CountNC
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:K" & CountNC
                    Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("I").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata05(nj, cwn, nyc, li, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - FTTH Cell Rename " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:J" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:I2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:J" & CountNJS
                    RFI = "J3:J" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:J" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("J").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:J" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:I2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:J" & CountCWN
                    RFI = "J3:J" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:J" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("J").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:J" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:I2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:J" & CountNYC
                    RFI = "J3:J" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:J" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("J").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("E:E").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("R:R").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:J" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:I2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("J2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 65535
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:J" & CountLI
                    RFI = "J3:J" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:J" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("J").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata06(nj, cwn, nyc, li, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - EOL Test Sheet " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("B:B").AutoFilter Field:=2, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("T:T").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("U:U").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountNJS
                    RFI = "F3:F" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:H" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("B:B").AutoFilter Field:=2, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("T:T").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("U:U").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountCWN
                    RFI = "F3:F" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:H" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("B:B").AutoFilter Field:=2, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("T:T").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("U:U").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountNYC
                    RFI = "F3:F" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:H" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("N:N").AutoFilter Field:=14, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("B:B").AutoFilter Field:=2, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("D:D").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("N:N").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("S:S").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("T:T").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("U:U").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 8421631
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountLI
                    RFI = "F3:F" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:H" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata07(nj, cwn, nyc, li, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - HFC Fiber Node Asbuilt " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("AI:AI").AutoFilter Field:=35, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("X:X").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("L1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AD:AD").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("M1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:M" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"

                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:L2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("M2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:L" & CountNJS
                    RFI = "M3:M" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:M" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("M").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("AI:AI").AutoFilter Field:=35, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("X:X").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("L1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AD:AD").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("M1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:M" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:L2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("M2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:L" & CountCWN
                    RFI = "M3:M" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:M" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("M").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("AI:AI").AutoFilter Field:=35, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("X:X").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("L1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AD:AD").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("M1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:M" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:L2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("M2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:L" & CountNYC
                    RFI = "M3:M" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:M" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("M").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("AI:AI").AutoFilter Field:=35, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("F:F").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("I:I").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("I1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("J1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("K1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("X:X").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("L1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AD:AD").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("M1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:M" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:L2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("M2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:L" & CountLI
                    RFI = "M3:M" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:M" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("M").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata08(nj, cwn, nyc, li, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - Node Split Design " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AF:AF").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AG:AG").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets("VBA").Range("A1:AV1").Delete
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:E" & CountNJS
                    RFI = "F3:F" & CountNJS
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:H" & CountNJS
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AF:AF").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AG:AG").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Workbooks(Filename).Sheets("VBA").Range("A1:AV1").Delete
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountCWN
                    RFI = "F3:F" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:H" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AF:AF").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AG:AG").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(Filename).Sheets("VBA").Range("A1:AV1").Delete
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountNYC
                    RFI = "F3:F" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:H" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AF:AF").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("AG:AG").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("H1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(Filename).Sheets("VBA").Range("A1:AV1").Delete
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:H" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("H:H").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountLI
                    RFI = "F3:F" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:H" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("F").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata09(nj, cwn, nyc, li, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - Coax Design " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
        ' Creating conditions
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("O:O").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:F2").Interior 'Update
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("G2").Interior 'Update
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:F" & CountNJS 'Update
                    RFI = "G3:G" & CountNJS        'Update
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:G" & CountNJS  'Update
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("O:O").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:F2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("G2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:F" & CountCWN
                    RFI = "G3:G" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:G" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("O:O").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:F2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("G2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:F" & CountNYC
                    RFI = "G3:G" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:G" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Worksheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("G:G").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("J:J").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("K:K").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Worksheets(Sheetname).Columns("O:O").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountLI
                    RFI = "F3:F" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:H" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub rfidata10(nj, cwn, nyc, li, total, Sheetname, Filename)
    'Creating file
        Call CheckDataSheet(Filename, "VBA")
        If total <> 0 Then
            TodayDate = Format(Date, "mmddyyyy")
            OutputfileRFI = "RFI - Coax Asbuilt " & TodayDate & ".xlsx"
            Call Check_if_workbook_is_open(OutputfileRFI)
            Application.DisplayAlerts = False
            Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputfileRFI 'Please Change "E:\OneDrive\Desktop\ExcelTestFiles\"
            Application.DisplayAlerts = True
            
            If nj <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "NJN", "NJS", "NJ"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").FormulaR1C1 = "NJ"
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding

                    CountNJS = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NJ").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("A2:F2").Interior 'Update
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range("G2").Interior 'Update
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:F" & CountNJS 'Update
                    RFI = "G3:G" & CountNJS        'Update
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With

                'TABLE
                    Table = "A2:G" & CountNJS  'Update
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NJ").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NJ").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If cwn <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "CWN"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "CT", "CWN", "WC"), Operator:=xlFilterValues
                    Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").FormulaR1C1 = "CWN"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountCWN = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("CWN").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("A2:F2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range("G2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:F" & CountCWN
                    RFI = "G3:G" & CountCWN
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
      
                'TABLE
                    Table = "A2:G" & CountCWN
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("CWN").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("CWN").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If nyc <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "NYC"
                    Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").FormulaR1C1 = "NYC"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountNYC = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("NYC").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("A2:F2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range("G2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:F" & CountNYC
                    RFI = "G3:G" & CountNYC
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
                    Table = "A2:G" & CountNYC
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("NYC").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("NYC").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            If li <> 0 Then
                'Filtering
                    Workbooks(OutputfileRFI).Sheets.Add.Name = "LI"
                    Workbooks(Filename).Sheets(Sheetname).Range("M:M").AutoFilter Field:=13, Operator:=xlFilterValues, _
                                                                Criteria1:="=" & "RFI"
                    Workbooks(Filename).Sheets(Sheetname).Range("A:A").AutoFilter Field:=1, Criteria1:=Array( _
                                                                "LIE", "LIW", "LI"), _
                                                                Operator:=xlFilterValues
                    Workbooks(Filename).Sheets(Sheetname).Columns("A:A").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("B:B").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("C:C").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("H:H").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("D1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("L:L").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("M:M").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("F1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).Columns("Q:Q").Copy 'Actual Delivery
                    Workbooks(Filename).Worksheets("VBA").Range("G1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                    Workbooks(Filename).Sheets(Sheetname).ShowAllData
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").FormulaR1C1 = "LI"
                    Count1 = WorksheetFunction.CountA(Workbooks(Filename).Worksheets("VBA").Range("A:A"))
                    Copy = "A1:G" & Count1
                    Workbooks(Filename).Worksheets("VBA").Range(Copy).Copy
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("A2").PasteSpecial Paste:=xlPasteValues
                    Workbooks(OutputfileRFI).Worksheets("LI").Range("B:B").NumberFormat = "d-mmm"
                'Creating Surrounding
                    CountLI = WorksheetFunction.CountA(Workbooks(OutputfileRFI).Sheets("LI").Range("A:A"))
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A1").Font
                                                                            .Name = "Calibri"
                                                                            .Size = 20
                                                                            .Strikethrough = False
                                                                            .Superscript = False
                                                                            .Subscript = False
                                                                            .OutlineFont = False
                                                                            .Shadow = False
                                                                            .Underline = xlUnderlineStyleSingle
                                                                            .ThemeColor = xlThemeColorLight1
                                                                            .TintAndShade = 0
                                                                            .ThemeFont = xlThemeFontMinor
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("A2:E2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 192
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range("F2:H2").Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 49407
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    OtherValue = "A3:H" & CountLI
                    RFI = "F3:F" & CountLI
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(OtherValue)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(RFI)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
                'TABLE
   
                    Table = "A2:H" & CountLI
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
                    Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Workbooks(OutputfileRFI).Worksheets("LI").Range(Table).Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireColumn.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("G").ColumnWidth = 40
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").EntireRow.AutoFit
                    Workbooks(OutputfileRFI).Worksheets("LI").Columns("A:W").HorizontalAlignment = xlCenter
            End If
            Application.DisplayAlerts = False
            Workbooks(OutputfileRFI).Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            Workbooks(OutputfileRFI).Save
        End If
End Sub

Sub CheckDataSheet(Filename, Sheetname)
    For Each Sheet In Workbooks(Filename).Worksheets
        If Sheet.Name = Sheetname Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
    Workbooks(Filename).Sheets.Add.Name = Sheetname
End Sub

Sub Check_if_workbook_is_open(OutputFileName)
    Dim wb As Workbook 'to test if workbook is open. No change here
        For Each wb In Workbooks
            If wb.Name = OutputFileName Then
                Workbooks(OutputFileName).Save
                Workbooks(OutputFileName).Close
            End If
        Next
End Sub



