'Version: 1.00
'CCI-Comcast
'For Echobroadband
'User: Rabiatul 
'By Farhat Abbas
Sub Arragmemt()
    'Part 1
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("B4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("E5:E43"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder _
            :="Ring,Rolt/HE Feeder,Distribution", DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

    'Part 2
            ActiveWorkbook.Worksheets("Sheet1").Columns("B:B").Copy 'Actual Delivery
            ActiveWorkbook.Worksheets("Sheet1").Range("L1").PasteSpecial Paste:=xlPasteValues
            ActiveWorkbook.Worksheets("Sheet1").Columns("E:E").Copy 'Actual Delivery
            ActiveWorkbook.Worksheets("Sheet1").Range("M1").PasteSpecial Paste:=xlPasteValues
            ActiveWorkbook.Worksheets("Sheet1").Range("L1:M4").Delete
            ActiveWorkbook.Worksheets("Sheet1").Columns("L:L").RemoveDuplicates Columns:=1, Header:=xlNo
            ActiveWorkbook.Worksheets("Sheet1").Columns("M:M").RemoveDuplicates Columns:=1, Header:=xlNo
            CountCC = WorksheetFunction.CountA(ActiveWorkbook.Worksheets("Sheet1").Range("L:L"))
                For X = 1 To CountCC
                    InL = "L" & X
                    OutN = "N" & X
                    ActiveWorkbook.Worksheets("Sheet1").Range(OutN) = ActiveWorkbook.Worksheets("Sheet1").Application.WorksheetFunction.CountIf(Worksheets("Sheet1").Range("B:B"), Worksheets("Sheet1").Range(InL))
                Next

            ActiveWorkbook.Worksheets("Sheet1").Range("O1") = ActiveWorkbook.Worksheets("Sheet1").Application.WorksheetFunction.CountIf(Worksheets("Sheet1").Range("E:E"), Worksheets("Sheet1").Range("M1"))
            ActiveWorkbook.Worksheets("Sheet1").Range("O2") = ActiveWorkbook.Worksheets("Sheet1").Application.WorksheetFunction.CountIf(Worksheets("Sheet1").Range("E:E"), Worksheets("Sheet1").Range("M2"))
            ActiveWorkbook.Worksheets("Sheet1").Range("O3") = ActiveWorkbook.Worksheets("Sheet1").Application.WorksheetFunction.CountIf(Worksheets("Sheet1").Range("E:E"), Worksheets("Sheet1").Range("M3"))
            
        '
            For X = 1 To CountCC
                FirstValue = "L" & X
                SecondValue = "L" & X + 1
                LValue = "N" & X
                myCellSetValue = ActiveWorkbook.Worksheets("Sheet1").Range(FirstValue)
                myCellSetValue1 = ActiveWorkbook.Worksheets("Sheet1").Range(SecondValue)
                If X = 1 Then
                FirstCellValue = ActiveWorkbook.Worksheets("Sheet1").Range(LValue) + 4
                End If
                If X <> 1 Then
                FirstCellValue = ActiveWorkbook.Worksheets("Sheet1").Range(LValue)
                End If
                FCV = FCV + FirstCellValue
                MyLen = Len(myCellSetValue)
                MyLen1 = Len(myCellSetValue)
                Z = MyLen - MyLen1
                Range1 = "A" & FCV & ":G" & FCV

                If Z = 0 Then 
                    For F = 1 To MyLen1

                        If F = 1 then
                            F1 = F 
                        End if 

                        IF F <> 1 Then 
                            F1 = F + 1
                        End if 
                        Word = Mid(myCellSetValue, F, 1)
                        Word1 = Mid(myCellSetValue1, F, 1)
                        intResult = StrComp(Word, Word1, vbTextCompare)
                        
                            If intResult = -1 Then
                                Exit For
                            End If
                        
                            If intResult = 1 Then
                                Exit For
                            End If

                    Next
                                        
                    G = MyLen1 - F

                    If G <> 0 Then

                        Type1 = 3
                        Call Boarder(Range1, Type1)
                    End If

                    If G = 0 Then
                        Type1 = 2
                        Call Boarder(Range1, Type1)
                    End If
                
                End If
                If Z <> 0 Then
                    Type1 = 3
                    Call Boarder(Range1, Type1)
                End If
            Next

            For X = 1 To 3 
                FirstValue1 = "M" & X
                SecondValue1 = "M" & X + 1
                LValue1 = "O" & X
                myCellSetValue = ActiveWorkbook.Worksheets("Sheet1").Range(FirstValue1)
                myCellSetValue1 = ActiveWorkbook.Worksheets("Sheet1").Range(SecondValue1)
                If X = 1 Then
                FirstCellValue1 = ActiveWorkbook.Worksheets("Sheet1").Range(LValue1) + 4
                End If
                If X <> 1 Then
                FirstCellValue1 = ActiveWorkbook.Worksheets("Sheet1").Range(LValue1)
                End If
                FCV1 = FCV1 + FirstCellValue1
                Range1 = "A" & FCV1 & ":G" & FCV1
                Type1 = 1
                Call Boarder(Range1, Type1)
            Next
            DeleteDataRange = "L1:O" & CountCC
            ActiveWorkbook.Worksheets("Sheet1").Range(DeleteDataRange).Delete
            ActiveWorkbook.Save
   
End Sub
Sub Boarder(Range1, Type1)
    'Thick
        If Type1 = 1 Then
            With ActiveWorkbook.Worksheets("Sheet1").Range(Range1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End If
    'Single Boarder
        If Type1 = 2 Then
            With ActiveWorkbook.Worksheets("Sheet1").Range(Range1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End If
    'Double Boarder
        If Type1 = 3 Then
            With ActiveWorkbook.Worksheets("Sheet1").Range(Range1).Borders(xlEdgeBottom)
                .LineStyle = xlDouble
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThick
            End With
        End If
End Sub



