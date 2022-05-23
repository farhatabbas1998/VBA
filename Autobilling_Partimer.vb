'Version: 1.00
'For Echobroadband
'By Farhat Abbas (Verified & Tested)

Sub Partimer_Billing()

  Application.DisplayAlerts = False
  Application.ScreenUpdating = False
 'Getting Data
  Dataws = "Billing_Partimer"
  Sparews = "Sparews"
  Masterwb = ThisWorkbook.Name
  Call CheckDataSheet(Masterwb, Sparews)
  
 'Input Workbook is represented as wb
  Dim mwb As Workbook
  Set mwb = Workbooks(Masterwb)
  Dim dws As Worksheet
  Set dws = mwb.Sheets(Dataws)
  Dim sws As Worksheet
  Set sws = mwb.Sheets(Sparews)
  
  
 'Getting Tracking File
  strUrl = dws.Cells(2, 3).Value
  Trackingws = dws.Cells(4, 12).Value
  Dim twb As Workbook
  Set twb = Application.Workbooks.Open(strUrl)
  Dim tws As Worksheet
  Set tws = twb.Sheets(Trackingws)
  Trackingwb = twb.Name
  strPath = ThisWorkbook.Path & "\Input files\" & Trackingwb
  twb.SaveAs Filename:=strPath
  

  Call StopAllFilters(Trackingwb)
  Call StartFilter(Trackingwb, Trackingws)
  tws.AutoFilter.Sort.SortFields.Clear
  tws.AutoFilter.Sort.SortFields.Add2 Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With tws.AutoFilter.Sort
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
  End With
  
 'Storing into an arrays
  TotalProjects = WorksheetFunction.CountA(dws.Range("B:B")) - 2
  Namelists = WorksheetFunction.CountA(dws.Range("G:G")) - 1
  TotalEntires = WorksheetFunction.CountA(tws.Range("A:A")) - 1
  Mini_Projects = WorksheetFunction.CountA(dws.Range("O:O")) - 1
  
  
  Dim Store_Project(1000) As Variant
  Dim Task_Project(1000) As Variant
  Dim Status_Project(1000) As Variant
  Dim Price_Project(1000) As Currency
  Dim Name_List(1000) As Variant
  Dim BillingName_List(1000) As Variant
  Dim Numbering(1000) As Variant
  Dim Activewb(1000) As Variant
  Dim Dilverydate(1000) As Variant
  Dim ByodWeek1(1000) As Variant
  Dim ByodWeek2(1000) As Variant
  Dim ByodWeek3(1000) As Variant
  Dim ByodWeek4(1000) As Variant
  Dim ByodWeek5(1000) As Variant
  Dim ByodPrice1(1000) As Currency
  Dim ByodPrice2(1000) As Currency
  Dim ByodPrice3(1000) As Currency
  Dim ByodPrice4(1000) As Currency
  Dim ByodPrice5(1000) As Currency
  Dim ByodBollen(1000) As Variant
  Dim NBIProjects(1000, 1000) As Variant
  Dim Total_Price(1000) As Currency
  Dim spac(1000) As Variant
  Dim ByodTotalPrice(1000) As Currency
  Dim Mini_Project_Price(1000) As Currency
  Dim Mini_Project_Details(1000) As Variant
  Dim Total_MiniTask(1000) As Variant
  Dim Total_MiniCount(1000) As Variant
  Dim TotalUniqueID(1000) As Variant
  
  
  For TotalProject = 1 To TotalProjects
    Store_Project(TotalProject) = dws.Cells(TotalProject + 4, 2).Value
    Task_Project(TotalProject) = dws.Cells(TotalProject + 4, 3).Value
    Status_Project(TotalProject) = dws.Cells(TotalProject + 4, 4).Value
    Price_Project(TotalProject) = dws.Cells(TotalProject + 4, 5).Value
  Next TotalProject
  For Namelist = 1 To Namelists
    Name_List(Namelist) = dws.Cells(Namelist + 4, 7).Value
    BillingName_List(Namelist) = dws.Cells(Namelist + 4, 8).Value & dws.Cells(5, 12).Value
    ByodBollen(Namelist) = dws.Cells(Namelist + 4, 9).Value
  Next Namelist
  For MiniProject = 1 To Mini_Projects
    Mini_Project_Price(MiniProject) = dws.Cells(MiniProject + 4, 15).Value
    Mini_Project_Details(MiniProject) = dws.Cells(MiniProject + 4, 14).Value
  Next MiniProject
  
 'Creating Files
    tws.Columns("A:A").Copy
    sws.Columns("A:A").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    sws.Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
    NametoProcesses = WorksheetFunction.CountA(sws.Range("A:A")) - 1
    For NametoProcess = 1 To NametoProcesses
      For Namelist = 1 To Namelists
        If sws.Cells(NametoProcess + 1, 1).Value = Name_List(Namelist) Then
          Totalbypass = 1 + Totalbypass
          Activewb(Totalbypass) = BillingName_List(Namelist)
          Call Creat_Name_File(BillingName_List(Namelist))
          With Workbooks(BillingName_List(Namelist)).Sheets(1)
            .Cells.RowHeight = 18
            .Cells(1, 2).Value = dws.Cells(6, 12).Value
            .Cells(3, 2).Value = "NO."
            .Cells(3, 3).Value = "NAME"
            .Cells(3, 4).Value = "DATE"
            .Cells(3, 5).Value = "PROJECT/SCOPE OF WORK"
            .Cells(3, 6).Value = "JOB ID"
            .Cells(3, 7).Value = "REMARKS"
            .Cells(3, 8).Value = "PRICE"
            Call DetailTable(BillingName_List(Namelist), 1, 3, 2, 3, 8)
            .Rows("1:3").Font.Bold = True
            .Columns(4).NumberFormat = "d-mmm"
            .Columns(8).NumberFormat = """RM""#,##0.00"
            .Cells(1, 2).Font.Underline = xlUnderlineStyleSingle
            With .Rows(3)
              .HorizontalAlignment = xlCenter
              .VerticalAlignment = xlCenter
              .ReadingOrder = xlContext
            End With
            .Activate
            ActiveWindow.DisplayGridlines = False
            With .Cells.Font
              .Name = "Times New Roman"
              .Size = 12
            End With
          End With
          Exit For
        End If
      Next Namelist
    Next NametoProcess
    sws.Cells.Clear
    
 'Rearranging Data
  For TotalEntire = 1 To TotalEntires
    For TotalProject = 1 To TotalProjects
      If UCase(Store_Project(TotalProject)) = UCase(tws.Cells(TotalEntire + 1, 3).Value) And UCase(Task_Project(TotalProject)) = UCase(tws.Cells(TotalEntire + 1, 4).Value) And UCase(Status_Project(TotalProject)) = UCase(tws.Cells(TotalEntire + 1, 10).Value) Then
        For Namelist = 1 To Namelists
          If UCase(Name_List(Namelist)) = UCase(tws.Cells(TotalEntire + 1, 1).Value) Then
            If UCase(tws.Cells(TotalEntire + 1, 3).Value) <> "NBI" Then
              Numbering(Namelist) = Numbering(Namelist) + 1
              With Workbooks(BillingName_List(Namelist)).Sheets(1)
                Dilverydate(Namelist) = tws.Cells(TotalEntire + 1, 2).Value
                .Cells(Numbering(Namelist) + 3, 2).Value = Numbering(Namelist)
                .Cells(Numbering(Namelist) + 3, 3).Value = Name_List(Namelist)
                .Cells(Numbering(Namelist) + 3, 4).Value = tws.Cells(TotalEntire + 1, 2).Value
                .Cells(Numbering(Namelist) + 3, 5).Value = tws.Cells(TotalEntire + 1, 3).Value
                .Cells(Numbering(Namelist) + 3, 6).Value = tws.Cells(TotalEntire + 1, 6).Value
                .Cells(Numbering(Namelist) + 3, 7).Value = tws.Cells(TotalEntire + 1, 16).Value
                Call DetailTable(BillingName_List(Namelist), 1, Numbering(Namelist) + 3, 2, Numbering(Namelist) + 3, 8)
                With .Rows(Numbering(Namelist) + 3)
                  .HorizontalAlignment = xlCenter
                  .VerticalAlignment = xlCenter
                  .ReadingOrder = xlContext
                End With
                If tws.Cells(TotalEntire + 1, 11).Value <> "" Or tws.Cells(TotalEntire + 1, 12).Value <> "" Or tws.Cells(TotalEntire + 1, 13).Value <> "" Or tws.Cells(TotalEntire + 1, 14).Value <> "" Or tws.Cells(TotalEntire + 1, 15).Value <> "" Then
                  Project_price = tws.Cells(TotalEntire + 1, 11).Value * Price_Project(TotalProject) + tws.Cells(TotalEntire + 1, 12).Value * Price_Project(TotalProject) + tws.Cells(TotalEntire + 1, 13).Value * Price_Project(TotalProject) + tws.Cells(TotalEntire + 1, 14).Value * Price_Project(TotalProject) + tws.Cells(TotalEntire + 1, 15).Value * Price_Project(TotalProject)
                  .Cells(Numbering(Namelist) + 3, 8).Value = Project_price
                  Total_Price(Namelist) = Total_Price(Namelist) + Project_price
                ElseIf tws.Cells(TotalEntire + 1, 3).Value <> "NBI" Then
                  Project_price = Price_Project(TotalProject)
                  .Cells(Numbering(Namelist) + 3, 8).Value = Project_price
                  Total_Price(Namelist) = Total_Price(Namelist) + Project_price
                Else
                  Project_price = 0
                  Total_Price(Namelist) = 0
                End If
                'heilight based on date
                If DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 6) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + 3, 1)
                  ByodWeek1(Namelist) = ByodWeek1(Namelist) + tws.Cells(TotalEntire + 1, 9).Value
                  ByodPrice5(Namelist) = ByodPrice1(Namelist) + Project_price
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 7) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 13) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + 3, 2)
                  ByodWeek2(Namelist) = ByodWeek2(Namelist) + tws.Cells(TotalEntire + 1, 9).Value
                  ByodPrice2(Namelist) = ByodPrice2(Namelist) + Project_price
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 14) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 20) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + 3, 1)
                  ByodWeek3(Namelist) = ByodWeek3(Namelist) + tws.Cells(TotalEntire + 1, 9).Value
                  ByodPrice3(Namelist) = ByodPrice3(Namelist) + Project_price
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 21) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 27) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + 3, 2)
                  ByodWeek4(Namelist) = ByodWeek4(Namelist) + tws.Cells(TotalEntire + 1, 9).Value
                  ByodPrice4(Namelist) = ByodPrice4(Namelist) + Project_price
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 28) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 34) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + 3, 1)
                  ByodWeek5(Namelist) = ByodWeek5(Namelist) + tws.Cells(TotalEntire + 1, 9).Value
                  ByodPrice5(Namelist) = ByodPrice5(Namelist) + Project_price
                End If
                .Columns("A").ColumnWidth = 0.88
                .Columns("B").ColumnWidth = 6
                .Columns("C:H").AutoFit
              End With
            ElseIf tws.Cells(TotalEntire + 1, 3).Value = "NBI" Then
              If UCase(NBIProjects(TotalUniqueID(Namelist), Namelist)) <> UCase(tws.Cells(TotalEntire + 1, 6).Value) Then
                With Workbooks(BillingName_List(Namelist)).Sheets(1)
                  TotalUniqueID(Namelist) = TotalUniqueID(Namelist) + 1
                  NBIProjects(TotalUniqueID(Namelist), Namelist) = tws.Cells(TotalEntire + 1, 6).Value
                  For MiniProject = 1 To Mini_Projects
                    Total_MiniCount(Namelist) = Total_MiniCount(Namelist) + 1
                    Total_MiniTask(Namelist) = Total_MiniTask(Namelist) + Mini_Project_Price(MiniProject)
                    .Cells(Total_MiniCount(Namelist) + 3, 2).Value = Total_MiniCount(Namelist)
                    .Cells(Total_MiniCount(Namelist) + 3, 3).Value = Name_List(Namelist)
                    .Cells(Total_MiniCount(Namelist) + 3, 4).Value = DateValue(tws.Cells(TotalEntire + 1, 2).Value)
                    .Cells(Total_MiniCount(Namelist) + 3, 5).Value = Mini_Project_Details(MiniProject)
                    .Cells(Total_MiniCount(Namelist) + 3, 6).Value = NBIProjects(TotalUniqueID(Namelist), Namelist)
                    .Cells(Total_MiniCount(Namelist) + 3, 8).Value = Mini_Project_Price(MiniProject)
                    Call DetailTable(BillingName_List(Namelist), 1, Total_MiniCount(Namelist) + 3, 2, Total_MiniCount(Namelist) + 3, 8)
                    With .Rows(Total_MiniCount(Namelist) + 3)
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .ReadingOrder = xlContext
                    End With
                'heilight based on date
                If DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 6) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Total_MiniCount(Namelist) + 3, 1)
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 7) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 13) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Total_MiniCount(Namelist) + 3, 2)
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 14) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 20) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Total_MiniCount(Namelist) + 3, 1)
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 21) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 27) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Total_MiniCount(Namelist) + 3, 2)
                ElseIf DateValue(tws.Cells(TotalEntire + 1, 2).Value) >= DateValue(dws.Cells(7, 12).Value + 28) And DateValue(tws.Cells(TotalEntire + 1, 2).Value) <= DateValue(dws.Cells(7, 12).Value + 34) Then
                  Call ColorCode(BillingName_List(Namelist), "Sheet1", Total_MiniCount(Namelist) + 3, 1)
                End If
                .Columns("A").ColumnWidth = 0.88
                .Columns("B").ColumnWidth = 6
                .Columns("C:H").AutoFit
                  Next MiniProject
                End With
              End If
            End If
          End If
        Next Namelist
      End If
    Next TotalProject
  Next TotalEntire
  
 'Sum or Byod & Saving
    For Namelist = 1 To Namelists
      If Numbering(Namelist) <> 0 Or Total_MiniCount(Namelist) <> 0 Then
        If ByodWeek1(Namelist) >= 20 And ByodBollen(Namelist) = 1 And ByodPrice1(Namelist) >= 360 Then
          spac(Namelist) = 1 + spac(Namelist)
          ByodTotalPrice(Namelist) = ((ByodPrice1(Namelist) / 18) * 45) / 40
          With Workbooks(BillingName_List(Namelist)).Sheets(1)
            If ByodTotalPrice(Namelist) >= 45 Then
              ByodTotalPrice(Namelist) = 45
              .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            Else: .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            End If
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Value = "BYOD (" & DateValue(dws.Cells(7, 12).Value) & " - " & DateValue(dws.Cells(7, 12).Value + 6) & ")"
            .Range("B" & Numbering(Namelist) + spac(Namelist) + 3 & ":H" & Numbering(Namelist) + spac(Namelist) + 3).Merge
            Call DetailTable(BillingName_List(Namelist), 1, Numbering(Namelist) + 3, 2, Numbering(Namelist) + 3, 8)
            Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + spac(Namelist) + 3, 1)
            Total_Price(Namelist) = Total_Price(Namelist) + ByodTotalPrice(Namelist)
          End With
        ElseIf ByodWeek2(Namelist) >= 20 And ByodBollen(Namelist) = 1 And ByodPrice2(Namelist) >= 360 Then
          spac(Namelist) = 1 + spac(Namelist)
          ByodTotalPrice(Namelist) = ((ByodPrice2(Namelist) / 18) * 45) / 40
          With Workbooks(BillingName_List(Namelist)).Sheets(1)
            If ByodTotalPrice(Namelist) >= 45 Then
              ByodTotalPrice(Namelist) = 45
              .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            Else: .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            End If
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodPrice2(Namelist)
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Value = "BYOD (" & DateValue(dws.Cells(7, 12).Value + 7) & " - " & DateValue(dws.Cells(7, 12).Value + 13) & ")"
            .Range("B" & Numbering(Namelist) + spac(Namelist) + 3 & ":G" & Numbering(Namelist) + spac(Namelist) + 3).Merge
            Call DetailTable(BillingName_List(Namelist), 1, Numbering(Namelist) + spac(Namelist) + 3, 2, Numbering(Namelist) + spac(Namelist) + 3, 8)
            Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + spac(Namelist) + 3, 2)
            Total_Price(Namelist) = Total_Price(Namelist) + ByodTotalPrice(Namelist)
          End With
        ElseIf ByodWeek3(Namelist) >= 20 And ByodBollen(Namelist) = 1 And ByodPrice3(Namelist) >= 360 Then
          spac(Namelist) = 1 + spac(Namelist)
          ByodTotalPrice(Namelist) = ((ByodPrice3(Namelist) / 18) * 45) / 40
          With Workbooks(BillingName_List(Namelist)).Sheets(1)
            If ByodTotalPrice(Namelist) >= 45 Then
              ByodTotalPrice(Namelist) = 45
              .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            Else: .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            End If
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Value = "BYOD (" & DateValue(dws.Cells(7, 12).Value + 14) & " - " & DateValue(dws.Cells(7, 12).Value + 20) & ")"
            .Range("B" & Numbering(Namelist) + spac(Namelist) + 3 & ":G" & Numbering(Namelist) + spac(Namelist) + 3).Merge
            Call DetailTable(BillingName_List(Namelist), 1, Numbering(Namelist) + spac(Namelist) + 3, 2, Numbering(Namelist) + spac(Namelist) + 3, 8)
            Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + spac(Namelist) + 3, 1)
            Total_Price(Namelist) = Total_Price(Namelist) + ByodTotalPrice(Namelist)
          End With
        ElseIf ByodWeek4(Namelist) >= 20 And ByodBollen(Namelist) = 1 And ByodPrice4(Namelist) >= 360 Then
          spac(Namelist) = 1 + spac(Namelist)
          ByodTotalPrice(Namelist) = ((ByodPrice4(Namelist) / 18) * 45) / 40
          With Workbooks(BillingName_List(Namelist)).Sheets(1)
            If ByodTotalPrice(Namelist) >= 45 Then
              ByodTotalPrice(Namelist) = 45
              .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            Else: .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            End If
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Value = "BYOD (" & DateValue(dws.Cells(7, 12).Value + 21) & " - " & DateValue(dws.Cells(7, 12).Value + 27) & ")"
            .Range("B" & Numbering(Namelist) + spac(Namelist) + 3 & ":G" & Numbering(Namelist) + spac(Namelist) + 3).Merge
            Call DetailTable(BillingName_List(Namelist), 1, Numbering(Namelist) + spac(Namelist) + 3, 2, Numbering(Namelist) + spac(Namelist) + 3, 8)
            Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + spac(Namelist) + 3, 2)
            Total_Price(Namelist) = Total_Price(Namelist) + ByodTotalPrice(Namelist)
          End With
        ElseIf ByodWeek5(Namelist) >= 20 And ByodBollen(Namelist) = 1 And ByodPrice5(Namelist) >= 360 Then
          spac(Namelist) = 1 + spac(Namelist)
          ByodTotalPrice(Namelist) = ((ByodPrice5(Namelist) / 18) * 45) / 40
          With Workbooks(BillingName_List(Namelist)).Sheets(1)
            If ByodTotalPrice(Namelist) >= 45 Then
              ByodTotalPrice(Namelist) = 45
              .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            Else: .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Value = ByodTotalPrice(Namelist)
            End If
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 8).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Font.Bold = True
            .Cells(Numbering(Namelist) + spac(Namelist) + 3, 2).Value = "BYOD (" & DateValue(dws.Cells(7, 12).Value + 28) & " - " & DateValue(dws.Cells(7, 12).Value + 34) & ")"
            .Range("B" & Numbering(Namelist) + spac(Namelist) + 3 & ":G" & Numbering(Namelist) + spac(Namelist) + 3).Merge
            Call DetailTable(BillingName_List(Namelist), 1, Numbering(Namelist) + spac(Namelist) + 3, 2, Numbering(Namelist) + spac(Namelist) + 3, 8)
            Call ColorCode(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + spac(Namelist) + 3, 1)
            Total_Price(Namelist) = Total_Price(Namelist) + ByodTotalPrice(Namelist)
          End With
        End If
   
        spac(Namelist) = 1 + spac(Namelist)
        With Workbooks(BillingName_List(Namelist)).Sheets(1)
          .Cells(Numbering(Namelist) + Total_MiniCount(Namelist) + spac(Namelist) + 3, 8).Value = Total_Price(Namelist) + Total_MiniTask(Namelist)
          .Cells(Numbering(Namelist) + Total_MiniCount(Namelist) + spac(Namelist) + 3, 8).Font.Bold = True
          .Columns("C:H").AutoFit
        End With
        Call TotalTable(BillingName_List(Namelist), "Sheet1", Numbering(Namelist) + Total_MiniCount(Namelist) + spac(Namelist) + 3, 8, Numbering(Namelist) + Total_MiniCount(Namelist) + spac(Namelist) + 3, 8)
        Workbooks(BillingName_List(Namelist)).Save
      End If
    Next Namelist
    
    Call DeleteDataSheet(Masterwb, Sparews)
    Call Check_if_workbook_is_open(Trackingwb)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    
End Sub
Function StopAllFilters(Filename)
  Dim ws As Worksheet
  For Each ws In Workbooks(Filename).Worksheets
   If ws.AutoFilterMode = True Then
      ws.AutoFilterMode = False
   End If
  Next ws
End Function
Function StartFilter(Filename, Sheetname)
  If Not Workbooks(Filename).Sheets(Sheetname).AutoFilterMode Then
     Workbooks(Filename).Sheets(Sheetname).Range("A1").AutoFilter
  End If
End Function
Private Function Creat_Name_File(Namelist)
    Application.DisplayAlerts = False
    Filename_Output = ThisWorkbook.Name
    Filename_Output = Namelist & CommonName & ".xlsx"
    Call Check_if_workbook_is_open(Filename_Output)
    Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\Output files\" & Filename_Output, CreateBackup:=False
    Application.DisplayAlerts = True
End Function

Private Function Check_if_workbook_is_open(Filename)
    Dim wb As Workbook 'to test if workbook is open. No change here
        For Each wb In Workbooks
            If wb.Name = Filename Then
                Workbooks(Filename).Save
                Workbooks(Filename).Close
            End If
        Next
End Function
Function CheckDataSheet(Filename, Sheetname)
    For Each Sheet In Workbooks(Filename).Worksheets ' Checking if VBA Sheet exist
        If Sheet.Name = Sheetname Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
    Workbooks(Filename).Sheets.Add.Name = Sheetname
End Function
Function DeleteDataSheet(Filename, Sheetname)
    For Each Sheet In Workbooks(Filename).Worksheets
        If Sheet.Name = Sheetname Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
End Function
Function DetailTable(Outputwb, Outputws, X1, Y1, X2, Y2)
  'Output Workbook
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlDiagonalDown).LineStyle = xlNone
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlDiagonalUp).LineStyle = xlNone
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Function
Function TotalTable(Outputwb, Outputws, X1, Y1, X2, Y2)
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlDiagonalDown).LineStyle = xlNone
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlDiagonalUp).LineStyle = xlNone
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeLeft).LineStyle = xlNone
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeRight).LineStyle = xlNone
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlInsideVertical).LineStyle = xlNone
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlInsideHorizontal).LineStyle = xlNone
End Function

Function ColorCode(Outputwb, Outputws, Rangeno, ColorType)
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    If ColorType = 1 Then
      With ows.Range("B" & Rangeno & ":H" & Rangeno).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
      End With
    ElseIf ColorType = 2 Then
      With ows.Range("B" & Rangeno & ":H" & Rangeno).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
      End With
    End If
End Function




