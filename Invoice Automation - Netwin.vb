'This program can take multiple input and create multiple output.
'Used for Data collection from different Wb to another.


Sub Main_Program()
  'Getting data File & Setup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim Output_workbook(100) As Variant
    Dim TablesFIlters As Variant
    Dim TableData(100) As Variant
    Dim Track_WB_Filename(100) As Variant

    
    
    'On Error Resume Next
    
    datawb = ActiveWorkbook.Name
    dataws = "Data Sheet"
    Tablews = "Table Enteries"
    Filews = "File Info"
    Appendws = "Appendix"
    date1 = Format(Date, "[$-en-US]d mmm yyyy;@")
    Dim dwb As Workbook
    Set dwb = Workbooks(datawb)
    Dim dws As Worksheet
    Set dws = dwb.Sheets(dataws)
    Dim dwt As Worksheet
    Set dwt = dwb.Sheets(Tablews)
    Dim dwf As Worksheet
    Set dwf = dwb.Sheets(Filews)
    Dim dwa As Worksheet
    Set dwa = dwb.Sheets(Appendws)
   
  'Calculate
    Tracking_Links = WorksheetFunction.CountA(dwf.Range("A:A")) - 1
    Tables_Height = WorksheetFunction.CountA(dwt.Range("B:B")) - 1
    Total_Region = WorksheetFunction.CountA(dwf.Range("B:B")) - 1
    NJN_Town_Total = WorksheetFunction.CountA(dwa.Range("D:D")) - 1
    NJS_Town_Total = WorksheetFunction.CountA(dwa.Range("E:E")) - 1
    Table_Width = 17
    
  'Creating Output files & Saving them
    For Region_Count = 1 To Total_Region
      Output_workbook(Region_Count) = dwf.Cells(Region_Count + 1, 3).Value & dwf.Cells(Region_Count + 1, 2).Value & date1 & ".xlsx"
      Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\Output files\" & Output_workbook(Region_Count), CreateBackup:=False
      ColumnLetter = Split(Cells(1, Region_Count + 4).Address, "$")(1)
      Sheet_Count_Total = WorksheetFunction.CountA(dwf.Range(ColumnLetter & ":" & ColumnLetter)) - 1
      For Sheet_Count = 1 To Sheet_Count_Total
        Workbooks(Output_workbook(Region_Count)).Sheets.Add.Name = dwf.Cells(1 + Sheet_Count, Region_Count + 4).Value
      Next
    Next
   
  'Extracting Input Files & Saving them
    For Tracking_Links_Count = 1 To Tracking_Links
      Link = dwf.Cells(1 + Tracking_Links_Count, 1).Value
      Dim twb As Workbook
      Set twb = Application.Workbooks.Open(Link)
      Track_WB_Filename(Tracking_Links_Count) = twb.Name
      strPath = ThisWorkbook.Path & "\Input files\" & Track_WB_Filename(Tracking_Links_Count)
      twb.SaveAs Filename:=strPath
    Next

  'Extracting data details from table entriers
    For Tables_Height_Count = 1 To Tables_Height
      For Table_Width_Count = 1 To Table_Width
        TableData(Table_Width_Count) = dwt.Cells(Tables_Height_Count + 1, Table_Width_Count).Value
      Next
      TablesFIlters = Array(TableData(8), TableData(9), TableData(10), TableData(11), TableData(12))
      For Total_Region_Count = 1 To Total_Region
        If TableData(14) = dwf.Cells(Total_Region_Count + 1, 2).Value Then
          CurrentWb = Output_workbook(Total_Region_Count)
        Exit For
        End If
      Next
      Call Master(TableData(1), TableData(2), TableData(3), TableData(4), TableData(5), TableData(6), TableData(7), TablesFIlters, TableData(13), CurrentWb, TableData(15), TableData(16), TableData(17), datawb, dataws, Filews)
    Next
    
  'Creating Output files & Saving them
    For Tracking_Links_Count = 1 To Tracking_Links
      Workbooks(Track_WB_Filename(Tracking_Links_Count)).Close SaveChanges:=True
    Next
End Sub
Function Master(Startcopyrange, Table_Column_Alpha, Total_Start_Count, Region_Table_Name, Track_Header, Status, Region, Filter_Column_Name As Variant, Common, Outputwb, Outputws, Trackingwb, Trackingws, datawb, dataws, Filews)
  'Declaring WB
    Dim dwb As Workbook
    Set dwb = Workbooks(datawb)
    Dim dws As Worksheet
    Set dws = dwb.Sheets(dataws)
    Dim dwf As Worksheet
    Set dwf = dwb.Sheets(Filews)
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    Dim twb As Workbook
    Set twb = Workbooks(Trackingwb)
    Dim tws As Worksheet
    Set tws = twb.Sheets(Trackingws)
  
  'Decalaring Varaibles
    Dim Header(100) As Variant
    Dim Header_Actual_Col(100) As Variant
    Dim Header_Title(100) As Variant
    Dim Header_Actual(100) As Variant
    Dim ColRef(100) As Variant
    Dim ColRefno(100) As Variant
    Dim Filter_Column_No(100) As Variant
    Dim Filter_Column_Alpha(100) As Variant
    Dim NJN_Town(200) As Variant
    Dim NJS_Town(200) As Variant
    
  'Starting and ending date
    Startingdate = Format(Date + dwf.Cells(2, 4).Value, "\>\=mm/dd/yyyy")
    Endingdate = Format(Date + dwf.Cells(3, 4).Value, "\<\=mm/dd/yyyy")
  
  'Getting Values
    Title_Col = Range(Table_Column_Alpha & 1).Column
    Shift = ows.Cells(Rows.Count, 4).End(xlUp).Row + 6
    Total_Column_Tws = 100
    Total_Filters = 4
    NJN_Town_Total = WorksheetFunction.CountA(dwa.Range("D:D")) - 1
    NJS_Town_Total = WorksheetFunction.CountA(dwa.Range("E:E")) - 1

   'Getting Town for NJ
    For NJN_Town_Count = 1 To NJN_Town_Total
      NJN_Town(NJN_Town_Count) = dwa.Cells(NJN_Town_Count + 1, 4).Value
    Next
    For NJS_Town_Count = 1 To NJS_Town_Total
      NJS_Town(NJS_Town_Count) = dwa.Cells(NJS_Town_Count + 1, 5).Value
    Next
    
  'Clearing filter and reapplying filter
    Call StopAllFilters(Trackingwb)
    Call StartFilter(Trackingwb, Trackingws)
    For Filtered_Count = 0 To Total_Filters
      For Total_Column_Tws_Count = 1 To Total_Column_Tws
        If tws.Cells(Track_Header, Total_Column_Tws_Count) = Filter_Column_Name(Filtered_Count) And tws.Cells(Track_Header, Total_Column_Tws_Count) <> Empty Then
          Filter_Column_Alpha(Filtered_Count) = Split(tws.Cells(Track_Header, Total_Column_Tws_Count).Address(True, False), "$")(0)
          Filter_Column_No(Filtered_Count) = Range(Filter_Column_Alpha(Filtered_Count) & 1).Column
          Exit For
        End If
      Next
    Next
    With tws.Range("A1")
      If Filter_Column_No(0) <> Empty Then
       .AutoFilter Field:=Filter_Column_No(0), Criteria1:=Region     'Region
      End If
      If Filter_Column_No(1) <> Empty Then
       .AutoFilter Field:=Filter_Column_No(1), Criteria1:="="  'Invoice colum
      End If
      If Filter_Column_No(2) <> Empty Then
       .AutoFilter Field:=Filter_Column_No(2), Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate 'date
      End If
      If Filter_Column_No(3) <> Empty Then
       .AutoFilter Field:=Filter_Column_No(3), Criteria1:=Array("COMPLETED", "COMPLETE", "DELIVERED", "DONE")   'Status
      End If
      If Filter_Column_No(4) <> Empty Then
        If Region_Table_Name = "New Jersey North Region" Then
          .AutoFilter Field:=Filter_Column_No(4), Criteria1:=NJN_Town
        ElseIf Region_Table_Name = "New Jersey South Region" Then
          .AutoFilter Field:=Filter_Column_No(4), Criteria1:=NJS_Town
        Else
          .AutoFilter Field:=Filter_Column_No(4), Criteria1:=Common
      End If
    End With
   
  'Counting Total Entries in data sheet
    Total_Headers = WorksheetFunction.CountA(dws.Range(Table_Column_Alpha & ":" & Table_Column_Alpha))
    Total_Enteries = tws.AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   
  'Storing datasheet
   For Total_Headers_Count = 1 To Total_Headers - 1
     Header_Title(Total_Headers_Count) = dws.Cells(Total_Headers_Count + 1, Title_Col).Value 'Header Title name
     Header_Actual(Total_Headers_Count) = dws.Cells(Total_Headers_Count + 1, Title_Col + 2).Value 'Header actual name
     ColRef(Total_Headers_Count) = dws.Cells(Total_Headers_Count + 1, 1).Value 'Ref Columns
     ColRefno(Total_Headers_Count) = Range(ColRef(Total_Headers_Count) & 1).Column 'Ref Number
   Next Total_Headers_Count
  
  'Main Features
    For Total_Headers_Count = 1 To Total_Headers - 1
      'Blank Feature
        If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = Empty Then
      'Numbering Feature
        ElseIf dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Number Ref" Then
          For Total_Enteries_Count = 1 To Total_Enteries
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Total_Enteries_Count
          Next
      'Copying Feature
        ElseIf dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value <> Empty Then
          For Total_Column_Tws_Count = 1 To Total_Column_Tws
            If tws.Cells(Track_Header, Total_Column_Tws_Count) = Header_Actual(Total_Headers_Count) Then
              Header_Actual_Col(Total_Headers_Count) = Split(dws.Cells(Track_Header, Total_Column_Tws_Count).Address(True, False), "$")(0)
              tws.Range(Header_Actual_Col(Total_Headers_Count) & Startcopyrange & ":" & Header_Actual_Col(Total_Headers_Count) & "99999").Copy
              ows.Range(ColRef(Total_Headers_Count) & Shift + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
            End If
          Next
    End If
    
  '1 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Mod" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty Then
        ElseIf ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value Mod 2 = 0 Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value / 2
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = ""
        ElseIf ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value Mod 2 = 1 Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value / 2
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = 1
        End If
      Next Total_Enteries_Count
    End If
  '2 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Spliting" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) - 1).Font.Bold = True
        'Disturbuting values
          If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value >= 1 And ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value <= 30 Then
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 1
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = Empty
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 2).Value = Empty
          ElseIf ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value >= 30 And ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value <= 60 Then
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = 1
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 2).Value = Empty
          ElseIf ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value >= 61 Then
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = Empty
            ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 2).Value = 1
          End If
      Next Total_Enteries_Count
    End If
    
  '3 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "If True" Then
      FirstSplit = Split(dws.Cells(Total_Headers_Count + 1, Title_Col + 2).Value, ",")(0)
      SecondSplit = Split(dws.Cells(Total_Headers_Count + 1, Title_Col + 2).Value, ",")(1)
      For Total_Column_Tws_Count = 1 To Total_Column_Tws
        If tws.Cells(Track_Header, Total_Column_Tws_Count) = FirstSplit Then
          FirstSplitKeyword = Split(dws.Cells(Track_Header, Total_Column_Tws_Count).Address(True, False), "$")(0)
          tws.Range(FirstSplitKeyword & Startcopyrange & ":" & FirstSplitKeyword & "99999").Copy
          ows.Range(ColRef(Total_Headers_Count) & Shift).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
      Next
      For Total_Column_Tws_Count = 1 To Total_Column_Tws
        If tws.Cells(Track_Header, Total_Column_Tws_Count) = SecondSplit Then
          SecondSplitKeyword = Split(dws.Cells(Track_Header, Total_Column_Tws_Count).Address(True, False), "$")(0)
          tws.Range(SecondSplitKeyword & Startcopyrange & ":" & SecondSplitKeyword & "99999").Copy
          ows.Range(ColRef(Total_Headers_Count + 1) & Shift).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        End If
      Next
      For Total_Enteries_Count = 1 To Total_Enteries
       'Disturbuting values
         If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value < 1 And ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value < 100 Then
           ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
           ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count + 1)).Value = Empty
         Else
           ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count + 1)).Value = Empty
         End If
      Next Total_Enteries_Count
    End If
    
  '4 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Package" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty Or ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 0 Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = 1
        Else
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 2).Value = 1
        End If
      Next Total_Enteries_Count
    End If
    
  '5 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Y/N" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = "Y" Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 1
        Else
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
        End If
      Next Total_Enteries_Count
    End If
    
  '6 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "RFI" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        RFI_Statement = InStr(ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value, "Design Modification to include feeder")
        If RFI_Statement = 1 Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 2).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 3).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 4).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 7).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 9).Value = 1
        Else
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 0.5
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 1).Value = 1
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count) + 2).Value = 1
        End If
      Next Total_Enteries_Count
    End If
    
  '7 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Copy_Sheetname" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Outputws
      Next Total_Enteries_Count
    End If
    
  '8 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "If > 0.5" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value > 0.49 Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 1
        Else
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
        End If
      Next Total_Enteries_Count
    End If
    
  '9 Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "TICK/X" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        If ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = "X" Or ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty Then
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = Empty
        Else
          ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = 1
        End If
      Next Total_Enteries_Count
    End If
    
  'Constant Feature
    If dws.Cells(Total_Headers_Count + 1, Title_Col + 1).Value = "Constant" Then
      For Total_Enteries_Count = 1 To Total_Enteries
        ows.Cells(Shift + Total_Enteries_Count, ColRefno(Total_Headers_Count)).Value = dws.Cells(Total_Headers_Count + 1, Title_Col + 2).Value
      Next Total_Enteries_Count
    End If
    
  'Renaming Header Title
    ows.Cells(Shift - 2, ColRefno(Total_Headers_Count)).Value = Header_Title(Total_Headers_Count) 'Header Names
    ows.Cells(Shift - 2, ColRefno(Total_Headers_Count)).Font.Bold = True
   Next
   
  'FInding total and making zero value blank
    For columncount = Total_Start_Count To Total_Headers - 1
      Total_Count = 0
      For Total_Enteries_Count = 1 To Total_Enteries
        If IsNumeric(ows.Cells(Shift + Total_Enteries_Count, columncount).Value) = True Then
          If ows.Cells(Shift + Total_Enteries_Count, columncount).Value = 0 Then
            ows.Cells(Shift + Total_Enteries_Count, columncount).Value = Empty
          Else
            Total_Count = Total_Count + ows.Cells(Shift + Total_Enteries_Count, columncount).Value
          End If
        Else
        End If
      Next Total_Enteries_Count
      If Total_Count <> 0 Then
        ows.Cells(Shift + Total_Enteries_Count + 1, columncount).Value = Total_Count
        ows.Cells(Shift + Total_Enteries_Count + 1, columncount).Font.Bold = True
      End If
      With ows.Cells(Shift + Total_Enteries_Count + 1, columncount).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.249977111117893
        .Bold = True
      End With
    Next columncount
   
  'Creating table and highlight and total title
    Call ThinBoarderInv(ColRef(1) & Shift - 1 & ":" & ColRef(Total_Headers - 1) & Shift + Total_Enteries + 1, Outputwb, Outputws)
    ows.Rows(Shift + Total_Enteries + 1 & ":" & Shift + Total_Enteries + 1).RowHeight = 2.25
    ows.Cells(Shift + Total_Enteries + 2, 4).Value = "Total:"
    With ows.Cells(Shift + Total_Enteries + 2, 4).Font
      .ThemeColor = xlThemeColorLight2
      .TintAndShade = -0.249977111117893
      .Bold = True
    End With
    ows.Cells.HorizontalAlignment = xlCenter
    ows.Cells.VerticalAlignment = xlCenter
    ows.Columns(5).HorizontalAlignment = xlLeft
    ows.Cells(Shift - 4, 3).Value = Region_Table_Name
    ows.Cells(Shift - 4, 3).Font.Bold = True
    ows.Cells(Shift - 4, 3).Font.Size = 16
    With ows.Range(ColRef(1) & Shift - 2 & ":" & ColRef(Total_Headers - 1) & Shift - 2).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    For Total_Headers_Count = 1 To Total_Headers
      ows.Columns(Total_Headers_Count).ColumnWidth = dws.Cells(Total_Headers_Count + 1, Title_Col + 3).Value
      ows.Cells(Shift, Total_Headers_Count).WrapText = True
    Next
End Function
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
Function ThinBoarderInv(Boarderange, Outputwb, Outputws)
  'Output Workbook
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    With ows
    .Range(Boarderange).Borders(xlDiagonalDown).LineStyle = xlNone
    .Range(Boarderange).Borders(xlDiagonalUp).LineStyle = xlNone
    With .Range(Boarderange).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Range(Boarderange).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Range(Boarderange).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Range(Boarderange).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Range(Boarderange).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Range(Boarderange).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   End With
End Function
