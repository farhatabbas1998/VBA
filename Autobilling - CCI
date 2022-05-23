'Version: 1.00
'For Echobroadband
'By Farhat Abbas (Verified & Tested)

Sub CCI_Billing()
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Masterwb = ThisWorkbook.Name 'Master Macro workbook name
  MasterwbLocation = ThisWorkbook.Path
  Dataws = "Billing_CCI"
  Sparews = "Sparews"
  OSparews = "Sheet1"
  On Error Resume Next
  'Online tracking
  strUrl = Workbooks(Masterwb).Sheets(Dataws).Range("D2").Value
  Dim wb As Workbook
  Set wb = Application.Workbooks.Open(strUrl)
  Trackingwb = wb.Name 'Tracking workbook name
  strPath = MasterwbLocation & "\Input files\" & Trackingwb
  wb.SaveAs Filename:=strPath
  
  Region = Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value
  Invoiceno = Workbooks(Masterwb).Sheets(Dataws).Cells(6, 9).Value
  Outputwb = Workbooks(Masterwb).Sheets(Dataws).Cells(7, 9).Value
  Outputws1 = Workbooks(Masterwb).Sheets(Dataws).Cells(8, 9).Value
  Outputws2 = Workbooks(Masterwb).Sheets(Dataws).Cells(9, 9).Value
  Trackingws1 = Workbooks(Masterwb).Sheets(Dataws).Cells(10, 9).Value
  Trackingws2 = Workbooks(Masterwb).Sheets(Dataws).Cells(11, 9).Value
  Trackingws3 = Workbooks(Masterwb).Sheets(Dataws).Cells(12, 9).Value
  Trackingws4 = Workbooks(Masterwb).Sheets(Dataws).Cells(13, 9).Value
  Trackingws5 = Workbooks(Masterwb).Sheets(Dataws).Cells(14, 9).Value
  
  
  Call Check_if_workbook_is_open(Outputwb)
  Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\Output files\" & Outputwb, CreateBackup:=False
    
  Call CheckDataSheet(Outputwb, Outputws1)
  Call CheckDataSheet(Outputwb, Outputws2)

  Call CheckDataSheet(Masterwb, Sparews)
  Call findCellColumnNo(Masterwb, Dataws, Trackingwb, Trackingws1, 23)
  Call findCellColumnNo(Masterwb, Dataws, Trackingwb, Trackingws2, 24)
  Call findCellColumnNo(Masterwb, Dataws, Trackingwb, Trackingws3, 25)
  Call findCellColumnNo(Masterwb, Dataws, Trackingwb, Trackingws4, 26)
  Call findCellColumnNo(Masterwb, Dataws, Trackingwb, Trackingws5, 27)
  Call Invoice_Summary_Templete(Masterwb, Dataws, Outputwb, Outputws1, Region, Invoiceno)
  Call CreatingODetailOutput(Masterwb, Dataws, Sparews, OSparews, Trackingwb, Trackingws1, Trackingws2, Trackingws3, Trackingws4, Trackingws5, Outputwb, Outputws1, Outputws2, Region, Invoiceno)
  Call DeleteDataSheet(Masterwb, Sparews)
  Call DeleteDataSheet(Outputwb, OSparews)
  Call pagesetupsetting(Outputwb, Outputws1)
  Workbooks(Masterwb).Sheets(Dataws).Columns("AE:AF").ClearContents
  Call Check_if_workbook_is_open(Trackingwb)
  Call titletext(Outputwb, Outputws2, Invoiceno, Region)
  Workbooks(Outputwb).Worksheets(Outputws1).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Worksheets(Outputws2).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Sheets(Outputws1).Range("I11").Value = Region
  
  Workbooks(Outputwb).Save
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
End Sub
Function CreatingODetailOutput(Masterwb, Dataws, Sparews, OSparews, Trackingwb, Trackingws1, Trackingws2, Trackingws3, Trackingws4, Trackingws5, Outputwb, Outputws1, Outputws2, Region, Invoiceno)
  'Input Workbook is represented as wb
    Dim mwb As Workbook
    Set mwb = Workbooks(Masterwb)
    Dim dws As Worksheet
    Set dws = mwb.Sheets(Dataws)
    Dim sws As Worksheet
    Set sws = mwb.Sheets(Sparews)

  'Input Workbook is represented as twb
    Dim twb As Workbook
    Set twb = Workbooks(Trackingwb)
  
  'input Worksheet is represented as tws
    Dim tws1 As Worksheet
    Set tws1 = twb.Sheets(Trackingws1)
    Dim tws2 As Worksheet
    Set tws2 = twb.Sheets(Trackingws2)
    Dim tws3 As Worksheet
    Set tws3 = twb.Sheets(Trackingws3)
    Dim tws4 As Worksheet
    Set tws4 = twb.Sheets(Trackingws4)
    Dim tws5 As Worksheet
    Set tws5 = twb.Sheets(Trackingws5)
  
  'Output Workbook is represented as owb
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
  
  'Output Worksheet is represented as ow
    Dim osws As Worksheet
    Dim ows1 As Worksheet
    Dim ows2 As Worksheet
    Set osws = owb.Sheets(OSparews)
    Set ows1 = owb.Sheets(Outputws1)
    Set ows2 = owb.Sheets(Outputws2)
    Dim dilverytype As Variant
    Call StopAllFilters(Trackingwb)
    Call StartFilter(Trackingwb, Trackingws1)
    Call StartFilter(Trackingwb, Trackingws2)
    Call StartFilter(Trackingwb, Trackingws3)
    Call StartFilter(Trackingwb, Trackingws4)
    Call StartFilter(Trackingwb, Trackingws5)
    
  '----------------------------------------------------------------------------

    'Status
      dilverytype = Array("Completed", "Completed 2", "Completed  2 (Valid)", "Correction-IP")

   'Table Naming
      CommTN1 = "Node Split Design & Asbuilt"
      CommTN2 = "Coax Design & Asbuild"
      CommTN3 = "SFU & MDU"
      CommTN4 = "Fiber Design & Asbuild"

    'X and Y shifting and space gap
      Yspacing = 3
      Xspacing = 4
      diff = 5
   '----------------------------------------------------------------------------

   TotalD = WorksheetFunction.CountA(dws.Range("C:C")) - 2
   'Table 1 --------------------------- NS
      With tws4.Range("A1")
        x = 3
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Criteria1:="<>="
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=" 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws4, Masterwb, Sparews, x)
      CountCom1 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 1, Xspacing, Yspacing, CommTN1, CountCom1, 6, 7)

      
    'Table 2 ---------------------------CED
      With tws1.Range("A1")
        x = 0
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Criteria1:="<>="
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=" 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws1, Outputwb, OSparews, x)
      CCom1 = WorksheetFunction.CountA(osws.Range("C:C"))
      With tws2.Range("A1")
        x = 1
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Criteria1:="<>="
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=" 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws2, Masterwb, Sparews, x)
      CCom2 = WorksheetFunction.CountA(sws.Range("C:C"))
      If CCom1 >= 2 Then
       TotalEntires = WorksheetFunction.CountA(dws.Range("T:T"))
       osws.Range(osws.Cells(6, 20).Value & "2:" & osws.Cells(6 + TotalEntires, 20).Value & CCom1).Copy
       sws.Range("A" & CCom2 + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      End If
      CountCom2 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 2, Xspacing + diff + CountCom1, Yspacing, CommTN2, CountCom2, 8, 9)
    'Table 3 --------------------------- NS
      With tws3.Range("A1")
        x = 2
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:="<>="
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=" 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws3, Masterwb, Sparews, x)
      CountCom3 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 3, Xspacing + diff * 2 + CountCom1 + CountCom2, Yspacing, CommTN3, CountCom3, 10, 11)
      
    'Table 4 --------------------------- FD
      With tws5.Range("A1")
        x = 4
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:="<>="
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=" 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws5, Masterwb, Sparews, x)
      CountCom4 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 4, Xspacing + diff * 3 + CountCom1 + CountCom2 + CountCom3, Yspacing, CommTN4, CountCom4, 12, 13)
   '-----------------------------------------------------------------------
   
   ows1.Cells(TotalD + 15, 12).Value = ows1.Cells(TotalD + 14, 7).Value + ows1.Cells(TotalD + 14, 9).Value + ows1.Cells(TotalD + 14, 11).Value + ows1.Cells(TotalD + 14, 13).Value
   
End Function
Function Invoice_Summary_Templete(Masterwb, Dataws, Outputwb, Outputws, Region, Invoiceno)
   'Master Macro Workbook
   Dim mwb As Workbook
   Set mwb = Workbooks(Masterwb)
   Dim dws As Worksheet
   Set dws = mwb.Sheets(Dataws)
   Dim sws As Worksheet
  'Output Workbook
   Dim wbo As Workbook
   Set wbo = Workbooks(Outputwb)
   Dim wo1 As Worksheet
   Set wo1 = wbo.Sheets(Outputws)
   TotalD = WorksheetFunction.CountA(dws.Range("C:C")) - 2
   Noofinvtab = WorksheetFunction.CountA(dws.Range("AC:AC")) - 2
 
  'Writing Values
   With wo1
    'Creating templete
      .Range("K1").FormulaR1C1 = "Invoice"
      .Range("K2").FormulaR1C1 = "DATE"
      .Range("K3").FormulaR1C1 = Date
      .Range("L3").NumberFormat = "d-mmm-yyyy"
      .Range("L2").FormulaR1C1 = "INVOICE #"
      .Range("L3").FormulaR1C1 = Invoiceno
      .Range("K5").FormulaR1C1 = "BILL TO"
      .Range("K6").FormulaR1C1 = "Echo Broadband, Inc"
      .Range("K7").FormulaR1C1 = "PO Box 1627"
      .Range("K8").FormulaR1C1 = "Broomfield, CO 80038"
      .Range("B11").FormulaR1C1 = "PROJECT: CCI / Comcast"
      .Range("D11").FormulaR1C1 = "** Data Processing and Provision of information**"
      .Range("B6").FormulaR1C1 = "ECHO Broadband Sdn Bhd"
      .Range("B7").FormulaR1C1 = "368-5-3 Bellisa Row"
      .Range("B8").FormulaR1C1 = "Jalan Burmah"
      .Range("B9").FormulaR1C1 = "10350 Penang"
      .Range("L11").FormulaR1C1 = "Terms"
      .Range("M11").FormulaR1C1 = "Net 90"
      .Range("H11").FormulaR1C1 = "Region"
      .Range("F12").FormulaR1C1 = "Node Splits "
      .Range("H12").FormulaR1C1 = "Coax Design & Asbuild"
      .Range("J12").FormulaR1C1 = "SFU & MDU"
      .Range("L12").FormulaR1C1 = "Fiber Design & Asbuild"
      .Range("L2:M2,L3:M3,K5:M5,K6:M6,K7:M7,K8:M8,K9:M9,D11:G11,F12:G12,H12:I12,J12:K12,L12:M12,L" & TotalD + 15 & ":M" & TotalD + 15).Merge
      .Range("F13,H13,J13,L13").FormulaR1C1 = "Qty"
      .Range("G13,I13,K13,M13").FormulaR1C1 = "Sub Total"
      .Rows(11).RowHeight = 30
      .Rows(12).RowHeight = 30
      Call Thin_OB("K5:M5,K6:M9", Outputwb, Outputws)
      Call ThinBoarderInv("K2:M3,H11:M12", Outputwb, Outputws)
      
    'calling function
        Call Copytovalueinvoice(Masterwb, Dataws, Outputwb, Outputws, "B13")
        Call ThinBoarderInv("B" & 14 & ":M" & TotalD + 13, Outputwb, Outputws)
        Call Thick_OB("K5:M9,K2:M3,B12:E12,B" & 13 & ":E" & TotalD + 14, Outputwb, Outputws)
        Call Thick_OB("F" & 14 & ":G" & TotalD + 14 & ",H" & 14 & ":I" & TotalD + 14 & ",J" & 14 & ":K" & TotalD + 14 & ",L" & 14 & ":M" & TotalD + 14 & ",L" & TotalD + 15 & ":M" & TotalD + 15, Outputwb, Outputws)
        Call Thick_OB("B11:M12,H11:I12,B13:M13,F12:G12,H12:I12,J12:K12,L12:M12,B14:M" & TotalD + 15, Outputwb, Outputws)
      .Range("E14:E" & TotalD + 15 & ",G14:G" & TotalD + 15 & ",I14:I" & TotalD + 15 & ",K14:K" & TotalD + 15 & ",M14:M" & TotalD + 15 & ",L" & TotalD + 15).NumberFormat = "[$$-en-US]#,##0.000"
      .Range("K" & TotalD + 15).FormulaR1C1 = "Invoice Total (USD)"
      dws.Cells(1, 32).Value = TotalD + 14
    'format
      .Range("A1:M13,H" & TotalD + 15).Font.Bold = True
      .Range("E12").Font.Bold = False
      .Range("A1:I" & TotalD * Noofinvtab + 14).Font.Size = 12
      .Range("B7:B9,I15:I" & TotalD + 14).Font.Size = 11
      .Range("K1").Font.Size = 20
      .Range("K3").Font.Name = "Arial"
      With .Range("K5").Font
      .Name = "Arial Black"
      .Size = 14
      End With
      With .Range("H6:H8").Font
      .Name = "Arial"
      .Size = 10
      End With
      .Columns("A").ColumnWidth = 0.88
      .Columns("B").ColumnWidth = 9
      .Columns("C").ColumnWidth = 32
      .Columns("D").ColumnWidth = 9
      .Columns("E").ColumnWidth = 10.2
      .Columns(6).ColumnWidth = 10
      .Columns(7).ColumnWidth = 11.2
      .Columns(8).ColumnWidth = 10
      .Columns(9).ColumnWidth = 11.2
      .Columns(10).ColumnWidth = 10
      .Columns(11).ColumnWidth = 11.2
      .Columns(12).ColumnWidth = 10
      .Columns(13).ColumnWidth = 11.2
      
      With .Range("B11:M12").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10020351
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      
      
      With .Range("H2:M5,E11:M12,D11,B12:M" & TotalD + 15)
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .ReadingOrder = xlContext
      End With
      With .Range("I1,K" & TotalD + 15)
          .HorizontalAlignment = xlRight
          .VerticalAlignment = xlCenter
          .ReadingOrder = xlContext
      End With
      With .Range("C13:C" & TotalD + 14)
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlCenter
          .WrapText = True
          .ReadingOrder = xlContext
      End With
      .Range("B13:I13,D11").WrapText = True
      End With
End Function
Function ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws, Outputws2, OutputValue, XShiftBlank, YShiftBlank, Tabletitle, Count, column1, column2)
  'Input Workbook is represented as wb
    Dim mwb As Workbook
    Set mwb = Workbooks(Masterwb)
    Dim dws As Worksheet
    Set dws = mwb.Sheets(Dataws)
    Dim sws As Worksheet
    Set sws = mwb.Sheets(Sparews)
  'Output Workbook is represented as owb
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
  
  'Output Worksheet is represented as ow
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    Dim ows2 As Worksheet
    Set ows2 = owb.Sheets(Outputws2)
    
    Numbering = 1
    TotalTask = WorksheetFunction.CountA(dws.Range("Q:Q")) - 1
    DTask = WorksheetFunction.Sum(dws.Range("Q:Q"))
    newcount = WorksheetFunction.CountA(sws.Range("C:C")) - 1
    With ows
         .Columns("A").ColumnWidth = 0.88
         .Columns("B").ColumnWidth = 3.75
         .Columns("C").ColumnWidth = 4.75
         .Columns("D").ColumnWidth = 12
         .Columns("E").ColumnWidth = 13
         .Columns("F").ColumnWidth = 21
         .Columns("G").ColumnWidth = 21
         .Columns("H").ColumnWidth = 28
         .Columns("I:L").ColumnWidth = 12.43
    End With
    If Count <= 1 Then
    Numbering = 0
    Count = 2
    End If

    For x = 1 To Count - 1
      XShift = x + XShiftBlank + 1
      y = 1
      STotal = 0
      YShift = YShiftBlank
      PriceTab = 4
      TotalPrice = 0
      For y = 1 To TotalTask
        'title Column
          If dws.Cells(5 + y, 17).Value = "S" Then
            ows.Cells(XShift, YShift + 1).Value = sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value
            YShift = YShift + 1
            STotal = STotal + 1
        'Single Column
          ElseIf dws.Cells(5 + y, 17).Value = 1 Then
            If sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value = "NA" Then
            ElseIf Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) <> 0 Then
              ows.Cells(XShift, YShift + 1).Value = sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value
              ows.Cells(XShift, YShift + DTask + 1).Value = dws.Cells(1 + PriceTab, 6).Value * ows.Cells(XShift, YShift + 1).Value
              TotalPrice = TotalPrice + ows.Cells(XShift, YShift + DTask + 1).Value
            End If
            YShift = YShift + 1
            PriceTab = 1 + PriceTab
        'Double Column
          ElseIf dws.Cells(5 + y, 17).Value = 2 Then
            If sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value = "NA" Then
            ElseIf Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) >= 1 And Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) <= dws.Cells(5 + y, 18).Value Then
              ows.Cells(XShift, YShift + 1).Value = 1
              ows.Cells(XShift, YShift + DTask + 1).Value = dws.Cells(5 + Total, 6).Value * 1
              TotalPrice = TotalPrice + ows.Cells(XShift, YShift + DTask + 1).Value
            ElseIf Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) >= dws.Cells(5 + y, 18).Value + 1 Then
              ows.Cells(XShift, YShift + 1).Value = 1
              ows.Cells(XShift, YShift + 2).Value = Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) - dws.Cells(5 + y, 18).Value
              ows.Cells(XShift, YShift + DTask + 1).Value = dws.Cells(1 + PriceTab, 6).Value * 1
              ows.Cells(XShift, YShift + DTask + 2).Value = dws.Cells(2 + PriceTab, 6).Value * ows.Cells(XShift, YShift + 1).Value
              TotalPrice = TotalPrice + ows.Cells(XShift, YShift + DTask + 1).Value + ows.Cells(XShift, YShift + DTask + 2).Value
            End If
            YShift = YShift + 2
            PriceTab = 2 + PriceTab
        'Triple Column
          ElseIf dws.Cells(5 + y, 17).Value = 3 Then
            If sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value = "NA" Then
            ElseIf Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) >= 1 And Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) <= dws.Cells(5 + y, 18).Value Then
              ows.Cells(XShift, YShift + 1).Value = 1
              ows.Cells(XShift, YShift + DTask + 1).Value = dws.Cells(1 + PriceTab, 6).Value * 1
              TotalPrice = TotalPrice + ows.Cells(XShift, YShift + DTask + 1).Value
            ElseIf Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) >= dws.Cells(5 + y, 18).Value + 1 And Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) <= dws.Cells(5 + y, 19).Value Then
              ows.Cells(XShift, YShift + 1).Value = 1
              ows.Cells(XShift, YShift + 2).Value = Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) - dws.Cells(5 + y, 18).Value
              ows.Cells(XShift, YShift + DTask + 1).Value = dws.Cells(1 + PriceTab, 6).Value * 1
              ows.Cells(XShift, YShift + DTask + 2).Value = dws.Cells(2 + PriceTab, 6).Value * ows.Cells(XShift, YShift + 2).Value
              TotalPrice = TotalPrice + ows.Cells(XShift, YShift + DTask + 1).Value + ows.Cells(XShift, YShift + DTask + 2).Value
            ElseIf Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) >= dws.Cells(5 + y, 19).Value + 1 Then
              ows.Cells(XShift, YShift + 1).Value = 1
              ows.Cells(XShift, YShift + 3).Value = Abs(sws.Range(dws.Cells(5 + y, 20).Value & x + 1).Value) - dws.Cells(5 + y, 18).Value
              ows.Cells(XShift, YShift + DTask + 1).Value = dws.Cells(1 + PriceTab, 6).Value * 1
              ows.Cells(XShift, YShift + DTask + 3).Value = dws.Cells(3 + PriceTab, 6).Value * ows.Cells(XShift, YShift + 3).Value
              TotalPrice = TotalPrice + ows.Cells(XShift, YShift + DTask + 3).Value + ows.Cells(XShift, YShift + DTask + 1).Value
            End If
            YShift = YShift + 3
            PriceTab = 3 + PriceTab
        End If
      Next
      If Numbering <> 0 Then
       ows.Cells(XShift, STotal + DTask * 2 + YShiftBlank + 1).Value = TotalPrice
       ows.Cells(XShift, YShiftBlank).Value = x
      End If
    Next
    scaling = 0
    scaling2 = 0
    For y = STotal + YShiftBlank + 1 To YShift + DTask + 1
      TotalPrice = 0
      For x = XShiftBlank + 2 To Count + XShiftBlank + 1
        TotalPrice = TotalPrice + ows.Cells(x, y).Value
      Next
      ows.Cells(XShift + 2, y).Value = TotalPrice
     
    Next
    For y = YShiftBlank To STotal + YShiftBlank - 1
      Titledes = 1 + Titledes
      ows.Cells(XShiftBlank + 1, YShiftBlank + Titledes).Value = dws.Cells(5 + Titledes, 11).Value
    Next
    headerdes = 0
    For y = STotal + YShiftBlank To YShift - 1
      headerdes = 1 + headerdes
      ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + headerdes).Value = dws.Cells(4 + headerdes, 4).Value
      ows.Cells(XShiftBlank, STotal + YShiftBlank + headerdes).Value = dws.Cells(4 + headerdes, 3).Value
      ows.Cells(XShiftBlank, STotal + YShiftBlank + DTask + headerdes).Value = dws.Cells(4 + headerdes, 3).Value
      ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + DTask + headerdes).Value = dws.Cells(4 + headerdes, 6).Value
    Next
    headerdes = 0
    For y = STotal + YShiftBlank To YShift - 1
      headerdes = 1 + headerdes
      ows2.Cells(13 + headerdes, column1).Value = ows.Cells(XShiftBlank + Count + 2, STotal + YShiftBlank + headerdes).Value
    Next
    headerdes = 0
    For y = STotal + YShiftBlank To YShift
      headerdes = 1 + headerdes
      ows2.Cells(13 + headerdes, column2).Value = ows.Cells(XShiftBlank + Count + 2, DTask + y + 1).Value
    Next
    
    ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + DTask + headerdes).Value = "Subtotal"
    ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + DTask + headerdes + 1).Value = "Remarks"
    ows.Cells(XShiftBlank + Count + 2, 7).Value = "Totals:"
    Call DetailTable(Outputwb, Outputws, 1, XShiftBlank + 1, YShiftBlank, XShiftBlank + Count + 1, STotal + YShiftBlank)
    Call DetailTable(Outputwb, Outputws, 1, XShiftBlank + 1, YShift + DTask, XShiftBlank + Count + 2, YShift + DTask + 2)
    Call DetailTable(Outputwb, Outputws, 2, XShiftBlank, STotal + YShiftBlank + 1, XShiftBlank + Count + 1, YShift)
    Call DetailTable(Outputwb, Outputws, 2, XShiftBlank, YShift + 1, XShiftBlank + Count + 2, YShift + DTask)
    ows.Range(XShiftBlank & ":" & XShiftBlank).Font.Bold = True
    ows.Range(Cells(XShiftBlank + 1, YShiftBlank + 1).Address, Cells(XShiftBlank + 1, YShiftBlank + 4).Address).Font.Bold = True
    ows.Range(XShiftBlank + Count + 2 & ":" & XShiftBlank + Count + 2).Font.Bold = True
    ows.Cells(XShiftBlank, YShiftBlank).Value = Tabletitle

    With ows.Range(Cells(XShiftBlank + 2, YShift + 1).Address, Cells(XShiftBlank + Count + 1, YShift + DTask).Address).Interior 'Blue color Highlighted
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent1
      .TintAndShade = 0.799981688894314
      .PatternTintAndShade = 0
    End With
    ows.Range(Cells(XShiftBlank + 1, YShift + 1).Address, Cells(XShiftBlank + Count + 3, YShift + DTask + 1).Address).NumberFormat = "[$$-en-US]#,##0.00"
    ows.Range(Cells(XShiftBlank + 1, YShiftBlank + 1).Address, Cells(XShiftBlank + Count + 1, YShiftBlank + 2).Address).NumberFormat = "d-mmm"
    With ows.Range(Cells(XShiftBlank, YShiftBlank).Address, Cells(XShiftBlank + Count + 3, YShift + DTask + 2).Address)
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = True
      .ReadingOrder = xlContext
    End With
    ows.Rows(XShiftBlank & ":" & XShiftBlank).RowHeight = 18
    ows.Rows(XShiftBlank + 1 & ":" & XShiftBlank + 1).RowHeight = 75
    ows.Rows(XShiftBlank + Count + 1 & ":" & XShiftBlank + Count + 1).RowHeight = 5.25
    With ows.Range(Cells(XShiftBlank, YShiftBlank).Address, Cells(XShiftBlank, YShiftBlank + 2).Address)
      .HorizontalAlignment = xlLeft
      .VerticalAlignment = xlCenter
      .WrapText = True
      .MergeCells = True
      .ReadingOrder = xlContext
    End With
    With ows.Range(Cells(XShift + 1, STotal + YShiftBlank + 1).Address, Cells(XShift + 1, YShift).Address).Font
      .Color = -10209504
      .TintAndShade = 0
    End With
    With ows.Range(Cells(XShiftBlank, YShiftBlank).Address, Cells(XShiftBlank, YShiftBlank + 1).Address).Font
      .Color = -4165632
      .TintAndShade = 0
      .Size = 14
    End With
    dws.Cells(x + 1, 32).Value = dws.Cells(y, 19).Value

End Function
Function DetailTable(Outputwb, Outputws, tabletype, X1, Y1, X2, Y2)
  'Output Workbook
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    If tabletype = 1 Then
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
    ElseIf tabletype = 2 Then
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlDiagonalDown).LineStyle = xlNone
    ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlDiagonalUp).LineStyle = xlNone
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ows.Range(Cells(X1, Y1).Address, Cells(X2, Y2).Address).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    End If
End Function
Function StoringData(Masterwb, Masterws, Trackingwb, Trackingws, Outputwb, Outputws, shift)
  Dim mwb As Workbook
  Set mwb = Workbooks(Masterwb)
      Taskrowcount = WorksheetFunction.CountA(mwb.Sheets(Masterws).Range("K:K")) - 2
      Workbooks(Outputwb).Sheets(Outputws).Cells.Clear
      For x = 1 To Taskrowcount
        If mwb.Sheets(Masterws).Cells(5 + x, 12 + shift).Value <> "" And mwb.Sheets(Masterws).Cells(5 + x, 20).Value <> "" Then
          Workbooks(Trackingwb).Sheets(Trackingws).Columns(mwb.Sheets(Masterws).Cells(5 + x, 12 + shift).Value & ":" & mwb.Sheets(Masterws).Cells(5 + x, 12 + shift).Value).Copy
          Workbooks(Outputwb).Sheets(Outputws).Columns(mwb.Sheets(Masterws).Cells(5 + x, 20).Value & ":" & mwb.Sheets(Masterws).Cells(5 + x, 20).Value).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End If
      Next
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
Function Thick_OB(Thick_outside_boarder, Outputwb, Outputws)
  'Output Workbook
  Dim owb As Workbook
  Set owb = Workbooks(Outputwb)
  Dim ows As Worksheet
  Set ows = owb.Sheets(Outputws)
  With ows
   With .Range(Thick_outside_boarder).Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlMedium
   End With
   With .Range(Thick_outside_boarder).Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlMedium
   End With
   With .Range(Thick_outside_boarder).Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlMedium
   End With
   With .Range(Thick_outside_boarder).Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlMedium
   End With
 End With
End Function
Function Thin_OB(Thin_outside_boarder, Outputwb, Outputws)
  'Output Workbook
  Dim owb As Workbook
  Set owb = Workbooks(Outputwb)
  Dim ows As Worksheet
  Set ows = owb.Sheets(Outputws)
  With ows
   With .Range(Thin_outside_boarder).Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
   End With
   With .Range(Thin_outside_boarder).Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
   End With
   With .Range(Thin_outside_boarder).Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
   End With
   With .Range(Thin_outside_boarder).Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .ColorIndex = 0
     .TintAndShade = 0
     .Weight = xlThin
   End With
 End With
End Function
Function Check_if_workbook_is_open(Filename)
    Dim wb As Workbook 'to test if workbook is open. No change here
        For Each wb In Workbooks
            If wb.Name = Filename Then
                Workbooks(Filename).Save
                Workbooks(Filename).Close
            End If
        Next
End Function
Function CheckDataSheet(Filename, Sheetname)
    For Each Sheet In Workbooks(Filename).Worksheets
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
Function Copytovalueinvoice(Masterwb, Dataws, Outputwb, Outputws, CopyRange)
  Dim wb As Workbook
  Set wb = Workbooks(Masterwb)
  Dim wbo As Workbook
  Set wbo = Workbooks(Outputwb)
  Dim wsPro As Worksheet
  Set wsPro = wbo.Sheets(Outputws)
  TotalD = WorksheetFunction.CountA(wb.Sheets(Dataws).Range("C:C")) + 2
  With wsPro
    wb.Sheets(Dataws).Range("C4:F" & TotalD).Copy
    .Range(CopyRange).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  End With
End Function
Function CopyfinalValue(Outputwb, Outputws, CopyRange, Pasterange)
  '***Please note that update is required in this function***
  Dim wbo As Workbook
  Set wbo = Workbooks(Outputwb)
  With wsPro
    wbo.Sheets(Outputws).Range(CopyRange).Copy
    wbo.Sheets(Outputws1).Range(Pasterange).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
  End With
End Function
Function pagesetupsetting(Outputwb, Outputws)
    With Workbooks(Outputwb).Sheets(Outputws).PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "Page &P of &N"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(1.8)
        .RightMargin = Application.InchesToPoints(1.8)
        .TopMargin = Application.InchesToPoints(1.9)
        .BottomMargin = Application.InchesToPoints(1.9)
        .HeaderMargin = Application.InchesToPoints(0.8)
        .FooterMargin = Application.InchesToPoints(0.8)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With

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
Function findCellColumnNo(Masterwb, Dataws, Outputwb, Outputws, rowno)
    Dim mwb As Workbook
    Set mwb = Workbooks(Masterwb)
    Dim dws As Worksheet
    Set dws = mwb.Sheets(Dataws)
  'Output Workbook is represented as owb
    Dim owb As Workbook
    Set owb = Workbooks(Outputwb)
  
  'Output Worksheet is represented as ow
    Dim ows As Worksheet
    Set ows = owb.Sheets(Outputws)
    Dim strSearch As String
    Dim aCell As Range

    strSearch = "Invoice No"

    Set aCell = ows.Cells.Find(What:=strSearch, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCell Is Nothing Then
        dws.Cells(8, rowno).Value = aCell.Column
    End If
End Function

Function titletext(Outputwb, Outputws, Invoiceno, Region)
  Dim owb As Workbook
  Set owb = Workbooks(Outputwb)
  
  'Output Worksheet is represented as ow
  Dim ows As Worksheet
  Set ows = owb.Sheets(Outputws)
  ows.Range("C1").Value = "Invoice:"
  ows.Range("D1").Value = Invoiceno
  ows.Range("H1").Value = "Region:"
  ows.Range("I1").Value = Region
  With ows.Range("C1:I1").Font
        .Color = -10477568
        .TintAndShade = 0
        .FontStyle = "Bold"
  End With
      With ows.Range("C1,H1")
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
    With ows.Range("D1,I1")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ows.Columns("C:C").EntireColumn.AutoFit
    ows.Columns("H:H").ColumnWidth = 7
End Function
