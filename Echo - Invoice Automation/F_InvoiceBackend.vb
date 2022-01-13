'Version: 1.00
'Ika - Invoice Backend
'For Echobroadband
'User: Ika
'By Farhat Abbas & Ika




Sub main()
    'Closingforms
     Application.DisplayAlerts = False
   
    'Declaring Variables
     filedate = Format(Date, "ddmmyyyy")
     OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
     Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Value
     region = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("I11").Value
     Call Invoice_Details(Filename, region)
     Application.DisplayAlerts = True
      
  End Sub
   
  Function Invoice_Details(Filename, region)
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
   
    'Setting Sheetnames
     Sheetname1 = "Node Split"
     Sheetname2 = "Commercial_ Expense Design"
     Sheetname3 = "Asbuilt Coax & Fiber"
     Sheetname4 = "SFU&MDU Design"
     Sheetname5 = "ME Design,Asbuit&Desktop Srvy"
     Sheetname6 = "DataProcess"
   
    'Creating Output Sheet1
     Outputsheet1 = "Invoice Details" 'wo1
     Outputsheet2 = "Invoice Summary" 'wo2
   
    'Input Workbook is represented as wb
     Dim wb As Workbook
     Set wb = Workbooks(Filename)
   
    'input Worksheet is represented as ws
     Dim ws1 As Worksheet
     Dim ws2 As Worksheet
     Dim ws3 As Worksheet
     Dim ws4 As Worksheet
     Dim ws5 As Worksheet
     Dim wsPro As Worksheet
     Set ws1 = wb.Sheets(Sheetname1)
     Set ws2 = wb.Sheets(Sheetname2)
     Set ws3 = wb.Sheets(Sheetname3)
     Set ws4 = wb.Sheets(Sheetname4)
     Set ws5 = wb.Sheets(Sheetname5)
     Set wsPro = wb.Sheets(Sheetname6)
   
   
    'Output Workbook is represented as wbo
     Dim wbo As Workbook
     Set wbo = Workbooks(OutputFileName)
   
    'Output Worksheet is represented as wo
     Dim wo1 As Worksheet
     Set wo1 = wbo.Sheets(Outputsheet1) 'detail
     Dim wo2 As Worksheet
     Set wo2 = wbo.Sheets(Outputsheet2) 'summary
     
    'Filtering the data Sheet 1 'NODE SPLIT
     ws1.Range("BL:BL").AutoFilter Field:=64, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws1.Range("D:D").AutoFilter Field:=4, Criteria1:=region  'Region
     ws1.Range("BM:BM").AutoFilter Field:=65, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws1.Range("BK:BK").AutoFilter Field:=63, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
     Count1 = ws1.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   
    'Filtering the data Sheet 4 'CED
     ws2.Range("AY:AY").AutoFilter Field:=51, Operator:=xlFilterValues, Criteria1:="<>="   'Blank Dilvery date
     ws2.Range("D:D").AutoFilter Field:=4, Criteria1:=region   'Region
     ws2.Range("AZ:AZ").AutoFilter Field:=52, Criteria1:="=", Operator:=xlFilterValues  'Invoice colum
     ws2.Range("AX:AX").AutoFilter Field:=50, Criteria1:=Array("Completed"), Operator:=xlFilterValues   'QB Dilevery
    'Count
     Count2 = ws2.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   
    'Filtering the data Sheet 5
     ws3.Range("AV:AV").AutoFilter Field:=48, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws3.Range("D:D").AutoFilter Field:=4, Criteria1:=region  'Region
     ws3.Range("AW:AW").AutoFilter Field:=49, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws3.Range("AU:AU").AutoFilter Field:=47, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
    'Count
     Count3 = ws3.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
     
    'Filtering the data Sheet 3 'SFU
     ws4.Range("BD:BD").AutoFilter Field:=56, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws4.Range("D:D").AutoFilter Field:=4, Criteria1:=region  'Region
     ws4.Range("BE:BE").AutoFilter Field:=57, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws4.Range("BC:BC").AutoFilter Field:=55, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
    'Count
     Count4 = ws4.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
     
     'Filtering the data Sheet 2 'ME
     ws5.Range("AT:AT").AutoFilter Field:=46, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws5.Range("D:D").AutoFilter Field:=4, Criteria1:=region 'Region
     ws5.Range("AU:AU").AutoFilter Field:=47, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws5.Range("AS:AS").AutoFilter Field:=45, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
    
    'Count
     Count5 = ws5.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
     
    'Format
      wo1.Columns("C:D").NumberFormat = "[$-en-US]d-mmm;@"
      With wo1.Columns("A:BT")
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = True
      End With
    'Sheet1 Calculation and pasting
      Call Copytodatasheet1'Update for d
      Call ProcessingValues(Count1, 4)'Update for d
      Call CalculateTotal(5, 4 + Count1, Count1)'Update for d
      Call detailtemplate(0, region, 1) 'Update for d
      Call TableArrangment("G3:BR3") 'Update for d
      Call TableArrangment("C4:BT4") 'Update for d
      If Count1 = 0 Then
          Count1 = 1
      End If
      Call TableArrangmentData("B5:BT" & Count1 + 4)'Update for d
      Call thickline1("AM3:AM" & 5 + Count1)'Update for d
      Call boardermoney(3, 5 + Count1)'Update for d
      wo1.Range("G" & 5 + Count1 & ":AL" & 5 + Count1).Copy'Update for d
      wo2.Range("F14:F45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True'Update for d
      wo1.Range("AM" & 5 + Count1 & ":BR" & 5 + Count1).Copy'Update for d
      wo2.Range("G14:G45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True'Update for d
      wo2.Range("F46").Value = Application.WorksheetFunction.Sum(wo2.Range("G14:G45"))'Update for d
  
    'Sheet 2 & Sheet 3
      Call Copytodatasheet2
      Call ProcessingValues(Count2, 10 + Count1)
      Call Copytodatasheet3
      Call ProcessingValues(Count3, 10 + Count1 + Count2)
      Call CalculateTotal(10 + Count1, 10 + Count1 + Count2 + Count3, Count2 + Count3)
      Call detailtemplate(6 + Count1, region, 2)
      Call TableArrangment("G" & 9 + Count1 & ":BR" & 9 + Count1 + Count2 + Count3)
      Call TableArrangment("C" & 10 + Count1 & ":BT" & 10 + Count1 + Count2 + Count3)
      If Count2 + Count3 = 0 Then
          Count3 = 1
      End If
      Call TableArrangmentData("B" & 11 + Count1 & ":BT" & 10 + Count1 + Count2 + Count3)
      Call thickline1("AM" & 9 + Count1 & ":AM" & 11 + Count1 + Count2 + Count3)
      Call boardermoney(9 + Count1, 11 + Count1 + Count2 + Count3)
      wo1.Range("G" & 11 + Count1 + Count2 + Count3 & ":AL" & 11 + Count1 + Count2 + Count3).Copy
      wo2.Range("H14:H45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
      wo1.Range("AM" & 11 + Count1 + Count2 + Count3 & ":BR" & 11 + Count1 + Count2 + Count3).Copy
      wo2.Range("I14:I45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
      wo2.Range("H46").Value = Application.WorksheetFunction.Sum(wo2.Range("I14:I45"))
    'Sheet 4
      Call Copytodatasheet4
      Call ProcessingValues(Count4, 16 + Count1 + Count2 + Count3)
      Call detailtemplate(12 + Count1 + Count2 + Count3, region, 3)
      Call CalculateTotal(17 + Count1 + Count2 + Count3, 16 + Count1 + Count2 + Count3 + Count4, Count4)
      Call TableArrangment("G" & 15 + Count1 + Count2 + Count3 & ":BR" & 15 + Count1 + Count2 + Count3 + Count4)
      Call TableArrangment("C" & 16 + Count1 + Count2 + Count3 & ":BT" & 16 + Count1 + Count2 + Count3 + Count4)
      If Count4 = 0 Then
          Count4 = 1
      End If
      Call TableArrangmentData("B" & 17 + Count1 + Count2 + Count3 & ":BT" & 16 + Count1 + Count2 + Count3 + Count4)
      Call thickline1("AM" & 15 + Count1 + Count2 + Count3 & ":AM" & 17 + Count1 + Count2 + Count3 + Count4)
      Call boardermoney(15 + Count1 + Count2 + Count3, 18 + Count1 + Count2 + Count3 + Count4)
      wo1.Range("G" & 17 + Count1 + Count2 + Count3 + Count4 & ":AL" & 17 + Count1 + Count2 + Count3 + Count4).Copy
      wo2.Range("J14:J45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
      wo1.Range("AM" & 17 + Count1 + Count2 + Count3 + Count4 & ":BR" & 17 + Count1 + Count2 + Count3 + Count4).Copy
      wo2.Range("K14:K45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
      wo2.Range("J46").Value = Application.WorksheetFunction.Sum(wo2.Range("K14:K45"))
  
    'Sheet 5
      Call Copytodatasheet5
      Call ProcessingValues(Count5, 22 + Count1 + Count2 + Count3 + Count4)
      Call detailtemplate(Count1 + Count2 + Count3 + Count4 + 18, region, 4)
      Call CalculateTotal(23 + Count1 + Count2 + Count3 + Count4, 23 + Count1 + Count2 + Count3 + Count5, Count5)
      Call TableArrangment("G" & 21 + Count1 + Count2 + Count3 + Count4 & ":BR" & 21 + Count1 + Count2 + Count3 + Count4 + Count5)
      Call TableArrangment("C" & 22 + Count1 + Count2 + Count3 + Count4 & ":BT" & 22 + Count1 + Count2 + Count3 + Count4 + Count5)
      If Count5 = 0 Then
          Count5 = 1
      End If
      Call TableArrangmentData("B" & 23 + Count1 + Count2 + Count3 + Count4 & ":BT" & 22 + Count1 + Count2 + Count3 + Count4 + Count5)
      Call thickline1("AM" & Count1 + Count2 + Count3 + Count4 + 21 & ":AM" & 24 + Count1 + Count2 + Count3 + Count5)
      Call boardermoney(Count1 + Count2 + Count3 + Count4 + 21, 24 + Count1 + Count2 + Count3 + Count5)
      wo1.Range("G" & 24 + Count1 + Count2 + Count3 + Count5 & ":AL" & 24 + Count1 + Count2 + Count3 + Count5).Copy
      wo2.Range("L14:L45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
      wo1.Range("AM" & 24 + Count1 + Count2 + Count3 + Count5 & ":BR" & 24 + Count1 + Count2 + Count3 + Count5).Copy
      wo2.Range("M14:M45").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
      wo2.Range("L46").Value = Application.WorksheetFunction.Sum(wo2.Range("M14:M45"))

    'Format
      wo2.Range("L47").Value = wo2.Range("F46").Value + wo2.Range("H46").Value + wo2.Range("J46").Value + wo2.Range("L46").Value
      wo1.Range("AM:BS").NumberFormat = "[$$-en-US]#,##0.00"
      wo1.Columns("A").ColumnWidth = 8
      wo1.Columns("B").ColumnWidth = 8
      wo1.Columns("C:D").ColumnWidth = 13.5
      wo1.Columns("E:F").ColumnWidth = 25.5
      wo1.Columns("G:AL").ColumnWidth = 13.5
      wo1.Columns("AM:BS").ColumnWidth = 11.5
      wo1.Columns("BT").ColumnWidth = 46.8
      wo1.Rows("1:1").RowHeight = 15.75
      wo1.Rows("2:2").RowHeight = 10
      wo1.Rows("3:3").RowHeight = 18
      wo1.Rows(8 + Count1 & ":" & 8 + Count1).RowHeight = 10
      wo1.Rows(9 + Count1 & ":" & 9 + Count1).RowHeight = 18
      wo1.Rows(14 + Count1 + Count2 + Count3 & ":" & 14 + Count1 + Count2 + Count3).RowHeight = 10
      wo1.Rows(15 + Count1 + Count2 + Count3 & ":" & 15 + Count1 + Count2 + Count3).RowHeight = 18
      wo1.Rows(Count1 + Count2 + Count3 + Count4 + 20 & ":" & Count1 + Count2 + Count3 + Count4 + 20).RowHeight = 10
      wo1.Rows(Count1 + Count2 + Count3 + Count4 + 21 & ":" & Count1 + Count2 + Count3 + Count4 + 21).RowHeight = 18
  End Function
  Function boardermoney(firstrow, lastrow)
      filedate = Format(Date, "ddmmyyyy")
      OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
      Outputsheetname1 = "Invoice Details"
  
   'Output Workbook is represented as wbo
      Dim wbo As Workbook
      Set wbo = Workbooks(OutputFileName)
  
   'Output Worksheet is represented as wo
      Dim wo1 As Worksheet
      Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
  
   'Invoice detail template
    With wo1
      .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlDiagonalDown).LineStyle = xlNone
      .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlDiagonalUp).LineStyle = xlNone
      With .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlEdgeTop)
          .LineStyle = xlDouble
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThick
      End With
      With .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range("AN" & lastrow & ":BS" & lastrow).Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      .Range("AN" & lastrow & ":BR" & lastrow).Borders(xlInsideHorizontal).LineStyle = xlNone 'money row 'change for d
      With .Range("AM" & firstrow & ":BR" & lastrow).Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent1
      .TintAndShade = 0.799981688894314
      .PatternTintAndShade = 0
      End With
      With .Range("BS" & firstrow + 1 & ":BS" & lastrow).Interior 'highlight'change for d
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorAccent1
      .TintAndShade = 0.799981688894314
      .PatternTintAndShade = 0
      End With
      .Range("AM" & firstrow & ":BT" & firstrow + 1).Font.Bold = True 'bold'change for d
      .Range("AM" & lastrow & ":BT" & lastrow).Font.Bold = True 'bold'change for d
   End With
  End Function
  Function thickline1(thickline)
    'Data
      filedate = Format(Date, "ddmmyyyy")
      OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
      Outputsheetname1 = "Invoice Details"
  
   'Output Workbook is represented as wbo
      Dim wbo As Workbook
      Set wbo = Workbooks(OutputFileName)
  
   'Output Worksheet is represented as wo
      Dim wo1 As Worksheet
      Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
  
   'Invoice detail template
    With wo1
      .Range(thickline).Borders(xlDiagonalDown).LineStyle = xlNone
      .Range(thickline).Borders(xlDiagonalUp).LineStyle = xlNone
      With .Range(thickline).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThick
      End With
      With .Range(thickline).Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(thickline).Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(thickline).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      .Range(thickline).Borders(xlInsideVertical).LineStyle = xlNone
      End With
  End Function
  
  Function detailtemplate(x, region, y)'change for d
    'Setting Sheetnames
     Sheetname1 = "Node Split Design & Asbuilt"
     Sheetname2 = "Coax Design & Asbuild"
     Sheetname3 = "SFU & MDU"
     Sheetname4 = "Fiber Design & Asbuild"
    'Data
     filedate = Format(Date, "ddmmyyyy")
     OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
     Outputsheetname1 = "Invoice Details"
     Outputsheetname2 = "Invoice Summary"
   
    'Output Workbook is represented as wbo
     Dim wbo As Workbook
     Set wbo = Workbooks(OutputFileName)
   
    'Output Worksheet is represented as wo
     Dim wo1 As Worksheet
     Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
    'Output Worksheet is represented as wo
     Dim wo2 As Worksheet
     Set wo2 = wbo.Sheets(Outputsheetname2) 'detail
   
    'Invoice detail template
     With wo1
      If y = 1 Then
        .Range("C" & x + 1).FormulaR1C1 = "Invoice:"
        .Range("D" & x + 1).FormulaR1C1 = wo2.Range("L3").Value
        .Range("G" & x + 1).FormulaR1C1 = "Region:"
        .Range("H" & x + 1).FormulaR1C1 = region
        .Range("C" & x + 3).FormulaR1C1 = Sheetname1
      ElseIf y = 2 Then
        .Range("C" & x + 3).FormulaR1C1 = Sheetname2
      ElseIf y = 3 Then
        .Range("C" & x + 3).FormulaR1C1 = Sheetname3
      ElseIf y = 4 Then
        .Range("C" & x + 3).FormulaR1C1 = Sheetname4
      End If
      .Range("C" & x + 3 & ":E" & x + 3).Merge
      With .Range("C" & x + 3 & ", D1, H1")
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlCenter
          .WrapText = True
          .Orientation = 0
          .AddIndent = False
          .IndentLevel = 0
          .ShrinkToFit = False
          .ReadingOrder = xlContext
      End With
      With .Range("C1, G1")
          .HorizontalAlignment = xlRight
          .VerticalAlignment = xlCenter
          .WrapText = True
          .Orientation = 0
          .AddIndent = False
          .IndentLevel = 0
          .ShrinkToFit = False
          .ReadingOrder = xlContext
      End With
      With .Range("C1:H1").Font
          .Color = -10477568
          .TintAndShade = 0
          .Size = 12
          .Bold = True
      End With
      With .Range("C" & x + 3).Font
          .ThemeColor = xlThemeColorAccent2
          .TintAndShade = -0.499984740745262
          .Size = 14
          .Bold = True
      End With
      .Range("D1:E1").Merge
      .Range("G" & x + 3).FormulaR1C1 = "D1"
      .Range("H" & x + 3).FormulaR1C1 = "D2"
      .Range("I" & x + 3).FormulaR1C1 = "D3"
      .Range("J" & x + 3).FormulaR1C1 = "D4"
      .Range("K" & x + 3).FormulaR1C1 = "D5"
      .Range("L" & x + 3).FormulaR1C1 = "D6"
      .Range("M" & x + 3).FormulaR1C1 = "D7"
      .Range("N" & x + 3).FormulaR1C1 = "D8"
      .Range("O" & x + 3).FormulaR1C1 = "D9"
      .Range("P" & x + 3).FormulaR1C1 = "D10"
      .Range("Q" & x + 3).FormulaR1C1 = "D11"
      .Range("R" & x + 3).FormulaR1C1 = "D12"
      .Range("S" & x + 3).FormulaR1C1 = "D13"
      .Range("T" & x + 3).FormulaR1C1 = "D14"
      .Range("U" & x + 3).FormulaR1C1 = "D15"
      .Range("V" & x + 3).FormulaR1C1 = "D16"
      .Range("W" & x + 3).FormulaR1C1 = "D17"
      .Range("X" & x + 3).FormulaR1C1 = "D18"
      .Range("Y" & x + 3).FormulaR1C1 = "D19"
      .Range("Z" & x + 3).FormulaR1C1 = "D24"
      .Range("AA" & x + 3).FormulaR1C1 = "D25"
      .Range("AB" & x + 3).FormulaR1C1 = "D26"
      .Range("AC" & x + 3).FormulaR1C1 = "D27"
      .Range("AD" & x + 3).FormulaR1C1 = "D29"
      .Range("AE" & x + 3).FormulaR1C1 = "D31"
      .Range("AF" & x + 3).FormulaR1C1 = "D32"
      .Range("AG" & x + 3).FormulaR1C1 = "D33"
      .Range("AH" & x + 3).FormulaR1C1 = "D34"
      .Range("AI" & x + 3).FormulaR1C1 = "D35"
      .Range("AJ" & x + 3).FormulaR1C1 = "D36"
      .Range("AK" & x + 3).FormulaR1C1 = "D37"
      .Range("AL" & x + 3).FormulaR1C1 = "D38"
      .Range("AM" & x + 3).FormulaR1C1 = "D1"
      .Range("AN" & x + 3).FormulaR1C1 = "D2"
      .Range("AO" & x + 3).FormulaR1C1 = "D3"
      .Range("AP" & x + 3).FormulaR1C1 = "D4"
      .Range("AQ" & x + 3).FormulaR1C1 = "D5"
      .Range("AR" & x + 3).FormulaR1C1 = "D6"
      .Range("AS" & x + 3).FormulaR1C1 = "D7"
      .Range("AT" & x + 3).FormulaR1C1 = "D8"
      .Range("AU" & x + 3).FormulaR1C1 = "D9"
      .Range("AV" & x + 3).FormulaR1C1 = "D10"
      .Range("AW" & x + 3).FormulaR1C1 = "D11"
      .Range("AX" & x + 3).FormulaR1C1 = "D12"
      .Range("AY" & x + 3).FormulaR1C1 = "D13"
      .Range("AZ" & x + 3).FormulaR1C1 = "D14"
      .Range("BA" & x + 3).FormulaR1C1 = "D15"
      .Range("BB" & x + 3).FormulaR1C1 = "D16"
      .Range("BC" & x + 3).FormulaR1C1 = "D17"
      .Range("BD" & x + 3).FormulaR1C1 = "D18"
      .Range("BE" & x + 3).FormulaR1C1 = "D19"
      .Range("BF" & x + 3).FormulaR1C1 = "D24"
      .Range("BG" & x + 3).FormulaR1C1 = "D25"
      .Range("BH" & x + 3).FormulaR1C1 = "D26"
      .Range("BI" & x + 3).FormulaR1C1 = "D27"
      .Range("BJ" & x + 3).FormulaR1C1 = "D29"
      .Range("BK" & x + 3).FormulaR1C1 = "D31"
      .Range("BL" & x + 3).FormulaR1C1 = "D32"
      .Range("BM" & x + 3).FormulaR1C1 = "D33"
      .Range("BN" & x + 3).FormulaR1C1 = "D34"
      .Range("BO" & x + 3).FormulaR1C1 = "D35"
      .Range("BP" & x + 3).FormulaR1C1 = "D36"
      .Range("BQ" & x + 3).FormulaR1C1 = "D37"
      .Range("BR" & x + 3).FormulaR1C1 = "D38"
      .Range("C" & x + 4 & ":F" & x + 4).Font.Bold = True
      .Range("C" & x + 4).FormulaR1C1 = "Date Created"
      .Range("D" & x + 4).FormulaR1C1 = "Delivery Date"
      .Range("E" & x + 4).FormulaR1C1 = "Job Number"
      .Range("F" & x + 4).FormulaR1C1 = "Type"
      .Range("G" & x + 4).FormulaR1C1 = "Route Drafting <2,000'"
      .Range("H" & x + 4).FormulaR1C1 = "Route Drafting >2,000'"
      .Range("I" & x + 4).FormulaR1C1 = "Large Route Drafting >20,000'"
      .Range("J" & x + 4).FormulaR1C1 = "Coax Design <2,000'"
      .Range("K" & x + 4).FormulaR1C1 = "Coax Design >2,000'"
      .Range("L" & x + 4).FormulaR1C1 = "Large Coax Design >20,000'"
      .Range("M" & x + 4).FormulaR1C1 = "Fiber Design <2,000'"
      .Range("N" & x + 4).FormulaR1C1 = "Fiber Design >2,000''"
      .Range("O" & x + 4).FormulaR1C1 = "Large Fiber Design >20,000''"
      .Range("P" & x + 4).FormulaR1C1 = "Parcel Creation AutoCAD'"
      .Range("Q" & x + 4).FormulaR1C1 = "Parcel Creation Other'"
      .Range("R" & x + 4).FormulaR1C1 = "Asbuild <2,000''"
      .Range("S" & x + 4).FormulaR1C1 = "Asbuild >2,000''"
      .Range("T" & x + 4).FormulaR1C1 = "Large Asbuilt >20,000''"
      .Range("U" & x + 4).FormulaR1C1 = "Node Split Asbuild'"
      .Range("V" & x + 4).FormulaR1C1 = "Coax MDU Up to 25'"
      .Range("W" & x + 4).FormulaR1C1 = "Coax MDU >25'"
      .Range("X" & x + 4).FormulaR1C1 = "Large Coax MDU >20,000''"
      .Range("Y" & x + 4).FormulaR1C1 = "MDU Detail Insert'"
      .Range("Z" & x + 4).FormulaR1C1 = "Forced Relo 2,000''"
      .Range("AA" & x + 4).FormulaR1C1 = "Forced Relo >2,000''"
      .Range("AB" & x + 4).FormulaR1C1 = "Node Split - Fiber and Coax'"
      .Range("AC" & x + 4).FormulaR1C1 = "Node Split Load Balance'"
      .Range("AD" & x + 4).FormulaR1C1 = "Plant Map Update'"
      .Range("AE" & x + 4).FormulaR1C1 = "Fiber trunk tree-modify'"
      .Range("AF" & x + 4).FormulaR1C1 = "Fiber Trunk Tree New'"
      .Range("AG" & x + 4).FormulaR1C1 = "Wavelength Res Req'"
      .Range("AH" & x + 4).FormulaR1C1 = "Ladder Report'"
      .Range("AI" & x + 4).FormulaR1C1 = "Fiber Route Trace'"
      .Range("AJ" & x + 4).FormulaR1C1 = "Splice Updates'"
      .Range("AK" & x + 4).FormulaR1C1 = "Splice Addition'"
      .Range("AL" & x + 4).FormulaR1C1 = "Misc Hourly Work '"
      .Range("AM" & x + 4).FormulaR1C1 = wo2.Range("E" & 14).Value
      .Range("AN" & x + 4).FormulaR1C1 = wo2.Range("E" & 15).Value
      .Range("AO" & x + 4).FormulaR1C1 = wo2.Range("E" & 16).Value
      .Range("AP" & x + 4).FormulaR1C1 = wo2.Range("E" & 17).Value
      .Range("AQ" & x + 4).FormulaR1C1 = wo2.Range("E" & 18).Value
      .Range("AR" & x + 4).FormulaR1C1 = wo2.Range("E" & 19).Value
      .Range("AS" & x + 4).FormulaR1C1 = wo2.Range("E" & 20).Value
      .Range("AT" & x + 4).FormulaR1C1 = wo2.Range("E" & 21).Value
      .Range("AU" & x + 4).FormulaR1C1 = wo2.Range("E" & 22).Value
      .Range("AV" & x + 4).FormulaR1C1 = wo2.Range("E" & 23).Value
      .Range("AW" & x + 4).FormulaR1C1 = wo2.Range("E" & 24).Value
      .Range("AX" & x + 4).FormulaR1C1 = wo2.Range("E" & 25).Value
      .Range("AY" & x + 4).FormulaR1C1 = wo2.Range("E" & 26).Value
      .Range("AZ" & x + 4).FormulaR1C1 = wo2.Range("E" & 27).Value
      .Range("BA" & x + 4).FormulaR1C1 = wo2.Range("E" & 28).Value
      .Range("BB" & x + 4).FormulaR1C1 = wo2.Range("E" & 29).Value
      .Range("BC" & x + 4).FormulaR1C1 = wo2.Range("E" & 30).Value
      .Range("BD" & x + 4).FormulaR1C1 = wo2.Range("E" & 31).Value
      .Range("BE" & x + 4).FormulaR1C1 = wo2.Range("E" & 32).Value
      .Range("BF" & x + 4).FormulaR1C1 = wo2.Range("E" & 33).Value
      .Range("BG" & x + 4).FormulaR1C1 = wo2.Range("E" & 34).Value
      .Range("BH" & x + 4).FormulaR1C1 = wo2.Range("E" & 35).Value
      .Range("BI" & x + 4).FormulaR1C1 = wo2.Range("E" & 36).Value
      .Range("BJ" & x + 4).FormulaR1C1 = wo2.Range("E" & 37).Value
      .Range("BK" & x + 4).FormulaR1C1 = wo2.Range("E" & 38).Value
      .Range("BL" & x + 4).FormulaR1C1 = wo2.Range("E" & 39).Value
      .Range("BM" & x + 4).FormulaR1C1 = wo2.Range("E" & 40).Value
      .Range("BN" & x + 4).FormulaR1C1 = wo2.Range("E" & 41).Value
      .Range("BO" & x + 4).FormulaR1C1 = wo2.Range("E" & 42).Value
      .Range("BP" & x + 4).FormulaR1C1 = wo2.Range("E" & 43).Value
      .Range("BQ" & x + 4).FormulaR1C1 = wo2.Range("E" & 44).Value
      .Range("BR" & x + 4).FormulaR1C1 = wo2.Range("E" & 45).Value
      .Range("BS" & x + 4).FormulaR1C1 = "Subtotal"
      .Range("BT" & x + 4).FormulaR1C1 = "Remark"
     End With
  End Function
  Function CalculateTotal(StartingValue, EndValue, Count)
      If Count = 0 Then
          EndingValue = EndValue + 1
      Else
          EndingValue = EndValue
      End If
  
      'Processing data
      filedate = Format(Date, "ddmmyyyy")
      OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
      Sheetname = "Invoice Details"
      Sheetname1 = "Invoice Summary"
      Dim wb As Workbook
      Set wb = Workbooks(OutputFileName)
      Dim wsPro As Worksheet
      Set wsPro = wb.Sheets(Sheetname)
      Dim wo1 As Worksheet
      Set wo1 = wb.Sheets(Sheetname1)
      
      'Calculating Total
      With wsPro
        .Range("G" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("G" & StartingValue & ":G" & EndingValue))
        .Range("H" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("H" & StartingValue & ":H" & EndingValue))
        .Range("I" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("I" & StartingValue & ":I" & EndingValue))
        .Range("J" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("J" & StartingValue & ":J" & EndingValue))
        .Range("K" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("K" & StartingValue & ":K" & EndingValue))
        .Range("L" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("L" & StartingValue & ":L" & EndingValue))
        .Range("M" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("M" & StartingValue & ":M" & EndingValue))
        .Range("N" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("N" & StartingValue & ":N" & EndingValue))
        .Range("O" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("O" & StartingValue & ":O" & EndingValue))
        .Range("P" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("P" & StartingValue & ":P" & EndingValue))
        .Range("Q" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("Q" & StartingValue & ":Q" & EndingValue))
        .Range("R" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("R" & StartingValue & ":R" & EndingValue))
        .Range("S" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("S" & StartingValue & ":S" & EndingValue))
        .Range("T" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("T" & StartingValue & ":T" & EndingValue))
        .Range("U" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("U" & StartingValue & ":U" & EndingValue))
        .Range("V" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("V" & StartingValue & ":V" & EndingValue))
        .Range("W" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("W" & StartingValue & ":W" & EndingValue))
        .Range("X" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("X" & StartingValue & ":X" & EndingValue))
        .Range("Y" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("Y" & StartingValue & ":Y" & EndingValue))
        .Range("Z" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("Z" & StartingValue & ":Z" & EndingValue))
        .Range("AA" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AA" & StartingValue & ":AA" & EndingValue))
        .Range("AB" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AB" & StartingValue & ":AB" & EndingValue))
        .Range("AC" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AC" & StartingValue & ":AC" & EndingValue))
        .Range("AD" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AD" & StartingValue & ":AD" & EndingValue))
        .Range("AE" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AE" & StartingValue & ":AE" & EndingValue))
        .Range("AF" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AF" & StartingValue & ":AF" & EndingValue))
        .Range("AG" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AG" & StartingValue & ":AG" & EndingValue))
        .Range("AH" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AH" & StartingValue & ":AH" & EndingValue))
        .Range("AI" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AI" & StartingValue & ":AI" & EndingValue))
        .Range("AJ" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AJ" & StartingValue & ":AJ" & EndingValue))
        .Range("AK" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AK" & StartingValue & ":AK" & EndingValue))
        .Range("AL" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AL" & StartingValue & ":AL" & EndingValue))
        .Range("AM" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AM" & StartingValue & ":AM" & EndingValue))
        .Range("AN" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AN" & StartingValue & ":AN" & EndingValue))
        .Range("AO" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AO" & StartingValue & ":AO" & EndingValue))
        .Range("AP" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AP" & StartingValue & ":AP" & EndingValue))
        .Range("AQ" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AQ" & StartingValue & ":AQ" & EndingValue))
        .Range("AR" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AR" & StartingValue & ":AR" & EndingValue))
        .Range("AS" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AS" & StartingValue & ":AS" & EndingValue))
        .Range("AT" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AT" & StartingValue & ":AT" & EndingValue))
        .Range("AU" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AU" & StartingValue & ":AU" & EndingValue))
        .Range("AV" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AV" & StartingValue & ":AV" & EndingValue))
        .Range("AW" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AW" & StartingValue & ":AW" & EndingValue))
        .Range("AX" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AX" & StartingValue & ":AX" & EndingValue))
        .Range("AY" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AY" & StartingValue & ":AY" & EndingValue))
        .Range("AZ" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("AZ" & StartingValue & ":AZ" & EndingValue))
        .Range("BA" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BA" & StartingValue & ":BA" & EndingValue))
        .Range("BB" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BB" & StartingValue & ":BB" & EndingValue))
        .Range("BC" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BC" & StartingValue & ":BC" & EndingValue))
        .Range("BD" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BD" & StartingValue & ":BD" & EndingValue))
        .Range("BE" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BE" & StartingValue & ":BE" & EndingValue))
        .Range("BF" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BF" & StartingValue & ":BF" & EndingValue))
        .Range("BG" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BG" & StartingValue & ":BG" & EndingValue))
        .Range("BH" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BH" & StartingValue & ":BH" & EndingValue))
        .Range("BI" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BI" & StartingValue & ":BI" & EndingValue))
        .Range("BJ" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BJ" & StartingValue & ":BJ" & EndingValue))
        .Range("BK" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BK" & StartingValue & ":BK" & EndingValue))
        .Range("BL" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BL" & StartingValue & ":BL" & EndingValue))
        .Range("BM" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BM" & StartingValue & ":BM" & EndingValue))
        .Range("BN" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BN" & StartingValue & ":BN" & EndingValue))
        .Range("BO" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BO" & StartingValue & ":BO" & EndingValue))
        .Range("BP" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BP" & StartingValue & ":BP" & EndingValue))
        .Range("BQ" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BQ" & StartingValue & ":BQ" & EndingValue))
        .Range("BR" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BR" & StartingValue & ":BR" & EndingValue))
        .Range("BS" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BS" & StartingValue & ":BS" & EndingValue)) 'add for d
        .Range(EndingValue + 1 & ":" & EndingValue + 1).Font.Bold = True
        With .Range("G" & EndingValue + 1 & ":AL" & EndingValue + 1).Font'change for d
          .Color = -10477568
          .TintAndShade = 0
          .Size = 11
          .Bold = True
        End With
        For x = 1 To Count
        .Cells(EndingValue - Count + x, 2).Value = x
        .Rows(EndingValue - Count + x & ":" & EndingValue - Count + x).RowHeight = 18 'change for d
        Next
        .Rows(EndingValue - Count - 2 & ":" & EndingValue - Count - 2).RowHeight = 34.5 'change for d
        
     End With
  End Function
  Function TableArrangment(tablesize)
          filedate = Format(Date, "ddmmyyyy")
          OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
          Sheet = "Invoice Details"
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
      End With
  End Function
  Function TableArrangmentData(tablesize)
      filedate = Format(Date, "ddmmyyyy")
      OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
      Sheet = "Invoice Details"
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeBottom)
          .LineStyle = xlDouble
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThick
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
      End With
  End Function
  Function TableArrangmentDataCalcu(tablesize)
      filedate = Format(Date, "ddmmyyyy")
      OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
      Sheet = "Invoice Details"
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeLeft)
        .LineStyle = xlThick
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
      End With
      With Workbooks(OutputFileName).Worksheets(Sheet).Range(tablesize).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
      End With
  End Function
  Function Copytodatasheet1()
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Value
    Sheetname = "Node Split"
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    With wsPro
          wb.Sheets(Sheetname).Columns("A:A").Copy 'Date created
          .Columns("A:A").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("BL:BL").Copy 'delivery date
      .Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("G:G").Copy 'Job number
      .Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("E:E").Copy 'Type
      .Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("N:N").Copy 'D1-D3
      .Columns("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("F:F").Clear
          wb.Sheets(Sheetname).Columns("O:O").Copy 'D7-D9
      .Columns("G:G").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("H:H").Clear
      .Columns("I:I").Clear
          wb.Sheets(Sheetname).Columns("P:P").Copy 'D12-D14
      .Columns("J:J").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("K:K").Copy 'D15
      .Columns("K:K").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("L:L").Clear
      .Columns("M:M").Clear
          wb.Sheets(Sheetname).Columns("Q:Q").Copy 'D19
      .Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("O:O").Clear
          wb.Sheets(Sheetname).Columns("L:L").Copy 'D26
      .Columns("P:P").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("M:M").Copy 'D27
      .Columns("Q:Q").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("R:R").Clear
          wb.Sheets(Sheetname).Columns("T:T").Copy 'D31
      .Columns("S:S").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("T:T").Clear
          wb.Sheets(Sheetname).Columns("U:U").Copy 'D33
      .Columns("U:U").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("V:V").Copy 'D34
      .Columns("V:V").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("W:W").Copy 'D35
      .Columns("W:W").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("X:X").Clear
          wb.Sheets(Sheetname).Columns("R:R").Copy 'D37
      .Columns("Y:Y").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("X:X").Copy 'D38
      .Columns("Z:Z").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
  End Function
  Function Copytodatasheet2()
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Value
    Sheetname = "Commercial_ Expense Design"
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    With wsPro
      wb.Sheets(Sheetname).Columns("A:A").Copy '
      .Columns("A:A").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("AY:AY").Copy '
      .Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("G:G").Copy '
      .Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("E:E").Copy '
      .Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("J:J").Copy 'D1-D3
      .Columns("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("K:K").Copy 'D4-D6
      .Columns("F:F").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("L:L").Copy 'D7-D9
      .Columns("G:G").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("N:N").Copy 'D10
      .Columns("H:H").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("M:M").Copy 'D11
      .Columns("I:I").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("O:O").Copy 'D12-D14
      .Columns("J:J").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("K:K").Clear
          wb.Sheets(Sheetname).Columns("P:P").Copy 'D16-D17
      .Columns("L:L").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("M:M").Clear
          wb.Sheets(Sheetname).Columns("Q:Q").Copy 'D19
      .Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("R:R").Copy 'D24-D25
      .Columns("O:O").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("P:P").Clear
      .Columns("Q:Q").Clear
          wb.Sheets(Sheetname).Columns("T:T").Copy 'D29
      .Columns("R:R").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("S:S").Clear
          wb.Sheets(Sheetname).Columns("W:W").Copy 'D31
      .Columns("T:T").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("U:U").Clear
          wb.Sheets(Sheetname).Columns("X:X").Copy 'D33
      .Columns("V:V").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("W:W").Clear
          wb.Sheets(Sheetname).Columns("U:U").Copy 'D36
      .Columns("X:X").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("V:V").Copy 'D37
      .Columns("Y:Y").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("Y:Y").Copy 'D38
      .Columns("Z:Z").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
  End Function
  Function Copytodatasheet3()
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Value
    Sheetname = "Asbuilt Coax & Fiber"
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    With wsPro
          wb.Sheets(Sheetname).Columns("A:A").Copy '
      .Columns("A:A").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("AV:AV").Copy '
      .Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("G:G").Copy '
      .Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("E:E").Copy '
      .Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("J:J").Copy 'D1-D3
      .Columns("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("F:F").Clear
      .Columns("G:G").Clear
          wb.Sheets(Sheetname).Columns("L:L").Copy 'D10
      .Columns("H:H").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("K:K").Copy 'D11
      .Columns("I:I").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("M:M").Copy 'D12-D14
      .Columns("J:J").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("K:K").Clear
          wb.Sheets(Sheetname).Columns("N:N").Copy 'D16-D17
      .Columns("L:L").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("M:M").Clear
          wb.Sheets(Sheetname).Columns("O:O").Copy 'D19
      .Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("O:O").Clear
      .Columns("P:P").Clear
      .Columns("Q:Q").Clear
      .Columns("R:R").Clear
          wb.Sheets(Sheetname).Columns("Q:Q").Copy 'D31
      .Columns("S:S").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("T:T").Clear
          wb.Sheets(Sheetname).Columns("R:R").Copy 'D33
      .Columns("U:U").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("V:V").Clear
      .Columns("W:W").Clear
          wb.Sheets(Sheetname).Columns("S:S").Copy 'D36
      .Columns("X:X").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("T:T").Copy 'D37
      .Columns("Y:Y").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           wb.Sheets(Sheetname).Columns("U:U").Copy 'D38
      .Columns("Z:Z").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
  End Function
  Function Copytodatasheet4()
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Value
    Sheetname = "SFU&MDU Design"
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    With wsPro
          wb.Sheets(Sheetname).Columns("A:A").Copy
      .Columns("A:A").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("BD:BD").Copy
      .Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("G:G").Copy
      .Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("E:E").Copy
      .Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("J:J").Copy '
      .Columns("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("K:K").Copy '
      .Columns("F:F").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("L:L").Copy '
      .Columns("G:G").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("N:N").Copy '
      .Columns("H:H").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("M:M").Copy '
      .Columns("I:I").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("O:O").Copy '
      .Columns("J:J").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("K:K").Clear
          wb.Sheets(Sheetname).Columns("R:R").Copy '
      .Columns("L:L").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("M:M").Clear
          wb.Sheets(Sheetname).Columns("S:S").Copy '
      .Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("O:O").Clear
      .Columns("P:P").Clear
      .Columns("Q:Q").Clear
      .Columns("R:R").Clear
          wb.Sheets(Sheetname).Columns("U:U").Copy 'D31
      .Columns("S:S").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("T:T").Clear
          wb.Sheets(Sheetname).Columns("V:V").Copy 'D33
      .Columns("U:U").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("W:W").Copy 'D34
      .Columns("V:V").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("W:W").Clear
          wb.Sheets(Sheetname).Columns("P:P").Copy 'D36
      .Columns("X:X").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("Q:Q").Copy 'D37
      .Columns("Y:Y").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("X:X").Copy 'D38
      .Columns("Z:Z").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
  End Function
  Function Copytodatasheet5()
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Value
    Sheetname = "ME Design,Asbuit&Desktop Srvy"
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    With wsPro
          wb.Sheets(Sheetname).Columns("A:A").Copy
          .Columns("A:A").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("AT:AT").Copy
          .Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("G:G").Copy
          .Columns("C:C").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("E:E").Copy
      .Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("J:J").Copy 'D1-D3
      .Columns("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("L:L").Copy 'D4-D6
      .Columns("F:F").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("K:K").Copy 'D7-D9
      .Columns("G:G").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("H:H").Clear
      .Columns("I:I").Clear
          wb.Sheets(Sheetname).Columns("M:M").Copy 'D12-D14
      .Columns("J:J").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("K:K").Clear
      .Columns("L:L").Clear
      .Columns("M:M").Clear
          wb.Sheets(Sheetname).Columns("N:N").Copy 'D19
      .Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("O:O").Clear
      .Columns("P:P").Clear
      .Columns("Q:Q").Clear
      .Columns("R:R").Clear
          wb.Sheets(Sheetname).Columns("S:S").Copy 'D31
      .Columns("S:S").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("T:T").Clear
          wb.Sheets(Sheetname).Columns("O:O").Copy 'D33
      .Columns("U:U").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      .Columns("V:V").Clear
      .Columns("W:W").Clear
          wb.Sheets(Sheetname).Columns("P:P").Copy 'D36
      .Columns("X:X").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("Q:Q").Copy 'D37
      .Columns("Y:Y").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          wb.Sheets(Sheetname).Columns("T:T").Copy 'D38
      .Columns("Z:Z").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
  End Function
  Function ProcessingValues(Count, linespace)
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Outputsheet1 = "Invoice Details" 'wo1
    Outputsheet2 = "Invoice Summary" 'wo2
    DataSheetname = "DataProcess"
    Filename = Workbooks(OutputFileName).Sheets(Outputsheet2).Range("Z" & 1).Value
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    Dim wbo As Workbook
    Set wbo = Workbooks(OutputFileName)
    Dim wo1 As Worksheet
    Set wo1 = wbo.Sheets(Outputsheet1)
    Dim wo2 As Worksheet
    Set wo2 = wbo.Sheets(Outputsheet2)
    For x = 1 To Count
      linespacedata = x + linespace
      'Details
        wo1.Cells(linespacedata, 3).Value = wsPro.Range("A" & x + 1).Value 'Date created
        wo1.Cells(linespacedata, 4).Value = wsPro.Range("B" & x + 1).Value 'delivery date
        wo1.Cells(linespacedata, 5).Value = wsPro.Range("C" & x + 1).Value 'Job number
        wo1.Cells(linespacedata, 6).Value = wsPro.Range("D" & x + 1).Value 'Type
      'Multiple Column D1 to D3 Templete1
        If wsPro.Range("E" & x + 1).Value >= 1 And wsPro.Range("E" & x + 1).Value <= 2000 Then 'cell(x,A)
          wo1.Cells(linespacedata, 7).Value = 1 ' we adding 3 to push data down in invoice sheet. wo.Cells(row,column)
          wo1.Cells(linespacedata, 39).Value = wo2.Cells(14, 5).Value * 1
          wo1.Cells(linespacedata, 8).Value = 0
          wo1.Cells(linespacedata, 9).Value = 0
          wo1.Cells(linespacedata, 40).Value = 0
          wo1.Cells(linespacedata, 41).Value = 0
        ElseIf wsPro.Range("E" & x + 1).Value >= 2001 And wsPro.Range("E" & x + 1).Value <= 20000 Then
          wo1.Cells(linespacedata, 7).Value = 1
          wo1.Cells(linespacedata, 8).Value = wsPro.Range("E" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 39).Value = wo2.Cells(14, 5).Value * 1
          wo1.Cells(linespacedata, 40).Value = wo2.Cells(15, 5).Value * wo1.Cells(linespacedata, 8).Value
          wo1.Cells(linespacedata, 9).Value = 0
          wo1.Cells(linespacedata, 41).Value = 0
        ElseIf wsPro.Range("E" & x + 1).Value >= 20001 Then
          wo1.Cells(linespacedata, 7).Value = 1
          wo1.Cells(linespacedata, 8).Value = 0
          wo1.Cells(linespacedata, 9).Value = wsPro.Range("E" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 39).Value = wo2.Cells(14, 5).Value * 1
          wo1.Cells(linespacedata, 40).Value = 0
          wo1.Cells(linespacedata, 41).Value = wo2.Cells(16, 5).Value * wo1.Cells(linespacedata, 9).Value
        Else
          wo1.Cells(linespacedata, 7).Value = 0
          wo1.Cells(linespacedata, 8).Value = 0
          wo1.Cells(linespacedata, 9).Value = 0
          wo1.Cells(linespacedata, 39).Value = 0
          wo1.Cells(linespacedata, 40).Value = 0
          wo1.Cells(linespacedata, 41).Value = 0
        End If
      'Multiple Column D4 to D6
        If wsPro.Range("F" & x + 1).Value >= 1 And wsPro.Range("F" & x + 1).Value <= 2000 Then
          wo1.Cells(linespacedata, 10).Value = 1
          wo1.Cells(linespacedata, 42).Value = wo2.Cells(17, 5).Value * 1
          wo1.Cells(linespacedata, 11).Value = 0
          wo1.Cells(linespacedata, 12).Value = 0
          wo1.Cells(linespacedata, 43).Value = 0
          wo1.Cells(linespacedata, 44).Value = 0
        ElseIf wsPro.Range("F" & x + 1).Value >= 2001 And wsPro.Range("F" & x + 1).Value <= 20000 Then
          wo1.Cells(linespacedata, 10).Value = 1
          wo1.Cells(linespacedata, 11).Value = wsPro.Range("F" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 42).Value = wo2.Cells(17, 5).Value * 1
          wo1.Cells(linespacedata, 43).Value = wo2.Cells(18, 5).Value * wo1.Cells(linespacedata, 11).Value
          wo1.Cells(linespacedata, 12).Value = 0
          wo1.Cells(linespacedata, 44).Value = 0
        ElseIf wsPro.Range("F" & x + 1).Value >= 20001 Then
          wo1.Cells(linespacedata, 10).Value = 1
          wo1.Cells(linespacedata, 12).Value = wsPro.Range("F" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 42).Value = wo2.Cells(17, 5).Value * 1
          wo1.Cells(linespacedata, 44).Value = wo2.Cells(19, 5).Value * wo1.Cells(linespacedata, 12).Value
          wo1.Cells(linespacedata, 11).Value = 0
          wo1.Cells(linespacedata, 43).Value = 0
        Else
          wo1.Cells(linespacedata, 10).Value = 0
          wo1.Cells(linespacedata, 11).Value = 0
          wo1.Cells(linespacedata, 12).Value = 0
          wo1.Cells(linespacedata, 42).Value = 0
          wo1.Cells(linespacedata, 43).Value = 0
          wo1.Cells(linespacedata, 44).Value = 0
        End If
      'Multiple Column D7 to D9
        If wsPro.Range("G" & x + 1).Value >= 1 And wsPro.Range("G" & x + 1).Value <= 2000 Then
          wo1.Cells(linespacedata, 13).Value = 1
          wo1.Cells(linespacedata, 45).Value = wo2.Cells(20, 5).Value * 1
                  wo1.Cells(linespacedata, 15).Value = 0
          wo1.Cells(linespacedata, 47).Value = 0
                  wo1.Cells(linespacedata, 14).Value = 0
          wo1.Cells(linespacedata, 46).Value = 0
        ElseIf wsPro.Range("G" & x + 1).Value >= 2001 And wsPro.Range("G" & x + 1).Value <= 20000 Then
          wo1.Cells(linespacedata, 13).Value = 1
          wo1.Cells(linespacedata, 14).Value = wsPro.Range("G" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 45).Value = wo2.Cells(20, 5).Value * 1
          wo1.Cells(linespacedata, 46).Value = wo2.Cells(21, 5).Value * wo1.Cells(linespacedata, 14).Value
          wo1.Cells(linespacedata, 15).Value = 0
          wo1.Cells(linespacedata, 47).Value = 0
        ElseIf wsPro.Range("G" & x + 1).Value >= 20001 Then
          wo1.Cells(linespacedata, 13).Value = 1
          wo1.Cells(linespacedata, 15).Value = wsPro.Range("G" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 45).Value = wo2.Cells(20, 5).Value * 1
          wo1.Cells(linespacedata, 47).Value = wo2.Cells(22, 5).Value * wo1.Cells(linespacedata, 14).Value
                  wo1.Cells(linespacedata, 14).Value = 0
          wo1.Cells(linespacedata, 46).Value = 0
        Else
          wo1.Cells(linespacedata, 13).Value = 0
          wo1.Cells(linespacedata, 14).Value = 0
          wo1.Cells(linespacedata, 15).Value = 0
          wo1.Cells(linespacedata, 45).Value = 0
          wo1.Cells(linespacedata, 46).Value = 0
          wo1.Cells(linespacedata, 47).Value = 0
        End If
      'Single Column D10
          If wsPro.Range("H" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 16).Value = wsPro.Range("H" & x + 1).Value
            wo1.Cells(linespacedata, 48).Value = wo2.Cells(23, 5).Value * wsPro.Range("H" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 16).Value = 0
                  wo1.Cells(linespacedata, 48).Value = 0
              End If
          'Single Column D11
          If wsPro.Range("I" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 17).Value = wsPro.Range("I" & x + 1).Value
            wo1.Cells(linespacedata, 49).Value = wo2.Cells(24, 5).Value * wsPro.Range("I" & x + 1).Value
              Else
              wo1.Cells(linespacedata, 17).Value = 0
              wo1.Cells(linespacedata, 49).Value = 0
              End If
          'Multiple Column D12 to D14
        If wsPro.Range("J" & x + 1).Value >= 1 And wsPro.Range("J" & x + 1).Value <= 2000 Then  'Change A, cell(x,A )
          wo1.Cells(linespacedata, 18).Value = 1
          wo1.Cells(linespacedata, 50).Value = wo2.Cells(25, 5).Value * 1
          wo1.Cells(linespacedata, 20).Value = 0
          wo1.Cells(linespacedata, 52).Value = 0
          wo1.Cells(linespacedata, 19).Value = 0
          wo1.Cells(linespacedata, 51).Value = 0
        ElseIf wsPro.Range("J" & x + 1).Value >= 2001 And wsPro.Range("J" & x + 1).Value <= 20000 Then
          wo1.Cells(linespacedata, 18).Value = 1
          wo1.Cells(linespacedata, 19).Value = wsPro.Range("J" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 50).Value = wo2.Cells(25, 5).Value * 1
          wo1.Cells(linespacedata, 51).Value = wo2.Cells(26, 5).Value * wo1.Cells(linespacedata, 19).Value
          wo1.Cells(linespacedata, 20).Value = 0
          wo1.Cells(linespacedata, 52).Value = 0
        ElseIf wsPro.Range("J" & x + 1).Value >= 20001 Then
          wo1.Cells(linespacedata, 18).Value = 1
          wo1.Cells(linespacedata, 20).Value = wsPro.Range("J" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 50).Value = wo2.Cells(25, 5).Value * 1
          wo1.Cells(linespacedata, 52).Value = wo2.Cells(27, 5).Value * wo1.Cells(linespacedata, 20).Value
          wo1.Cells(linespacedata, 19).Value = 0
          wo1.Cells(linespacedata, 51).Value = 0
        Else
          wo1.Cells(linespacedata, 18).Value = 0
          wo1.Cells(linespacedata, 19).Value = 0
          wo1.Cells(linespacedata, 20).Value = 0
          wo1.Cells(linespacedata, 50).Value = 0
          wo1.Cells(linespacedata, 51).Value = 0
          wo1.Cells(linespacedata, 52).Value = 0
        End If
      'Single Column D15
          If wsPro.Range("K" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 21).Value = wsPro.Range("K" & x + 1).Value
            wo1.Cells(linespacedata, 53).Value = wo2.Cells(28, 5).Value * wsPro.Range("K" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 21).Value = 0
                  wo1.Cells(linespacedata, 53).Value = 0
              End If
          'Multiple Column D16 to D17
        If wsPro.Range("L" & x + 1).Value >= 1 And wsPro.Range("L" & x + 1).Value <= 2000 Then 'Change A, cell(x,A )
          wo1.Cells(linespacedata, 22).Value = 1
          wo1.Cells(linespacedata, 54).Value = wo2.Cells(29, 5).Value * 1
          wo1.Cells(linespacedata, 23).Value = 0
          wo1.Cells(linespacedata, 55).Value = 0
        ElseIf wsPro.Range("L" & x + 1).Value >= 2001 And wsPro.Range("L" & x + 1).Value <= 20000 Then
          wo1.Cells(linespacedata, 22).Value = 1
          wo1.Cells(linespacedata, 23).Value = wsPro.Range("L" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 54).Value = wo2.Cells(29, 5).Value * 1
          wo1.Cells(linespacedata, 55).Value = wo2.Cells(30, 5).Value * wo1.Cells(linespacedata, 23).Value
        Else
          wo1.Cells(linespacedata, 22).Value = 0
          wo1.Cells(linespacedata, 23).Value = 0
          wo1.Cells(linespacedata, 54).Value = 0
          wo1.Cells(linespacedata, 55).Value = 0
        End If
      'Single Column D18
          If wsPro.Range("M" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 24).Value = wsPro.Range("M" & x + 1).Value
            wo1.Cells(linespacedata, 56).Value = wo2.Cells(31, 5).Value * wsPro.Range("M" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 24).Value = 0
                  wo1.Cells(linespacedata, 56).Value = 0
              End If
          'Single Column D19
          If wsPro.Range("N" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 25).Value = wsPro.Range("N" & x + 1).Value
            wo1.Cells(linespacedata, 57).Value = wo2.Cells(32, 5).Value * wsPro.Range("N" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 25).Value = 0
                  wo1.Cells(linespacedata, 57).Value = 0
              End If
          'Multiple Column D24 to D25
        If wsPro.Range("O" & x + 1).Value >= 1 And wsPro.Range("O" & x + 1).Value <= 2000 Then
          wo1.Cells(linespacedata, 26).Value = 1
          wo1.Cells(linespacedata, 58).Value = wo2.Cells(33, 5).Value * 1
          wo1.Cells(linespacedata, 27).Value = 0
          wo1.Cells(linespacedata, 59).Value = 0
        ElseIf wsPro.Range("O" & x + 1).Value >= 2001 And wsPro.Range("O" & x + 1).Value <= 20000 Then
          wo1.Cells(linespacedata, 26).Value = 1
          wo1.Cells(linespacedata, 27).Value = wsPro.Range("O" & x + 1).Value - 2000
          wo1.Cells(linespacedata, 58).Value = wo2.Cells(33, 5).Value * 1
          wo1.Cells(linespacedata, 59).Value = wo2.Cells(34, 5).Value * wo1.Cells(linespacedata, 27).Value
        Else
          wo1.Cells(linespacedata, 26).Value = 0
          wo1.Cells(linespacedata, 27).Value = 0
          wo1.Cells(linespacedata, 58).Value = 0
          wo1.Cells(linespacedata, 59).Value = 0
        End If
      'Single Column D26
          If wsPro.Range("P" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 28).Value = wsPro.Range("P" & x + 1).Value
            wo1.Cells(linespacedata, 60).Value = wo2.Cells(35, 5).Value * wsPro.Range("P" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 28).Value = 0
                  wo1.Cells(linespacedata, 60).Value = 0
              End If
          'Single Column D27
          If wsPro.Range("Q" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 29).Value = wsPro.Range("Q" & x + 1).Value
            wo1.Cells(linespacedata, 61).Value = wo2.Cells(36, 5).Value * wsPro.Range("Q" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 29).Value = 0
                  wo1.Cells(linespacedata, 61).Value = 0
              End If
          'Single Column D29
          If wsPro.Range("R" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 30).Value = wsPro.Range("R" & x + 1).Value
            wo1.Cells(linespacedata, 62).Value = wo2.Cells(37, 5).Value * wsPro.Range("R" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 30).Value = 0
                  wo1.Cells(linespacedata, 62).Value = 0
              End If
          'Single Column D31
          If wsPro.Range("S" & x + 1).Value <> 0 Then
                wo1.Cells(linespacedata, 31).Value = wsPro.Range("S" & x + 1).Value
            wo1.Cells(linespacedata, 63).Value = wo2.Cells(38, 5).Value * wsPro.Range("S" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 31).Value = 0
                  wo1.Cells(linespacedata, 63).Value = 0
              End If
          'Single Column D32
          If wsPro.Range("T" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 32).Value = wsPro.Range("T" & x + 1).Value
            wo1.Cells(linespacedata, 64).Value = wo2.Cells(39, 5).Value * wsPro.Range("T" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 32).Value = 0
                  wo1.Cells(linespacedata, 64).Value = 0
              End If
          'Single Column D33
          If wsPro.Range("U" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 33).Value = wsPro.Range("U" & x + 1).Value
            wo1.Cells(linespacedata, 65).Value = wo2.Cells(40, 5).Value * wsPro.Range("U" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 33).Value = 0
                  wo1.Cells(linespacedata, 65).Value = 0
              End If
          'Single Column D34
          If wsPro.Range("V" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 34).Value = wsPro.Range("V" & x + 1).Value
            wo1.Cells(linespacedata, 66).Value = wo2.Cells(41, 5).Value * wsPro.Range("V" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 34).Value = 0
                  wo1.Cells(linespacedata, 66).Value = 0
              End If
          'Single Column D35
          If wsPro.Range("W" & x + 1).Value <> 0 Then
            wo1.Cells(linespacedata, 35).Value = wsPro.Range("W" & x + 1).Value
            wo1.Cells(linespacedata, 67).Value = wo2.Cells(42, 5).Value * wsPro.Range("W" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 35).Value = 0
                  wo1.Cells(linespacedata, 67).Value = 0
              End If
      'Single Column D36
          If wsPro.Range("X" & x + 1).Value <> 0 Then
              wo1.Cells(linespacedata, 36).Value = wsPro.Range("X" & x + 1).Value
          wo1.Cells(linespacedata, 68).Value = wo2.Cells(43, 5).Value * wsPro.Range("X" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 36).Value = 0
                  wo1.Cells(linespacedata, 68).Value = 0
              End If
      'Single Column D37
          If wsPro.Range("Y" & x + 1).Value <> 0 Then
                  wo1.Cells(linespacedata, 37).Value = wsPro.Range("Y" & x + 1).Value
            wo1.Cells(linespacedata, 69).Value = wo2.Cells(44, 5).Value * wsPro.Range("Y" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 37).Value = 0
                  wo1.Cells(linespacedata, 69).Value = 0
              End If
      'Single Column D38
          If wsPro.Range("Z" & x + 1).Value <> 0 Then
                  wo1.Cells(linespacedata, 38).Value = wsPro.Range("Z" & x + 1).Value
            wo1.Cells(linespacedata, 70).Value = wo2.Cells(45, 5).Value * wsPro.Range("Z" & x + 1).Value
              Else
                  wo1.Cells(linespacedata, 38).Value = 0
                  wo1.Cells(linespacedata, 70).Value = 0
              End If
      'Calculating total
        wo1.Cells(linespacedata, 71).Value = WorksheetFunction.Sum(wo1.Range("AM" & linespacedata & ":BR" & linespacedata))
      Next
End Function
  
  
  
  