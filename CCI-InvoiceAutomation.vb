'Version: 1.00
'Ika - Invoice
'For Echobroadband
'User: Ika
'By Farhat Abbas & Ika


Sub F_Invoice_Automation_CCI()
 Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 filedate = Format(Date, "ddmmyyyy")
 OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
 Call Check_if_workbook_is_open(OutputFileName)
 Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\" & OutputFileName, CreateBackup:=False
 Filename = ThisWorkbook.name
 Call CheckDataSheet(OutputFileName, "Invoice Details")
 Call CheckDataSheet(OutputFileName, "Invoice Summary")
 Call CheckDataSheet(Filename, "DataProcess")
 Call DeleteDataSheet(OutputFileName, "Sheet1")
 Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").FormulaR1C1 = Filename
 Call Invoice_Summary
 Call main
 'Call Invoice_Image(ThisWorkbook.Path & "\" & "Echologo.png")
 Call DeleteDataSheet(Filename, "DataProcess")
 Workbooks(OutputFileName).Worksheets("Invoice Summary").Activate
 ActiveWindow.DisplayGridlines = False
 Workbooks(OutputFileName).Worksheets("Invoice Details").Activate
 ActiveWindow.DisplayGridlines = False
 Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Delete
 Application.DisplayAlerts = True
 Call pagesetupsetting
 Workbooks(OutputFileName).Save
 Application.ScreenUpdating = True
End Sub
Function Check_if_workbook_is_open(OutputFileName)
    Dim wb As Workbook 'to test if workbook is open. No change here
        For Each wb In Workbooks
            If wb.name = OutputFileName Then
                Workbooks(OutputFileName).Save
                Workbooks(OutputFileName).Close
            End If
        Next
End Function
Function CheckDataSheet(Filename, Sheetname)
    For Each Sheet In Workbooks(Filename).Worksheets ' Checking if VBA Sheet exist
        If Sheet.name = Sheetname Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
    Workbooks(Filename).Sheets.Add.name = Sheetname
End Function
Function DeleteDataSheet(Filename, Sheetname)
    For Each Sheet In Workbooks(Filename).Worksheets ' Checking if VBA Sheet exist
        If Sheet.name = Sheetname Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
End Function
Function Copytodvalue()
  '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "Report"
  DataSheetname = "Invoice Summary"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wbo As Workbook
  Set wbo = Workbooks(OutputFileName)
  Dim wsPro As Worksheet
  Set wsPro = wbo.Sheets(DataSheetname)
  TotalD = WorksheetFunction.CountA(wb.Sheets(Sheetname).Range("S:S"))
  With wsPro
    wb.Sheets(Sheetname).Range("S4:S" & TotalD).Copy 'D19
    .Range("B14:B" & TotalD + 10).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  End With
End Function
Function CopytodPrice()
  '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "Report"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  
  
  DataSheetname = "Invoice Summary"
  Dim wbo As Workbook
  Set wbo = Workbooks(OutputFileName)
  Dim wsPro As Worksheet
  Set wsPro = wbo.Sheets(DataSheetname)
  TotalD = WorksheetFunction.CountA(wb.Sheets(Sheetname).Range("S:S"))
  With wsPro
    .Range("I11").FormulaR1C1 = wb.Sheets(Sheetname).Cells(2, 20).Value
    .Range("L3").FormulaR1C1 = "USL20522001-" & wb.Sheets(Sheetname).Cells(3, 20).Value
    wb.Sheets(Sheetname).Range("T4:T" & TotalD).Copy 'D19
    .Range("E14:E" & TotalD + 10).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  End With
End Function
Function Invoice_Summary()
  'Data
   filedate = Format(Date, "ddmmyyyy")
   OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
   Outputsheetname1 = "Invoice Summary"
 
  'Output Workbook is represented as wbo
   Dim wbo As Workbook
   Set wbo = Workbooks(OutputFileName)
 
  'Output Worksheet is represented as wo
   Dim wo1 As Worksheet
   Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
   Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
   TotalD = WorksheetFunction.CountA(Workbooks(Filename).Sheets("Report").Range("S:S"))
   
 
 
  'Writing Values
   With wo1
     .Range("B11:C11,D11:G11,B12:E12,F12:G12,H12:I12,J12:K12").Merge
     .Range("L12:M12,K9:M9,K5:M5,K1:M1").Merge
     .Range("L2:M2,L3:M3,K6:M6,K7:M7,K8:M8").Merge
     .Range("B" & TotalD + 11 & ":E" & TotalD + 11 & ",F" & TotalD + 11 & ":G" & TotalD + 11 & ",H" & TotalD + 11 & ":I" & TotalD + 11 & ",J" & TotalD + 11 & ":K" & TotalD + 11 & ",L" & TotalD + 11 & ":M" & TotalD + 11).Merge 'Subtotal
     .Range("L" & TotalD + 12 & ":M" & TotalD + 12 & ",B" & TotalD + 12 & ":K" & TotalD + 12).Merge 'Invoice Total
     .Range("K1").FormulaR1C1 = "Invoice"
     .Range("K2").FormulaR1C1 = "DATE"
     .Range("K3").FormulaR1C1 = Date
     .Range("K3").NumberFormat = "d-mmm-yyyy"
     .Range("L2").FormulaR1C1 = "INVOICE #"
     .Range("K5").FormulaR1C1 = "BILL TO"
     .Range("K6").FormulaR1C1 = "Echo Broadband, Inc"
     .Range("K7").FormulaR1C1 = "Attn:  Accounts Payable"
     .Range("K8").FormulaR1C1 = "PO Box 1627"
     .Range("K9").FormulaR1C1 = "Broomfield, CO 80038"
     .Range("B11").FormulaR1C1 = "PROJECT: CCI / Comcast"
     .Range("C11").FormulaR1C1 = "CCI / Comcast"
     .Range("B6").FormulaR1C1 = "ECHO Broadband Sdn Bhd"
     .Range("B7").FormulaR1C1 = "368-5-3 Bellisa Row"
     .Range("B8").FormulaR1C1 = "Jalan Burmah"
     .Range("B9").FormulaR1C1 = "10350 Penang"
     .Range("D11").FormulaR1C1 = "** Data Processing and Provision of information**"
     .Range("H11").FormulaR1C1 = "Region :"
     .Range("L11").FormulaR1C1 = "Terms"
     .Range("M11").FormulaR1C1 = "Net 90"
     .Range("F12").FormulaR1C1 = "Node Splits"
     .Range("H12").FormulaR1C1 = "Coax Design & Asbuild"
     .Range("J12").FormulaR1C1 = "SFU & MDU"
     .Range("L12").FormulaR1C1 = "Fiber Design & Asbuild"
     .Range("B13").FormulaR1C1 = "Task No."
     .Range("C13").FormulaR1C1 = "Description"
     .Range("D13").FormulaR1C1 = "Unit"
     .Range("E13").FormulaR1C1 = "Rate(USD)"
     .Range("C14").FormulaR1C1 = "Route Drating<2000"
     .Range("C15").FormulaR1C1 = "Route Drating>2000"
     .Range("C16").FormulaR1C1 = "Large Route Drafting>20000"
     .Range("C17").FormulaR1C1 = "Coax Design<2000"
     .Range("C18").FormulaR1C1 = "Coax Design>2000"
     .Range("C19").FormulaR1C1 = "Large Coax Design>20000"
     .Range("C20").FormulaR1C1 = "Fiber Design<2000"
     .Range("C21").FormulaR1C1 = "Fiber Design>2000"
     .Range("C22").FormulaR1C1 = "Large Fiber Design>20000"
     .Range("C23").FormulaR1C1 = "Parcel Creation AutoCAD"
     .Range("C24").FormulaR1C1 = "Parcel Creation Others"
     .Range("C25").FormulaR1C1 = "Asbuild<2000"
     .Range("C26").FormulaR1C1 = "Asbuild>2000"
     .Range("C27").FormulaR1C1 = "Large Asbuild>20000"
     .Range("C28").FormulaR1C1 = "Node Split Asbuild"
     .Range("C29").FormulaR1C1 = "Coax MDU Up to 25"
     .Range("C30").FormulaR1C1 = "Coax MDU> 25"
     .Range("C31").FormulaR1C1 = "Large Coax MDU>20000"
     .Range("C32").FormulaR1C1 = "MDU Detail Insert"
     .Range("C33").FormulaR1C1 = "Forced Relo <2000"
     .Range("C34").FormulaR1C1 = "Forced Relo >2000"
     .Range("C35").FormulaR1C1 = "Node Split - fiber & Coax"
     .Range("C36").FormulaR1C1 = "Tier 2 Tombstone & EOL Update"
     .Range("C37").FormulaR1C1 = "Node Split Load Rebalance"
     .Range("C38").FormulaR1C1 = "Misc Repowering of node"
     .Range("C39").FormulaR1C1 = "Plant Map Update"
     .Range("C40").FormulaR1C1 = "Fiber Trunk Tree Modify"
     .Range("C41").FormulaR1C1 = "Fiber Trunk Tree New"
     .Range("C42").FormulaR1C1 = "Wavelength Res Req"
     .Range("C43").FormulaR1C1 = "Ladder Report"
     .Range("C44").FormulaR1C1 = "Fiber Route Trace"
     .Range("C45").FormulaR1C1 = "Splice Update"
     .Range("C46").FormulaR1C1 = "Splice Addition"
     .Range("C47").FormulaR1C1 = "Misc Hourly work"
     .Range("F13,H13,J13,L13").FormulaR1C1 = "Qty"
     .Range("I13,G13,K13,M13").FormulaR1C1 = "Sub Total"
     .Range("D43").FormulaR1C1 = "Report"
     .Range("D23,D24").FormulaR1C1 = "Parcel"
     .Range("D30,D36").FormulaR1C1 = "Unit"
     .Range("D41,D42").FormulaR1C1 = "Drawing"
     .Range("D47").FormulaR1C1 = "Hour"
     .Range("D14,D17,D20,D25,D28,D29,D33,D35,D37,D38").FormulaR1C1 = "Project"
     .Range("D15,D16,D18,D19,D21,D22,D26,D27,D31,D34").FormulaR1C1 = "Feet"
     .Range("D32,D39,D40,D44").FormulaR1C1 = "Each"
     .Range("D45,D46").FormulaR1C1 = "Insert"
     
     Call Copytodvalue
     Call CopytodPrice
     '----------------------------------------------------------------------------------
     '**** Update Required ****
     'Add New line here for D follow this pattern for Description & Unit
     '.Range("C48").FormulaR1C1 = "Route Drating<2000" 'Description
     '.Range("D48").FormulaR1C1 = "Project"
     
     
     
     
     
     '------------------------------------------------
     .Range("B" & TotalD + 11).FormulaR1C1 = "Subtotals"
     .Range("B" & TotalD + 12).FormulaR1C1 = "Invoice Total ( USD)"
     
     .Range("E14:E" & TotalD + 10 & ",G14:G" & TotalD + 10 & ",I14:I" & TotalD + 10 & ",K14:K" & TotalD + 100 & ",M14:M" & TotalD + 10 & ",F" & TotalD + 11 & ":M" & TotalD + 12).NumberFormat = "[$$-en-US]#,##0.00"
     .Range("A1:M9,A12:M13,B:B,B" & TotalD + 11 & ":M" & TotalD + 12 & ",B11:H11,L11").Font.FontStyle = "Bold"
     
     .Columns("A").ColumnWidth = 0.88
     .Columns("B").ColumnWidth = 8
     .Columns("C").ColumnWidth = 24.38
     .Columns("D").ColumnWidth = 7.5
     .Columns("E").ColumnWidth = 9.5
     .Columns("F").ColumnWidth = 7.6
     .Columns("H").ColumnWidth = 7.6
     .Columns("L").ColumnWidth = 7.6
     .Columns("J").ColumnWidth = 7.6
     .Columns("K").ColumnWidth = 11.14
     .Columns("G").ColumnWidth = 11.14
     .Columns("I").ColumnWidth = 11.14
     .Columns("M").ColumnWidth = 11.14
     With .Range("B6:B9,K6:K10").Font
     .name = "Arial"
     .Size = 10
     End With
     With .Range("K5").Font
     .name = "Arial Black"
     .Size = 14
     End With
     With .Range("L3").Font
     .name = "Arial"
     .Size = 11
     End With
     With .Range("K5").Font
     .name = "Arial Black"
     .Size = 14
     End With
     With .Range("B11").Font
     .name = "Calibri"
     .Size = 13
     End With
     With .Range("K1").Font
     .name = "Calibri"
     .Size = 20
     End With
     With .Range("K2,L2,D11,C13,B13:B" & TotalD + 10 & ",B" & TotalD + 11 & ":M" & TotalD + 12).Font
     .name = "Calibri"
     .Size = 12
     End With
     With .Range("D11:M" & TotalD + 12 & ",B13,K1:M3,K5")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .ReadingOrder = xlContext
     End With
     With .Range("D11")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .ReadingOrder = xlContext
     End With
     With .Range("C11,C13:C" & TotalD + 10 & ",C11")
       .HorizontalAlignment = xlGeneral
       .VerticalAlignment = xlCenter
       .ReadingOrder = xlContext
     End With
     With .Range("B14:B" & TotalD + 12)
       .HorizontalAlignment = xlRight
       .VerticalAlignment = xlCenter
       .ReadingOrder = xlContext
   End With
   'normal table
     tablesize = "B11:M" & TotalD + 12 & ",K2:M3,K5:M5"
     With .Range(tablesize).Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .ColorIndex = xlAutomatic
     .TintAndShade = 0
     .Weight = xlThin
     End With
     With .Range(tablesize).Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .ColorIndex = xlAutomatic
         .TintAndShade = 0
         .Weight = xlThin
     End With
     With .Range(tablesize).Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .ColorIndex = xlAutomatic
         .TintAndShade = 0
         .Weight = xlThin
     End With
     With .Range(tablesize).Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .ColorIndex = xlAutomatic
         .TintAndShade = 0
         .Weight = xlThin
     End With
     With .Range(tablesize).Borders(xlInsideVertical)
         .LineStyle = xlContinuous
         .ColorIndex = xlAutomatic
         .TintAndShade = 0
         .Weight = xlThin
     End With
     With .Range(tablesize).Borders(xlInsideHorizontal)
         .LineStyle = xlContinuous
         .ColorIndex = xlAutomatic
         .TintAndShade = 0
         .Weight = xlThin
     End With
     With .Range("B11:M12").Interior
       .Pattern = xlSolid
       .PatternColorIndex = xlAutomatic
       .Color = 10020351
       .TintAndShade = 0
       .PatternTintAndShade = 0
     End With
   'Erasing line boarder
     Thick_outside_boarder = "K2:M3,K5:M9,B11:G11,H11:I11,J11:K11,L11:M11,B11:E11,F11:G11,H11:I11,J11:K11,L11:M11,B12:E12,B" & TotalD + 12 & ":K" & TotalD + 12 & ",L" & TotalD + 11 & ":M" & TotalD + 12
     Call Thick_OB(Thick_outside_boarder)
     Thick_outside_boarder = "F12:G12,H12:I12,J12:K12,L12:M12,B13:E" & TotalD + 11 & ",F13:G" & TotalD + 11 & ",H13:I" & TotalD + 11 & ",J13:K" & TotalD + 11 & ",L13:M" & TotalD + 11 & ",B" & TotalD + 12 & ":E" & TotalD + 12 & ",J" & TotalD + 12 & ":K" & TotalD + 12 & ",F" & TotalD + 12 & ":G" & TotalD + 12 & ",H" & TotalD + 12 & ":I" & TotalD + 12 & ",L" & TotalD + 12 & ":M" & TotalD + 12
     Call Thick_OB(Thick_outside_boarder)
     .Range("C13,E13,G13,I13,K13,M13").Borders(xlDiagonalDown).LineStyle = xlNone
     .Range("C13,E13,G13,I13,K13,M13").Borders(xlDiagonalUp).LineStyle = xlNone
     .Range("C13,E13,G13,I13,K13,M13").Borders(xlEdgeLeft).LineStyle = xlNone
     With .Range("C13,E13,G13,I13,K13,M13").Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .ColorIndex = 0
         .TintAndShade = 0
         .Weight = xlMedium
     End With
     With .Range("C13,E13,G13,I13,K13,M13").Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .ColorIndex = 0
         .TintAndShade = 0
         .Weight = xlThin
     End With
     With .Range("C13,E13,G13,I13,K13,M13").Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .ColorIndex = 0
         .TintAndShade = 0
         .Weight = xlMedium
     End With
     .Range("C13,E13,G13,I13,K13,M13").Borders(xlInsideVertical).LineStyle = xlNone
     .Range("C13,E13,G13,I13,K13,M13").Borders(xlInsideHorizontal).LineStyle = xlNone
     Call Clearup
     For x = 14 To 47
     .Rows(x & ":" & x).RowHeight = 19.5
     Next x
     .Rows("12:12").RowHeight = 30
     .Rows("11:11").RowHeight = 30
     .Rows(TotalD + 11 & ":" & TotalD + 11).RowHeight = 21
     .Rows(TotalD + 12 & ":" & TotalD + 12).RowHeight = 27
     
     End With
 
   
 End Function
   
 Function Thick_OB(Thick_outside_boarder)
  'Data
   filedate = Format(Date, "ddmmyyyy")
   OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
   Outputsheetname1 = "Invoice Summary"
 
  'Output Workbook is represented as wbo
   Dim wbo As Workbook
   Set wbo = Workbooks(OutputFileName)
 
  'Output Worksheet is represented as wo
   Dim wo1 As Worksheet
   Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
 
  'Writing Values
   With wo1
 
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
 Function Clearup()
  'Data
   filedate = Format(Date, "ddmmyyyy")
   OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
   Outputsheetname1 = "Invoice Summary"
 
  'Output Workbook is represented as wbo
   Dim wbo As Workbook
   Set wbo = Workbooks(OutputFileName)
 
  'Output Worksheet is represented as wo
   Dim wo1 As Worksheet
   Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
 
  'Writing Values
   With wo1
   .Range("D13").Borders(xlDiagonalDown).LineStyle = xlNone
   .Range("D13").Borders(xlDiagonalUp).LineStyle = xlNone
   .Range("D13").Borders(xlEdgeLeft).LineStyle = xlNone
   With .Range("D13").Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlMedium
   End With
   With .Range("D13").Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlThin
   End With
   .Range("D13").Borders(xlEdgeRight).LineStyle = xlNone
   .Range("D13").Borders(xlInsideVertical).LineStyle = xlNone
   .Range("D13").Borders(xlInsideHorizontal).LineStyle = xlNone
 
   .Range("D11:G11").Borders(xlDiagonalDown).LineStyle = xlNone
   .Range("D11:G11").Borders(xlDiagonalUp).LineStyle = xlNone
   .Range("D11:G11").Borders(xlEdgeLeft).LineStyle = xlNone
   With .Range("D11:G11").Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlMedium
   End With
   With .Range("D11:G11").Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlMedium
   End With
   With .Range("D11:G11").Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .ColorIndex = 0
       .TintAndShade = 0
       .Weight = xlMedium
   End With
   .Range("D11:G11").Borders(xlInsideHorizontal).LineStyle = xlNone
 End With
End Function
 

Function main()

  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  region = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("I11").Value
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
   TotalD = WorksheetFunction.CountA(wb.Sheets("Report").Range("S:S"))
    '----------------------------------------------------------------------------------
    '**** Update Required ****
  
    'Filtering the data Sheet 1 'NODE SPLIT
     ws1.Range("BM:BM").AutoFilter Field:=65, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws1.Range("D:D").AutoFilter Field:=4, Criteria1:=region  'Region
     ws1.Range("BN:BN").AutoFilter Field:=66, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws1.Range("BL:BL").AutoFilter Field:=64, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
    
    'Filtering the data Sheet 4 'CED
     ws2.Range("AY:AY").AutoFilter Field:=51, Operator:=xlFilterValues, Criteria1:="<>="   'Blank Dilvery date
     ws2.Range("D:D").AutoFilter Field:=4, Criteria1:=region   'Region
     ws2.Range("AZ:AZ").AutoFilter Field:=52, Criteria1:="=", Operator:=xlFilterValues  'Invoice colum
     ws2.Range("AX:AX").AutoFilter Field:=50, Criteria1:=Array("Completed"), Operator:=xlFilterValues   'QB Dilevery
    
    'Filtering the data Sheet 5
     ws3.Range("AV:AV").AutoFilter Field:=48, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws3.Range("D:D").AutoFilter Field:=4, Criteria1:=region  'Region
     ws3.Range("AW:AW").AutoFilter Field:=49, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws3.Range("AU:AU").AutoFilter Field:=47, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
    
    'Filtering the data Sheet 3 'SFU
     ws4.Range("BD:BD").AutoFilter Field:=56, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws4.Range("D:D").AutoFilter Field:=4, Criteria1:=region  'Region
     ws4.Range("BE:BE").AutoFilter Field:=57, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws4.Range("BC:BC").AutoFilter Field:=55, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
    
     'Filtering the data Sheet 2 'ME
     ws5.Range("AT:AT").AutoFilter Field:=46, Operator:=xlFilterValues, Criteria1:="<>="  'Blank Dilvery date
     ws5.Range("D:D").AutoFilter Field:=4, Criteria1:=region 'Region
     ws5.Range("AU:AU").AutoFilter Field:=47, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
     ws5.Range("AS:AS").AutoFilter Field:=45, Criteria1:=Array("Completed"), Operator:=xlFilterValues  'QB Dilevery
     
    '----------------------------------------------------------------------------------

  'Count
   Count1 = ws1.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   Count2 = ws2.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   count3 = ws3.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   count4 = ws4.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
   count5 = ws5.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
  
   count12 = Count1 + Count2
   count123 = Count1 + Count2 + count3
   count1234 = Count1 + Count2 + count3 + count4
   count12345 = Count1 + Count2 + count3 + count4 + count5
  
  
  
  'Format
    wo1.Columns("C:D").NumberFormat = "[$-en-US]d-mmm;@"
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    
    With wo1.Columns("A:BX")
    
    
    '----------------------------------------------------------------------------------
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = True
    End With
  
  'Sheet1 Calculation and pasting
    'Calling Functions
      Call Copytodatasheet1 'Updating the function - Copying Data from tracking to vba sheet
      Call ProcessingValues(Count1, 4) 'Updating the function - Calculating for D1 till D...
      Call detailtemplate(0, region, 1, Count1, "F", "G") 'Updating the function - Creating templete for detail
      
  'Sheet 2 & Sheet 3
    'Calling Functions
      Call Copytodatasheet2
      Call ProcessingValues(Count2, 10 + Count1)
      Call Copytodatasheet3
      Call ProcessingValues(count3, 10 + count12)
      Call detailtemplate(6 + Count1, region, 2, Count2 + count3, "H", "I")
      
  'Sheet 4
    'Calling Functions
      Call Copytodatasheet4
      Call ProcessingValues(count4, 16 + count123)
      Call detailtemplate(12 + count123, region, 3, count4, "J", "K")
      
  'Sheet 5
    'Calling Functions
      Call Copytodatasheet5
      Call ProcessingValues(count5, 22 + count1234)
      Call detailtemplate(count1234 + 18, region, 4, count5, "L", "M")
      
      
  'Format
    wo2.Range("L" & TotalD + 12).Value = wo2.Range("F" & TotalD + 11).Value + wo2.Range("H" & TotalD + 11).Value + wo2.Range("J" & TotalD + 11).Value + wo2.Range("L" & TotalD + 11).Value

    wo1.Columns("A").ColumnWidth = 8
    wo1.Columns("B").ColumnWidth = 8
    wo1.Columns("C:D").ColumnWidth = 13.5
    wo1.Columns("E:F").ColumnWidth = 25.5
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    
    wo1.Range("AO:BW").NumberFormat = "[$$-en-US]#,##0.00"
    wo1.Columns("G:AL").ColumnWidth = 13.5
    wo1.Columns("AO:BW").ColumnWidth = 11.5
    wo1.Columns("BX").ColumnWidth = 46.8
    
    '----------------------------------------------------------------------------------
    wo1.Rows("1:1").RowHeight = 15.75
    wo1.Rows("2:2").RowHeight = 10 'table 1

End Function
Function boardermoney(firstrow, lastrow)
 '***Please note that update is required in this function***
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
    '----------------------------------------------------------------------------------
    '**** Update Required ****

    Boarder_LastRow = "AP" & lastrow & ":BW" & lastrow

    '----------------------------------------------------------------------------------
    'table
      .Range(Boarder_LastRow).Borders(xlDiagonalDown).LineStyle = xlNone
      .Range(Boarder_LastRow).Borders(xlDiagonalUp).LineStyle = xlNone
      With .Range(Boarder_LastRow).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(Boarder_LastRow).Borders(xlEdgeTop)
          .LineStyle = xlDouble
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThick
      End With
      With .Range(Boarder_LastRow).Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(Boarder_LastRow).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      With .Range(Boarder_LastRow).Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .ColorIndex = 0
          .TintAndShade = 0
          .Weight = xlThin
      End With
      .Range(Boarder_LastRow).Borders(xlInsideHorizontal).LineStyle = xlNone
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    
    
    With .Range("AO" & firstrow & ":BV" & lastrow).Interior 'Blue color Highlighted
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
    End With
    With .Range("BW" & firstrow + 1 & ":BW" & lastrow).Interior 'Blue color Highlighted last column Subtotal
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
    End With
    .Range("AO" & firstrow & ":BX" & firstrow + 1).Font.Bold = True 'First second row to bold
    .Range("AO" & lastrow & ":BX" & lastrow).Font.Bold = True 'Second second row to bold
    
    '----------------------------------------------------------------------------------
 End With
End Function
Function thickline1(thickline)
  '***Please note that update is required in this function***
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
Function detailtemplate(x, region, y, Count, QtyCol, SubtotalCol)
  If Count = 0 Then
      NewCount = Count + 1
  Else
      NewCount = Count
  End If

  'Data
  
   filedate = Format(Date, "ddmmyyyy")
   OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
   Outputsheetname1 = "Invoice Details"
   Outputsheetname2 = "Invoice Summary"
  'Input sheet data
   Filename = Workbooks(OutputFileName).Sheets(Outputsheetname2).Range("Z1").Value
   Dim wb As Workbook
   Set wb = Workbooks(Filename)
   Dim win As Worksheet
   Set win = wb.Sheets("Report")

  'Output Workbook is represented as wbo
   Dim wbo As Workbook
   Set wbo = Workbooks(OutputFileName)

  'Output Worksheet is represented as wo
   Dim wo1 As Worksheet
   Set wo1 = wbo.Sheets(Outputsheetname1) 'detail
  'Output Worksheet is represented as wo
   Dim wo2 As Worksheet
   Set wo2 = wbo.Sheets(Outputsheetname2) 'Summary

  'Invoice detail template
   With wo1
   TotalD = WorksheetFunction.CountA(wb.Sheets("Report").Range("S:S"))
    If y = 1 Then
      .Range("C" & x + 1).FormulaR1C1 = "Invoice:"
      .Range("D" & x + 1).FormulaR1C1 = wo2.Range("L3").Value
      .Range("G" & x + 1).FormulaR1C1 = "Region:"
      .Range("H" & x + 1).FormulaR1C1 = region
      .Range("C" & x + 3).FormulaR1C1 = "Node Split Design & Asbuilt"
    ElseIf y = 2 Then
      .Range("C" & x + 3).FormulaR1C1 = "Coax Design & Asbuild"
    ElseIf y = 3 Then
      .Range("C" & x + 3).FormulaR1C1 = "SFU & MDU"
    ElseIf y = 4 Then
      .Range("C" & x + 3).FormulaR1C1 = "Fiber Design & Asbuild"
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
    .Range("D1:E1").Merge 'x + 3

    .Range("C" & x + 4 & ":F" & x + 4).Font.Bold = True
    .Range("C" & x + 4).FormulaR1C1 = "Date Created"
    .Range("D" & x + 4).FormulaR1C1 = "Delivery Date"
    .Range("E" & x + 4).FormulaR1C1 = "Job Number"
    .Range("F" & x + 4).FormulaR1C1 = "Type"
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    
    
    .Range("BW" & x + 4).FormulaR1C1 = "Subtotal"
    .Range("BX" & x + 4).FormulaR1C1 = "Remark"
    Call CopytodvalueTracking("G" & x + 3 & ":AN" & x + 3)
    Call CopytodvalueTracking("AO" & x + 3 & ":BV" & x + 3)
    Call CopytoddetailsTracking("G" & x + 4 & ":AN" & x + 4)
    Call CopytodPriceTracking("AO" & x + 4 & ":BV" & x + 4)
    Call TableArrangment("G" & x + 3 & ":BV" & x + 3)
    Call TableArrangment("C" & x + 4 & ":BX" & x + 4)
    Call TableArrangmentData("B" & x + 5 & ":BX" & x + 4 + NewCount)
    Call CalculateTotal(x + 5, x + 4 + NewCount)
    Call thickline1("AO" & x + 3 & ":AO" & x + 5 + NewCount)
    Call boardermoney(x + 3, x + 5 + NewCount) 'Updating the function   - Creating Boarder for money
    With .Range("G" & x + 5 + NewCount & ":AN" & x + 5 + NewCount).Font   'Blue font color
    
    '----------------------------------------------------------------------------------
        .Color = -10477568
        .TintAndShade = 0
        .Size = 11
        .Bold = True
    End With
    For J = 1 To Count
      .Range("B" & x + 4 + J).Value = J  'Numbering the data
    Next
    For Z = 1 To NewCount
      .Rows(x + 4 + Z & ":" & x + 4 + Z).RowHeight = 18 'Changing the row height
    Next
    .Rows(x + 2 & ":" & x + 2).RowHeight = 15 'Changing the row height
    .Rows(x + 3 & ":" & x + 3).RowHeight = 18 'Changing the row height
    .Rows(x + 4 & ":" & x + 4).RowHeight = 46 'Changing the row height
    

    
    
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    
    .Range("G" & x + 5 + NewCount & ":AN" & x + 5 + NewCount).Copy '---------------
    wo2.Range(QtyCol & "14:" & QtyCol & TotalD + 10).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    .Range("AO" & x + 5 + NewCount & ":BV" & x + 5 + NewCount).Copy '---------------
    wo2.Range(SubtotalCol & "14:" & SubtotalCol & TotalD + 10).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    wo2.Range(QtyCol & TotalD + 11).Value = Application.WorksheetFunction.Sum(wo2.Range(SubtotalCol & "14:" & SubtotalCol & TotalD + 10))
    
    
    
    '----------------------------------------------------------------------------------
   End With
End Function
Function CalculateTotal(StartingValue, EndingValue)
  '***Please note that update is required in this function***
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
    .Range("G" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("G" & StartingValue & ":G" & EndingValue)) 'Summing up the Column
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
    .Range("BS" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BS" & StartingValue & ":BS" & EndingValue))
    .Range("BT" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BT" & StartingValue & ":BT" & EndingValue))
    .Range("BU" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BU" & StartingValue & ":BU" & EndingValue))
    .Range("BV" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BV" & StartingValue & ":BV" & EndingValue))
    .Range("BW" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BW" & StartingValue & ":BW" & EndingValue))
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'add next line change "BW" to next column
    '.Range("BX" & EndingValue + 1).Value = Application.WorksheetFunction.Sum(.Range("BX" & StartingValue & ":BX" & EndingValue))
    
    
    '----------------------------------------------------------------------------------
    
    .Range(EndingValue + 1 & ":" & EndingValue + 1).Font.Bold = True 'Making pricing to bold

    
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
Function COPYPASTETracking(COPY_VALUE, PASTE_VALUE, Sheetname)
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
    With wsPro
    wb.Sheets(Sheetname).Columns(COPY_VALUE & ":" & COPY_VALUE).Copy
    .Columns(PASTE_VALUE & ":" & PASTE_VALUE).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End With
End Function
Function CopytodvalueTracking(PASTE_VALUE)
  '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "Report"
  DataSheetname = "Invoice Details"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wbo As Workbook
  Set wbo = Workbooks(OutputFileName)
  Dim wsPro As Worksheet
  Set wsPro = wbo.Sheets(DataSheetname)
  With wsPro
    TotalD = WorksheetFunction.CountA(wb.Sheets(Sheetname).Range("S:S"))
    wb.Sheets(Sheetname).Range("S4:S" & TotalD).Copy
    .Range(PASTE_VALUE).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
  End With
End Function
Function Copytodatasheet1()
  'Setting Data
    filedate = Format(Date, "ddmmyyyy")
    OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
    Sheetname = "Node Split"
    DataSheetname = "DataProcess"
    Dim wb As Workbook
    Set wb = Workbooks(Filename)
    Dim wsPro As Worksheet
    Set wsPro = wb.Sheets(DataSheetname)
  With wsPro
      .Cells.Clear
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'Call COPYPASTETracking("Column letter from tracking", "letter design for data base", Sheetname)
    'for example D39
    'Call COPYPASTETracking("Column letter from tracking", "AC", Sheetname) 'D38
    
      Call COPYPASTETracking("A", "A", Sheetname) 'Date created - 1
      Call COPYPASTETracking("BM", "B", Sheetname) 'delivery date - 2
      Call COPYPASTETracking("G", "C", Sheetname) 'Job number - 3
      Call COPYPASTETracking("E", "D", Sheetname) 'Type - 4
      Call COPYPASTETracking("O", "E", Sheetname) 'D1 to D3
      Call COPYPASTETracking("P", "G", Sheetname) 'D7 to D9
      Call COPYPASTETracking("Q", "J", Sheetname) 'D12-D14
      Call COPYPASTETracking("K", "K", Sheetname) 'D15
      Call COPYPASTETracking("R", "N", Sheetname) 'D19
      Call COPYPASTETracking("L", "P", Sheetname) 'D26
      Call COPYPASTETracking("N", "Q", Sheetname) 'D26A
      Call COPYPASTETracking("M", "R", Sheetname) 'D27
      Call COPYPASTETracking("U", "U", Sheetname) 'D31
      Call COPYPASTETracking("V", "W", Sheetname) 'D33
      Call COPYPASTETracking("W", "X", Sheetname) 'D34
      Call COPYPASTETracking("X", "Y", Sheetname) 'D35
      Call COPYPASTETracking("S", "Z", Sheetname) 'D36
      Call COPYPASTETracking("T", "AA", Sheetname) 'D37
      Call COPYPASTETracking("Y", "AB", Sheetname) 'D38


    '----------------------------------------------------------------------------------
  End With
End Function
Function Copytodatasheet2()
    '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "Commercial_ Expense Design"
  DataSheetname = "DataProcess"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wsPro As Worksheet
  Set wsPro = wb.Sheets(DataSheetname)
  With wsPro
      .Cells.Clear
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'Call COPYPASTETracking("Column letter from tracking", "letter design for data base", Sheetname)
    'for example D39
    'Call COPYPASTETracking("Column letter from tracking", "AC", Sheetname) 'D38
    
      Call COPYPASTETracking("A", "A", Sheetname) 'Date created - 1
      Call COPYPASTETracking("AY", "B", Sheetname) 'delivery date - 2
      Call COPYPASTETracking("G", "C", Sheetname) 'Job number - 3
      Call COPYPASTETracking("E", "D", Sheetname) 'Type - 4
      Call COPYPASTETracking("J", "E", Sheetname) 'D1 to D3
      Call COPYPASTETracking("K", "F", Sheetname) 'D4 to D6
      Call COPYPASTETracking("L", "G", Sheetname) 'D7 to D9
      Call COPYPASTETracking("N", "H", Sheetname) 'D10
      Call COPYPASTETracking("M", "I", Sheetname) 'D11
      Call COPYPASTETracking("O", "J", Sheetname) 'D12-D14
      Call COPYPASTETracking("P", "L", Sheetname) 'D16-D17
      Call COPYPASTETracking("Q", "N", Sheetname) 'D19
      Call COPYPASTETracking("R", "O", Sheetname) 'D24-D25
      Call COPYPASTETracking("T", "T", Sheetname) 'D29
      Call COPYPASTETracking("W", "U", Sheetname) 'D31
      Call COPYPASTETracking("X", "W", Sheetname) 'D33
      Call COPYPASTETracking("U", "Z", Sheetname) 'D36
      Call COPYPASTETracking("V", "AA", Sheetname) 'D37
      Call COPYPASTETracking("Y", "AB", Sheetname) 'D38
      

    '----------------------------------------------------------------------------------
  End With
End Function
Function Copytodatasheet3()
    '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "Asbuilt Coax & Fiber"
  DataSheetname = "DataProcess"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wsPro As Worksheet
  Set wsPro = wb.Sheets(DataSheetname)
  With wsPro
      .Cells.Clear
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'Call COPYPASTETracking("Column letter from tracking", "letter design for data base", Sheetname)
    'for example D39
    'Call COPYPASTETracking("Column letter from tracking", "AC", Sheetname) 'D38
      Call COPYPASTETracking("A", "A", Sheetname) 'Date created - 1
      Call COPYPASTETracking("AV", "B", Sheetname) 'delivery date - 2
      Call COPYPASTETracking("G", "C", Sheetname) 'Job number - 3
      Call COPYPASTETracking("E", "D", Sheetname) 'Type - 4
      Call COPYPASTETracking("J", "E", Sheetname) 'D1 to D3
      Call COPYPASTETracking("L", "H", Sheetname) 'D10
      Call COPYPASTETracking("K", "I", Sheetname) 'D11
      Call COPYPASTETracking("M", "J", Sheetname) 'D12-D14
      Call COPYPASTETracking("N", "L", Sheetname) 'D16-D17
      Call COPYPASTETracking("O", "N", Sheetname) 'D19
      Call COPYPASTETracking("P", "T", Sheetname) 'D29
      Call COPYPASTETracking("Q", "U", Sheetname) 'D31
      Call COPYPASTETracking("R", "W", Sheetname) 'D33
      Call COPYPASTETracking("S", "Z", Sheetname) 'D36
      Call COPYPASTETracking("T", "AA", Sheetname) 'D37
      Call COPYPASTETracking("U", "AB", Sheetname) 'D38
      
      

    '----------------------------------------------------------------------------------

  End With
End Function
Function Copytodatasheet4()
    '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "SFU&MDU Design"
  DataSheetname = "DataProcess"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wsPro As Worksheet
  Set wsPro = wb.Sheets(DataSheetname)
  With wsPro
      .Cells.Clear
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'Call COPYPASTETracking("Column letter from tracking", "letter design for data base", Sheetname)
    'for example D39
    'Call COPYPASTETracking("Column letter from tracking", "AC", Sheetname) 'D38
    
      Call COPYPASTETracking("A", "A", Sheetname) 'Date created - 1
      Call COPYPASTETracking("BD", "B", Sheetname) 'delivery date - 2
      Call COPYPASTETracking("G", "C", Sheetname) 'Job number - 3
      Call COPYPASTETracking("E", "D", Sheetname) 'Type - 4
      Call COPYPASTETracking("J", "E", Sheetname) 'D1 to D3
      Call COPYPASTETracking("K", "F", Sheetname) 'D4 to D6
      Call COPYPASTETracking("L", "G", Sheetname) 'D7 to D9
      Call COPYPASTETracking("N", "H", Sheetname) 'D10
      Call COPYPASTETracking("M", "I", Sheetname) 'D11
      Call COPYPASTETracking("O", "J", Sheetname) 'D12-D14
      Call COPYPASTETracking("R", "L", Sheetname) 'D16-D17
      Call COPYPASTETracking("S", "N", Sheetname) 'D19
      Call COPYPASTETracking("T", "S", Sheetname) 'D28
      Call COPYPASTETracking("U", "U", Sheetname) 'D31
      Call COPYPASTETracking("V", "W", Sheetname) 'D33
      Call COPYPASTETracking("W", "X", Sheetname) 'D34
      Call COPYPASTETracking("P", "Z", Sheetname) 'D36
      Call COPYPASTETracking("Q", "AA", Sheetname) 'D37
      Call COPYPASTETracking("X", "AB", Sheetname) 'D38
      
      
    '----------------------------------------------------------------------------------
  End With
End Function
Function Copytodatasheet5()
    '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "ME Design,Asbuit&Desktop Srvy"
  DataSheetname = "DataProcess"
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wsPro As Worksheet
  Set wsPro = wb.Sheets(DataSheetname)
  With wsPro
      .Cells.Clear
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'Call COPYPASTETracking("Column letter from tracking", "letter design for data base", Sheetname)
    'for example D39
    'Call COPYPASTETracking("Column letter from tracking", "AC", Sheetname) 'D38
    
      Call COPYPASTETracking("A", "A", Sheetname) 'Date created - 1
      Call COPYPASTETracking("AT", "B", Sheetname) 'delivery date - 2
      Call COPYPASTETracking("G", "C", Sheetname) 'Job number - 3
      Call COPYPASTETracking("E", "D", Sheetname) 'Type - 4
      Call COPYPASTETracking("J", "E", Sheetname) 'D1 to D3
      Call COPYPASTETracking("L", "F", Sheetname) 'D4 to D6
      Call COPYPASTETracking("K", "G", Sheetname) 'D7 to D9
      Call COPYPASTETracking("M", "J", Sheetname) 'D12-D14
      Call COPYPASTETracking("N", "N", Sheetname) 'D19
      Call COPYPASTETracking("S", "U", Sheetname) 'D31
      Call COPYPASTETracking("O", "W", Sheetname) 'D33
      Call COPYPASTETracking("P", "Z", Sheetname) 'D36
      Call COPYPASTETracking("Q", "AA", Sheetname) 'D37
      Call COPYPASTETracking("T", "AB", Sheetname) 'D38
      
      
      
    '----------------------------------------------------------------------------------
  End With
End Function
Function ProcessingValues(Count, linespace)
    '***Please note that update is required in this function***
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Outputsheet1 = "Invoice Details" 'wo1
  Outputsheet2 = "Invoice Summary" 'wo2
  DataSheetname = "DataProcess"
  Filename = Workbooks(OutputFileName).Sheets(Outputsheet2).Range("Z1").Value
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wsPro As Worksheet
  Set wsPro = wb.Sheets(DataSheetname)
  Dim win As Worksheet
  Set win = wb.Sheets("Report")
  Dim wbo As Workbook
  Set wbo = Workbooks(OutputFileName)
  Dim wo1 As Worksheet
  Set wo1 = wbo.Sheets(Outputsheet1)
  TotalD = WorksheetFunction.CountA(wb.Sheets("Report").Range("S:S")) - 3
  StartDValue = 6
  DiffDValue = TotalD
  For x = 1 To Count
    linespacedata = x + linespace
    'Details
      wo1.Cells(linespacedata, 3).Value = wsPro.Range("A" & x + 1).Value 'Date created
      wo1.Cells(linespacedata, 4).Value = wsPro.Range("B" & x + 1).Value 'delivery date
      wo1.Cells(linespacedata, 5).Value = wsPro.Range("C" & x + 1).Value 'Job number
      wo1.Cells(linespacedata, 6).Value = wsPro.Range("D" & x + 1).Value 'Type
    'Multiple Column D1 to D3 Templete1
      If wsPro.Range("E" & x + 1).Value >= 1 And wsPro.Range("E" & x + 1).Value <= 2000 Then 'cell(x,A)
        wo1.Cells(linespacedata, StartDValue + 1).Value = 1 ' we adding 3 to push data down in invoice sheet. wo.Cells(row,column)
        wo1.Cells(linespacedata, StartDValue + 2).Value = 0
        wo1.Cells(linespacedata, StartDValue + 3).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 1).Value = win.Cells(4, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 2).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 3).Value = 0
      ElseIf wsPro.Range("E" & x + 1).Value >= 2001 And wsPro.Range("E" & x + 1).Value <= 20000 Then
        wo1.Cells(linespacedata, StartDValue + 1).Value = 1
        wo1.Cells(linespacedata, StartDValue + 2).Value = wsPro.Range("E" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + 3).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 1).Value = win.Cells(4, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 2).Value = win.Cells(5, 20).Value * wo1.Cells(linespacedata, StartDValue + 2).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 3).Value = 0
      ElseIf wsPro.Range("E" & x + 1).Value >= 20001 Then
        wo1.Cells(linespacedata, StartDValue + 1).Value = 1
        wo1.Cells(linespacedata, StartDValue + 2).Value = 0
        wo1.Cells(linespacedata, StartDValue + 3).Value = wsPro.Range("E" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 1).Value = win.Cells(4, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 2).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 3).Value = win.Cells(6, 20).Value * wo1.Cells(linespacedata, StartDValue + 3).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 1).Value = 0
        wo1.Cells(linespacedata, StartDValue + 2).Value = 0
        wo1.Cells(linespacedata, StartDValue + 3).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 1).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 2).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 3).Value = 0
      End If
    'Multiple Column D4 to D6
      If wsPro.Range("F" & x + 1).Value >= 1 And wsPro.Range("F" & x + 1).Value <= 2000 Then
        wo1.Cells(linespacedata, StartDValue + 4).Value = 1
        wo1.Cells(linespacedata, StartDValue + 5).Value = 0
        wo1.Cells(linespacedata, StartDValue + 6).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 4).Value = win.Cells(7, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 5).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 6).Value = 0
      ElseIf wsPro.Range("F" & x + 1).Value >= 2001 And wsPro.Range("F" & x + 1).Value <= 20000 Then
        wo1.Cells(linespacedata, StartDValue + 4).Value = 1
        wo1.Cells(linespacedata, StartDValue + 5).Value = wsPro.Range("F" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + 6).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 4).Value = win.Cells(7, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 5).Value = win.Cells(8, 20).Value * wo1.Cells(linespacedata, StartDValue + 5).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 6).Value = 0
      ElseIf wsPro.Range("F" & x + 1).Value >= 20001 Then
        wo1.Cells(linespacedata, StartDValue + 4).Value = 1
        wo1.Cells(linespacedata, StartDValue + 5).Value = 0
        wo1.Cells(linespacedata, StartDValue + 6).Value = wsPro.Range("F" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 4).Value = win.Cells(7, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 5).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 6).Value = win.Cells(9, 20).Value * wo1.Cells(linespacedata, StartDValue + 6).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 4).Value = 0
        wo1.Cells(linespacedata, StartDValue + 5).Value = 0
        wo1.Cells(linespacedata, StartDValue + 6).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 4).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 5).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 6).Value = 0
      End If
    'Multiple Column D7 to D9
      If wsPro.Range("G" & x + 1).Value >= 1 And wsPro.Range("G" & x + 1).Value <= 2000 Then
        wo1.Cells(linespacedata, StartDValue + 7).Value = 1
        wo1.Cells(linespacedata, StartDValue + 8).Value = 0
        wo1.Cells(linespacedata, StartDValue + 9).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 7).Value = win.Cells(10, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 8).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 9).Value = 0
      ElseIf wsPro.Range("G" & x + 1).Value >= 2001 And wsPro.Range("G" & x + 1).Value <= 20000 Then
        wo1.Cells(linespacedata, StartDValue + 7).Value = 1
        wo1.Cells(linespacedata, StartDValue + 8).Value = wsPro.Range("G" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + 9).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 7).Value = win.Cells(10, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 8).Value = win.Cells(11, 20).Value * wo1.Cells(linespacedata, StartDValue + 8).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 9).Value = 0
      ElseIf wsPro.Range("G" & x + 1).Value >= 20001 Then
        wo1.Cells(linespacedata, StartDValue + 7).Value = 1
        wo1.Cells(linespacedata, StartDValue + 8).Value = 0
        wo1.Cells(linespacedata, StartDValue + 9).Value = wsPro.Range("G" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 7).Value = win.Cells(10, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 8).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 9).Value = win.Cells(12, 20).Value * wo1.Cells(linespacedata, StartDValue + 9).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 7).Value = 0
        wo1.Cells(linespacedata, StartDValue + 8).Value = 0
        wo1.Cells(linespacedata, StartDValue + 9).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 7).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 8).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 9).Value = 0
      End If
    'Single Column D10
      If wsPro.Range("H" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 10).Value = wsPro.Range("H" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 10).Value = win.Cells(13, 20).Value * wsPro.Range("H" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 10).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 10).Value = 0
      End If
    'Single Column D11
      If wsPro.Range("I" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 11).Value = wsPro.Range("I" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 11).Value = win.Cells(14, 20).Value * wsPro.Range("I" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 11).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 11).Value = 0
      End If
    'Multiple Column D12 to D14
      If wsPro.Range("J" & x + 1).Value >= 1 And wsPro.Range("J" & x + 1).Value <= 2000 Then  'Change A, cell(x,A )
        wo1.Cells(linespacedata, StartDValue + 12).Value = 1
        wo1.Cells(linespacedata, StartDValue + 13).Value = 0
        wo1.Cells(linespacedata, StartDValue + 14).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 12).Value = win.Cells(15, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 13).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 14).Value = 0
      ElseIf wsPro.Range("J" & x + 1).Value >= 2001 And wsPro.Range("J" & x + 1).Value <= 20000 Then
        wo1.Cells(linespacedata, StartDValue + 12).Value = 1
        wo1.Cells(linespacedata, StartDValue + 13).Value = wsPro.Range("J" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + 14).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 12).Value = win.Cells(15, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 13).Value = win.Cells(16, 20).Value * wo1.Cells(linespacedata, StartDValue + 13).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 14).Value = 0
      ElseIf wsPro.Range("J" & x + 1).Value >= 20001 Then
        wo1.Cells(linespacedata, StartDValue + 12).Value = 1
        wo1.Cells(linespacedata, StartDValue + 13).Value = 0
        wo1.Cells(linespacedata, StartDValue + 14).Value = wsPro.Range("J" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 12).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 13).Value = win.Cells(15, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 14).Value = win.Cells(17, 20).Value * wo1.Cells(linespacedata, StartDValue + 14).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 12).Value = 0
        wo1.Cells(linespacedata, StartDValue + 13).Value = 0
        wo1.Cells(linespacedata, StartDValue + 14).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 12).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 13).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 14).Value = 0
      End If
    'Single Column D15
      If wsPro.Range("K" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 15).Value = wsPro.Range("K" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 15).Value = win.Cells(18, 20).Value * wsPro.Range("K" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 15).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 15).Value = 0
      End If
    'Multiple Column D16 to D17
      If wsPro.Range("L" & x + 1).Value >= 1 And wsPro.Range("L" & x + 1).Value <= 25 Then 'Change A, cell(x,A )
        wo1.Cells(linespacedata, StartDValue + 16).Value = 1
        wo1.Cells(linespacedata, StartDValue + 17).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 16).Value = win.Cells(19, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 17).Value = 0
      ElseIf wsPro.Range("L" & x + 1).Value >= 26 Then
        wo1.Cells(linespacedata, StartDValue + 16).Value = 1
        wo1.Cells(linespacedata, StartDValue + 17).Value = wsPro.Range("L" & x + 1).Value - 25
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 16).Value = win.Cells(19, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 17).Value = win.Cells(20, 20).Value * wo1.Cells(linespacedata, StartDValue + 17).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 16).Value = 0
        wo1.Cells(linespacedata, StartDValue + 17).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 16).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 17).Value = 0
      End If
    'Single Column D18
      If wsPro.Range("M" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 18).Value = wsPro.Range("M" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 18).Value = win.Cells(21, 20).Value * wsPro.Range("M" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 18).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 18).Value = 0
      End If
    'Single Column D19
      If wsPro.Range("N" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 19).Value = wsPro.Range("N" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 19).Value = win.Cells(22, 20).Value * wsPro.Range("N" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 19).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 19).Value = 0
      End If
    'Multiple Column D24 to D25
      If wsPro.Range("O" & x + 1).Value >= 1 And wsPro.Range("O" & x + 1).Value <= 2000 Then
        wo1.Cells(linespacedata, StartDValue + 20).Value = 1
        wo1.Cells(linespacedata, StartDValue + 21).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 20).Value = win.Cells(23, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 21).Value = 0
      ElseIf wsPro.Range("O" & x + 1).Value >= 2001 Then
        wo1.Cells(linespacedata, StartDValue + 20).Value = 1
        wo1.Cells(linespacedata, StartDValue + 21).Value = wsPro.Range("O" & x + 1).Value - 2000
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 20).Value = win.Cells(23, 20).Value * 1
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 21).Value = win.Cells(24, 20).Value * wo1.Cells(linespacedata, StartDValue + 21).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 20).Value = 0
        wo1.Cells(linespacedata, StartDValue + 21).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 20).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 21).Value = 0
      End If
    'Single Column D26
      If wsPro.Range("P" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 22).Value = wsPro.Range("P" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 22).Value = win.Cells(25, 20).Value * wsPro.Range("P" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 22).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 22).Value = 0
      End If
    'Single Column D26A
    If wsPro.Range("Q" & x + 1).Value <> 0 Then
      wo1.Cells(linespacedata, StartDValue + 23).Value = wsPro.Range("Q" & x + 1).Value
      wo1.Cells(linespacedata, StartDValue + DiffDValue + 23).Value = win.Cells(26, 20).Value * wsPro.Range("Q" & x + 1).Value
    Else
      wo1.Cells(linespacedata, StartDValue + 23).Value = 0
      wo1.Cells(linespacedata, StartDValue + DiffDValue + 23).Value = 0
    End If
    'Single Column D27
      If wsPro.Range("R" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 24).Value = wsPro.Range("R" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 24).Value = win.Cells(27, 20).Value * wsPro.Range("R" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 24).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 24).Value = 0
      End If
    'Single Column D28
    If wsPro.Range("S" & x + 1).Value <> 0 Then
      wo1.Cells(linespacedata, StartDValue + 25).Value = wsPro.Range("S" & x + 1).Value
      wo1.Cells(linespacedata, StartDValue + DiffDValue + 25).Value = win.Cells(28, 20).Value * wsPro.Range("S" & x + 1).Value
    Else
      wo1.Cells(linespacedata, StartDValue + 25).Value = 0
      wo1.Cells(linespacedata, StartDValue + DiffDValue + 25).Value = 0
    End If
    'Single Column D29
      If wsPro.Range("T" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 26).Value = wsPro.Range("T" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 26).Value = win.Cells(29, 20).Value * wsPro.Range("T" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 26).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 26).Value = 0
      End If
    'Single Column D31
      If wsPro.Range("U" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 27).Value = wsPro.Range("U" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 27).Value = win.Cells(30, 20).Value * wsPro.Range("U" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 27).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 27).Value = 0
      End If
    'Single Column D32
      If wsPro.Range("V" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 28).Value = wsPro.Range("V" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 28).Value = win.Cells(31, 20).Value * wsPro.Range("V" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 28).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 28).Value = 0
      End If
    'Single Column D33
      If wsPro.Range("W" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 29).Value = wsPro.Range("W" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 29).Value = win.Cells(32, 20).Value * wsPro.Range("W" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 29).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 29).Value = 0
      End If
    'Single Column D34
      If wsPro.Range("X" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 30).Value = wsPro.Range("X" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 30).Value = win.Cells(33, 20).Value * wsPro.Range("X" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 30).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 30).Value = 0
      End If
    'Single Column D35
      If wsPro.Range("Y" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 31).Value = wsPro.Range("Y" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 31).Value = win.Cells(34, 20).Value * wsPro.Range("Y" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 31).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 31).Value = 0
      End If
    'Single Column D36
      If wsPro.Range("Z" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 32).Value = wsPro.Range("Z" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 32).Value = win.Cells(35, 20).Value * wsPro.Range("Z" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 32).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 32).Value = 0
      End If
    'Single Column D37
      If wsPro.Range("AA" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 33).Value = wsPro.Range("AA" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 33).Value = win.Cells(36, 20).Value * wsPro.Range("AA" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 33).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 33).Value = 0
      End If
    'Single Column D38
      If wsPro.Range("AB" & x + 1).Value <> 0 Then
        wo1.Cells(linespacedata, StartDValue + 34).Value = wsPro.Range("AB" & x + 1).Value
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 34).Value = win.Cells(37, 20).Value * wsPro.Range("AB" & x + 1).Value
      Else
        wo1.Cells(linespacedata, StartDValue + 34).Value = 0
        wo1.Cells(linespacedata, StartDValue + DiffDValue + 34).Value = 0
      End If
    '----------------------------------------------------------------------------------
    '**** Update Required ****
    'adding new column
    ''Single Column D39
    '  If wsPro.Range("AC" & x + 1).Value <> 0 Then
    '    wo1.Cells(linespacedata, StartDValue + 35).Value = wsPro.Range("AC" & x + 1).Value
    '    wo1.Cells(linespacedata, StartDValue + DiffDValue + 35).Value = win.Cells(38, 20).Value * wsPro.Range("AC" & x + 1).Value
    '  Else
    '    wo1.Cells(linespacedata, StartDValue + 35).Value = 0
    '    wo1.Cells(linespacedata, StartDValue + DiffDValue + 35).Value = 0
    '  End If
      
    'Calculating total
      'AO and BV Need to be updated as it will shift
      
      wo1.Cells(linespacedata, DiffDValue + DiffDValue + StartDValue + 1).Value = WorksheetFunction.Sum(wo1.Range("AO" & linespacedata & ":BV" & linespacedata))
      
      
    '----------------------------------------------------------------------------------
  Next
End Function
Function CopytodPriceTracking(PASTE_VALUE)
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Sheetname = "Report"
  DataSheetname = "Invoice Details"
  Sheetname = "Report"
 
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wbo As Workbook
  Set wbo = Workbooks(OutputFileName)
  Dim wsPro As Worksheet
  Set wsPro = wbo.Sheets(DataSheetname)
  With wsPro
    TotalD = WorksheetFunction.CountA(wb.Sheets(Sheetname).Range("S:S"))
    wb.Sheets(Sheetname).Range("T4:T" & TotalD).Copy 'D19
    .Range(PASTE_VALUE).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
  End With
End Function
Function CopytoddetailsTracking(PASTE_VALUE)
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
  Sheetname = "Invoice Summary"
  DataSheetname = "Invoice Details"
  Filename = Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z1").Value
  Dim wb As Workbook
  Set wb = Workbooks(Filename)
  Dim wbo As Workbook
  Set wbo = Workbooks(OutputFileName)
  Dim wsPro As Worksheet
  Set wsPro = wbo.Sheets(DataSheetname)
  With wsPro
    TotalD = WorksheetFunction.CountA(wb.Sheets("Report").Range("S:S"))
    wbo.Sheets(Sheetname).Range("C14:C" & TotalD + 10).Copy
    .Range(PASTE_VALUE).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
  End With
End Function

Function pagesetupsetting()
  filedate = Format(Date, "ddmmyyyy")
  OutputFileName = "Ika-Invoice " & filedate & ".xlsx"
    With Workbooks(OutputFileName).Sheets("Invoice Summary").PageSetup
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




