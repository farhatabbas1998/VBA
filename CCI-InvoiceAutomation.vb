'Version: 1.00
'For Echobroadband
'By Farhat Abbas (Verified & Tested)

Sub Comcast_Billing()
  Application.ScreenUpdating = False
  Application.DisplayAlerts = False
  Masterwb = ThisWorkbook.Name 'Master Macro workbook name
  Dataws = "Billing_Comcast"
  Sparews = "Sparews"
  OSparews = "Sheet1"
  'Online tracking
  strUrl = Workbooks(Masterwb).Sheets(Dataws).Range("D2").Value
  Set wb = Application.Workbooks.Open(strUrl)
  Trackingwb = wb.Name 'Tracking workbook name
  strPath = ThisWorkbook.Path & "\Input files\" & Trackingwb
  wb.SaveAs Filename:=strPath
  region = Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value
  Invoiceno = Workbooks(Masterwb).Sheets(Dataws).Cells(6, 9).Value
  Outputwb = Workbooks(Masterwb).Sheets(Dataws).Cells(7, 9).Value
  Outputws1 = Workbooks(Masterwb).Sheets(Dataws).Cells(8, 9).Value
  Outputws2 = Workbooks(Masterwb).Sheets(Dataws).Cells(9, 9).Value
  Outputws3 = Workbooks(Masterwb).Sheets(Dataws).Cells(10, 9).Value
  Outputws4 = Workbooks(Masterwb).Sheets(Dataws).Cells(11, 9).Value
  Outputws5 = Workbooks(Masterwb).Sheets(Dataws).Cells(12, 9).Value
  Outputws6 = Workbooks(Masterwb).Sheets(Dataws).Cells(13, 9).Value
  Trackingws1 = Workbooks(Masterwb).Sheets(Dataws).Cells(14, 9).Value
  Trackingws2 = Workbooks(Masterwb).Sheets(Dataws).Cells(15, 9).Value
  Trackingws3 = Workbooks(Masterwb).Sheets(Dataws).Cells(16, 9).Value
  Trackingws4 = Workbooks(Masterwb).Sheets(Dataws).Cells(17, 9).Value
  Trackingws5 = Workbooks(Masterwb).Sheets(Dataws).Cells(18, 9).Value

  Call Check_if_workbook_is_open(Outputwb)
  Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\Output files\" & Outputwb, CreateBackup:=False

  Call CheckDataSheet(Outputwb, Outputws1)
  Call CheckDataSheet(Outputwb, Outputws2)
  Call CheckDataSheet(Outputwb, Outputws3)
  Call CheckDataSheet(Outputwb, Outputws4)
  Call CheckDataSheet(Outputwb, Outputws5)
  Call CheckDataSheet(Outputwb, Outputws6)
  Call CheckDataSheet(Masterwb, Sparews)
  Call Invoice_Summary_Templete(Masterwb, Dataws, Outputwb, Outputws1, region, Invoiceno)
  Call CreatingODetailOutput(Masterwb, Dataws, Sparews, OSparews, Trackingwb, Trackingws1, Trackingws2, Trackingws3, Trackingws4, Trackingws5, Outputwb, Outputws1, Outputws2, Outputws3, Outputws4, Outputws5, Outputws6, region, Invoiceno)
  Call DeleteDataSheet(Masterwb, Sparews)
  Call DeleteDataSheet(Outputwb, OSparews)
  Call pagesetupsetting(Outputwb, Outputws1)
  Workbooks(Masterwb).Sheets(Dataws).Columns("AE:AF").ClearContents
  Call Check_if_workbook_is_open(Trackingwb)
  
  Workbooks(Outputwb).Worksheets(Outputws1).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Worksheets(Outputws2).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Worksheets(Outputws3).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Worksheets(Outputws4).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Worksheets(Outputws5).Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(Outputwb).Worksheets(Outputws6).Activate
  ActiveWindow.DisplayGridlines = False
  
  Workbooks(Outputwb).Save
  Application.DisplayAlerts = True
  Application.ScreenUpdating = True
End Sub
Function CreatingODetailOutput(Masterwb, Dataws, Sparews, OSparews, Trackingwb, Trackingws1, Trackingws2, Trackingws3, Trackingws4, Trackingws5, Outputwb, Outputws1, Outputws2, Outputws3, Outputws4, Outputws5, Outputws6, region, Invoiceno)
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
    Dim ows3 As Worksheet
    Dim ows4 As Worksheet
    Dim ows5 As Worksheet
    Dim ows6 As Worksheet
    Set osws = owb.Sheets(OSparews)
    Set ows1 = owb.Sheets(Outputws1)
    Set ows2 = owb.Sheets(Outputws2)
    Set ows3 = owb.Sheets(Outputws3)
    Set ows4 = owb.Sheets(Outputws4)
    Set ows5 = owb.Sheets(Outputws5)
    Set ows6 = owb.Sheets(Outputws6)
    Dim dilverytype As Variant
    Dim Commtype1   As Variant
    Dim Commtype2   As Variant
    Dim Commtype3   As Variant
    Dim Commtype4   As Variant
    Dim Resitype1   As Variant
    Dim Resitype2   As Variant
    Dim Metroetype1 As Variant
    Dim Metroetype2 As Variant
    Dim NStype1     As Variant
    Dim FDtype1     As Variant
    Call StopAllFilters(Trackingwb)
    Call StartFilter(Trackingwb, Trackingws1)
    Call StartFilter(Trackingwb, Trackingws2)
    Call StartFilter(Trackingwb, Trackingws3)
    Call StartFilter(Trackingwb, Trackingws4)
    Call StartFilter(Trackingwb, Trackingws5)
    
  '----------------------------------------------------------------------------
    'job types
      Commtype1 = Array("Comm Design", "Hyperbuilt Design", "Span Replacement Design", "Probuild Design")
      Commtype2 = Array("Comm Asbuilt", "Span Replacement Asbuilt", "Hyperbuilt Asbuilt")
      Commtype3 = Array("Force Relocation Asbuilt,Force Relocation Design")
      Commtype4 = Array("Spatial Asbuilt")
      Resitype1 = Array("Res Design")
      Resitype2 = Array("Res Asbuilt")
      Metroetype1 = Array("Metro E Design")
      Metroetype2 = Array("Metro E Asbuilt", "Fiber Maintenance")
      NStype1 = Array("NS Design", "NS Asbuilt")
      FDtype1 = Array("FD Design", "FD Asbuilt", "Res design (FD)")

    'Status
      dilverytype = Array("Completed", "Completed 2", "Completed  2 (Valid)")

   'Table Naming
      CommTN1 = "Commercial Design"
      CommTN2 = "Commercial Asbuilt"
      CommTN3 = "Commercial Force Relocate"
      CommTN4 = "Spatial As built New Node"
      ResiTN1 = "Residential Design"
      ResiTN2 = "Residential Asbuilt"
      MetroeTN1 = "Metro E Design"
      MetroeTN2 = "Metro E  Asbuilt"
      NSTN1 = "Node Split"
      FDTN1 = "Fiber Deep"
    'X and Y shifting and space gap
      Yspacing = 3
      Xspacing = 4
      diff = 5
   '----------------------------------------------------------------------------

   TotalD = WorksheetFunction.CountA(dws.Range("C:C")) - 2
   Startingdate = Format(Date - 7, "\>\=mm/dd/yyyy")
   Endingdate = Format(Date - 1, "\>\=mm/dd/yyyy")
   'Sheet 1
      With tws2.Range("A1")
        x = 1
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Commtype1, Operator:=xlFilterValues  'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws2, Masterwb, Sparews, 1)
      CountCom1 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 1, Xspacing, Yspacing, CommTN1, CountCom1)
      '--------------------------------2
      With tws1.Range("A1")
        x = 0
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Commtype2, Operator:=xlFilterValues  'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws1, Masterwb, Sparews, 0)
      CountCom2 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 2, Xspacing + diff + CountCom1, Yspacing, CommTN2, CountCom2)
      '--------------------------------3
      With tws1.Range("A1")
        x = 0
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Commtype3, Operator:=xlFilterValues  'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws1, Outputwb, OSparews, 0)
      CCom1 = WorksheetFunction.CountA(osws.Range("C:C"))
      With tws2.Range("A1")
        x = 1
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Commtype4, Operator:=xlFilterValues  'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws2, Masterwb, Sparews, 1)
      CCom2 = WorksheetFunction.CountA(sws.Range("C:C"))
      If CCom1 >= 2 Then
       TotalEntires = WorksheetFunction.CountA(dws.Range("T:T"))
       osws.Range(osws.Cells(6, 20).Value & "2:" & osws.Cells(6 + TotalEntires, 20).Value & CCom1).Copy
       sws.Range("A" & CCom2 + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      End If
      CountCom3 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 3, Xspacing + diff * 2 + CountCom1 + CountCom2, Yspacing, CommTN3, CountCom3)
      '--------------------------------4
      With tws1.Range("A1")
        x = 0
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Commtype4, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws1, Masterwb, Sparews, 0)
      CountCom4 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws2, Outputws1, 4, Xspacing + diff * 3 + CountCom1 + CountCom2 + CountCom3, Yspacing, CommTN4, CountCom4)
   '-----------------------------------------------------------------------
   'Sheet2
      With tws2.Range("A1")
        x = 1
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Resitype1, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws2, Masterwb, Sparews, 1)
      CountRes1 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws3, Outputws1, 1, Xspacing, Yspacing, ResiTN1, CountRes1)
      '--------------------------------2
      With tws1.Range("A1")
        x = 0
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Resitype2, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws1, Masterwb, Sparews, 0)
      CountRes2 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws3, Outputws1, 2, Xspacing + diff + CountRes1, Yspacing, ResiTN2, CountRes2)
   '-----------------------------------------------------------------------
   'Sheet3
      With tws3.Range("A1")
        x = 2
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Metroetype1, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws3, Masterwb, Sparews, 2)
      CountMetroe1 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws4, Outputws1, 5, Xspacing, Yspacing, MetroeTN1, CountMetroe1)
      '--------------------------------2
      With tws3.Range("A1")
         x = 2
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=Metroetype2, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws3, Masterwb, Sparews, 2)
      CountMetroe2 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws4, Outputws1, 6, Xspacing + diff + CountMetroe1, Yspacing, MetroeTN2, CountMetroe2)
   '-----------------------------------------------------------------------
   'Sheet4
      With tws4.Range("A1")
        x = 3
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=NStype1, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws4, Masterwb, Sparews, 3)
      CountNS1 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws5, Outputws1, 7, Xspacing, Yspacing, NSTN1, CountNS1)
   '--------------------------------2
   'Sheet5
      With tws5.Range("A1")
        x = 4
        .AutoFilter Field:=dws.Cells(6, 23 + x).Value, Criteria1:=dilverytype, Operator:=xlFilterValues 'Dilevery Status
        .AutoFilter Field:=dws.Cells(7, 23 + x).Value, Operator:=xlFilterValues, Criteria1:=Startingdate, Criteria2:=Endingdate
        .AutoFilter Field:=dws.Cells(8, 23 + x).Value, Criteria1:="=", Operator:=xlFilterValues 'Invoice colum
        .AutoFilter Field:=dws.Cells(9, 23 + x).Value, Criteria1:=Workbooks(Masterwb).Sheets(Dataws).Cells(5, 9).Value  'Region
        .AutoFilter Field:=dws.Cells(10, 23 + x).Value, Criteria1:=FDtype1, Operator:=xlFilterValues   'Job Type
      End With
      Call StoringData(Masterwb, Dataws, Trackingwb, Trackingws5, Masterwb, Sparews, 4)
      CountFD1 = WorksheetFunction.CountA(sws.Range("C:C"))
      Call ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws6, Outputws1, 8, Xspacing, Yspacing, FDTN1, CountFD1)
      ows1.Range("I" & dws.Cells(1, 32).Value).Value = WorksheetFunction.Sum(ows1.Range("I" & dws.Cells(1, 31).Value - 1 & ":I" & dws.Cells(1, 32).Value - 1))
End Function
Function Invoice_Summary_Templete(Masterwb, Dataws, Outputwb, Outputws, region, Invoiceno)
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
      .Range("I1").FormulaR1C1 = "Invoice"
      .Range("H2").FormulaR1C1 = "DATE"
      .Range("H3").FormulaR1C1 = Date
      .Range("H3").NumberFormat = "d-mmm-yyyy"
      .Range("I2").FormulaR1C1 = "INVOICE #"
      .Range("I3").FormulaR1C1 = Invoiceno
      .Range("H5").FormulaR1C1 = "BILL TO"
      .Range("H6").FormulaR1C1 = "Echo Broadband, Inc"
      .Range("H7").FormulaR1C1 = "PO Box 1627"
      .Range("H8").FormulaR1C1 = "Broomfield, CO 80038"
      .Range("B11").FormulaR1C1 = "PROJECT: ARRIS / Comcast"
      .Range("B12").FormulaR1C1 = "Statement of Work Number : Comcast Design and Asbuild"
      .Range("B6").FormulaR1C1 = "ECHO Broadband Sdn Bhd"
      .Range("B7").FormulaR1C1 = "368-5-3 Bellisa Row"
      .Range("B8").FormulaR1C1 = "Jalan Burmah"
      .Range("B9").FormulaR1C1 = "10350 Penang"
      .Range("E12").FormulaR1C1 = "** Data Processing and Provision of information**"
      .Range("I11").FormulaR1C1 = "Terms"
      .Range("I12").FormulaR1C1 = "Net 90"
      .Range("H11").FormulaR1C1 = "PO No."
      .Range("B13").FormulaR1C1 = "Item"
      .Range("C13").FormulaR1C1 = "Description"
      .Range("D13").FormulaR1C1 = "Rate"
      .Range("E13").FormulaR1C1 = "Unit"
      .Range("F13").FormulaR1C1 = "Unit Price (USD)"
      .Range("G13").FormulaR1C1 = "QTY"
      .Range("H13").FormulaR1C1 = "Sub Total"
      .Range("I13").FormulaR1C1 = "Scope Totals"
      .Range("B14").FormulaR1C1 = "Region:"
      .Range("C14").FormulaR1C1 = region
      Call Thin_OB("H5:I5,H6:I9", Outputwb, Outputws)
      Call ThinBoarderInv("H2:I3,H11:I12,B13:I13,B14", Outputwb, Outputws)
      .Range("H5:I5,H6:I6,H7:I7,H8:I8,H9:I9").Merge
    'calling function
      For x = 0 To Noofinvtab - 1
        Sgap = Sgap + 2
        .Range("C" & TotalD * x + Gap + 15).FormulaR1C1 = dws.Cells(6 + x, 29).Value
        .Range("C" & TotalD * x + Gap + 15).Font.Bold = True
        Call Copytovalueinvoice(Masterwb, Dataws, Outputwb, Outputws, "B" & TotalD * x + Gap + 16)
        Call ThinBoarderInv("B" & TotalD * x + Gap + 15 & ":I" & TotalD + TotalD * x + Gap + 15, Outputwb, Outputws)
        Call Thick_OB("B" & TotalD * x + Gap + 15 & ":I" & TotalD + TotalD * x + Gap + 15, Outputwb, Outputws)
        .Range("I" & TotalD * x + Gap + 15 & ":I" & TotalD + TotalD * x + Gap + 15).Merge
        dws.Cells(x + 1, 31).Value = TotalD * x + Gap + 16
        Gap = Gap + 2
      Next
      Call Thick_OB("B11:G12,H11:H12,I11:I12,B13:I13,B14:I" & TotalD * Noofinvtab + Gap + 14 & ",B" & TotalD * Noofinvtab + Gap + 14 & ":H" & TotalD * Noofinvtab + Gap + 14 & ",I" & TotalD * Noofinvtab + Gap + 14, Outputwb, Outputws)
      .Range("F:F,H15:I" & TotalD * Noofinvtab + Gap + 14).NumberFormat = "[$$-en-US]#,##0.00"
      .Range("H" & TotalD * Noofinvtab + Gap + 14).FormulaR1C1 = "Invoice Total (USD)"
      dws.Cells(1, 32).Value = TotalD * Noofinvtab + Gap + 14
    'format
      .Range("A1:I15,I:I,H" & TotalD * Noofinvtab + Gap + 14).Font.Bold = True
      .Range("E12").Font.Bold = False
      .Range("A1:I" & TotalD * Noofinvtab + Gap + 14).Font.Size = 12
      .Range("B7:B9,I15:I" & TotalD * Noofinvtab + Gap + 14).Font.Size = 11
      .Range("B14:C14").Font.Size = 16
      .Range("I1").Font.Size = 24
      .Range("I3").Font.Name = "Arial"
      With .Range("H5").Font
      .Name = "Arial Black"
      .Size = 14
      End With
      With .Range("H6:H8").Font
      .Name = "Arial"
      .Size = 10
      End With
      .Columns("A").ColumnWidth = 0.88
      .Columns("B").ColumnWidth = 8
      .Columns("C").ColumnWidth = 60
      .Columns("D").ColumnWidth = 5
      .Columns("E").ColumnWidth = 11.25
      .Columns("F").ColumnWidth = 9
      .Columns("G").ColumnWidth = 9.5
      .Columns("H").ColumnWidth = 15.25
      .Columns("I").ColumnWidth = 22.43
      
      With .Range("H2:I5,E11:I12,B13:I13,B15:I" & TotalD * Noofinvtab + Gap + 14)
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .ReadingOrder = xlContext
      End With
      With .Range("I1,H" & TotalD * Noofinvtab + Gap + 14)
          .HorizontalAlignment = xlRight
          .VerticalAlignment = xlCenter
          .ReadingOrder = xlContext
      End With
      With .Range("C15:C" & TotalD * Noofinvtab + Gap + 14)
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlCenter
          .WrapText = True
          .ReadingOrder = xlContext
      End With
      .Range("B13:I13").WrapText = True
      End With
End Function
Function ProcessCellValue(Masterwb, Dataws, Sparews, Outputwb, Outputws, Outputws2, OutputValue, XShiftBlank, YShiftBlank, Tabletitle, Count)
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
      If scaling <= DTask - 1 Then
        ows2.Range("G" & dws.Cells(OutputValue, 31).Value + scaling).Value = TotalPrice + ows2.Range("G" & dws.Cells(OutputValue, 31).Value + scaling).Value
        scaling = 1 + scaling
      ElseIf scaling > DTask - 1 And scaling + scaling2 <= DTask * 2 - 1 Then
        ows2.Range("H" & dws.Cells(OutputValue, 31).Value + scaling2).Value = TotalPrice + ows2.Range("H" & dws.Cells(OutputValue, 31).Value + scaling2).Value
        scaling2 = 1 + scaling2
      Else
        ows2.Range("I" & dws.Cells(OutputValue, 31).Value - 1).Value = TotalPrice + ows2.Range("I" & dws.Cells(OutputValue, 31).Value - 1).Value
      End If
    Next
    For y = YShiftBlank To STotal + YShiftBlank - 1
      Titledes = 1 + Titledes
      ows.Cells(XShiftBlank + 1, YShiftBlank + Titledes).Value = dws.Cells(5 + Titledes, 11).Value
    Next
    headerdes = 0
    For y = STotal + YShiftBlank To YShift - 1
      headerdes = 1 + headerdes
      ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + headerdes).Value = dws.Cells(4 + headerdes, 3).Value
      ows.Cells(XShiftBlank, STotal + YShiftBlank + headerdes).Value = dws.Cells(4 + headerdes, 4).Value
      ows.Cells(XShiftBlank, STotal + YShiftBlank + DTask + headerdes).Value = dws.Cells(4 + headerdes, 4).Value
      ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + DTask + headerdes).Value = dws.Cells(4 + headerdes, 6).Value
    Next
    ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + DTask + headerdes + 1).Value = "Subtotal"
    ows.Cells(XShiftBlank + 1, STotal + YShiftBlank + DTask + headerdes + 2).Value = "Remarks"
    ows.Cells(XShiftBlank + Count + 2, 7).Value = "Totals:"
    Call DetailTable(Outputwb, Outputws, 1, XShiftBlank + 1, YShiftBlank, XShiftBlank + Count + 1, STotal + YShiftBlank)
    Call DetailTable(Outputwb, Outputws, 1, XShiftBlank + 1, YShift + DTask, XShiftBlank + Count + 2, YShift + DTask + 2)
    Call DetailTable(Outputwb, Outputws, 2, XShiftBlank, STotal + YShiftBlank + 1, XShiftBlank + Count + 1, YShift)
    Call DetailTable(Outputwb, Outputws, 2, XShiftBlank, YShift + 1, XShiftBlank + Count + 2, YShift + DTask)
    ows.Range(XShiftBlank & ":" & XShiftBlank).Font.Bold = True
    ows.Range(Cells(XShiftBlank + 1, YShiftBlank + 1).Address, Cells(XShiftBlank + 1, YShiftBlank + 5).Address).Font.Bold = True
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
    wb.Sheets(Dataws).Range("B5:F" & TotalD).Copy
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



