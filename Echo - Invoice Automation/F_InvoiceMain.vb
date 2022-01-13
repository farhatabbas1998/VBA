'Version: 1.00
'Ika - Invoice main
'For Echobroadband
'User: Ika
'By Farhat Abbas & Ika


Sub Invoicemain()

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
  Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).FormulaR1C1 = Filename
  F_InvoiceFrontend.Show
  F_InvoiceBackend.main
  Call Invoice_Image(ThisWorkbook.Path & "\Echologo.png")
  Call DeleteDataSheet(Filename, "DataProcess")
  Workbooks(OutputFileName).Worksheets("Invoice Summary").Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(OutputFileName).Worksheets("Invoice Details").Activate
  ActiveWindow.DisplayGridlines = False
  Workbooks(OutputFileName).Sheets("Invoice Summary").Range("Z" & 1).Delete
  Application.DisplayAlerts = True
  Workbooks(OutputFileName).Save
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
 Function Invoice_Image(imagePath)
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
  'Image File
 
 
   Dim imgLeft As Double
   Dim imgTop As Double
 
  'Writing Values
   With wo1
     imgLeft = .Range("B2").Left
     imgTop = .Range("B2").Left
     wo1.Shapes.AddPicture Filename:=imagePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=imgLeft, Top:=imgTop, Width:=-1, Height:=-1
   End With
 End Function
 