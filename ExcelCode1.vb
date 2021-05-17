
Sub EOLSHEET()

Dim sDate As Long
Dim eDate As Long
Dim GetValue As String





Workbooks.Add
x = InputBox("Give me some date")
'RegionBF = InputBox("Which region! the long name")
'RegionSF = InputBox("Which region! the Short name(CAPS)")
OutputFileName = InputBox("Give me name of file") + ".xls"



Workbooks.Add.SaveAs Filename:="C:\Users\farhat\Desktop\" + OutputFileName
Sheetname_2 = "EOL Sheet"

Workbooks(OutputFileName).Sheets(1).Name = Sheetname_2


eDate = Date - 1

sDate = Date - x




MsgBox "SDate is: " & Format(sDate, "d mmmm, yyyy")
MsgBox "EDate is: " & Format(eDate, "d mmmm, yyyy")
Filename = "ECHO-NETWIN DESIGN Tracking List 2021 (13).xlsx"
Sheetname_1 = "DESIGN"
Sheetname_3 = "REDESIGN"
Workbooks.Open "C:\Users\farhat\Downloads\" + Filename 'Open Data file location



'Filterring data


Workbooks(Filename).Sheets(Sheetname_1).Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, _
                                                Criteria1:=">=" & Format(sDate, "d/mmmm/yyyy"), _
                                                Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                        
                        

GetValue = Workbooks(Filename).Sheets(Sheetname_1).Range("A:A").Value

MsgBox GetValue
                        
                        


Workbooks(Filename).Worksheets(Sheetname_1).Columns("P:P").Copy 'Actual Delivery
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("C4").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    
                        
Workbooks(Filename).Worksheets(Sheetname_1).Columns("G:G").Copy 'NETWIN CELL NAME
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("E4").PasteSpecial Paste:=xlPasteValues

Workbooks(Filename).Worksheets(Sheetname_1).Columns("D:D").Copy 'TOWN
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("D4").PasteSpecial Paste:=xlPasteValues

Dim lRow As Long
Dim lCol As Long

    'Find the last non-blank cell in column A(1)
    lRow = Workbooks(OutputFileName).Worksheets(Sheetname_2).Cells(Rows.Count, 3).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Workbooks(OutputFileName).Worksheets(Sheetname_2).Cells(3, Columns.Count).End(xlToLeft).Column
    LastCount = lRow - 3
    Firstnumber = 1

   
'     MsgBox "Last Row: " & LastCount
     NumSeq = "B" & LastCount

    
      NumEndB = "B5:B" & lRow
      NumEndF = "F5:F" & lRow
      NumEndG = "G5:G" & lRow
      Table = "B5:G" & lRow
      Table3 = "C3:C7"
      

Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("A4:G4").FormulaR1C1 = " "

Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("F5").FormulaR1C1 = "Complete"

Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("F5").FormulaR1C1 = "Complete"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("G5").FormulaR1C1 = "Sent Through RFI Sites"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("C3").FormulaR1C1 = "Date Delivered"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("D3").FormulaR1C1 = "Town"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("E3").FormulaR1C1 = "Netwin Cell Name"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("F3").FormulaR1C1 = "EOL Sheets"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("G3").FormulaR1C1 = "Remarks"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("A3:G3").Font.FontStyle = "Bold"






                        
                        

    
    Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
    Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
    With Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(Table).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
        


   
'Workbooks(Filename).Close SaveChanges:=True 'Closing data file



Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B5:B5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndB), Type:=xlFillSeries
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("F5:F5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndF), Type:=xlCopySeries
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("G5:G5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndG), Type:=xlCopySeries

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
Workbooks(OutputFileName).Sheets.Add.Name = Sheetname_3

Workbooks(Filename).Sheets(Sheetname_3).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                Criteria1:="=" & "CT"
                                                
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("C3").FormulaR1C1 = "Date Delivered"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("D3").FormulaR1C1 = "Town"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("E3").FormulaR1C1 = "Netwin Cell Name"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("F3").FormulaR1C1 = "HTTP Design & Draft"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("G3").FormulaR1C1 = "Schematic Drawings"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("H3").FormulaR1C1 = "BOM & EOL Sheets"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("I3").FormulaR1C1 = "Splicing Matrix"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("J3").FormulaR1C1 = "Remarks"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("F5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("G5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("H5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("J5").FormulaR1C1 = "Sent Through RFI Sites"

Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("B5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_3).Range("A3:G3").Font.FontStyle = "Bold"
                                                
                                                
                                                

End Sub









