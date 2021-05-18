Sub USL()


    'version 1.01

    'inputs dates, file name & region
    '18 may 2021
    'Create a new sheet in Inputfile and name it "VBA" leave it blank can be deleted later!
    'replace save file location and open file location line and as well as name



    Dim sDate As Long
Dim eDate As Long
Dim GetValue As String



Workbooks.Add
x = InputBox("Give me number of day")

OutputFileName = InputBox("Give me new file name") + ".xlsx"

Workbooks.Add.SaveAs Filename:="E:\OneDrive\Desktop\ExcelTestFiles\" + OutputFileName
Sheetname_2 = "EOL Sheet"

Workbooks(OutputFileName).Sheets(1).Name = Sheetname_2


eDate = Date - 1

sDate = Date - x





MsgBox "SDate is: " & Format(sDate, "d mmmm, yyyy")
'MsgBox "EDate is: " & Format(eDate, "d mmmm, yyyy")
Filename = "ECHO-NETWIN DESIGN Tracking List 2021 (13).xlsx"
Sheetname_1 = "DESIGN"
Sheetname_3 = "REDESIGN"
Sheetname_5 = "VBA"
Workbooks.Open "D:\Downloads\" + Filename 'Open Data file location


'Workbooks(Filename).Sheets.Add.Name = Sheetname_5

'Filterring data

Workbooks(Filename).Sheets(Sheetname_1).ShowAllData
Workbooks(Filename).Sheets(Sheetname_3).ShowAllData


Workbooks(Filename).Sheets(Sheetname_1).Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, _
                                                Criteria1:=">=" & Format(sDate, "d/mmmm/yyyy"), _
                                                Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                                                
Workbooks(Filename).Worksheets(Sheetname_1).Columns("A:A").Copy 'Actual Delivery
Workbooks(Filename).Worksheets(Sheetname_5).Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Dim r As Range, ary
    Set r = Workbooks(Filename).Worksheets(Sheetname_5).Columns("A:A")
    With Application
        MsgBox .TextJoin(" ", True, r)
    End With

                                                
                                                
region = InputBox("Give me Region for EOL Sheet")
region = UCase(region)
                                                
                                                
                        
Workbooks(Filename).Sheets(Sheetname_1).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                Criteria1:="=" & region
                                                    


'MsgBox GetValue
                        
'MsgBox "Region Availble is " & Workbooks(Filename).Sheets(Sheetname_1).Range("A:A").Value


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
    NewCount = lRow + 2
    NumEndD = "D" & NewCount
    NumEndE = "E" & NewCount
    LastCount = lRow - 4
    
'    MsgBox lRow
      
      

    
    NumEndB = "B5:B" & lRow
    NumEndF = "F5:F" & lRow
    NumEndG = "G5:G" & lRow
    GEnds = "A3:G" & lRow
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
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndD).FormulaR1C1 = "Total Cells: "
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndE).FormulaR1C1 = LastCount
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndD).Font.Color = vbBlue
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndE).Font.Color = vbBlue
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndE).Font.FontStyle = "Bold"
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndD).Font.FontStyle = "Bold"





  If region = "CT" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").FormulaR1C1 = "Connecticut Region" 'more state add here
  ElseIf region = "LIE" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").FormulaR1C1 = "Michigan Region"
  ElseIf region = "NJN" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").FormulaR1C1 = "New Jersey North  Region"
  ElseIf region = "NJS" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").FormulaR1C1 = "New Jersey South  Region"
  ElseIf region = "NYC" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").FormulaR1C1 = "New York City Region"
  ElseIf region = "WC" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").FormulaR1C1 = "Connecticut"
  End If
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2:G3").Font.Size = 12
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").Font.Color = vbBlue
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B2").Font.FontStyle = "Bold"
                        

    
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
        


   


If Count < 1 Then

Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("B5:B5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndB), Type:=xlFillSeries
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("F5:F5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndF), Type:=xlCopySeries
Workbooks(OutputFileName).Worksheets(Sheetname_2).Range("G5:G5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_2).Range(NumEndG), Type:=xlCopySeries
End If


Workbooks(OutputFileName).Worksheets(Sheetname_2).Columns("A:W").EntireColumn.AutoFit

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
Sheetname_4 = "HTTP design"
Workbooks(OutputFileName).Sheets.Add.Name = Sheetname_4

Workbooks(Filename).Sheets(Sheetname_3).Range("P:P").AutoFilter Field:=16, Operator:=xlFilterValues, _
                                                Criteria1:=">=" & Format(sDate, "d/mmmm/yyyy"), _
                                                Criteria2:="<=" & Format(eDate, "d/mmmm/yyyy")
                                                

Workbooks(Filename).Worksheets(Sheetname_3).Columns("A:A").Copy 'Actual Delivery
Workbooks(Filename).Worksheets(Sheetname_5).Range("B1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    Set r = Workbooks(Filename).Worksheets(Sheetname_5).Columns("B:B")
    With Application
        MsgBox .TextJoin(" ", True, r)
    End With

                                                
                                                
region = InputBox("Give me Region for HTTP design")
region = UCase(region)
                                                                         
                                                
Workbooks(Filename).Sheets(Sheetname_3).Range("A:A").AutoFilter Field:=1, Operator:=xlFilterValues, _
                                                Criteria1:="=" & region

                                                
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("C3").FormulaR1C1 = "Date Delivered"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("D3").FormulaR1C1 = "Town"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("E3").FormulaR1C1 = "Netwin Cell Name"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("F3").FormulaR1C1 = "HTTP Design & Draft"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("G3").FormulaR1C1 = "Schematic Drawings"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("H3").FormulaR1C1 = "BOM & EOL Sheets"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("I3").FormulaR1C1 = "Splicing Matrix"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("J3").FormulaR1C1 = "Remarks"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("F5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("G5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("H5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("J5").FormulaR1C1 = "Sent Through RFI Sites"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("I5").FormulaR1C1 = "1"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("A3:J3").Font.FontStyle = "Bold"

Workbooks(Filename).Worksheets(Sheetname_3).Columns("P:P").Copy 'Actual Delivery
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("C4").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    
                        
Workbooks(Filename).Worksheets(Sheetname_3).Columns("G:G").Copy 'NETWIN CELL NAME
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("E4").PasteSpecial Paste:=xlPasteValues

Workbooks(Filename).Worksheets(Sheetname_3).Columns("D:D").Copy 'TOWN
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("D4").PasteSpecial Paste:=xlPasteValues

    'Find the last non-blank cell in column A(1)
     lRow = Workbooks(OutputFileName).Worksheets(Sheetname_4).Cells(Rows.Count, 3).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
     lCol = Workbooks(OutputFileName).Worksheets(Sheetname_4).Cells(3, Columns.Count).End(xlToLeft).Column
     LastCount = lRow - 4
     Firstnumber = 1

   
    'MsgBox "Last Row: " & LastCount
      NumSeq = "B" & LastCount

      NewCount = lRow + 2
      NumEndD = "D" & NewCount
      NumEndE = "E" & NewCount
    
      NumEndB = "B5:B" & lRow
      NumEndF = "F5:F" & lRow
      NumEndG = "G5:G" & lRow
      NumEndH = "H5:H" & lRow
      NumEndI = "I5:I" & lRow
      NumEndJ = "J5:J" & lRow
      Table = "B5:J" & lRow
      Table3 = "C3:C7"
      JEnds = "A3:J" & lRow
If Count < 1 Then
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B5:B5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndB), Type:=xlFillSeries
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("F5:F5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndF), Type:=xlCopySeries
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("G5:G5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndG), Type:=xlCopySeries
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("H5:H5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndH), Type:=xlCopySeries
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("I5:I5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndI), Type:=xlCopySeries
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("J5:J5").AutoFill Destination:=Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndJ), Type:=xlCopySeries
End If

'Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("A4:J7").Clear

    Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlDiagonalDown).LineStyle = xlNone
    Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlDiagonalUp).LineStyle = xlNone
    With Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(Table).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndD).FormulaR1C1 = "Total Cells: "
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndD).Font.Color = vbBlue
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndE).FormulaR1C1 = LastCount
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndE).Font.Color = vbBlue
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndE).Font.FontStyle = "Bold"
Workbooks(OutputFileName).Worksheets(Sheetname_4).Range(NumEndD).Font.FontStyle = "Bold"




  If region = "CT" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").FormulaR1C1 = "Connecticut Region" 'more state add here
  ElseIf region = "LIE" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").FormulaR1C1 = "Michigan Region"
  ElseIf region = "NJN" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").FormulaR1C1 = "New Jersey North  Region"
  ElseIf region = "NJS" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").FormulaR1C1 = "New Jersey South  Region"
  ElseIf region = "NYC" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").FormulaR1C1 = "New York City Region"
  ElseIf region = "WC" Then
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").FormulaR1C1 = "Connecticut"
  End If
  
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2:J3").Font.Size = 12
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2:J3").Font.FontStyle = "Bold"
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("B2").Font.Color = vbBlue
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Range("A4:J4").Clear
  Workbooks(OutputFileName).Worksheets(Sheetname_4).Columns("A:W").HorizontalAlignment = xlCenter
  Workbooks(OutputFileName).Worksheets(Sheetname_2).Columns("A:W").HorizontalAlignment = xlCenter

Workbooks(OutputFileName).Worksheets(Sheetname_4).Columns("A:W").EntireColumn.AutoFit

'Workbooks(OutputFileName).Worksheets(Sheetname_4).DisplayGridlines = False
'Workbooks(OutputFileName).Worksheets(Sheetname_2).DisplayGridlines = False

Workbooks(OutputFileName).Close SaveChanges:=True 'Closing data file

End Sub







