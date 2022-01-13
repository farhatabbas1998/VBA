VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_InvoiceFrontend 
   Caption         =   "Invoice Automation Pricing"
   ClientHeight    =   10770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   OleObjectBlob   =   "F_InvoiceFrontend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_InvoiceFrontend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version: 1.00
'Ika - Invoice Backend
'For Echobroadband
'User: Ika
'By Farhat Abbas & Ika

Private Sub Continue_Click()
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


 If Region1 = True Then
 .Range("I11").FormulaR1C1 = "Atlanta"
 Call Reg1
 ElseIf Region2 = True Then
 .Range("I11").FormulaR1C1 = "Beltway"
 Call Reg2
 ElseIf Region3 = True Then
 .Range("I11").FormulaR1C1 = "California"
 Call Reg3
 ElseIf Region4 = True Then
 .Range("I11").FormulaR1C1 = "Chicago"
 Call Reg4
 ElseIf Region5 = True Then
 .Range("I11").FormulaR1C1 = "Twin City"
 Call Reg5
 ElseIf Region6 = True Then
 .Range("I11").FormulaR1C1 = "Houston"
 Call Reg6
 ElseIf Region7 = True Then
 .Range("I11").FormulaR1C1 = "Seattle"
 Call Reg7
 ElseIf Region8 = True Then
 .Range("I11").FormulaR1C1 = "Florida"
 Call Reg8
 End If
  .Range("L" & 3).FormulaR1C1 = "USL20522001-" & InvoiceNo.Value
  .Range("E" & 14).FormulaR1C1 = D1.Value
  .Range("E" & 15).FormulaR1C1 = D2.Value
  .Range("E" & 16).FormulaR1C1 = D3.Value
  .Range("E" & 17).FormulaR1C1 = D4.Value
  .Range("E" & 18).FormulaR1C1 = D5.Value
  .Range("E" & 19).FormulaR1C1 = D6.Value
  .Range("E" & 20).FormulaR1C1 = D7.Value
  .Range("E" & 21).FormulaR1C1 = D8.Value
  .Range("E" & 22).FormulaR1C1 = D9.Value
  .Range("E" & 23).FormulaR1C1 = D10.Value
  .Range("E" & 24).FormulaR1C1 = D11.Value
  .Range("E" & 25).FormulaR1C1 = D12.Value
  .Range("E" & 26).FormulaR1C1 = D13.Value
  .Range("E" & 27).FormulaR1C1 = D14.Value
  .Range("E" & 28).FormulaR1C1 = D15.Value
  .Range("E" & 29).FormulaR1C1 = D16.Value
  .Range("E" & 30).FormulaR1C1 = D17.Value
  .Range("E" & 31).FormulaR1C1 = D18.Value
  .Range("E" & 32).FormulaR1C1 = D19.Value
  .Range("E" & 33).FormulaR1C1 = D24.Value
  .Range("E" & 34).FormulaR1C1 = D25.Value
  .Range("E" & 35).FormulaR1C1 = D26.Value
  .Range("E" & 36).FormulaR1C1 = D27.Value
  .Range("E" & 37).FormulaR1C1 = D29.Value
  .Range("E" & 38).FormulaR1C1 = D31.Value
  .Range("E" & 39).FormulaR1C1 = D32.Value
  .Range("E" & 40).FormulaR1C1 = D33.Value
  .Range("E" & 41).FormulaR1C1 = D34.Value
  .Range("E" & 42).FormulaR1C1 = D35.Value
  .Range("E" & 43).FormulaR1C1 = D36.Value
  .Range("E" & 44).FormulaR1C1 = D37.Value
  .Range("E" & 45).FormulaR1C1 = D38.Value
 End With
   Call Invoice_Summary
End Sub


Private Sub Reg1()
'setting default
D1.Value = "1"
D2.Value = "1"
D3.Value = "1"
D4.Value = "1"
D5.Value = "1"
D6.Value = "1"
D7.Value = "1"
D8.Value = "1"
D9.Value = "1"
D10.Value = "1"
D11.Value = "1"
D12.Value = "1"
D13.Value = "1"
D14.Value = "1"
D15.Value = "1"
D16.Value = "1"
D17.Value = "1"
D18.Value = "1"
D19.Value = "1"
D24.Value = "1"
D25.Value = "1"
D26.Value = "1"
D27.Value = "1"
D29.Value = "1"
D31.Value = "1"
D32.Value = "1"
D33.Value = "1"
D34.Value = "1"
D35.Value = "1"
D36.Value = "1"
D37.Value = "1"
D38.Value = "1"
End Sub
Private Sub Reg2()
  'setting default
  D1.Value = "2"
  D2.Value = "2"
  D3.Value = "2"
  D4.Value = "2"
  D5.Value = "2"
  D6.Value = "2"
  D7.Value = "2"
  D8.Value = "2"
  D9.Value = "2"
  D10.Value = "2"
  D11.Value = "2"
  D12.Value = "2"
  D13.Value = "2"
  D14.Value = "2"
  D15.Value = "2"
  D16.Value = "2"
  D17.Value = "2"
  D18.Value = "2"
  D19.Value = "2"
  D24.Value = "2"
  D25.Value = "2"
  D26.Value = "2"
  D27.Value = "2"
  D29.Value = "2"
  D31.Value = "2"
  D32.Value = "2"
  D33.Value = "2"
  D34.Value = "2"
  D35.Value = "2"
  D36.Value = "2"
  D37.Value = "2"
  D38.Value = "2"
  End Sub
  Private Sub Reg3()
  'setting default
  D1.Value = "3"
  D2.Value = "3"
  D3.Value = "3"
  D4.Value = "3"
  D5.Value = "3"
  D6.Value = "3"
  D7.Value = "3"
  D8.Value = "3"
  D9.Value = "3"
  D10.Value = "3"
  D11.Value = "3"
  D12.Value = "3"
  D13.Value = "3"
  D14.Value = "3"
  D15.Value = "3"
  D16.Value = "3"
  D17.Value = "3"
  D18.Value = "3"
  D19.Value = "3"
  D24.Value = "3"
  D25.Value = "3"
  D26.Value = "3"
  D27.Value = "3"
  D29.Value = "3"
  D31.Value = "3"
  D32.Value = "3"
  D33.Value = "3"
  D34.Value = "3"
  D35.Value = "3"
  D36.Value = "3"
  D37.Value = "3"
  D38.Value = "3"
  End Sub
  Private Sub Reg4()
  'setting default
  D1.Value = "4"
  D2.Value = "4"
  D3.Value = "4"
  D4.Value = "4"
  D5.Value = "4"
  D6.Value = "4"
  D7.Value = "4"
  D8.Value = "4"
  D9.Value = "4"
  D10.Value = "4"
  D11.Value = "4"
  D12.Value = "4"
  D13.Value = "4"
  D14.Value = "4"
  D15.Value = "4"
  D16.Value = "4"
  D17.Value = "4"
  D18.Value = "4"
  D19.Value = "4"
  D24.Value = "4"
  D25.Value = "4"
  D26.Value = "4"
  D27.Value = "4"
  D29.Value = "4"
  D31.Value = "4"
  D32.Value = "4"
  D33.Value = "4"
  D34.Value = "4"
  D35.Value = "4"
  D36.Value = "4"
  D37.Value = "4"
  D38.Value = "4"
  End Sub
  Private Sub Reg5()
  'setting default
  D1.Value = "5"
  D2.Value = "5"
  D3.Value = "5"
  D4.Value = "5"
  D5.Value = "5"
  D6.Value = "5"
  D7.Value = "5"
  D8.Value = "5"
  D9.Value = "5"
  D10.Value = "5"
  D11.Value = "5"
  D12.Value = "5"
  D13.Value = "5"
  D14.Value = "5"
  D15.Value = "5"
  D16.Value = "5"
  D17.Value = "5"
  D18.Value = "5"
  D19.Value = "5"
  D24.Value = "5"
  D25.Value = "5"
  D26.Value = "5"
  D27.Value = "5"
  D29.Value = "5"
  D31.Value = "5"
  D32.Value = "5"
  D33.Value = "5"
  D34.Value = "5"
  D35.Value = "5"
  D36.Value = "5"
  D37.Value = "5"
  D38.Value = "5"
  End Sub
  Private Sub Reg6()
  'setting default
  D1.Value = "6"
  D2.Value = "6"
  D3.Value = "6"
  D4.Value = "6"
  D5.Value = "6"
  D6.Value = "6"
  D7.Value = "6"
  D8.Value = "6"
  D9.Value = "6"
  D10.Value = "6"
  D11.Value = "6"
  D12.Value = "6"
  D13.Value = "6"
  D14.Value = "6"
  D15.Value = "6"
  D16.Value = "6"
  D17.Value = "6"
  D18.Value = "6"
  D19.Value = "6"
  D24.Value = "6"
  D25.Value = "6"
  D26.Value = "6"
  D27.Value = "6"
  D29.Value = "6"
  D31.Value = "6"
  D32.Value = "6"
  D33.Value = "6"
  D34.Value = "6"
  D35.Value = "6"
  D36.Value = "6"
  D37.Value = "6"
  D38.Value = "6"
  End Sub
  Private Sub Reg7()
  'setting default
  D1.Value = "7"
  D2.Value = "7"
  D3.Value = "7"
  D4.Value = "7"
  D5.Value = "7"
  D6.Value = "7"
  D7.Value = "7"
  D8.Value = "7"
  D9.Value = "7"
  D10.Value = "7"
  D11.Value = "7"
  D12.Value = "7"
  D13.Value = "7"
  D14.Value = "7"
  D15.Value = "7"
  D16.Value = "7"
  D17.Value = "7"
  D18.Value = "7"
  D19.Value = "7"
  D24.Value = "7"
  D25.Value = "7"
  D26.Value = "7"
  D27.Value = "7"
  D29.Value = "7"
  D31.Value = "7"
  D32.Value = "7"
  D33.Value = "7"
  D34.Value = "7"
  D35.Value = "7"
  D36.Value = "7"
  D37.Value = "7"
  D38.Value = "7"
  End Sub
Private Sub Reg8()
'setting default
D1.Value = "8"
D2.Value = "8"
D3.Value = "8"
D4.Value = "8"
D5.Value = "8"
D6.Value = "8"
D7.Value = "8"
D8.Value = "8"
D9.Value = "8"
D10.Value = "8"
D11.Value = "8"
D12.Value = "8"
D13.Value = "8"
D14.Value = "8"
D15.Value = "8"
D16.Value = "8"
D17.Value = "8"
D18.Value = "8"
D19.Value = "8"
D24.Value = "8"
D25.Value = "8"
D26.Value = "8"
D27.Value = "8"
D29.Value = "8"
D31.Value = "8"
D32.Value = "8"
D33.Value = "8"
D34.Value = "8"
D35.Value = "8"
D36.Value = "8"
D37.Value = "8"
D38.Value = "8"
End Sub


Sub Invoice_Summary()
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
    .Range("B11:C11,D11:G11,B12:E12,F12:G12,H12:I12,J12:K12").Merge
    .Range("L12:M12,B47:E47,F46:G46,H46:I46,J46:K46,L46:M46").Merge
    .Range("L47:M47,B47:K47,L2:M2,L3:M3,K6:M6,K7:M7,K8:M8").Merge
    .Range("K9:M9,K5:M5,B46:E46,K1:M1").Merge
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
    .Range("B14").FormulaR1C1 = "D1"
    .Range("B15").FormulaR1C1 = "D2"
    .Range("B16").FormulaR1C1 = "D3"
    .Range("B17").FormulaR1C1 = "D4"
    .Range("B18").FormulaR1C1 = "D5"
    .Range("B19").FormulaR1C1 = "D6"
    .Range("B20").FormulaR1C1 = "D7"
    .Range("B21").FormulaR1C1 = "D8"
    .Range("B22").FormulaR1C1 = "D9"
    .Range("B23").FormulaR1C1 = "D10"
    .Range("B24").FormulaR1C1 = "D11"
    .Range("B25").FormulaR1C1 = "D12"
    .Range("B26").FormulaR1C1 = "D13"
    .Range("B27").FormulaR1C1 = "D14"
    .Range("B28").FormulaR1C1 = "D15"
    .Range("B29").FormulaR1C1 = "D16"
    .Range("B30").FormulaR1C1 = "D17"
    .Range("B31").FormulaR1C1 = "D18"
    .Range("B32").FormulaR1C1 = "D19"
    .Range("B33").FormulaR1C1 = "D24"
    .Range("B34").FormulaR1C1 = "D25"
    .Range("B35").FormulaR1C1 = "D26"
    .Range("B36").FormulaR1C1 = "D27"
    .Range("B37").FormulaR1C1 = "D29"
    .Range("B38").FormulaR1C1 = "D31"
    .Range("B39").FormulaR1C1 = "D32"
    .Range("B40").FormulaR1C1 = "D33"
    .Range("B41").FormulaR1C1 = "D34"
    .Range("B42").FormulaR1C1 = "D35"
    .Range("B43").FormulaR1C1 = "D36"
    .Range("B44").FormulaR1C1 = "D37"
    .Range("B45").FormulaR1C1 = "D38"
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
    .Range("C36").FormulaR1C1 = "Node Split Load Rebalance"
    .Range("C37").FormulaR1C1 = "Plant Map Update"
    .Range("C38").FormulaR1C1 = "Fiber Trunk Tree Modify"
    .Range("C39").FormulaR1C1 = "Fiber Trunk Tree New"
    .Range("C40").FormulaR1C1 = "Wavelength Res Req"
    .Range("C41").FormulaR1C1 = "Ladder Report"
    .Range("C42").FormulaR1C1 = "Fiber Route Trace"
    .Range("C43").FormulaR1C1 = "Splice Update"
    .Range("C44").FormulaR1C1 = "Splice Addition"
    .Range("C45").FormulaR1C1 = "Misc Hourly work"
    .Range("D23,D24").FormulaR1C1 = "Parcel"
    .Range("D30").FormulaR1C1 = "Unit"
    .Range("D38,D39,D40").FormulaR1C1 = "Drawing"
    .Range("D41").FormulaR1C1 = "Report"
    .Range("D43,D44").FormulaR1C1 = "Insert"
    .Range("D45").FormulaR1C1 = "Hour"
    .Range("B46").FormulaR1C1 = "Subtotals"
    .Range("B47").FormulaR1C1 = "Invoice Total ( USD)"
    .Range("F13,H13,J13,L13").FormulaR1C1 = "Qty"
    .Range("I13,G13,K13,M13").FormulaR1C1 = "Sub Total"
    .Range("D14,D17,D20,D25,D28,D29,D33,D35,D36").FormulaR1C1 = "Project"
    .Range("D15,D16,D18,D19,D21,D22,D26,D27,D31,D34").FormulaR1C1 = "Feet"
    .Range("D32,D37,D38,D37,D42").FormulaR1C1 = "Each"
    'ActiveWindow.DisplayGridlines = False
    .Range("E14:E45,G14:G46,I14:I46,K14:K46,M14:M47,F46:M46,l47").NumberFormat = "[$$-en-US]#,##0.00"
    .Range("A1:M9,A12:M13,B:B,F47:M48,D46:M46,B11:H11,L11").Font.FontStyle = "Bold"
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
    With .Range("K2,L2,D11,C13,B13:B45,B46:M47").Font
    .name = "Calibri"
    .Size = 12
    End With
    With .Range("D11:M47,B13,K1:M3,K5")
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
    With .Range("C11,C13:C45,C11")
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .ReadingOrder = xlContext
    End With
    With .Range("B14:B47")
      .HorizontalAlignment = xlRight
      .VerticalAlignment = xlCenter
      .ReadingOrder = xlContext
  End With
  'normal table
    tablesize = "B11:M47,K2:M3,K5:M5"
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
    Thick_outside_boarder = "K2:M3,K5:M9,B11:G11,H11:I11,J11:K11,L11:M11,B11:E11,F11:G11,H11:I11,J11:K11,L11:M11,B12:E12,B47:K47,L47:M47"
    Call Thick_OB(Thick_outside_boarder)
    Thick_outside_boarder = "F12:G12,H12:I12,J12:K12,L12:M12,B13:E45,F13:G45,H13:I45,J13:K45,L13:M45,B46:E46,J46:K46,F46:G46,H46:I46,L46:M46"
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
    For x = 14 To 45
    .Rows(x & ":" & x).RowHeight = 19.5
    Next x
    .Rows("12:12").RowHeight = 30
    .Rows("11:11").RowHeight = 30
    .Rows("46:46").RowHeight = 21
    .Rows("47:47").RowHeight = 27
    
    End With

  
End Sub
  
Sub Thick_OB(Thick_outside_boarder)
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
End Sub
Sub Clearup()
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
End Sub
