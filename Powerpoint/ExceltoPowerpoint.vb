Sub pptfromexcel()

  'Excel
   Filename = ThisWorkbook.Name
   Sheetname1 = ThisWorkbook.Sheets(1).Name
   Sheetname2 = ThisWorkbook.Sheets(2).Name
   Dim wb As Workbook
   Dim ws1 As Worksheet
   Set wb = Workbooks(Filename)
   Set ws1 = wb.Sheets(Sheetname1)
   Set ws2 = wb.Sheets(Sheetname2)
   
  'PowerPoint
   Dim pptapp As PowerPoint.Application
   Dim pptppt As PowerPoint.Presentation
   Dim pptsld As PowerPoint.Slide
   Set pptapp = New PowerPoint.Application
   Set pptppt = pptapp.Presentations.Add
   pptapp.Visible = True
   pptapp.Activate
   
  'Getting data
   LeftFlowerImagePath = ws2.Cells(5, 6).Value
   RightFlowerImagePath = ws2.Cells(6, 6).Value
   Jun1stWinnerPath = ws2.Cells(7, 6).Value
   Jun2ndWinnerPath = ws2.Cells(8, 6).Value
   Jun3rdWinnerPath = ws2.Cells(9, 6).Value
   Nov1stWinnerPath = ws2.Cells(10, 6).Value
   Nov2ndWinnerPath = ws2.Cells(11, 6).Value
   Nov3rdWinnerPath = ws2.Cells(12, 6).Value
   Pro1stWinnerPath = ws2.Cells(13, 6).Value
   Pro2ndWinnerPath = ws2.Cells(14, 6).Value
   Pro3rdWinnerPath = ws2.Cells(15, 6).Value
   RGB_R1_1 = ws2.Cells(5, 8).Value
   RGB_G1_1 = ws2.Cells(6, 8).Value
   RGB_B1_1 = ws2.Cells(7, 8).Value
   RGB_R1_2 = ws2.Cells(8, 8).Value
   RGB_G1_2 = ws2.Cells(9, 8).Value
   RGB_B1_2 = ws2.Cells(10, 8).Value
   RGB_R2 = ws2.Cells(11, 8).Value
   RGB_G2 = ws2.Cells(12, 8).Value
   RGB_B2 = ws2.Cells(13, 8).Value
   RGB_R3 = ws2.Cells(14, 8).Value
   RGB_G3 = ws2.Cells(15, 8).Value
   RGB_B3 = ws2.Cells(16, 8).Value
   FontName1 = ws2.Cells(17, 8).Value
   FontName2 = ws2.Cells(18, 8).Value
   FontName3 = ws2.Cells(19, 8).Value
   FontSize1 = ws2.Cells(20, 8).Value
   FontSize2 = ws2.Cells(21, 8).Value
   FontSize3 = ws2.Cells(22, 8).Value
   FontBold1 = ws2.Cells(23, 8).Value
   FontBold2 = ws2.Cells(24, 8).Value
   FontBold3 = ws2.Cells(25, 8).Value
  'Counting Data
   countdata = Application.WorksheetFunction.CountA(ws1.Range("B:B")) - 1
   
    'looping for each row into slide
      For x = 1 To countdata
        Set pptsld = pptppt.Slides.Add(x, ppLayoutTitle)
        pptsld.Shapes(1).TextFrame.TextRange = ws1.Cells(x + 1, 2).Value
        pptsld.Shapes(2).TextFrame.TextRange = ws1.Cells(x + 1, 5).Value & vbNewLine & " - " & ws1.Cells(x + 1, 8).Value & " - "
        
        If ws1.Cells(x + 1, 9).Value = 1 And ws1.Cells(x + 1, 4).Value = "PRO" Then
          pptsld.Shapes.AddPicture Filename:=Pro1stWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        ElseIf ws1.Cells(x + 1, 9).Value = 2 And ws1.Cells(x + 1, 4).Value = "PRO" Then
          pptsld.Shapes.AddPicture Filename:=Pro2ndWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        ElseIf ws1.Cells(x + 1, 9).Value = 3 And ws1.Cells(x + 1, 4).Value = "PRO" Then
          pptsld.Shapes.AddPicture Filename:=Pro3rdWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        End If
        
        If ws1.Cells(x + 1, 9).Value = 1 And ws1.Cells(x + 1, 4).Value = "JUNIOR" Then
          pptsld.Shapes.AddPicture Filename:=Jun1stWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        ElseIf ws1.Cells(x + 1, 9).Value = 2 And ws1.Cells(x + 1, 4).Value = "JUNIOR" Then
          pptsld.Shapes.AddPicture Filename:=Jun2ndWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        ElseIf ws1.Cells(x + 1, 9).Value = 3 And ws1.Cells(x + 1, 4).Value = "JUNIOR" Then
          pptsld.Shapes.AddPicture Filename:=Jun3rdWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        End If
        
        If ws1.Cells(x + 1, 9).Value = 1 And ws1.Cells(x + 1, 4).Value = "NOV" Then
          pptsld.Shapes.AddPicture Filename:=Nov1stWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        ElseIf ws1.Cells(x + 1, 9).Value = 2 And ws1.Cells(x + 1, 4).Value = "NOV" Then
          pptsld.Shapes.AddPicture Filename:=Nov2ndWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        ElseIf ws1.Cells(x + 1, 9).Value = 3 And ws1.Cells(x + 1, 4).Value = "NOV" Then
          pptsld.Shapes.AddPicture Filename:=Nov3rdWinnerPath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=0, Top:=0, Width:=960, Height:=540
        End If
        
        SchoolLen = Len(ws1.Cells(x + 1, 5).Value)
        TeamLen = Len(ws1.Cells(x + 1, 2).Value)
        Totalen = Len(ws1.Cells(x + 1, 5).Value & vbNewLine & " - " & ws1.Cells(x + 1, 8).Value & " - ")
        pptsld.Shapes(1).TextFrame.TextRange.Font.Name = FontName1
        pptsld.Shapes(1).TextFrame.TextRange.Font.Size = FontSize1
        pptsld.Shapes(1).TextFrame.TextRange.Font.Color.RGB = RGB(RGB_R1_1, RGB_G1_1, RGB_B1_1)
        If FontBold1 = 1 Then
         pptsld.Shapes(1).TextFrame.TextRange.Characters(1, TeamLen).Font.Bold = msoTrue
        End If
        pptsld.Shapes(1).TextFrame2.TextRange.Font.Fill.TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        
        pptsld.Shapes(2).TextFrame.TextRange.Font.Name = FontName2
        pptsld.Shapes(2).TextFrame.TextRange.Font.Size = FontSize2
        pptsld.Shapes(2).TextFrame.TextRange.Font.Color.RGB = RGB(RGB_R2, RGB_G2, RGB_B2)
        If FontBold2 = 1 Then
         pptsld.Shapes(2).TextFrame.TextRange.Characters(1, SchoolLen).Font.Bold = msoTrue
        End If
        pptsld.Shapes(2).TextFrame.TextRange.Characters(SchoolLen + 1, Totalen).Font.Name = FontName3
        pptsld.Shapes(2).TextFrame.TextRange.Characters(SchoolLen + 1, Totalen).Font.Size = FontSize3
        pptsld.Shapes(2).TextFrame.TextRange.Characters(SchoolLen + 1, Totalen).Font.Color.RGB = RGB(RGB_R3, RGB_G3, RGB_B3)
        If FontBold3 = 1 Then
         pptsld.Shapes(2).TextFrame.TextRange.Characters(SchoolLen + 1, Totalen).Font.Bold = msoTrue
        End If
        pptsld.Shapes.AddPicture Filename:=RightFlowerImagePath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=820, Top:=230, Width:=60, Height:=120
        pptsld.Shapes.AddPicture Filename:=LeftFlowerImagePath, linktofile:=msoTrue, SaveWithDocument:=msoTrue, Left:=80, Top:=230, Width:=60, Height:=120
        
        pptsld.Shapes(1).TextFrame2.TextRange.Font.Fill.GradientStops.Insert RGB(RGB_R1_2, RGB_G1_2, RGB_B1_2), 0.5
        pptsld.Shapes(1).TextFrame2.TextRange.Font.Fill.GradientStops.Insert RGB(RGB_R1_1, RGB_G1_1, RGB_B1_1), 0.5

    Next
 
End Sub

