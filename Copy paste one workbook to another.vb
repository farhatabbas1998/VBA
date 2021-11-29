Sub Claim()
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        

        filedate = Format(Date, "ddmmyyyy")
        OutputFileName = "Claim " & filedate & ".csv"
        Call Check_if_workbook_is_open(OutputFileName)
        
        Workbooks.Add.SaveAs Filename:=ThisWorkbook.Path & "\" & OutputFileName, FileFormat:=xlCSVUTF8, CreateBackup:=False
        Filename = ThisWorkbook.Name
        Sheetname = "Sheet1"
        Claimsheet = "Claim"

        Call CheckDataSheet(OutputFileName)
        Call Deletesheet1(OutputFileName)
        Dim wb As Workbook
        Dim ws As Worksheet
        Set wb = Workbooks(OutputFileName)
        Set ws = wb.Sheets(Claimsheet)
        With ws

        .Range("A:A").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("A:A").Value
        .Range("B:B").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("B:B").Value
        .Range("C:C").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("C:C").Value
        .Range("D:D").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("D:D").Value
        .Range("E:E").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("E:E").Value
        .Range("F:F").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("F:F").Value
        .Range("G:G").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("G:G").Value
        .Range("H:H").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("H:H").Value
        .Range("I:I").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("I:I").Value
        .Range("J:J").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("J:J").Value
        .Range("K:K").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("K:K").Value
        .Range("L:L").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("L:L").Value
        .Range("M:M").FormulaR1C1 = Workbooks(Filename).Sheets(Sheetname).Range("M:M").Value


        End With
        
        Workbooks(OutputFileName).Worksheets(Claimsheet).Columns("A:W").EntireColumn.AutoFit
        Workbooks(OutputFileName).Save
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
End Sub


Sub Check_if_workbook_is_open(OutputFileName)
    Dim wb As Workbook 'to test if workbook is open. No change here
        For Each wb In Workbooks
            If wb.Name = OutputFileName Then
                Workbooks(OutputFileName).Save
                Workbooks(OutputFileName).Close
            End If
        Next
End Sub
Sub CheckDataSheet(Filename)
    For Each Sheet In Workbooks(Filename).Worksheets ' Checking if VBA Sheet exist
        If Sheet.Name = "Claim" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
    Workbooks(Filename).Sheets.Add.Name = "Claim"
End Sub
Sub Deletesheet1(Filename)
    For Each Sheet In Workbooks(Filename).Worksheets ' Delete sheet1
        If Sheet.Name <> "Claim" Then
            Application.DisplayAlerts = False
            Sheet.Delete
            Application.DisplayAlerts = True
        End If
    Next Sheet
End Sub
