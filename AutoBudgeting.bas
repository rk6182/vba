Attribute VB_Name = "AutoBudgeting"
Sub AutBudgt()

  Application.Visible = True
  
  Sheet2.Visible = True
  
  Sheet2.Select
  
  Sheet2.Buttons.Delete
  
  ActiveSheet.Name = "Budgeting"


  Dim FYcnt As Integer
       Dim i As Integer
       Dim FY As String
       Dim Selc As Boolean
       Dim colno As Integer
       Dim lcol As Integer
       Dim rowno As Integer
       Dim j As Integer
       Dim k As Integer
       Dim Samemonthlastyear As Double
       Dim increper As Double
       Dim CurrLrow As Integer
       Dim currLcol As Integer
       Dim FYSearch As Integer

       
       
        Range("A5").Value = "AccountID"
        Range("B5").Value = "Account Name"
        Range("C3").Value = "Fiscal Year"
        Range("E1").Value = "Increment %"
        Range("E1:F1").Merge
        Range("E1:G1").Borders.LineStyle = xlContinuous
        Range("G2").Value = "Reset"
        Range("G2").Borders.LineStyle = xlContinuous
        
        
        lrow = Sheet5.Cells(Rows.Count, 1).End(xlUp).Row
         
         Sheet5.Visible = True
         
         Sheet5.Select
         
         Sheet5.Range(Cells(4, 1), Cells(lrow, 2)).Copy
         
         Sheet5.Visible = xlSheetVeryHidden

         Sheet2.Select
         Range("A6").Select
              
         Selection.PasteSpecial Paste:=xlPasteValues
         
         Columns("A:B").AutoFit
         
         
         
        CurrLrow = Sheet2.Cells(Rows.Count, 1).End(xlUp).Row
          
        On Error Resume Next
            
          
        FYcnt = Range("A1").Value
        
        
        If FYcnt = 0 Then
        
        
        MsgBox "Please select no of FY"
        
        Budgeting.Show
        
        Exit Sub
        
        Else
        
        
        End If
        
        tday = Now()
        
        
        FyFormat = Format(tday, "mmm'YY")
        
               On Error Resume Next
        FYSearch = 1
        
        FYSearch = Application.WorksheetFunction.Match(FyFormat, Sheet1.Range("E:E"), 0)
         
        
      
        
        For i = 0 To FYcnt - 1
            
        FY = Replace(Sheet1.Cells(FYSearch + i * 12, 1).Value, "FY 20", "FY ")
      
        
            If FY <> "" Then
            

            colno = 1

            colno = Application.WorksheetFunction.Match(FY, Range("BIData!A2:CC2"), 0)
        
            
        
            lcol = Sheets("Budgeting").Cells(5, Columns.Count).End(xlToLeft).Column
            
            
            Cells(4, lcol + 1).Value = FY
            
            Range("E3").Value = "Updating for " & FY
            
            Cells(4, lcol + 1).Borders.LineStyle = xlContinuous
            
        ''    Range(Cells(6, lcol + 1), Cells(6, lcol + 12)).Group
        

        
        
        Else
        
        MsgBox "Please update Versioning after " & Right(Range("E3").Value, 8)
        
        GoTo insertbudgeting
        
        End If
        
            rowno = Application.WorksheetFunction.Match(Sheet1.Cells(FYSearch + i * 12, 1).Value, Range("Versioning!A:A"), 0)
            
            For j = 0 To 11
            
                Sheets("Budgeting").Cells(5, lcol + 1 + j).Value = Left(Sheet1.Cells(rowno + j, 5).Value, 3)
                Sheets("Budgeting").Cells(5, lcol + 1 + j).Select
          
                
                    For k = 1 To lrow - 3
                        
                        If colno = 1 Then
                        
                           
                           
                           Samemonthlastyear = Sheets("Budgeting").Cells(5 + k, lcol - 11 + j).Value
                           increper = Range("G1").Value
                            
                            Sheets("Budgeting").Cells(5 + k, lcol + 1 + j).Value = Samemonthlastyear * increper + Samemonthlastyear
                            
                            If Range("C1").Value = 1 Then
                            
                             Cells(5 + k, lcol + 1 + j).AddComment increper * 100 & " % increment of " & Round(Samemonthlastyear, 0) & " for " & FY & " " & Left(Sheet1.Cells(rowno + j, 5).Value, 3)
                            Else
                            
                            End If
             
                         Else
                            
                            If Left(Sheet1.Cells(rowno + j, 5).Value, 3) = Sheet5.Cells(3, colno - 12 + j).Value Then
                            
                                Samemonthlastyear = Sheet5.Cells(3 + k, colno - 12 + j).Value
                           
                            Else
                            
                                Samemonthlastyear = 0
                                
                            End If
                           
                           increper = Range("G1").Value
                            
                            Sheets("Budgeting").Cells(5 + k, lcol + 1 + j).Value = Samemonthlastyear * increper + Samemonthlastyear
                            
                            If Range("C1").Value = 1 Then
                            
                            Cells(5 + k, lcol + 1 + j).AddComment increper * 100 & " % increment of " & Round(Samemonthlastyear, 0) & " for " & FY & " " & Left(Sheet1.Cells(rowno + j, 5).Value, 3)
                            
                            Else
                            
                           '' Application.ScreenUpdating = False
                            
                            End If
                            
                            
                        End If
                   
                    Next k
            
            Next j
                

        
        Next i
        

insertbudgeting:
                
        
        currLcol = Cells(5, Columns.Count).End(xlToLeft).Column
        
        Range(Cells(5, 1), Cells(CurrLrow, currLcol)).Borders.LineStyle = xlContinuous
        
         Range(Cells(5, 1), Cells(CurrLrow, 1)).NumberFormat = "@"
        
        
        Range("E3").Clear
        Range("C1").Clear
        
        
        Range("A1").Value = "Logo"
        
        

        
        Range("A1:B2").Merge
        Range("A1:B2").Borders.LineStyle = xlContinuous
        
        
        ActiveWorkbook.Save
        
        
        ActiveSheet.Buttons.Add(520, 2, 69, 21.5).Select
        Selection.OnAction = "confirmation"
        Selection.Characters.Text = "Push in BI"
        
        
       

        
        Range("A1").Select
        
        Application.ScreenUpdating = True
                        
MsgBox "Done"


End Sub
