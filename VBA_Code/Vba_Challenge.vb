Sub StockData()

 For Each WS In Worksheets
  
   'Set the values
     Dim WSName As String
     Dim CurrentRow As Long
     Dim StartRow As Long
     Dim TickCount As Long
     Dim LastRowA As Long
     Dim PercentChange As Double
     
     'Create the headers
     
     WS.Range("I1").Value = "Ticker"
     WS.Range("J1").Value = "Yearly Change"
     WS.Range("K1").Value = "Percent Change"
     WS.Range("L1").Value = "Total Stock Volume"
     
     'Set TickCount
     
     TickCount = 2
     
     'Set Start Row
     
     StartRow = 2
     
     'Last non blank cell in Colum A
     LastRowA = WS.Cells(Rows.Count, 1).End(xlUp).Row
     'I used the MsgBox to make sure the values were correct before moving forward
     
     For CurrentRow = 2 To LastRowA
       
        If WS.Cells(CurrentRow + 1, 1).Value <> WS.Cells(CurrentRow, 1).Value Then
       
         WS.Cells(TickCount, 9).Value = WS.Cells(CurrentRow, 1).Value
          
         'Percentage change (F-C)
         
         WS.Cells(TickCount, 10).Value = WS.Cells(CurrentRow, 6).Value - WS.Cells(StartRow, 3).Value
          
          'Add color to the cells
             If WS.Cells(TickCount, 10).Value > 0 Then
             WS.Cells(TickCount, 10).Interior.ColorIndex = 4
             Else
             WS.Cells(TickCount, 10).Interior.ColorIndex = 3
             End If
              
         'Percent Change
         WS.Cells(TickCount, 11).Value = Format(((WS.Cells(CurrentRow, 6).Value - WS.Cells(StartRow, 3).Value) / WS.Cells(StartRow, 3).Value), "Percent")
          
         'Total Volume
         WS.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(WS.Cells(StartRow, 7), WS.Cells(CurrentRow, 7)))
          
          
         TickCount = TickCount + 1
         StartRow = CurrentRow + 1
          
        
          
         End If
           
         Next CurrentRow
          
    'Second part, create new headers
     WS.Range("N2").Value = "Greatest % increase"
     WS.Range("N3").Value = "Greatest % Decrease"
     WS.Range("N4").Value = "Greatest Total Volume"
     WS.Range("O1").Value = "Ticker"
     WS.Range("P1").Value = "Value"
     
     'Declare de variables
     Dim Great_inc As Double
     Dim Great_Decre As Double
     Dim Great_TotalV As Double
     Dim LastRowI As Long
     
     'Find last non-blank column in I
     
     LastRowI = WS.Cells(Rows.Count, 9).End(xlUp).Row
     'I used the MsgBox to make sure the values were correct before moving forward
     
     
     'set the initial values
     Great_inc = WS.Cells(2, 11).Value
     Great_Decre = WS.Cells(2, 11).Value
     Great_TotalV = WS.Cells(2, 12).Value
     
     'Loop thru data to find the value
     For CurrentRow = 2 To LastRowI
     
     If WS.Cells(CurrentRow, 11).Value > Great_inc Then
     
         Great_inc = WS.Cells(CurrentRow, 11).Value
         WS.Cells(2, 15).Value = WS.Cells(CurrentRow, 9).Value
         
         Else
         
        Great_inc = Great_inc
        
        End If
        
        If WS.Cells(CurrentRow, 11).Value < Great_Decre Then
     
         Great_Decre = WS.Cells(CurrentRow, 11).Value
         WS.Cells(3, 15).Value = WS.Cells(CurrentRow, 9).Value
         
         Else
         
        Great_Decre = Great_Decre
        
       End If
        
        If WS.Cells(CurrentRow, 12).Value > Great_TotalV Then
     
         Great_TotalV = WS.Cells(CurrentRow, 12).Value
         WS.Cells(4, 15).Value = WS.Cells(CurrentRow, 9).Value
         
         Else
         
        Great_TotalV = Great_TotalV
        
       End If
        
       'print the values from my if stament in the worksheet
       WS.Cells(2, 16).Value = Format(Great_inc, "Percent")
       WS.Cells(3, 16).Value = Format(Great_Decre, "Percent")
       WS.Cells(4, 16).Value = Great_TotalV
       
        
    
        
   Next CurrentRow
   
      
    
          'Adjust column
          WS.Cells.EntireColumn.AutoFit
          
              
 Next WS
   
   

End Sub


