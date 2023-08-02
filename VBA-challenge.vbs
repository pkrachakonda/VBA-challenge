Attribute VB_Name = "Module1"
Sub Stock():

    'Declaration of Variables'
        Dim Stock_Name As String
        Dim Stock_Volume As Double
        Dim Summary_Row As Integer
        Dim Open_Price As Double
        Dim Close_Price As Double
    
    For Each ws In Worksheets ' Loop through each worksheet'
    
    ' Initiating variable '
        Summary_Row = 2
        Stock_Volume = 0
        
        Open_Price = ws.Range("C2").Value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
       ' Defining Cell Names'
       
        ws.Range("I1").Value = "Tinker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Tinker"
        ws.Range("Q1").Value = "Value"
               
        For i = 2 To LastRow ' Looping from second row till last row '
        
            If ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value Then ' To check whether current cell stock name is same as in the next cell
            Stock_Name = ws.Range("A" & i).Value ' Assigning the name of stock name to variable "Stock Name (Tinker)"'
            Stock_Volume = Stock_Volume + ws.Range("G" & i).Value ' Estimating the Stock Volume '
            Close_Price = ws.Range("F" & i).Value ' Estimaing the Close Price of Stock'
            
            ws.Range("I" & Summary_Row).Value = Stock_Name 'Assigning the Name of Stock Name (Tinker) to Column I in each worksheet '
            ws.Range("J" & Summary_Row).Value = (Close_Price - Open_Price) 'Estimating yearly change of each Stock and assign it to Column J in each worksheet '
            ws.Range("K" & Summary_Row).Value = (Close_Price - Open_Price) / (Open_Price) 'Estimating Percent change of each Stock and assign it to Column K in each worksheet '
            ws.Range("L" & Summary_Row).Value = Stock_Volume  ' Estimating the Total Stock Volume and assigning it Column L in each worksheet'
            
            ' Assigning colour Red (ColorIndex = 3) if values in J and K Columns are negative, otherwise Green (ColorIndex = 4) if values are postive for each worksheet '
                If ws.Range("J" & Summary_Row).Value < 0 Then
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 3
                    ws.Range("K" & Summary_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
                    ws.Range("K" & Summary_Row).Interior.ColorIndex = 4
                End If
                
            Open_Price = ws.Range("C" & i + 1) 'Assigning open price of new Tinker (Stock) in each worksheet'
            Summary_Row = Summary_Row + 1 ' Moving the Stock assigning row by one '
            
            Stock_Volume = 0 'Resetting the Stock volume '
            Else
            Stock_Volume = Stock_Volume + ws.Range("G" & i).Value ' Adding Daily Stock Volume in each worksheet'
            End If
            
        Next i ' Moving to Next Row in the Worksheet'
        
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K:K")) ' Finding Stock (Tinker) with highest yearly precent increase using worksheet built in function in each worksheet'
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K:K")) ' Finding Stock (Tinker) with lowest  yearly precent increase using worksheet built in function in each worksheet'
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L:L")) ' Finding Stock (Tinker) with highest total yearly volume using worksheet built in function in each worksheet'
        ws.Range("J2:J" & LastRow).NumberFormat = "#,##0.00"       ' Assigning Numbering style (two digits after decimal point) for Yearly Change Column in each worksheet'
        ws.Range("L2:L" & LastRow).NumberFormat = "#,##0"          ' Assigning Numbering style (thousand sepearor) for Total Stock Volume Column in each worksheet '
        ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"          ' Assigning Numbering style(Percentage) for Percent Change Column in each worksheet '
        ws.Range("Q2:Q3").NumberFormat = "0.00%"                   ' Assigning Numbering style(Percentage) for Value Column in each worksheet '
        ws.Range("Q4").NumberFormat = "0.00E+##"                   ' Assigning Numbering style(Scientific) for Greatest Total Volume Value row in each worksheet'

        
        For i = 2 To LastRow ' Looping from second row till last row '

            If ws.Range("K" & i).Value = ws.Range("Q2").Value Then 'Finding Stock Name which has highest yearly percent increase and assign that value to Cell "P2", in each worksheet '
                ws.Range("P2").Value = ws.Range("I" & i)
            ElseIf ws.Range("K" & i).Value = ws.Range("Q3").Value Then 'Finding Stock Name which has lowest yearly percent increase and assign that value to Cell "P3", in each worksheet '
                ws.Range("P3").Value = ws.Range("I" & i)
            ElseIf ws.Range("L" & i).Value = ws.Range("Q4").Value Then 'Finding Stock Name which has highest Yearly total stock volume and assign that value to Cell "P4", in each worksheet '
                ws.Range("P4").Value = ws.Range("I" & i)
            End If

        Next i ' Moving to Next Row in the Worksheet'
    
    ws.Columns("A:Z").AutoFit    ' Autofit all columns in the worksheet'
    
    Next ws ' Moving to next worksheet '
    
End Sub
    
