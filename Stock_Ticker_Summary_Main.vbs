Attribute VB_Name = "Module1"
Sub StockTicker()

    ' Loop through each worksheet
    Dim Ws As Worksheet
    For Each Ws In ThisWorkbook.Sheets
    
        
        ' Determine the Last Row
        Dim lastrow As Long
        lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        

        ' Add the Column Headers for the summary table
        Ws.Cells(1, 9).Value = "Ticker"
        Ws.Cells(1, 10).Value = "Yearly Change"
        Ws.Cells(1, 11).Value = "Percent Change"
        Ws.Cells(1, 12).Value = "Total Stock Volume"
        
      ' Set an initial variable for holding the ticker name
      Dim Ticker As String
           
      ' Set an initial variable for holding the total yearly change
      Dim Yearly_Change As Double
      Yearly_Change = 0
      
      ' Set an initial variable for holding the total Open value
      Dim Open_Value As Double
      Open_Value = 0
    
      ' Set an initial variable for holding the Percent Change
      Dim Percent_Change As Double
      
      
      ' Set an initial variable for holding the total volume
      Dim Total_Volume As Double
      Total_Volume = 0
      
      ' Keep track of the row for each ticker name in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
    
          ' Loop through all rows
          For I = 2 To lastrow
                       
            ' Check if we are still within the same ticker name, if it is not...
            If Ws.Cells(I + 1, 1).Value <> Ws.Cells(I, 1).Value Then
        
              ' Set the Ticker name in the summary table
              Ticker = Ws.Cells(I, 1).Value
        
              ' Add to the Yearly Change
              Yearly_Change = Yearly_Change + (Ws.Cells(I, 6).Value - Ws.Cells(I, 3).Value)
              
              ' Add to the Open Value
              Open_Value = Open_Value + (Ws.Cells(I, 3).Value)
              
              ' Calculate the Percent Change
              If Open_Value <> 0 Then
              Percent_Change = (Yearly_Change / Open_Value)
              Else: Percent_Change = 0
              End If
              
              ' Add to the Total Volume
              Total_Volume = Total_Volume + (Ws.Cells(I, 7).Value)
              
              ' Print the Ticker name in the Summary Table
              Ws.Range("I" & Summary_Table_Row).Value = Ticker
        
              ' Print the Yearly Change Amount to the Summary Table
              Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                                    
              ' Print Percent Change and format to a percentage
              Ws.Range("K" & Summary_Table_Row).Value = Percent_Change
              Ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
              
              ' Print the Total Volume to the Summary Table
              Ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                              
              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              
              ' Reset the Yearly Change
              Yearly_Change = 0
              
              ' Reset Open Value
              Open_Value = 0
           
              ' Reset the Total Volume
              Total_Volume = 0
                    
                    
            Else
        
              ' Add to the Yearly Change
              Yearly_Change = Yearly_Change + (Ws.Cells(I, 6).Value - Ws.Cells(I, 3).Value)
              
              ' Add to the Open Value
              Open_Value = Open_Value + (Ws.Cells(I, 3).Value)
                     
              ' Add to the Total Volume
              Total_Volume = Total_Volume + (Ws.Cells(I, 7).Value)
            
            End If
                
               
          ' Next Row
          Next I
          
            ' Creating table headers for min/max values
            Ws.Cells(1, 15).Value = "Ticker Name"
            Ws.Cells(2, 14).Value = "Greatest Percent Increase"
            Ws.Cells(3, 14).Value = "Greatest Percent Decrease"
            Ws.Cells(4, 14).Value = "Greatest Total Volume"
            
            ' Set variable and initial value for maximum percent change
            Dim MaxVal As Double
            MaxVal = 0
            ' Set variable and initial value for minimum percent change
            Dim MinVal As Double
            MinVal = 0
            ' Set variable and initial value for maximum volume
            Dim MaxTot As LongLong
            MaxTot = 0
                 
            ' Loop through the summary table rows
            For j = 2 To Summary_Table_Row
                ' Format the Yearly Changes in the Summary Table
                ' Positive changes are shaded green
                 If Ws.Cells(j, 10).Value > 0 Then
                     Ws.Cells(j, 10).Interior.ColorIndex = 4
                ' Negative changes are shaded red
                 ElseIf Ws.Cells(j, 10).Value < 0 Then
                     Ws.Cells(j, 10).Interior.ColorIndex = 3
                 
                 End If
                 
                 ' Determine if the row value for percent change is larger than current max
                 If Ws.Cells(j, 11).Value > MaxVal Then
                 ' Put the associated ticker name in the min/max table
                 Ws.Cells(2, 15) = Ws.Cells(j, 9).Value
                 ' Put the value in the min/max table and format it
                 Ws.Cells(2, 16) = Ws.Cells(j, 11).Value
                 Ws.Cells(2, 16).NumberFormat = "0.00%"
                 ' Set the maxval to the current max for the next itteration
                 MaxVal = Ws.Cells(2, 16).Value
                 
                 ' Determine if the row value for percent change is smaller than current min
                 ElseIf Ws.Cells(j, 11).Value < MinVal Then
                 ' Put the associated ticker name in the min/max table
                 Ws.Cells(3, 15) = Ws.Cells(j, 9).Value
                 ' Put the value in the min/max table and format it
                 Ws.Cells(3, 16) = Ws.Cells(j, 11).Value
                 Ws.Cells(3, 16).NumberFormat = "0.00%"
                 ' Set the minval to the current min for the next itteration
                 MinVal = Ws.Cells(3, 16).Value
                 
                 End If
                 
                 ' Determine if the row value for total volume is larger than current max
                 If Ws.Cells(j, 12).Value > MaxTot Then
                 ' Put the associated ticker name in the min/max table
                 Ws.Cells(4, 15) = Ws.Cells(j, 9).Value
                 ' Put the value in the min/max table
                 Ws.Cells(4, 16) = Ws.Cells(j, 12).Value
                 ' Set the maxtot to the current max for the next itteration
                 MaxTot = Ws.Cells(4, 16).Value
                 
                 End If
                 
            ' Next Row
            Next j
            
            
    ' Move to next Worksheet
    Next Ws
        
End Sub
