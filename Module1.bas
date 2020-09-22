Attribute VB_Name = "Module1"
Sub Multiple_year_stock()

  
Dim Ws As Worksheet
  
  
For Each Ws In Worksheets
    Ws.Activate
    
  ' Set an initial variable for holding the Ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for Yearly Change
  Dim Yearly_Change As Double
  Yearly_Change = 0
  
  ' Set an initial variable for Percent Change
  Dim Percent_Change As String
  Percent_Change = 0
  
  ' Set an initial variable for Total Stock Value
  Dim Total_Stock_Value As Double
  Total_Stock_Value = 0
  
  ' Keep track of the location for each Ticker name in the Change Row
  Dim Yearly_Change_Row As Double
  Yearly_Change_Row = 2
  
  ' Keep track of the location for each Ticker name in the Percent Change Row
  Dim Percent_Change_Row As String
  Percent_Change_Row = 2
  
  ' Keep track of the location for each Ticker name in the Total Stock Value Row
  Dim Total_Stock_Value_Row As Double
  Total_Stock_Value_Row = 2


  ' Add Header Values
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Value"
  
  
  
  
    ' Loop through Ticker names
    For i = 2 To 800000

        ' Check if we are still within the same Ticker name, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the Ticker name
        Ticker_Name = Cells(i, 1).Value
      
      
        ' Add to the Yearly change
        Yearly_Change = Yearly_Change + Cells(i, 3).Value - Cells(i, 6).Value
      
      
        ' Add to the Total Stock Value
        Total_Stock_Value = Total_Stock_Value + Cells(i, 7).Value
      
      
      
        ' Print the Ticker name in the Tickere row
        Range("I" & Yearly_Change_Row).Value = Ticker_Name

        ' Print the Yearly Change value to the Yearly Change row
        Range("J" & Yearly_Change_Row).Value = Yearly_Change
            
            
      
        ' Print the Percent Change to the Percent Change row
        Range("K" & Percent_Change_Row).Value = Percent_Change
      
        ' Change Percent Change numbers to %
        Range("K" & Percent_Change_Row).NumberFormat = "0.00%"
      
        ' Print the Total Stock Value to the Total Stock Value row
        Range("L" & Total_Stock_Value_Row).Value = Total_Stock_Value
      
      

        ' Add one to the Yearly Change row
        Yearly_Change_Row = Yearly_Change_Row + 1
      
        ' Add one to the Percent Change row
        Percent_Change_Row = Percent_Change_Row + 1
      
        ' Add one to Total Stock Value row
        Total_Stock_Value_Row = Total_Stock_Value_Row + 1
      
      
      
        ' Reset the Yearly change
        Yearly_Change = 0
      
        ' Reset the Percentage change
        Percent_Change = 0
      
        ' Reset the Total Stock Value
        Total_Stock_Value = 0


        ' If the cell immediately following a row is the same Ticker...
        Else

        ' Add to the Yearly change
        Yearly_Change = Yearly_Change + Cells(i, 3).Value - Cells(i, 6).Value
        
            If Percent_Change = 0 Then
            Percent_Change = 0
            
            Else
      
            ' Percent change calculation
            Percent_Change = Percent_Change + Cells(i, 6).Value - Cells(i, 3).Value
      
            Percent_Change = Percent_Change / Cells(i, 3).Value
            
            End If
      
        ' Add to the Total Stock Value
        Total_Stock_Value = Total_Stock_Value + Cells(i, 7).Value
        
    

        End If
        
    Next i
    
            Dim k As Long
            

            
            
            For k = 2 To Rows.Count
            
            
            If Cells(k, 10).Value > 0 And IsEmpty(Cells(i, 10).Value) Then
                Cells(k, 10).Interior.ColorIndex = 4
                
            ElseIf Cells(k, 10).Value < 0 Then
                Cells(k, 10).Interior.ColorIndex = 3
            End If
            
            Next k
            
Next
 
End Sub


