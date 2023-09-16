Attribute VB_Name = "Module1"
Sub Stock_Symbol()

  Dim wsMySheet As Worksheet
  
  Application.ScreenUpdating = False
  
  
  For Each wsMySheet In ThisWorkbook.Sheets
  
  wsMySheet.Select
  
  ' Set an initial variable for holding the Ticker Symbol name
  Dim Ticker_Name As String
  
  ' Set initial variable for holding opening amount
  Dim Open_Price As Double
  Open_Price = Cells(2, 3).Value
  
  ' Set initial variable to hold closing amount
  Dim Close_Price As Double
  Close_Price = 0

    ' Set initial variable to hold Price Change
  Dim Price_Change As Double
  Price_Change = 0

 ' Counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Set an initial variable for holding the total per Ticker Symbol
  Dim Ticker_Total As Double
  Ticker_Total = 0

  ' Keep track of the location for each Stock Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Stock Ticker Symbol
  For i = 2 To lastrow

    ' Check if we are still within the same Stock Ticker Symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      
       ' Add to the  Close Price
      Close_Price = Cells(i, 6).Value
      
      'Calculate Change in Price
      Price_Change = Close_Price - Open_Price
      
      ' Print the Price Change to the Summary Table
      Range("j" & Summary_Table_Row).Value = Price_Change
      
      'Calculate Percentage Change
    Range("k" & Summary_Table_Row).Value = Price_Change / Open_Price


      ' Print the Stock Ticker Symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0
      
       ' Reset the Close Price
      Close_Price = 0
      
      ' Reset Price Change
      Price_Change = 0
      
    ' Reset the Open Price
      Open_Price = Cells(i + 1, 3).Value
      


    ' If the cell immediately following a row is the same Stock Ticker Symbol...
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value

    End If

  Next i
  
  'Set initial variable for Stock Symbol for Greatest Volumne
  Dim Ticker_Vol As String
  Ticker_Vol = Cells(2, 9).Value
  
  'Set initial variable for Greatest Volumne Total
  Dim Great_Vol As Double
  Great_Vol = Cells(2, 12).Value
  
  ' Set initial Variable for Greatest % Increase Ticker Symbol
  Dim GR_Inc_Sym As String
  
   ' Set initial Variable for Greatest % Increase
  Dim GR_Inc_Chg As Double
  
  ' Set initial Variable for Greatest % Decrease Ticker Symbol
  Dim GR_Dec_Sym As String
  
   ' Set initial Variable for Greatest % Decrease
  Dim GR_Dec_Chg As Double
  
  ' Counts the number of rows
  lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
  
   ' Loop through all Stock Ticker Summary
  For i = 2 To lastrow2
  
 ' Format Change cells by positive or negative change
  If Cells(i, 10).Value >= 0 Then
Cells(i, 10).Interior.Color = vbGreen

ElseIf Cells(i, 10).Value < 0 Then
Cells(i, 10).Interior.Color = vbRed

End If

' Check if the volume for a given stock is the greatest
If Cells(i + 1, 12).Value > Great_Vol Then

Ticker_Vol = Cells(i + 1, 9).Value

Great_Vol = Cells(i + 1, 12).Value

End If

'Check for Greastest Percentage Increase
If Cells(i, 11).Value >= 0 And Cells(i, 11).Value > GR_Inc_Chg Then
    
    GR_Inc_Chg = Cells(i, 11).Value
    
    GR_Inc_Sym = Cells(i, 9).Value
    
End If
    
'Check for Greastest Percentage Decrease
If Cells(i, 11).Value < 0 And Cells(i, 11).Value < GR_Dec_Chg Then
    
    GR_Dec_Chg = Cells(i, 11).Value
    
    GR_Dec_Sym = Cells(i, 9).Value

End If

Next i

' Print Greatest Increase, Greatest Decrease, and Greatest Volumne
Cells(2, 16).Value = GR_Inc_Sym
Cells(2, 17).Value = GR_Inc_Chg
Cells(3, 16).Value = GR_Dec_Sym
Cells(3, 17).Value = GR_Dec_Chg
Cells(4, 16).Value = Ticker_Vol
Cells(4, 17).Value = Great_Vol


Next wsMySheet

Application.ScreenUpdating = True

End Sub


