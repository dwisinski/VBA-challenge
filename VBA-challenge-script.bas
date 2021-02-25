Attribute VB_Name = "Module1"
' Main Script
Sub stockData()
    
    ' Definitions
    Dim row As Long
    Dim col As Integer
    Dim totalVolume As LongLong
    Dim tableRow As Long
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim triggerOpen As Double
    Dim triggerClose As Double
    Dim tickerlastrow As Long
    Dim formatlastrow As Integer
    Dim wsName As String
    
    ' Insert column headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    ' Define and insert Worksheet name
    wsName = ActiveSheet.Name
    
    Range("N1") = wsName
        
    ' Get last row for Column A
    tickerlastrow = Cells(Rows.Count, 1).End(xlUp).row
    
    ' Prepare data table and set first ticker row trigger
    tableRow = 2
    trigger = 1
    
    ' Define range parameters
    For row = 2 To tickerlastrow
    
    ' This conditional is necessary to throw out stocks with "0" values which otherwise would cause an error
    If Not Cells(row, 3).Value = 0 Then
    
    ' For first ticker row trigger, record value of yearly opening price
         If trigger = 1 Then
            triggerOpen = Cells(row, 3).Value
            trigger = 0
                   
         End If
  
    ' Create running total of volume for each ticker
    totalVolume = Cells(row, 7).Value + totalVolume
        
            ' Create conditionals tied to last ticker row trigger
            If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
        
            ' For last ticker row trigger, record value of yearly closing price
            triggerClose = Cells(row, 6).Value
            ' Calculate yearly change as yearly closing price - yearly opening price
            yearlyChange = (triggerClose - triggerOpen)
                                       
            ' Calculate yearly percentage change as yearly change / yearly opening price
            percentChange = (yearlyChange / triggerOpen)
                                           
            ' Insert recorded values into data table
            Cells(tableRow, 9).Value = Cells(row, 1).Value
            Cells(tableRow, 10).Value = yearlyChange
            Cells(tableRow, 11).Value = percentChange
            Cells(tableRow, 11).NumberFormat = "0.00%"
            Cells(tableRow, 12).Value = totalVolume
            
            ' Iterate through the data table rows
            tableRow = tableRow + 1
            
            ' Reset the volume count
            totalVolume = 0
            
            ' Reset the recorded value of yearly opening price
            triggerOpen = 0
            
            ' Reset the first ticker row trigger
            trigger = 1
            
            End If
                            
        Else
        
        End If
        
    Next row

    ' Conditional on Column K and other misc. formatting
    formatlastrow = Cells(Rows.Count, 11).End(xlUp).row
    
    For row = 2 To formatlastrow
        For col = 11 To 11
        
        If Cells(row, col).Value < 0 Then
            Cells(row, col).Interior.ColorIndex = 3
        Else
            Cells(row, col).Interior.ColorIndex = 10
        
        End If
        
        Next col
    
    Next row
  
    Call getAdditional
    
End Sub
' Additional Table Script
Sub getAdditional()

    ' Definitions
    Dim row As Long
    Dim additionallastRow As Long
    Dim minvalue As Double
    Dim maxValue As Double
    Dim maxVolume As LongLong
    Dim tickerMinValue As String
    Dim tickerMaxValue As String
    Dim tickerMaxVolume As String

    ' Get last row for additonal table
    additionallastRow = Cells(Rows.Count, 9).End(xlUp).row
    
    ' Set initial values for use in loop
    maxValue = Cells(2, 11).Value
    minvalue = Cells(2, 11).Value
    maxVolume = Cells(2, 12).Value
        
    ' Iterate through additonal table rows
    For row = 2 To (additionallastRow - 1)
    
            If Cells(row + 1, 11).Value > maxValue Then
            maxValue = Cells(row + 1, 11).Value
            End If
            
            If Cells(row + 1, 11).Value = maxValue Then
            tickerMaxValue = Cells(row + 1, 9).Value
            End If
            
            If Cells(row + 1, 11).Value < minvalue Then
            minvalue = Cells(row + 1, 11).Value
            End If
            
            If Cells(row + 1, 11).Value = minvalue Then
            tickerMinValue = Cells(row + 1, 9).Value
            End If
            
            If Cells(row + 1, 12).Value > maxVolume Then
            maxVolume = Cells(row + 1, 12).Value
            End If
            
            If Cells(row + 1, 12).Value = maxVolume Then
            tickerMaxVolume = Cells(row + 1, 9).Value
            End If
                       
    Next row
    
    ' Insert values into additonal table
    Range("O2") = tickerMaxValue
    Range("P2") = maxValue
    Range("P2").NumberFormat = "0.00%"

    Range("O3") = tickerMinValue
    Range("P3") = minvalue
    Range("P3").NumberFormat = "0.00%"
    
    Range("O4") = tickerMaxVolume
    Range("P4") = maxVolume

    ' Auto-fit column widths after all data is inserted
    Range("A:P").Columns.AutoFit
    
End Sub
