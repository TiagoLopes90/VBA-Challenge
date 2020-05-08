Sub Homework02():
    For Each ws In Worksheets

        'Selecting & Sort Cells by Header First Column - just to make surea all data is correclty sorted
        ws.Range("A1", ws.Range("A1").End(xlDown).End(xlToRight)).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Key2:=ws.Range("B1"), Order2:=xlAscending, Header:=xlYes

        'Header on the table
        ws.Cells(1,9).Value = "Ticker"
        ws.Cells(1,10).Value = "Yearly Change"
        ws.Cells(1,11).Value = "Percentage Change"
        ws.Cells(1,12).Value = "Total Volume Stock"

        Dim TVS as Double 'define Total Volume Stock Variable
        Dim Ybegin as Double 'define Year Begining Variable
        Dim Yend as Double 'define Year End Variable
        Dim tableRow as integer 'define tablerow to iterate summary table
        Dim yearChange as Double 'Variable to calculate difference between YearB and Year end
        'Defining Ticker and Value of the Greatest Increase
        Dim bigger as Double
        Dim bigger_ticket as String
        'Defining Ticker and Value of the Greatest Decrease
        Dim lower as Double
        Dim lower_ticket as String
        'Defining Ticker and Value of Higher Volume
        Dim biggest as Double
        Dim biggest_ticket as String

        'To create Summary Table: row count
        tableRow = 2

        'Determine the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


        'Define the first Year Begin
        Ybegin = ws.Cells(2,3).Value
        TVS = ws.Cells(2,7).Value

        'Cycle across all data
        For i = 2 to lastRow:  
            'Check either we are under the same ticker or not
            If ws.Cells(i+1,1).Value <> ws.Cells(i,1).Value Then

                'Defining Year End Variable and calculate last TVS
                Yend = ws.Cells(i,6).Value
                TVS = TVS+ws.Cells(i,7)

                'Print all Results in the table and calculate Year Change and % of Change
                yearChange = Yend - Ybegin
                ws.Cells(tableRow,10).Value = yearChange
                
                'Format Color based on Year Change Value
                If yearChange >=0 Then
                    ws.Cells(tableRow,10).Interior.ColorIndex = 4
                Else
                    ws.Cells(tableRow,10).Interior.ColorIndex = 3
                End If

                'Calculate Year Change % and adjust in case Ybegin = 0
                If Ybegin = 0 Then 
                    ws.Cells(tableRow,11).Value= 0
                Else
                    ws.Cells(tableRow,11).Value= yearChange/Ybegin
                    ws.Cells(tableRow,11).NumberFormat = "0.00%"
                End If

                'Formating Variables
                ws.Cells(tableRow,12).Value = TVS
                ws.Cells(tableRow,12).NumberFormat = "#,##0"
                ws.Cells(tableRow,9).Value = ws.Cells(i,1).Value

                'Reset all Variables
                TVS = 0
                Ybegin = ws.Cells(i+1,3).Value
                'Increment 1 Row
                tableRow = tableRow+1
                
            Else
                'Continue calculating TVS
                TVS = TVS + ws.Cells(i,7).Value
            End If

        Next i

        'Challenge - create the second small table:

        'Determine the last row of summary table
        lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Define Names of Columns & Lines
        ws.Cells(2,15).Value = "Greatest % Decrease"
        ws.Cells(3,15).Value = "Greatest % Increase"
        ws.Cells(4,15).Value = "Greatest Total Volume"
        ws.Cells(1,16).Value = "Ticker"
        ws.Cells(1,17).Value = "Value"

        'Define first Variables & Tickets
        bigger = ws.Cells(2,11).Value
        bigger_ticket = ws.Cells(2,9).Value
        lower = ws.Cells(2,11).Value
        lower_ticket = ws.Cells(2,9).Value
        biggest = ws.Cells(2,12).Value
        biggest_ticket = ws.Cells(2,9).Value

        'Cycle to define the Greatest Increase, Decrease and Higher Volume
        For j = 3 to lastRow2:
            if ws.Cells(j,11).Value>bigger Then
                bigger = ws.Cells(j,11).Value
                bigger_ticket = ws.Cells(j,9).Value
            End If
            if ws.Cells(j,11).Value<lower Then
                lower = ws.Cells(j,11).Value
                lower_ticket = ws.Cells(j,9).Value
            End If
            if ws.Cells(j,11).Value>biggest Then
                biggest = ws.Cells(j,11).Value
                biggest_ticket = ws.Cells(j,9).Value
            End If
        Next j

        'Defining Challenge Table and Formatting
        ws.Cells(2,16).Value = bigger_ticket
        ws.Cells(3,16).Value = lower_ticket
        ws.Cells(4,16).Value = biggest_ticket
        ws.Cells(2,17).Value = bigger
        ws.Cells(2,17).NumberFormat = "0.00%"
        ws.Cells(3,17).Value = lower
        ws.Cells(3,17).NumberFormat = "0.00%"
        ws.Cells(4,17).Value = biggest
        ws.Cells(4,17).NumberFormat = "#,##0"

    Next ws
End Sub