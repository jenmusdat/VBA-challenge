Attribute VB_Name = "Module1"
Sub Stocks()
'define all my variable
Dim i As Long
Dim total As Double
Dim row_count As Long
Dim summary_row As Long
Dim closing As Double
Dim opening As Double
Dim change As Double
Dim Percent_change As Double
Dim xsh As Worksheet
Dim ws_name As String



    'how to name each header in all my worksheets
    Application.ScreenUpdating = False
        For Each xsh In Worksheets
                xsh.Select
                Range("I1:L1") = Array("Ticker Name", "Yearly Change", "Percent Change", "Total Stock Change")
                MsgBox (ws_name)
        Application.ScreenUpdating = True

    'summary row data will start in row 2
    summary_row = 2
    total = 0

    'This will get me to the end without having to know the number of rows
    row_count = Cells(Rows.Count, "A").End(xlUp).Row
    
    'my rows are going to start in row 2 to the end
    For i = 2 To row_count

        'this will give me the total in column 7
        total = total + Cells(i, 7).Value

        'If column a changes then i will grab new data
        If Cells(i, 1) <> Cells(i + 1, 1) Then

            'how to calulate the number for the yearly change column
            closing = Cells(i, 6).Value
            opening = Cells(i, 3).Value
            change = (closing - opening)
            If current_cell <> next_cell Then change2 = Cells(i, 6).Value
            If opening > 0 Then

                'how to calcualte the percent change from the opening to closing
                Percent_change = Round(change / opening, 4)
                NumberFormat = "0.00%"
                Else
                Percent_change = 0
                End If

                opening = Cells(i + 1, 3).Value

            'this colors the cell
            If change > 0 Then
            Cells(summary_row, 10).Interior.ColorIndex = 4
            ElseIf change < 0 Then
            Cells(summary_row, 10).Interior.ColorIndex = 3
            End If

    'where the results will go
    Cells(summary_row, 9) = Cells(i, 1)
    Cells(summary_row, 10) = change
    Cells(summary_row, 11) = Percent_change
    Cells(summary_row, 12) = total
    summary_row = summary_row + 1
    total = 0

    End If

Next i

Next xsh

End Sub
