Sub ExcelIterate()
"    'In response to a question stating "Is there an Excel formula that will extract or highlight outstanding values in a column, where values within that column (when added together) equal a certain value?"
    Dim c As Range
    invoiceAmount = 0
    paymentAmount = 0

    For Each c In Range("B:B") ' Change this to your row size or leave for the entire column. If you have a header row, you'll need to skip it.
        ' This IF statement gets the invoiceAmount
        If c.Value < 0 Then 'I assume all values less than 0 are invoices?
            invoiceAmount = c.Value 'Reset the invoice amount each time we encounter a negative value. Assumption is that the cells are in order and paymentAmounts only map to a preceeding invoiceAmount.
            paymentAmount = 0  'If we find an invoice, I assume we zero out the payments? (See assumption above)
        End If
    
    'This IF statement adds up paymentAmounts and checks to see when the paymentAmounts get above 0 (after which, cells are filled in Yellow).
        If c.Value >= 0 Then 'Assume all values of 0 or more are payments
            paymentAmount = paymentAmount + c.Value 'Increment the paymentAmount
            If paymentAmount + invoiceAmount > 0 Then  'Only highlight cells when the paymentAmount starts going above the invoiceAmount. Highlight all cells where the balance is positive
                c.Interior.Color = RGB(255, 255, 0) 'When the condition is met, highlight the cell in yellow. This will highlight all cells where the balance is greater than 0
            End If
        End If
    Next c
End Sub

