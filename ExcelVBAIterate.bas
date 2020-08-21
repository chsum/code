Sub ExcelIterate()
    'In response to https://twitter.com/yatesamantha/status/1296822923358085127?s=20
    Dim c As Range
    invoiceAmount = 0
    paymentAmount = 0

    For Each c In Range("B2:B31") ' Change this to your row size
        If c.Value < 0 Then 'I assume values less than 0 are invoices?
            invoiceAmount = c.Value
            paymentAmount = 0  'If we find an invoice, I assume we zero out the payments?
        End If
        If c.Value >= 0 Then
            paymentAmount = paymentAmount + c.Value
            If paymentAmount + invoiceAmount >= 0 Then
                c.Interior.Color = RGB(255, 255, 0)
            End If
        End If
    Next c
End Sub
