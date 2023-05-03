Public  Sub newForm()
    Dim newFormWorkbook As Workbook
    Dim newForm As Worksheet
    Dim dataSheet As Worksheet

    Dim lastEntryRow As Long
    
    Dim itemNo As String
    Dim itemName(1 To 2) As String
    Dim itemBrand(1 To 2) As String
    Dim itemStatus(1 To 2) As String
    Dim itemLocation(1 To 2) As String
    Dim itemSupplier(1 To 2) As String
    Dim itemOtherInfo(1 To 2) As String
    Dim itemQty As Integer
    Dim itemModel As String
    Dim itemSerialNo As String
    Dim itemSize As String
    Dim itemWeight As String
    Dim itemReceivedDate As Date
    Dim itemPrice As Double
    Dim itemCountedBy As String
    
    Set newFormWorkbook = ActiveWorkbook
    ' check if opened workbook is stock count workbook
    If newFormWorkbook.Name <> "StockCount.xlsx" Then
        Set newFormWorkbook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\StockCount.xlsx")
    End If

    ' set newForm
    Set newForm = newFormWorkbook.Sheets("Entry Form")
    itemNo = newForm.Range("B2").Value
    itemName(1) = newForm.Range("B4").Value
    itemName(2) = newForm.Range("C4").Value
    itemBrand(1) = newForm.Range("B5").Value
    itemBrand(2) = newForm.Range("C5").Value
    itemStatus(1) = newForm.Range("B6").Value
    itemStatus(2) = newForm.Range("C6").Value
    itemLocation(1) = newForm.Range("B7").Value
    itemLocation(2) = newForm.Range("C7").Value
    itemSupplier(1) = newForm.Range("B8").Value
    itemSupplier(2) = newForm.Range("C8").Value
    itemOtherInfo(1) = newForm.Range("B9").Value
    itemOtherInfo(2) = newForm.Range("C9").Value
    itemQty = newForm.Range("B10").Value
    itemModel = newForm.Range("B11").Value
    itemSerialNo = newForm.Range("B12").Value
    itemSize = newForm.Range("B13").Value
    itemWeight = newForm.Range("B14").Value
    itemReceivedDate = newForm.Range("B15").Value
    itemPrice = newForm.Range("B16").Value
    itemCountedBy = newForm.Range("B17").Value

    Set dataSheet = newFormWorkbook.Sheets("Data")
    ' find last row entry row in data sheet
    lastEntryRow = dataSheet.Range("A" & Rows.Count).End(xlUp).Row + 1

    ' write data to data sheet
    dataSheet.Range("A" & lastEntryRow).Value = itemNo
    dataSheet.Range("B" & lastEntryRow).Value = itemName(1)
    dataSheet.Range("C" & lastEntryRow).Value = itemName(2)
    dataSheet.Range("D" & lastEntryRow).Value = itemBrand(1)
    dataSheet.Range("E" & lastEntryRow).Value = itemBrand(2)
    dataSheet.Range("F" & lastEntryRow).Value = itemStatus(1)
    dataSheet.Range("G" & lastEntryRow).Value = itemStatus(2)
    dataSheet.Range("H" & lastEntryRow).Value = itemLocation(1)
    dataSheet.Range("I" & lastEntryRow).Value = itemLocation(2)
    dataSheet.Range("J" & lastEntryRow).Value = itemSupplier(1)
    dataSheet.Range("K" & lastEntryRow).Value = itemSupplier(2)
    dataSheet.Range("L" & lastEntryRow).Value = itemOtherInfo(1)
    dataSheet.Range("M" & lastEntryRow).Value = itemOtherInfo(2)
    dataSheet.Range("N" & lastEntryRow).Value = itemQty
    dataSheet.Range("O" & lastEntryRow).Value = itemModel
    dataSheet.Range("P" & lastEntryRow).Value = itemSerialNo
    dataSheet.Range("Q" & lastEntryRow).Value = itemSize
    dataSheet.Range("R" & lastEntryRow).Value = itemWeight
    dataSheet.Range("S" & lastEntryRow).Value = itemReceivedDate
    dataSheet.Range("T" & lastEntryRow).Value = itemPrice
    dataSheet.Range("U" & lastEntryRow).Value = itemCountedBy
    newFormWorkbook.Save
End Sub