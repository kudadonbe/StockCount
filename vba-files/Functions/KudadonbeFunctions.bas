Attribute VB_Name = "KudadonbeFunctions"
Public  Function updateStockCount()
    ' after testing implement data loading here

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
        ' openStockWorkbook
        Dim stockCountWorkbook As Workbook
        Dim stockCountSheet As Worksheet
        
        Set stockCountWorkbook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\StockCount.xlsx")
        ' set stockCountSheet
        Set stockCountSheet = stockCountWorkbook.Sheets("Form")
        itemNo = stockCountSheet.Range("B2").value
        itemName(1) = stockCountSheet.Range("B4").value
        itemName(2) = stockCountSheet.Range("C4").value
        itemBrand(1) = stockCountSheet.Range("B5").value
        itemBrand(2) = stockCountSheet.Range("C5").value
        itemStatus(1) = stockCountSheet.Range("B6").value
        itemStatus(2) = stockCountSheet.Range("C6").value
        itemLocation(1) = stockCountSheet.Range("B7").value
        itemLocation(2) = stockCountSheet.Range("C7").value
        itemSupplier(1) = stockCountSheet.Range("B8").value
        itemSupplier(2) = stockCountSheet.Range("C8").value
        itemOtherInfo(1) = stockCountSheet.Range("B9").value
        itemOtherInfo(2) = stockCountSheet.Range("C9").value
        itemQty = stockCountSheet.Range("B10").value
        itemModel = stockCountSheet.Range("B11").value
        itemSerialNo = stockCountSheet.Range("B12").value
        itemSize = stockCountSheet.Range("B13").value
        itemWeight = stockCountSheet.Range("B14").value
        itemReceivedDate = stockCountSheet.Range("B15").value
        itemPrice = stockCountSheet.Range("B16").value
        itemCountedBy = stockCountSheet.Range("B17").value
    
        ' open stock workbook
        Dim stockWorkbook As Workbook
        Set stockWorkbook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\stock_update_v6_2022.xlsx")
        ' navigate to stock sheet
        Dim stockSheet As Worksheet
        Dim contentSheet As Worksheet
        ' select items to update
        Set contentSheet = stockWorkbook.Sheets("Content")
        ' find item in stock sheet
        Dim foundItemRow As Long
        foundItemRow = Application.Match(CLng(itemNo), contentSheet.Range("B:B"), 0)
        ' if item is found update the item Information
        If Not IsError(foundItemRow) Then
            ' update values
            If itemBrand(1) <> "" Then contentSheet.Range("F" & foundItemRow).value = itemBrand(1)
            If itemStatus(1) <> "" Then contentSheet.Range("G" & foundItemRow).value = itemStatus(1)
            If itemLocation(1) <> "" Then contentSheet.Range("H" & foundItemRow).value = itemLocation(1)
            If itemSupplier(1) <> "" Then contentSheet.Range("I" & foundItemRow).value = itemSupplier(1)
            If itemOtherInfo(1) <> "" Then contentSheet.Range("J" & foundItemRow).value = itemOtherInfo(1)
            If itemModel <> "" Then contentSheet.Range("K" & foundItemRow).value = itemModel
            If itemSerialNo <> "" Then contentSheet.Range("L" & foundItemRow).value = itemSerialNo
            If itemSize <> "" Then contentSheet.Range("M" & foundItemRow).value = itemSize
            If itemWeight <> "" Then contentSheet.Range("N" & foundItemRow).value = itemWeight
            If itemPrice <> 0 Then contentSheet.Range("O" & foundItemRow).value = itemPrice
            If itemName(2) <> "" Then contentSheet.Range("R" & foundItemRow).value = itemName(2)
            If itemBrand(2) <> "" Then contentSheet.Range("S" & foundItemRow).value = itemBrand(2)
            If itemStatus(2) <> "" Then contentSheet.Range("T" & foundItemRow).value = itemStatus(2)
            If itemLocation(2) <> "" Then contentSheet.Range("U" & foundItemRow).value = itemLocation(2)
            If itemSupplier(2) <> "" Then contentSheet.Range("V" & foundItemRow).value = itemSupplier(2)
            If itemOtherInfo(2) <> "" Then contentSheet.Range("W" & foundItemRow).value = itemOtherInfo(2)
    
            ' after updating the item information from Content sheet goto the item sheet and update the stock count
            Set stockSheet = stockWorkbook.Sheets(itemNo)
            Dim systemCount As Integer
            Dim physicalCount As Integer
            Dim lastEntryRow As Long
            Dim countDifference As Integer
            physicalCount = itemQty
    
            lastEntryRow = stockSheet.Range("A" & Rows.Count).End(xlUp).Row
            systemCount = stockSheet.Range("H" & lastEntryRow).value
    
            countDifference = physicalCount - systemCount
            If physicalCount < systemCount Then
                stockSheet.Range("I" & lastEntryRow).value = physicalCount
                stockSheet.Range("J" & lastEntryRow).Formula = "=I" & lastEntryRow & "-H" & lastEntryRow
                If countDifference < 0 Then countDifference = (countDifference * -1)
                stockSheet.Range("G" & lastEntryRow + 1).value = countDifference
                stockSheet.Range("A" & lastEntryRow + 1).value = itemReceivedDate
                stockSheet.Range("C" & lastEntryRow + 1).value = itemOtherInfo(1)
                stockSheet.Range("D" & lastEntryRow + 1).value = itemCountedBy
            ElseIf physicalCount > systemCount Then
    
                stockSheet.Range("I" & lastEntryRow).value = physicalCount
                stockSheet.Range("J" & lastEntryRow).Formula = "=I" & lastEntryRow & "-H" & lastEntryRow
    
                stockSheet.Range("F" & lastEntryRow + 1).value = countDifference
                stockSheet.Range("A" & lastEntryRow + 1).value = itemReceivedDate
                stockSheet.Range("C" & lastEntryRow + 1).value = itemOtherInfo(1)
                stockSheet.Range("D" & lastEntryRow + 1).value = itemCountedBy
            Else
                stockSheet.Range("I" & lastEntryRow).value = physicalCount
                stockSheet.Range("J" & lastEntryRow).Formula = "=I" & lastEntryRow & "-H" & lastEntryRow
                stockSheet.Range("K" & lastEntryRow).value = itemReceivedDate
            End If
        Else
            ' if item is not found add the item to the stock sheet
            ' call newItem function
        End If
        
  
    stockWorkbook.Save ' save the stock workbook
End Function