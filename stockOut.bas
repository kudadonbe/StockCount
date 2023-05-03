
Sub stockOut()

    Dim pCode As Integer
    Dim reDate As String
    Dim reqNo As String
    Dim reqUser As String
    Dim reqQty As Integer

    Dim grf As Workbook
    Dim upDatingRow As Integer

    Set grf = ActiveWorkbook

    upDatingRow = ActiveCell.Row

    reDate = Range("A33").Value
    reqNo = Range("A5").Text
    reqUser = Range("E33").Text
    'MsgBox reqUser

    'Exit Sub

    pCode = Range("H" & upDatingRow).Value
    reqQty = Range("I" & upDatingRow).Value



    Dim stockBook As Workbook
    'Application.ScreenUpdating = False
    Set stockBook = Workbooks.Open("\\server\sections\Co-operate Affairs\Stock\stock_update_v6_2022.xlsx")
    'MsgBox pCode
    Set stockSheet = stockBook.Worksheets(pCode)

    stockSheet.Activate

    Dim lastRow As Long
    lastRow = Range("C" & Rows.Count).End(xlUp).Row
    newRow = lastRow + 1
    stockSheet.Range("A" & newRow).Value = reDate
    stockSheet.Range("C" & newRow).Value = reqNo
    stockSheet.Range("D" & newRow).Value = reqUser
    stockSheet.Range("G" & newRow).Value = reqQty
    'stockSheet.Range("G" & newRow).Select
    'stockSheet.Range("G" & newRow).Activate

    'MsgBox "LastRow " & lastRow
    stockBook.Save
    'stockBook.Close

    grf.Activate
    Set grfSheet = grf.Worksheets("Goods Requisition")
    grfSheet.Range("L" & upDatingRow).Value = ChrW(10003)
    grf.Save



End Sub