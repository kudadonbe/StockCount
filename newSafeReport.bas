
Public Sub safeReport()
  Dim SafeData As Workbook
  Dim safeDataSheet As Worksheet
  
  Dim TemplateReport As Workbook
  Dim TemplateReportSheet As Worksheet

  Dim startDate As Date, endDate As Date
  

  Dim recDateArr(10000) As String
  Dim detailsArr(10000) As String
  Dim recNumArr(10000) As String
  Dim totalArr(10000) As String
  Dim isCanceledArr(10000) As String

  Dim glCodeArr(10000) As String
  Dim incomeCodeArr(10000) As String
  Dim activityArr(10000) As String
  Dim totalSum(10000) As Double

  Dim isWeekly As Boolean

  Dim isCanceled As String

  Dim recDate As String
  Dim recNum As String
  Dim details As String
  Dim total As Double
  Dim totalIncome As Double

  Dim glCode As String
  Dim incomeCode As String
  Dim activity As String

  Dim lastRow As Integer
  Dim i As Integer
  Dim filteredDataIndex As Integer
  Dim indexOfFoundIncomeCode As Integer
  Dim newIncomeRecord As Integer
  
  

  Dim payType As String


  Dim dialogBox As FileDialog
  Set dialogBox = Application.FileDialog(msoFileDialogFilePicker)
  
  ' Allow the user to select only one file
  dialogBox.AllowMultiSelect = False
  
  ' Set the title of the dialog box
  dialogBox.Title = "Select the data sheet file"
  
  ' Set the filters to show only Excel files
  dialogBox.Filters.Clear
  dialogBox.Filters.Add "Excel files", "*.xlsx; *.xlsm; *.xls"
  
  ' Show the dialog box and get the file path
  If dialogBox.Show = True Then
      Dim filePath As String
      filePath = dialogBox.SelectedItems(1)
      ' MsgBox "Selected file path: " & filePath
  Else
      MsgBox "No file selected."
  End If

 
  
  payType = "Cash"

  startDate = InputBox("Enter Starting date as DD/MM/YYYY")

  If IsDate(startDate) Then
      ' Do something with the date
    '   MsgBox startDate
  Else
      ' Handle invalid input
      MsgBox "Date should be in DD/MM/YYYY eg: 01/12/2023"
  End If

  isWeekly = MsgBox("Do you want a Weekly report?", vbYesNo) = vbYes
  Dim ReportType As String
  

  If (isWeekly) Then
      ' MsgBox "Generating Weekly report"
      endDate = DateAdd("d", 6, startDate) ' Set end date here
      ReportType = "WeeklyReport"
  Else
      ' MsgBox "Generating Monthlyreport"
      endDate = DateSerial(Year(startDate), Month(startDate) + 1, 0)
      ReportType = "MonthlyReport"
  End If

 
  Set SafeData = Workbooks.Open(filePath)
  Set safeDataSheet = SafeData.Worksheets("Sheet1")

  lastRow = safeDataSheet.Cells(safeDataSheet.Rows.Count, 1).End(xlUp).Row

  ' Loop through the rows in the worksheet and extract data between the date range
  filteredDataIndex = 0
  newIncomeRecord = 0
  For i = 2 To lastRow ' Assuming data starts from row 2

      If (safeDataSheet.Cells(i, "D") >= startDate And safeDataSheet.Cells(i, "D") <= endDate And safeDataSheet.Cells(i, "Y") = payType) Then
          
          recDate = safeDataSheet.Cells(i, "D")
          recNum = safeDataSheet.Cells(i, "E")
          details = safeDataSheet.Cells(i, "O")
          total = safeDataSheet.Cells(i, "U")
          isCanceled = safeDataSheet.Cells(i, "W")

          glCode = safeDataSheet.Cells(i, "J")
          incomeCode = safeDataSheet.Cells(i, "G")
          activity = safeDataSheet.Cells(i, "H")


          recDateArr(filteredDataIndex) = recDate
          recNumArr(filteredDataIndex) = recNum
          detailsArr(filteredDataIndex) = details
          totalArr(filteredDataIndex) = total
          isCanceledArr(filteredDataIndex) = isCanceled

          filteredDataIndex = filteredDataIndex + 1

          For indexOfFoundIncomeCode = 0 To UBound(incomeCodeArr)
              ' Debug.Print "looking for " + incomeCode
              If incomeCodeArr(indexOfFoundIncomeCode) = incomeCode Then
                  ' Found the income code, exit the loop and return the row index
                  Exit For
              End If
              
          Next indexOfFoundIncomeCode

          If indexOfFoundIncomeCode <= UBound(incomeCodeArr) Then
              ' Found the income code, do something with it (e.g. print the row data)
              'Debug.Print IncomeDetails(indexOfFoundIncomeCode, totalCol)
              totalSum(indexOfFoundIncomeCode) = totalSum(indexOfFoundIncomeCode) + total
          Else
              ' Debug.Print "New Income"
              glCodeArr(newIncomeRecord) = glCode
              incomeCodeArr(newIncomeRecord) = incomeCode
              activityArr(newIncomeRecord) = activity
              totalSum(newIncomeRecord) = total
              'Debug.Print "index of " & CStr(incomeCodeArr(newIncomeRecord)) & " is " & newIncomeRecord
              ' Income code not found

              newIncomeRecord = newIncomeRecord + 1
          End If
      End If
  Next i

  SafeData.Close SaveChanges:=False

  
  Dim totalIncomeRecord As Integer
  Dim totalTransRecord As Integer
  Dim ReportName As String
  Set TemplateReport = Workbooks.Open("S:\Co-operate Affairs\Safe\Templates\SafeReportTemplate.xltm")
  Set TemplateReportSheet = TemplateReport.Sheets(ReportType)

  ' to generate report first open weekly report sheet from tamplate
  ' put the start date at B9
  TemplateReportSheet.Range("B9").Value = startDate
  ' first Income records
  ' get total number of records
  ' copy formula at K10
  Dim summingFormulaRange As Range
  Dim summingFormula As String
  summingFormula = TemplateReportSheet.Range("K10").Formula
  
  
  ' MsgBox (newIncomeRecord) & " inserting to Income Records"
  ' insert rows at 17

  If (newIncomeRecord > 3) Then
      TemplateReportSheet.Rows("16:" & (16 + newIncomeRecord - 3)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  End If

  ' enter records
  Dim nextIncomeRecord As Integer
  For nextIncomeRecord = 0 To (newIncomeRecord - 1)
      'Debug.Print CStr(glCodeArr(y)) & " | " & CStr(incomeCodeArr(y)) & " | " & CStr(activityArr(y)) & " | " & CStr(totalSum(y))
      TemplateReportSheet.Range("B" & 16 + nextIncomeRecord).Value = glCodeArr(nextIncomeRecord)
      TemplateReportSheet.Range("C" & 16 + nextIncomeRecord).Value = incomeCodeArr(nextIncomeRecord)
      TemplateReportSheet.Range("D" & 16 + nextIncomeRecord).Value = activityArr(nextIncomeRecord)
      TemplateReportSheet.Range("K" & 16 + nextIncomeRecord).Value = totalSum(nextIncomeRecord)
  Next nextIncomeRecord


  ' than get number of transactions
  ' MsgBox filteredDataIndex & " inserting to Trans Records"
  ' insert row at 11
  If (filteredDataIndex > 3) Then
      TemplateReportSheet.Rows("10:" & (10 + filteredDataIndex - 3)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
  End If
  
  ' enter records with formula to K
  

  Dim newTransRecords As Integer
  For newTransRecords = 0 To (filteredDataIndex - 1)

      TemplateReportSheet.Range("B" & 10 + newTransRecords).Value = recDateArr(newTransRecords)
      TemplateReportSheet.Range("C" & 10 + newTransRecords).Value = recNumArr(newTransRecords)
      TemplateReportSheet.Range("D" & 10 + newTransRecords).Value = detailsArr(newTransRecords)
      TemplateReportSheet.Range("I" & 10 + newTransRecords).Value = totalArr(newTransRecords)

      If (isCanceledArr(newTransRecords) = "Yes") Then
          TemplateReportSheet.Range("B" & 10 + newTransRecords & ":I" & 10 + newTransRecords).Font.Color = vbRed 'change font color to red
          TemplateReportSheet.Range("B" & 10 + newTransRecords & ":I" & 10 + newTransRecords).Font.Strikethrough = True 'add strikethrough
      Else
          TemplateReportSheet.Range("B" & 10 + newTransRecords & ":I" & 10 + newTransRecords).Font.Color = vbBlack 'change font color to red
          TemplateReportSheet.Range("B" & 10 + newTransRecords & ":I" & 10 + newTransRecords).Font.Strikethrough = False 'add strikethrough
      End If
  Next newTransRecords

  
  Set summingFormulaRange = TemplateReportSheet.Range("K10:K" & (10 + filteredDataIndex))
  TemplateReportSheet.Range("K10").Formula = summingFormula
  TemplateReportSheet.Range("K10").Copy
  summingFormulaRange.PasteSpecial xlPasteFormulas

  

  ' copy report name from J3
  ReportName = TemplateReportSheet.Range("J3").Value
  Dim ReportFilePath As String
  
  ReportFilePath = "S:\Co-operate Affairs\Safe\2023\Reports\" & ReportName & ".xlsx"
  ' save report
  Dim newReport As Workbook
  Set newReport = Workbooks.Add
  TemplateReportSheet.Copy Before:=newReport.Sheets(1)

  ' While newReport.Sheets.Count > 1
  '     newReport.Sheets(newReport.Sheets.Count).Delete
  ' Wend

  newReport.SaveAs ReportFilePath
  TemplateReport.Close SaveChanges:=False

End Sub