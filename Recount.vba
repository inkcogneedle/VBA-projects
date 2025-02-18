Sub Recount_clearFunction()

'

' clearFunction Macro

'

    Dim ws As Worksheet

    Dim findWords As Variant, word As Variant

    Dim cell As Range, rng As Range                                                                                                                                                                                                                                           

    Dim startPos As Long


    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Application.Calculation = xlCalculationManual

    Application.EnableEvents = False

    Application.DisplayStatusBar = False
   

    ' clears all information

    Cells.Clear
   

    ' formats cells in column A back to normal

    With Range("A:A")

        .Font.Name = "Calibri"

        .Font.Size = 11

    End With
   

    ' default message for range of cells

    Range("A1").Value = "1. Before using, delete old excel files that have the name as the following file: Report"

    Range("A2").Value = "2. Go to All Items and Attributes on the SCI landing page, select the play button, and then Run Excel (example shown below)"

    Range("A19").Value = "3. In iLPNs UI, search using the TCIN and ASN provided on the SharePoint FDC Recount, and enter them into the fields (example shown below)"

    Range("A31").Value = "4. Within the iLPNs UI, click on 'Tools' and then 'Report' to save the excel file. Then click the Run button for result"
   

    Set ws = ThisWorkbook.Sheets("Recount") ' change to your sheet name

    Set rng = ws.Range("A1:A2") ' change to the range you want searched
   

    ' array of words to search and bold for cells A1 and A2

    findWords = Array("Report", "All Items and Attributes")
   

    ' Loop through each cell in the range

    For Each cell In rng

        If Not IsEmpty(cell.Value) Then

            ' Loop through each word to search for

            For Each word In findWords

                startPos = 1

                ' Find all occurrences of the word in the cell

                Do

                    startPos = InStr(startPos, cell.Value, word, vbTextCompare)

                    If startPos > 0 Then

                        ' Bold the word

                        cell.Characters(startPos, Len(word)).Font.Bold = True

                        startPos = startPos + Len(word) ' Move past the current word

                    End If

                Loop While startPos > 0

            Next word

        End If

    Next cell
   

    Set rng = ws.Range("A19") ' change to the range you want searched
   

    ' array of words to search and bold for cell A19

    findWords = Array("TCIN", "ASN")
   

    ' Loop through cell A19

    For Each cell In rng

        If Not IsEmpty(cell.Value) Then

            ' Loop through each word to search for

            For Each word In findWords

                startPos = 1

                ' Find all occurrences of the word in the cell

                Do

                    startPos = InStr(startPos, cell.Value, word, vbTextCompare)

                    If startPos > 0 Then

                        ' Bold the word

                        cell.Characters(startPos, Len(word)).Font.Bold = True

                        startPos = startPos + Len(word) ' Move past the current word

                    End If

                Loop While startPos > 0

            Next word

        End If

    Next cell
   

    ' checks to see if sheet exists and if so, makes sheet active and prompts for delete

    If Recount_doExists("Recount-Result") Then

        ThisWorkbook.Sheets("Recount-Result").Activate

        Recount_closeSheet

    End If
   

    ' Re-enable screen updating and alerts

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

    Application.Calculation = xlCalculationAutomatic

    Application.EnableEvents = True

    Application.DisplayStatusBar = True

 
    ' clears object variables from memory

    Set ws = Nothing

    Erase findWords

End Sub



Sub Recount_closeSheet()

'

' closeSheet Macro

'

    On Error Resume Next ' incase the sheet isn't there

    Sheets("Recount-Result").Delete

    On Error GoTo 0 ' reset error handling

End Sub



Sub Recount_copyData()

'

'   copyData Macro

'

    Dim targetWorkbook As Workbook

    Dim sourceSheet, targetSheet As Worksheet

    Dim sourceFile As String, userProfile As String

    Dim lastRow, lastCol, targetRow As Long
   

    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")


    ' Set the source file path (closed workbook)

    sourceFile = userProfile & "\Downloads\Report.xls"
   

    ' validate if file exists and if not, exit the Macro

    If Not Recount_fileExists(sourceFile) Then

        MsgBox "File not found or needs to be renamed as 'Report'"

        Exit Sub

    End If


    ' Set the target workbook (the workbook from which you are running this macro)

    Set targetWorkbook = ThisWorkbook


    ' checks to see if the sheet exist, if so, creates a new sheet or uses existing sheet

    If Not Recount_doExists("Recount-Result") Then

    ' Create a new sheet in the target workbook to paste the data

        Set targetSheet = targetWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Recount"))

        targetSheet.Name = "Recount-Result" ' Change to your target sheet name

        setupPrintMargins targetSheet ' sets the margins to 0 for left and right for print setup

    Else

        ' if the sheet exists already, then assigns to variable without creating a new sheet

        Set targetSheet = targetWorkbook.Sheets("Recount-Result")

    End If
  

    ' Open the source workbook in the background (without showing it)

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Application.Calculation = xlCalculationManual

    Application.EnableEvents = False

    Application.DisplayStatusBar = False


    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=False)


    ' Set the source and target worksheets

    Set sourceSheet = sourceWorkbook.Sheets(1) ' Change to your index or name if needed
   

    ' Delete rows 1 to 2 from source sheet

    sourceSheet.Rows("1:2").Delete


    Recount_searchAndOrganizeColumns

    Recount_searchAndRetrieveTiHi


    ' if the Macro cannot find the file "All Item Attributes" it'll end and close everything

    Dim sourceFilePath As String

    ' Define the path of the source workbook (external workbook) to check if the file exists

    sourceFilePath = userProfile & "\Downloads\All Item Attributes.xlsx" ' Update the path as needed

    ' validate if file "All Item Attributes exist", and if not, ends Macro

    If Not Recount_fileExists(sourceFilePath) Then

        On Error Resume Next ' incase the sheet isn't there

        targetSheet.Delete ' deletes the sheet if it was created

        ' Close the source workbook without saving

        sourceWorkbook.Close SaveChanges:=False

        On Error GoTo 0 ' reset error handling

        Exit Sub ' exits out of the main sub

    End If
  

    ' sorts the active location by ascending order from A1 to last row of column F

    With Range("A1:F" & Cells(Rows.Count, "A").End(xlUp).Row)

        .Sort key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes

    End With

 
    ' Find the last row with data in Column G of the source sheet

    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "H").End(xlUp).Row
   

    ' Copy data from A1 to G(lastRow) in the source sheet

   sourceSheet.Range("A1:H" & lastRow).Copy
   

    ' checks if the first row is empty, and if so, is the starting place to add information

    Set emptyCell = targetSheet.Range("A1")

    If emptyCell.Value = "" Then

        targetRow = 1

    Else

        ' find the next available row in the target sheet

        targetRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 3

    End If
   

    ' Paste the copied data into the target sheet

    targetSheet.Cells(targetRow, 1).PasteSpecial Paste:=xlPasteValues
   

    ' Close the source workbook without saving

    sourceWorkbook.Close SaveChanges:=False
   

    ' creates ti and hi header with format

    targetSheet.Range("G" & targetRow).Value = "Ti"

    targetSheet.Range("H" & targetRow).Value = "Hi"

    targetSheet.Range("G" & targetRow & ":H" & targetRow).Font.Bold = True

    targetSheet.Columns("G:H").HorizontalAlignment = xlCenter
   

    ' Skip 2 row and then add "Expected" and "Received" values

    targetRow = targetRow + lastRow - 1 + 2 ' Move 2 row down after the copied data


    ' adds cell value expected and received

    targetSheet.Cells(targetRow, 1).Value = "Expected"

    targetSheet.Cells(targetRow, 2).Value = "Received"

    
    ' styles the cell to be outlined all black

    Set cell = targetSheet.Range(targetSheet.Cells(targetRow, 1), targetSheet.Cells(targetRow + 1, 2))

    With cell.Borders

        .LineStyle = xlContinuous

        .Color = vbBlack

        .TintAndShade = 0

    End With

    cell.HorizontalAlignment = xlCenter


    Recount_findAndReplace ' find and replace headers

    targetSheet.Cells.EntireColumn.AutoFit ' widen all cells to fit content on the sheet


    ' Optional: Clear the clipboard to remove the copied data

    Application.CutCopyMode = False
   

    ' Re-enable screen updating and alerts

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

    Application.Calculation = xlCalculationAutomatic

    Application.EnableEvents = True

    Application.DisplayStatusBar = True
   

    ' check if the file exists before trying to delete it

    If Dir(sourceFile) <> "" Then

        ' delete the file

        Kill sourceFile

    End If

     setupPrintMargins targetSheet


    Set sourceWorksheet = Nothing

    Set targetWorksheet = Nothing

    Set sourceSheet = Nothing

    Set targetSheet = Nothing

End Sub



Function Recount_doExists(sheetName As String) As Boolean

    '

    ' doExists function

    '

    On Error Resume Next

    Recount_doExists = Not Worksheets(sheetName) Is Nothing

    On Error GoTo 0

End Function



Function Recount_fileExists(filePath As String) As Boolean

'

'   fileExists function

'

    If Dir(filePath) <> "" Then

        Recount_fileExists = True

    Else

        Recount_fileExists = False

    End If

End Function



Sub Recount_findAndReplace()

'

'   findAndReplace

'

    Dim ws As Worksheet

    Dim targetRange As Range

    Dim findWords() As Variant

    Dim replaceWords() As Variant

    Dim i As Integer
   

    ' Set the worksheet and range to search within

    Set ws = ThisWorkbook.Sheets("Recount-Result") ' Replace with your sheet name

    Set targetRange = ws.UsedRange ' Replace with your desired range
   

    ' Define the find words and replace words arrays

    findWords = Array("Current Location", "LPN Quantity", "Expiration date", "LPN Facility Status") ' Add words to find

    replaceWords = Array("Location", "Quantity", "Exp date", "Description") ' Corresponding words to replace with

 
    ' Loop through each word in the findWords array

    For i = LBound(findWords) To UBound(findWords)

        ' Replace each find word with its corresponding replace word

        targetRange.Replace What:=findWords(i), Replacement:=replaceWords(i), _

                            LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False

    Next i
   

    Set ws = Nothing

    Erase findWords

    Erase replaceWords

End Sub



Sub Recount_searchAndOrganizeColumns()

'

'   searchAndOrganizeColumns

'

    Dim sFile As String, userProfile As String

    Dim wb As Workbook

    Dim ws As Worksheet

    Dim searchWords As Variant

    Dim i As Integer, j As Integer

    Dim foundColumn As Range

    Dim destColumn As Integer

    Dim word As String
   

    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")

  
    ' Set the source file path (closed workbook)

    sFile = userProfile & "\Downloads\Report.xls"
   

    ' Define your worksheet

    Set wb = Workbooks.Open(sFile, ReadOnly:=False) ' Adjust Sheet name if needed


    Set ws = wb.Sheets(1)
   

    ' Define the search words

    searchWords = Array("Current Location", "LPN", "Item", "LPN Quantity", "Expiration date", "LPN Facility Status") ' Add your words here


    ' Initialize the destination column (where we will paste the found columns)

    destColumn = 1 ' Start pasting from column A (1 = A, 2 = B, etc.)
   

    ' Loop through each search word

    For i = LBound(searchWords) To UBound(searchWords)

        word = searchWords(i)


        ' Look for the word in the first row (or header row)

        Set foundColumn = Nothing

        For j = 1 To ws.UsedRange.Columns.Count

            If ws.Cells(1, j).Value = word Then

                Set foundColumn = ws.Columns(j)

                Exit For

            End If

        Next j
   

        ' If found, copy the entire column to the destination (A, B, C, etc.)

        If Not foundColumn Is Nothing Then

            foundColumn.Copy Destination:=ws.Columns(destColumn)

            destColumn = destColumn + 1 ' Move to the next column (A -> B -> C -> ...)

        End If

    Next i
   

    Set wb = Nothing

    Set ws = Nothing

    Erase searchWords

End Sub



Sub Recount_searchAndRetrieveTiHi()

'

' searchAndRetrieveTiHi Macro

'

    Dim sourceWorkbook As Workbook

    Dim targetWorkbook As Workbook

    Dim targetSheet As Worksheet

    Dim searchValue As String, prevValue As String

    Dim foundCell As Range

    Dim rowNum As Long

    Dim copyDescription As Range

    Dim copyRange As Range

    Dim targetRow As Long

    Dim ws As Worksheet

    Dim sourceFilePath As String, targetFilePath As String, userProfile As String
   

    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")
   

    ' Define the path of the source workbook (external workbook) to lookup description and TixHi

    sourceFilePath = userProfile & "\Downloads\All Item Attributes.xlsx" ' Update the path
   

    ' validate if file exists and if not, exit the Macro

    If Not Recount_fileExists(sourceFilePath) Then

        MsgBox "File not found or needs to be renamed as 'All Item Attributes'"

        Exit Sub

    End If


    ' Open the source workbook

    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
   

    ' Define the path of the target workbook (external workbook) to save description and TixHi

    targetFilePath = userProfile & "\Downloads\Report.xls"
    

    ' open the target workbook and sheet (change as needed)

    Set targetWorkbook = Workbooks.Open(targetFilePath) ' The workbook where the information will go

    Set targetSheet = targetWorkbook.Sheets(1) ' Change to your target sheet name

 
    ' Start pasting data in row 2 of the target sheet

    targetRow = 2
   

    Do While targetSheet.Cells(targetRow, 3).Value <> "" ' Continue until the cell is empty


        searchValue = targetSheet.Cells(targetRow, 3).Value ' assigns value of the next cell
       

        If searchValue = prevValue Then

            rowNum = foundCell.Row

            ' set the range to copy column F for description if row found

            Set copyDescription = ws.Range("F" & rowNum)

            ' Copy the data from columns F for description

            copyDescription.Copy

            ' Paste the copied data into the target sheet (starting at the target row)

            targetSheet.Cells(targetRow, 6).PasteSpecial Paste:=xlPasteValues

            ' Set the range to copy from columns S and T of the found row

            Set copyRange = ws.Range("S" & rowNum & ":T" & rowNum)

            ' Copy the data from columns S and T for TixHi

            copyRange.Copy

            ' Paste the copied data into the target sheet (starting at the target row)

            targetSheet.Cells(targetRow, 7).PasteSpecial Paste:=xlPasteValues

            ' Clear clipboard after copying

            Application.CutCopyMode = False

            ' Move to the next row

            targetRow = targetRow + 1

        Else

            For Each ws In sourceWorkbook.Sheets


                Set foundCell = ws.Cells.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
       

                If Not foundCell Is Nothing Then

                    rowNum = foundCell.Row

                    ' set the range to copy column F for description if row found

                    Set copyDescription = ws.Range("F" & rowNum)

                    ' Copy the data from columns F for description

                    copyDescription.Copy

                    ' Paste the copied data into the target sheet (starting at the target row)

                    targetSheet.Cells(targetRow, 6).PasteSpecial Paste:=xlPasteValues

                    ' Set the range to copy from columns S and T of the found row

                    Set copyRange = ws.Range("S" & rowNum & ":T" & rowNum)

                    ' Copy the data from columns S and T for TixHi

                    copyRange.Copy

                    ' Paste the copied data into the target sheet (starting at the target row)

                    targetSheet.Cells(targetRow, 7).PasteSpecial Paste:=xlPasteValues

                    ' Clear clipboard after copying

                    Application.CutCopyMode = False

                   ' Move to the next row

                    targetRow = targetRow + 1

                    prevValue = searchValue ' compares next search value with previous one

                    Exit For

                End If

            Next ws

         End If

    Loop
   

    ' Close the source workbook without saving

    sourceWorkbook.Close SaveChanges:=False


    Set sourceWorkbook = Nothing

    Set targetWorkbook = Nothing

    Set targetSheet = Nothing

    Set ws = Nothing
   

End Sub
