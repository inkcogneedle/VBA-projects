Sub Top10_clearFunction()

'

' clearFunction Macro

'

    Dim ws As Worksheet

    Dim findWords As Variant, word As Variant

    Dim startPos As Long

    Dim cell As Range, rng As Range, cTCINs As Range
   

    ' clears all information

    Cells.Clear
   

    ' formats cells A16 to A25 back to normal

    With Range("A15:A25")

        .Font.Name = "Calibri"

        .Font.Size = 11

    End With
   

    Range("A16:A25").Style = "Output" ' cell style
   

    ' formats cells A2 to I11 back to normal

    With Range("A2:I11")

        .Font.Name = "Calibri"

        .Font.Size = 11

    End With


    ' default message for cell A1

    Range("A1").Value = "Before using the Top 10 items button, download these 2 from the SCI landing page: DPCIs in Demand for Order Pool and All Items and Attributes"

   
    Range("A15").Value = "Enter 10 TCINs below:" ' sets value back incase it's deleted

    Range("A15").Style = "Heading 3" ' cell style
   

    Range("A27").Value = "combined TCINs ->" ' sets value back incase it's deleted

    Range("A27").Style = "Explanatory Text" ' cell style

    Range("A27").HorizontalAlignment = xlCenter ' center alignment for cell
   

    Set cTCINs = Range("B27:K27") ' set the range

    ' apply the "Gold, Accent 4, Lighter 40%" fill color

    cTCINs.Interior.ThemeColor = xlThemeColorAccent4

    cTCINs.Interior.TintAndShade = 0.4
   

    Range("A29").Value = "Copy and paste the combined TCINs into the Item field, set the Location class to 'Active', and check the Display zero quantity Locations box. Go to 'Tools' and pull 'Report', and then hit the Run Button"
   

    Set ws = ThisWorkbook.Sheets("Top10Audit") ' change to your sheet name

    Set rng = ws.Range("A1") ' change to the range you want searched
   

    ' array of words to search and bold for cells A1 and A2

    findWords = Array("DPCIs in Demand for Order Pool", "All Items and Attributes")
   

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


    Set rng = ws.Range("A29") ' change to the range you want searched

 
    ' array of words to search and bold for cells A1 and A2

    findWords = Array("Item", "Location class", "Display zero quantity Locations")
   

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
   

    ' checks to see if sheet exists and if so, makes sheet active and prompts for delete

    If Top10_sheetExists("Top10Audit-Result") Then

        ThisWorkbook.Sheets("Top10Audit-Result").Activate

        Top10_closeSheet

    End If
   

    ' clears object variables from memory

    Set ws = Nothing

    Set findWords = Nothing

End Sub



Sub Top10_closeSheet()

'

' closeSheet Macro

'

    On Error Resume Next ' incase the sheet isn't there

    Sheets("Top10Audit-result").Delete

    On Error GoTo 0 ' reset error handling

End Sub



Sub Top10_concatenateWithArray()

'

'  concatenateWithArray Macro

'

    Dim myArray() As Variant

    Dim rng As Range

    Dim arrayResult As String


    'initialize array with values for (A2:A11)

    Set rng = ThisWorkbook.Sheets("Top10Audit").Range("A16:A25")


    'Assign the values of the range to the array

    myArray = rng.Value
   

    Dim i As Integer

    For i = 1 To UBound(myArray, 1) 'loop through the rows in the immediate window

        If i = LBound(myArray) Then

            arrayResult = myArray(i, 1) 'don't add a comma before the first element

        Else

            arrayResult = arrayResult & ", " & myArray(i, 1) 'add a comma between elements

        End If

    Next i
   

    Range("B27").Value = arrayResult 'concatenated values set to cell
   

    ' Clears the array from memory

    Erase myArray

End Sub



Function Top10_fileExists(filePath As String) As Boolean

'

'   fileExists function

'

    If Dir(filePath) <> "" Then

        Top10_fileExists = True

    Else

        Top10_fileExists = False

    End If

End Function



Sub Top10_pullDataFromReport()

'

'   pullDataFromReport

'

    Dim sourceWorkbook As Workbook, targetWorkbook As Workbook

    Dim sourceSheet As Worksheet, targetSheet As Worksheet

    Dim sourceRange As Range

    Dim sourceFile As String, userProfile As String

    Dim lastRow As Long, lastCol As Long

   
    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")


    ' Set the source file path (closed workbook)

    sourceFile = userProfile & "\Downloads\Report.xls"
   

    ' validate if file exists and if not, exit the Macro

    If Not Top10_fileExists(sourceFile) Then

        MsgBox "File not found or needs to be renamed as 'Report'"

        Exit Sub

    End If
   

    ' Set the target workbook (the workbook from which you are running this macro)

    Set targetWorkbook = ThisWorkbook
   

    ' checks to see if the sheet exist, if so, prompts user to close sheet and ends function

    If Top10_sheetExists("Top10Audit-Result") Then

        MsgBox "Please close the sheet 'Top10Audit-Result' before running Macro"

        Exit Sub

    End If
  

    ' Create a new sheet in the target workbook to paste the data

    Set targetSheet = targetWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Top10Audit"))

    targetSheet.Name = "Top10Audit-Result" ' set name of new sheet

    setupPrintMargins targetSheet


    ' Open the source workbook in the background (without showing it)

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Application.Calculation = xlCalculationManual

    Application.EnableEvents = False

    Application.DisplayStatusBar = False


    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=False)
  

    ' Set the source sheet (adjust as necessary)

    Set sourceSheet = sourceWorkbook.Sheets(1) ' Change index or name if needed
   

    ' Delete rows 1 and 2 from source sheet

    sourceSheet.Rows("1:2").Delete
   

    ' Find the last row and column with data in the source sheet

    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row

    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
   

    ' Set the source range (all the data from the source sheet)

    Set sourceRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastCol))


    Top10_sortAndFormat

   
    ' Copy the content from the source range

    sourceRange.Copy


    ' Paste the content into the target sheet starting from A1

    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteValues

    targetSheet.Range("A1").PasteSpecial Paste:=xlPasteFormats ' Optional: to retain formatting
   

    ' creates ti and hi header with format

    targetSheet.Range("E1").Value = "Ti"

    targetSheet.Range("F1").Value = "Hi"

    targetSheet.Columns("E:F").HorizontalAlignment = xlCenter

    targetSheet.Range("E1:F1").Font.Bold = True


    ' sets button on new sheet to allow user to close

    Set Button = targetSheet.Buttons.Add(Top:=100, Left:=400, Width:=80, Height:=30)

    Button.Caption = "Close Sheet"

    Button.OnAction = "Top10_closeSheet"


    ' Close the source workbook without saving

    sourceWorkbook.Close SaveChanges:=False
   

    Top10_searchAndRetrieveTiHi ' opens up All Items Attribute and searches for Ti and Hi


    ' Define the path of the source workbook (external workbook)

    Dim sourceFilePath As String

    sourceFilePath = userProfile & "\Downloads\All Item Attributes.xlsx" ' Update the path


    ' this stops the macro from proceeding further since the function Top10_searchAndRetrieveTiHi could not find the file 'All Items Attributes'

    If Not Top10_fileExists(sourceFilePath) Then

        MsgBox "Macro is closing. Please download the excel file 'All Item Attributes' from the SCI landing page and try again."

        Exit Sub

    End If
 

    targetSheet.Cells.EntireColumn.AutoFit ' widen all cells to fit content on the sheet


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


    ' clears object variables from memory

    Set sourceWorkbook = Nothing

    Set targetWorkbook = Nothing

    Set sourceSheet = Nothing

    Set targetSheet = Nothing
 
End Sub



Sub Top10_pullTop10()

'

'   pullTop10 Macro

'

    Dim sourceWorkbook As Workbook, targetWorkbook As Workbook

    Dim sourceSheet As Worksheet, targetSheet As Worksheet

    Dim sourceRange As Range, foundCell As Range

    Dim sourceFile As String, userProfile As String

    Dim lastRow As Long


    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")
   

    ' Set the source file path (closed workbook)

    sourceFile = userProfile & "\Downloads\DPCIs in Demand for Drop.xlsx"


    ' validate if file exists and if not, exit the Macro

    If Not Top10_fileExists(sourceFile) Then

        MsgBox "File not found or needs to be named exactly as 'DPCIs in Demand for Drop'"

        Exit Sub

    End If
   

    ' Set the target workbook (the workbook from which you are running this macro)

    Set targetWorkbook = ThisWorkbook


    ' Open the source workbook in the background (without showing it)

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Application.Calculation = xlCalculationManual

    Application.EnableEvents = False

    Application.DisplayStatusBar = False
   

    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=False)


    ' Set the source sheet (adjust as necessary)

    Set sourceSheet = sourceWorkbook.Sheets(1) ' Change index or name if needed
 

    ' unmerge cells

    sourceSheet.Cells.UnMerge
   

    ' find the last row with data in column A (change to another column if needed)

    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
   

    ' check if there's data to delete

    If lastRow > 1 Then ' ensure at least one row remains

        sourceSheet.Rows(lastRow).Delete ' deletes the total count row in DPCIs in Demand for Drop

    Else

        MsgBox "There is no data to delete! Is this the right file?", vbExclamation

        Exit Sub

    End If


    ' define the search range (searches the entire sheet)

    Set searchRange = sourceSheet.Range("A1:A" & lastRow) ' Adjust column A if needed


    ' Find the cell containing "TCIN"

    Set foundCell = searchRange.Find(What:="TCIN", LookAt:=xlWhole, MatchCase:=False)


    ' Check if "TCIN" was found

    If Not foundCell Is Nothing Then

        ' Delete all rows above the found cell

        sourceSheet.Rows("1:" & foundCell.Row - 1).Delete

    Else

        MsgBox "Macro needs to be updated", vbExclamation
       

        ' Close the source workbook without saving

        sourceWorkbook.Close SaveChanges:=False


        ' Re-enable screen updating and alerts

        Application.ScreenUpdating = True

        Application.DisplayAlerts = True

        Application.Calculation = xlCalculationAutomatic

        Application.EnableEvents = True

        Application.DisplayStatusBar = True

        Exit Sub

    End If
 

    ' Delete columns D to E and I

    sourceSheet.Columns("D:E").Delete

    sourceSheet.Columns("G:G").Delete
   

    ' sorts the cancelled cases by descending order

    With Range("A1:I" & Cells(Rows.Count, "A").End(xlUp).Row)

        .Sort key1:=Range("F2"), Order1:=xlDescending, Header:=xlYes

    End With
 

    sourceSheet.Range("A1:I11").Copy ' sets the range to copy
  

    Set targetSheet = targetWorkbook.Sheets("Top10Audit")


    ' Paste the content into the target sheet starting from A1

    targetSheet.Range("A2").PasteSpecial Paste:=xlPasteValues ' sets the cell to begin the paste area

    targetSheet.Range("A2").PasteSpecial Paste:=xlPasteFormats ' Optional: to retain formatting

    targetSheet.Range("A3:I12").HorizontalAlignment = xlCenter ' sets alignment to center
  

    ' Close the source workbook without saving

    sourceWorkbook.Close SaveChanges:=False


    ' Re-enable screen updating and alerts

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

    Application.Calculation = xlCalculationAutomatic

    Application.EnableEvents = True

    Application.DisplayStatusBar = True
   

    Dim response As Integer
   

    ' check if the file exists before trying to delete it

    If Dir(sourceFile) <> "" Then

        response = MsgBox("Do you want to delete the excel file DPCIs in Demand for Drop?", vbYesNo, "Now that you've pulled the top 10")

            If response = vbYes Then

                ' delete the file

                Kill sourceFile

            End If

    End If
  

    ' display a Yes/No message box

    response = MsgBox("Do you want to add the TCINs to cells A16 - A25?", vbYesNo, "Confirmation")


    ' check if the user clicked "yes"

    If response = vbYes Then

        targetSheet.Range("A3:A12").Copy

        targetSheet.Range("A16:A25").PasteSpecial Paste:=xlPasteValues

        targetSheet.Range("A16:A25").HorizontalAlignment = xlCenter

        Application.CutCopyMode = False
       

    End If
  

    ' clears object variables from memory

    Set sourceWorkbook = Nothing

    Set targetWorkbook = Nothing

    Set sourceSheet = Nothing

    Set targetSheet = Nothing

End Sub



Sub Top10_searchAndRetrieveTiHi()

'

' searchAndRetrieveTiHi Macro

'

    Dim sourceWorkbook As Workbook, targetWorkbook As Workbook

    Dim targetSheet As Worksheet

    Dim searchValue As String, prevValue As String

    Dim foundCell As Range

    Dim rowNum As Long

    Dim copyRange As Range

    Dim targetRow As Long

    Dim ws As Worksheet

    Dim userProfile As String


    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")

 
    ' Define the path of the source workbook (external workbook)

    Dim sourceFilePath As String

    sourceFilePath = userProfile & "\Downloads\All Item Attributes.xlsx" ' Update the path


    ' validate if file exists and if not, exit the Macro

    If Not Top10_fileExists(sourceFilePath) Then

        MsgBox "File not found or needs to be named exactly as 'All Item Attributes'"

        Top10_closeSheet

        Exit Sub

    End If


    ' Open the source workbook

    Set sourceWorkbook = Workbooks.Open(sourceFilePath)
   

    ' Define the target workbook and sheet (change as needed)

    Set targetWorkbook = ThisWorkbook ' The workbook where this macro is running

    Set targetSheet = targetWorkbook.Sheets("Top10Audit-Result") ' Change to your target sheet name


    ' Start pasting data in row 2 of the target sheet

    targetRow = 2
  

    Do While targetSheet.Cells(targetRow, 2).Value <> "" ' Continue until the cell is empty


        searchValue = targetSheet.Cells(targetRow, 2).Value

 
        If searchValue = prevValue Then

            rowNum = foundCell.Row

            ' Set the range to copy from columns S and T of the found row

            Set copyRange = ws.Range("S" & rowNum & ":T" & rowNum)

            ' Copy the data from columns S and T

            copyRange.Copy

            ' Paste the copied data into the target sheet (starting at the target row)

            targetSheet.Cells(targetRow, 5).PasteSpecial Paste:=xlPasteValues

            ' Clear clipboard after copying

            Application.CutCopyMode = False

            ' Move to the next row

            targetRow = targetRow + 1

        Else

            For Each ws In sourceWorkbook.Sheets


                Set foundCell = ws.Cells.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
      

                If Not foundCell Is Nothing Then

                    rowNum = foundCell.Row

                    ' Set the range to copy from columns S and T of the found row

                    Set copyRange = ws.Range("S" & rowNum & ":T" & rowNum)

                    ' Copy the data from columns S and T

                    copyRange.Copy

                    ' Paste the copied data into the target sheet (starting at the target row)

                    targetSheet.Cells(targetRow, 5).PasteSpecial Paste:=xlPasteValues

                    ' Clear clipboard after copying

                    Application.CutCopyMode = False

                    targetRow = targetRow + 1 ' move to the next row

                    prevValue = searchValue ' compares next search value with previous one

                    Exit For

                End If

            Next ws

         End If

    Loop


    ' Close the source workbook without saving

    sourceWorkbook.Close SaveChanges:=False
                                                                                                                                                                                                                 

    ' clears object variables from memory

    Set sourceWorkbook = Nothing

    Set targetWorkbook = Nothing

    Set targetSheet = Nothing

    Set ws = Nothing

End Sub



Function Top10_sheetExists(sheetName As String) As Boolean

    '

    ' sheetExist function

    '

    On Error Resume Next

    Top10_sheetExists = Not Worksheets(sheetName) Is Nothing

    On Error GoTo 0

End Function



Sub Top10_sortAndFormat()

'

' sortAndFormat Macro

'

    ' Filter and Sorts the report of the top 10 items

    Columns("A:A").Delete

    Columns("E:I").Delete   


    With Range("A1:I" & Cells(Rows.Count, "A").End(xlUp).Row)

        .Sort key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes

    End With
  
End Sub
