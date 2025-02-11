Sub RA_clearFunction()

'

'   RA_clearFunction

'

    ' clears all information

    Cells.Clear

    ThisWorkbook.Sheets("Map").Range("A2").Value = ""                                                                                                                                                                                                                                           

End Sub



Sub RA_copyAllCellsBetweenSheets()

'

'   copyAllCellsBetweenSheets

'

    Dim sourceWorkbook As Workbook, targetWorkbook As Workbook

    Dim sourceSheet As Worksheet, targetSheet As Worksheet, targetMap As Worksheet

    Dim sourceFile As String, userProfile As String

    Dim headers As Variant, dataArray As Variant

    Dim foundCell As Range, searchRange As Range

    Dim lastRow As Long, i As Integer

    Dim targetRow As Long, targetCol As Integer



    ' Get the user profile directory (works on Windows systems)

    userProfile = Environ("USERPROFILE")


    ' Set the source file path (closed workbook)

    sourceFile = userProfile & "\Downloads\Report.xls"   


    ' validate if file exists and if not, exit the Macro

    If Not RA_fileExists(sourceFile) Then

        MsgBox "File not found or needs to be renamed as 'Report'"

        Exit Sub

    End If
   

    ' Open the source workbook in the background (without showing it)

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    Application.EnableEvents = False

    Application.DisplayStatusBar = False
   

    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=False) ' Use the sourceFile path to open workbook

    Set sourceSheet = sourceWorkbook.Sheets("InstantReport") ' Replace with your source sheet name
   

    ' Set the source and target workbooks and worksheets

    Set targetWorkbook = ThisWorkbook ' Use the current workbook for the target

    Set targetSheet = targetWorkbook.Sheets("ReserveAudit") ' Replace with your target sheet name


    ' Define headers to search for

    headers = Array("Location", "Location Class", "Dedication Type", "Dedicated Item", "Putaway Zone", "Pull Zone", "Pack Zone", "Maximum Quantity", "Quantity UOM", "Business Unit", "Current Quantity", "Putaway Lock", "Auto Inventory Lock")
   

    targetRow = 3 ' Start pasting from row 3 in the target sheet

    targetCol = 1 ' Start pasting in column A

 
    ' define the search range (all used cells)

    Set searchRange = sourceSheet.UsedRange

   
    For i = LBound(headers) To UBound(headers)

        ' Find the header anywhere in the source sheet(UsedRange)

        Set foundCell = searchRange.Find(What:=headers(i), LookAt:=xlWhole, MatchCase:=False)
       

        ' If found, copy the column data including the header

        If Not foundCell Is Nothing Then

            ' Determine the last row in the column

            lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, foundCell.Column).End(xlUp).Row
           

            ' Store the data in an array (including header)

            dataArray = sourceSheet.Range(foundCell, sourceSheet.Cells(lastRow, foundCell.Column)).Value
           

            ' Paste the array into the target sheet

            targetSheet.Cells(targetRow, targetCol).Resize(UBound(dataArray, 1), 1).Value = dataArray
           

            ' Move to the next column in the target sheet

            targetCol = targetCol + 1

        End If

    Next i
  

    ' Close the source workbook without saving

    sourceWorkbook.Close SaveChanges:=False
 

    Dim inputString As String

    Dim extractedString As String


    inputString = targetSheet.Range("A4")


    ' Extract the first 4 characters

    extractedString = Left(inputString, 4)


    Set targetMap = targetWorkbook.Sheets("Map")

    targetMap.Activate ' puts user on the sheet Map
   

    ' Place the extracted string into a specific cell (e.g., A1 on the active sheet)

    targetMap.Range("A2").Value = extractedString
 

    ' check if the file exists before trying to delete it

    If Dir(sourceFile) <> "" Then

        ' delete the file

        Kill sourceFile

    End If


    ' Re-enable screen updating and alerts

    Application.ScreenUpdating = True

    Application.DisplayAlerts = True

    Application.EnableEvents = True

    Application.DisplayStatusBar = True
   

    ' AUTO PRINT CHECK BOX

    If targetSheet.Range("P11").Value = True Then

        ' print executes here

        targetMap.PrintOut

    End If
   

    Set sourceWorkbook = Nothing

    Set targetWorkbook = Nothing

    Set sourceSheet = Nothing

    Set targetSheet = Nothing

    Set targetMap = Nothing

    Set foundCell = Nothing

    Erase dataArray

End Sub



Function RA_fileExists(filePath As String) As Boolean

'

'   fileExists function

'

    If Dir(filePath) <> "" Then

        RA_fileExists = True

    Else

        RA_fileExists = False

    End If

End Function
