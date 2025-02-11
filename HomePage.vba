‘ this code is the start of the home page to hold all the vba programs
Sub switchSheet()
'
'   switchSheet Macro
'
    Dim ws As Worksheet, wsMap As Worksheet
    Dim selectedSheet As String
    Dim cb As Object

    ' Get the selected sheet from the ComboBox
    Set cb = Sheets("Home").Shapes("ComboBox").OLEFormat.Object.Object ' Change ComboBox1 to your ComboBox name
    selectedSheet = cb.Value
   
    ' If no sheet is selected, exit the sub
    If selectedSheet <> "Recount" And selectedSheet <> "ReserveAudit" And selectedSheet <> "Top10Audit" Then
        MsgBox "Please select a sheet from the drop-down list."
        Exit Sub
    End If
 

    For Each ws In ThisWorkbook.Sheets
        If ws.Name = selectedSheet Then
            ws.Visible = xlSheetVisible ' Show the selected sheet
            ws.Activate ' Activate the selected sheet
            Exit For ' exit loop when sheet is found
        End If
    Next ws  

    ' loops through all sheets and hides the inactive ones
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> selectedSheet Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' if the selected sheet is ReserveAudit, then also show the Map sheet
    If selectedSheet = "ReserveAudit" Then
        Set wsMap = ThisWorkbook.Sheets("Map")
        wsMap.Visible = xlSheetVisible
    End If

    ' clears object variables from memory
    Set ws = Nothing
    Set wsMap = Nothing
End Sub



Sub returnToHome()
'
'   returnToHome Macro
'
    Dim ws As Worksheet
    Dim keepSheets As Object
    Dim sheetNames As Variant
    Dim i As Long
    Dim activeSheetName As String, compareSheet As String

    ' declares value for the sheet we're looking for (add new ones here as well)
    sheetNames = Array("Recount", "ReserveAudit", "Top10Audit")

    ' assigns that active sheet to a variable
    activeSheetName = ActiveSheet.Name  

    ' searches for the activeSheet and performs a memory clear (Add New Sheets here if there is clearFunction)
    For i = LBound(sheetNames) To UBound(sheetNames)
        If sheetNames(i) = activeSheetName And activeSheetName = "Recount" Then
            Recount_clearFunction
            Exit For ' exit For loop when active sheet is found
        ElseIf sheetNames(i) = activeSheetName And activeSheetName = "Top10Audit" Then
            Top10_clearFunction
            Exit For ' exit For loop when active sheet is found
        ElseIf sheetNames(i) = activeSheetName And activeSheetName = "ReserveAudit" Then
            RA_clearFunction
            Exit For ' exit For loop when active sheet is found
        End If
    Next i 

    ' this part keeps all sheet related to the Macro, any other created sheets will get deleted to avoid clutter

    ' Create a dictionary to store sheets to keep (Add new sheets here when expanding Macro)
    Set keepSheets = CreateObject("Scripting.Dictionary")
    keepSheets.Add "Home", True
    keepSheets.Add "Top10Audit", True
    keepSheets.Add "Recount", True
    keepSheets.Add "ReserveAudit", True
    keepSheets.Add "Map", True 

    ' Loop through all sheets in the workbook
    Application.DisplayAlerts = False ' Disable delete confirmation
    For Each ws In ThisWorkbook.Sheets
        compareSheet = ws.Name
        ' Check if the sheet is NOT in the keepSheets dictionary
        If Not keepSheets.exists(compareSheet) Then
            ws.Visible = xlSheetVisible ' if sheet is hidden, makes it visible
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True ' Re-enable alerts

    ' loop through all sheets to show Home and hide others
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Home" Then
            ws.Visible = xlSheetVisible
            ws.Activate ' return to Home and triggers Worksheet_Activate()
        End If       

        ' hide all sheets except Home
        If ws.Name <> "Home" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    ' clears object variable from memory
    Erase sheetNames

End Sub



Sub showTabs()
'
'   Temporary macro to show all tabs for debugging
'
    For Each ws In ThisWorkbook.Sheets
        ws.Visible = xlSheetVisible
    Next ws

End Sub
