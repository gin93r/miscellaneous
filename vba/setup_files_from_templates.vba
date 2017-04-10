Private Sub generateFiles()
    'Excel declarations
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim startingCell As Integer
    
    
    'Word declarations
    Dim wdApp As Object
    Dim wdPath As String
    Dim wdDoc As Object
    'Dim sel As word.Selection
    
    'Set the excel workbook and worksheets
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Sheet1")
    
    Dim saveDir As FileDialog
    Dim savePath As String
    
    'Set saveDir = Application.FileDialog(msoFileDialogFolderPicker)
    startingCell = Application.InputBox("Enter a starting row; eg: 2", Type:=1)
        
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select an output folder"
        If .Show <> 0 Then
            savePath = .SelectedItems(1)
            GoTo OpenTemplate
        Else
            MsgBox ("Script Cancelled")
            Exit Sub
        End If
        
    End With
   
OpenTemplate:
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select Word Template"
        .AllowMultiSelect = False
        If .Show <> 0 Then
            'Set wdDoc = wdApp.Documents.Open(Filename:=.SelectedItems(1), ReadOnly:=True)
            wdPath = .SelectedItems(1)
            GoTo CreateFile
        Else
            MsgBox ("Script Cancelled")
            Exit Sub
        End If
    End With
    
CreateFile:
    'Check if word is open
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")

    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    'open the word template file chose previously
    Set wdDoc = wdApp.Documents.Open(Filename:=wdPath, ReadOnly:=True)
    wdApp.Visible = True
    wdApp.Activate
    'make a new directory for the variable
    
    
    Dim ItemNumber, ItemName, strFullName, strFilePath As String
    Dim rngSelection As Range
    
    'loop through the rows in the excel document
    Application.ScreenUpdating = False
    NumRows = Range("A" & startingCell, Range("A" & startingCell).End(xlDown)).Rows.Count
    For i = startingCell To startingCell + NumRows - 1

        If Cells(i, 1) <> "" Then
            ItemNumber = Cells(i, 1)
            ItemName = Cells(i, 2)
            strFullName = ItemNumber & " - " & ItemName
            Set fs = CreateObject("Scripting.FileSystemObject")
            fs.CreateFolder (savePath & "\" & strFullName)
            
            'modify the word document
            With wdApp.Selection.Find
                .Text = "Item Name: "
                .Forward = True
                .MatchWholeWord = True
            End With
            wdApp.Selection.Find.Execute
            wdApp.Selection.InsertAfter (ItemName)
            wdApp.Selection.Collapse Direction:=wdCollapseEnd
            
            
            wdApp.Selection.Find.ClearFormatting
            With wdApp.Selection.Find
                .Text = "Item No.: "
                .Forward = True
                .MatchWholeWord = True
            End With
            wdApp.Selection.Find.Execute
            wdApp.Selection.InsertAfter (ItemNumber)
            
            strFilePath = savePath & "\" & strFullName & "\" & strFullName & ".docx"
            wdDoc.SaveAs (strFilePath)
            wdDoc.Undo 2
        Else
            MsgBox ("Empty Item Number - Exiting")
            Exit Sub
        End If
        
    Next i
    Application.ScreenUpdating = True

    wdDoc.Close (False)
    wdApp.Quit
    MsgBox ("Finished")
End Sub

