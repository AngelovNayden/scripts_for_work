Attribute VB_Name = "Module1"
Sub SaveEmailsAndAttachments()
    Dim objItem As Object
    Dim objAttachment As Attachment
    Dim saveFolder As String
    Dim filePath As String
    Dim inputValue As String
    Dim processedValue As String
    
    ' Prompt the user to enter a value
    inputValue = InputBox("Enter a path to a folder:", "Input Required")
    
    ' Check if the user entered something
    If inputValue <> "" Then
        ' Store or process the value in another variable
        processedValue = inputValue
        
        ' Ensure the folder path ends with a backslash
        If Right(processedValue, 1) <> "\" Then
            processedValue = processedValue & "\"
        End If
        
        ' Display the processed value (optional: for debugging)
        MsgBox "Path to save: " & processedValue
    Else
        MsgBox "No input provided."
        Exit Sub
    End If
    
    ' Set the folder where you want to save the attachments
    saveFolder = processedValue  ' This is the correct folder path
    
    ' Loop through selected items in the folder
    For Each objItem In Application.ActiveExplorer.Selection
        ' Loop through attachments and save them
        For Each objAttachment In objItem.Attachments
            ' Create the full file path by combining saveFolder and the attachment's file name
            filePath = saveFolder & objAttachment.FileName
            
            ' Save the attachment
            On Error Resume Next  ' In case of errors, skip to the next attachment
            objAttachment.SaveAsFile filePath
            If Err.Number <> 0 Then
                MsgBox "Error saving file: " & filePath & vbCrLf & Err.Description
                Err.Clear
            End If
            On Error GoTo 0  ' Reset error handling
        Next objAttachment
    Next objItem
End Sub

