Option Explicit
'*******************************************************************************
' Module: ProjectExporter
' Purpose: Exports all VBA components from the active workbook to a specified
'          directory.
' Author:  Giovanni Di Toro
' Date:    28-Feb-2025
' Version: 1.1
'*******************************************************************************

'*******************************************************************************
' Function: GetExportFolder
' Purpose: Prompts the user to confirm or select an export folder
' Inputs:
'   DefaultPath: The default export path to suggest
' Outputs:
'   String: The confirmed or selected export folder path, or empty string if canceled
'*******************************************************************************
Private Function GetExportFolder(DefaultPath As String) As String
    Dim folderPath As String
    Dim userChoice As VbMsgBoxResult
    Dim fDialog As Object
    
    ' Extract directory from the default path
    folderPath = Left(DefaultPath, InStrRev(DefaultPath, Application.PathSeparator))
    
    ' Suggest the default export location
    userChoice = MsgBox("Export to:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
                        "Click Yes to confirm this location" & vbCrLf & _
                        "Click No to choose another location" & vbCrLf & _
                        "Click Cancel to abort export", _
                        vbQuestion + vbYesNoCancel, "Confirm Export Location")
    
    Select Case userChoice
        Case vbYes
            ' User confirmed the default location
            GetExportFolder = folderPath
            
        Case vbNo
            ' User wants to choose another location
            On Error Resume Next
            ' FIX: Use proper error handling for folder dialog
            Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
            If Not fDialog Is Nothing Then
                With fDialog
                    .Title = "Select Export Folder"
                    .AllowMultiSelect = False
                    .InitialFileName = folderPath
                    
                    If .Show Then
                        ' Fix the crash issue by checking if .SelectedItems exists and has items
                        If .SelectedItems.Count > 0 Then
                            folderPath = .SelectedItems(1)
                            ' Ensure path ends with path separator
                            If Right(folderPath, 1) <> Application.PathSeparator Then
                                folderPath = folderPath & Application.PathSeparator
                            End If
                            GetExportFolder = folderPath
                        Else
                            GetExportFolder = ""
                        End If
                    Else
                        ' User cancelled folder dialog
                        GetExportFolder = ""
                    End If
                End With
            Else
                ' Fallback to default if dialog fails
                MsgBox "Could not open folder selection dialog. Using default folder.", vbExclamation
                GetExportFolder = folderPath
            End If
            On Error GoTo 0
            
        Case vbCancel
            ' User cancelled the export
            GetExportFolder = ""
            
        Case Else
            ' Handle unexpected result
            GetExportFolder = ""
    End Select
End Function

'*******************************************************************************
' Function: CreateDirectoryIfNotExists
' Purpose: Creates a directory if it doesn't exist
' Inputs:
'   DirPath: Directory path to create
' Returns:
'   Boolean: True if directory exists or was created successfully
'*******************************************************************************
Private Function CreateDirectoryIfNotExists(DirPath As String) As Boolean
    Dim fso As Object
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(DirPath) Then
        fso.CreateFolder DirPath
    End If
    
    CreateDirectoryIfNotExists = (Err.Number = 0)
    On Error GoTo 0
    
    Set fso = Nothing
End Function

'*******************************************************************************
' Function: SanitizeFileName
' Purpose: Removes invalid characters from file names
' Inputs:
'   FileName: String containing the file name to sanitize
' Returns:
'   String: Sanitized file name
'*******************************************************************************
Private Function SanitizeFileName(FileName As String) As String
    Dim result As String
    Dim invalidChars As String
    Dim i As Integer
    
    ' Define invalid filename characters
    invalidChars = "\/:*?""<>|"
    
    result = FileName
    
    ' Replace invalid characters with underscore
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    SanitizeFileName = result
End Function

'*******************************************************************************
' Function: ExportVbaProject
' Purpose: Exports all VBA components from the specified workbook to a directory.
' Inputs:
'   ProjectPath: The path to the directory where the VBA components will be
'                exported.
'   SourceWorkbook: The workbook from which the VBA components will be exported.
'   Overwrite: Optional boolean to specify if existing files should be
'              overwritten (default = False).
' Outputs:
'   Boolean: True if export was successful, False otherwise
'*******************************************************************************
Public Function ExportVbaProject(ProjectPath As String, SourceWorkbook As Workbook, Optional Overwrite As Boolean = False) As Boolean
    Dim objVbComp As VBComponent
    Dim exportDir As String
    Dim filename As String
    Dim WbDir As String
    Dim ExportPath As String
    Dim selectedFolder As String
    Dim exportCount As Long
    Dim skipCount As Long
    Dim totalCount As Long
    
    ' Initialize counters
    exportCount = 0
    skipCount = 0
    
    ' Get the base filename from the provided path using native VBA
    filename = Dir(ProjectPath, vbNormal)

    ' Remove extension from the filename
    If InStr(filename, ".") > 0 Then
        filename = Left(filename, InStrRev(filename, ".") - 1)
    End If

    ' Get the directory of the workbook
    WbDir = Left(SourceWorkbook.FullName, InStrRev(SourceWorkbook.FullName, Application.PathSeparator))

    ' Get export folder confirmation from user
    selectedFolder = GetExportFolder(WbDir)
    
    ' Check if user cancelled the export
    If selectedFolder = "" Then
        MsgBox "Export cancelled by user.", vbInformation
        ExportVbaProject = False
        Exit Function
    End If
    
    ' Construct the export directory path
    exportDir = selectedFolder & SanitizeFileName(filename) & " VBA Project" & Application.PathSeparator

    ' Create the directory if it doesn't exist
    If Not CreateDirectoryIfNotExists(exportDir) Then
        MsgBox "Error creating directory: " & exportDir, vbCritical
        ExportVbaProject = False
        Exit Function
    End If
    
    ' Count total components
    totalCount = SourceWorkbook.VBProject.VBComponents.Count
    
    ' Update status bar
    Application.StatusBar = "Exporting VBA project components..."
    
    ' Loop through all VBA components in the specified workbook
    For Each objVbComp In SourceWorkbook.VBProject.VBComponents
        ExportPath = exportDir & SanitizeFileName(objVbComp.Name)

        ' Determine the type of component and set the appropriate extension
        Select Case objVbComp.Type
            Case vbext_ct_StdModule
                ExportPath = ExportPath & ".bas"
            Case vbext_ct_ClassModule
                ExportPath = ExportPath & ".cls"
            Case vbext_ct_MSForm
                ExportPath = ExportPath & ".frm"
            Case vbext_ct_Document
                ExportPath = ExportPath & ".cls"
            Case Else
                ' For unknown types, use .txt as default
                ExportPath = ExportPath & ".txt"
        End Select

        ' Update status bar with current component
        Application.StatusBar = "Exporting: " & objVbComp.Name & " (" & (exportCount + 1) & " of " & totalCount & ")"

        ' Check if file exists and either overwrite or skip based on the Overwrite parameter
        If Dir(ExportPath) = "" Or Overwrite Then
            On Error Resume Next
            objVbComp.Export ExportPath
            
            If Err.Number = 0 Then
                exportCount = exportCount + 1
            Else
                Debug.Print "Error exporting component: " & objVbComp.Name & " - " & Err.Description
            End If
            On Error GoTo 0
        Else
            ' Log or notify that the file was skipped
            skipCount = skipCount + 1
            Debug.Print "File " & ExportPath & " already exists and was skipped."
        End If
    Next objVbComp

    ' Clear status bar
    Application.StatusBar = False
    
    ' Show export summary
    MsgBox "Export Summary:" & vbCrLf & _
           "Components Exported: " & exportCount & vbCrLf & _
           "Components Skipped: " & skipCount & vbCrLf & vbCrLf & _
           "Export Location: " & exportDir, _
           vbInformation, "Export Complete"
    
    ExportVbaProject = True
End Function

'*******************************************************************************
' Sub: RunExporter
' Purpose: Test function to export the VBA project of the active workbook
' The path will be the same as the active workbook with the VBA components exported to a subdirectory
' named after the workbook.
'*******************************************************************************
Public Sub RunExporter()
    Dim ExportPath As String
    Dim success As Boolean
    
    ' Check if workbook is saved
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook before exporting the VBA project.", vbExclamation
        Exit Sub
    End If
    
    ' This workbook path with the filename
    ExportPath = ThisWorkbook.FullName
    
    ' Export the VBA project of this workbook to the specified path
    success = ExportVbaProject(ExportPath, ThisWorkbook, True)
    
    ' Success message is shown in ExportVbaProject function
End Sub

'*******************************************************************************
' Sub: ExportSelectedComponents
' Purpose: Exports only the selected components from the VBA Project
'*******************************************************************************
Public Sub ExportSelectedComponents()
    ' Not implemented yet - future enhancement
    MsgBox "This feature is not implemented yet.", vbInformation
End Sub
