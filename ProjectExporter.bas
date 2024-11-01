Option Explicit

'*******************************************************************************
' Module: ProjectExporter
' Purpose: Exports all VBA components from the specified workbook to a directory.
'          Supports automatic sub-folder creation and custom filtering for components.
'*******************************************************************************

'*******************************************************************************
' Function: ExportVbaProject
' Purpose: Exports all VBA components from the specified workbook to a directory.
' Inputs:
'   ProjectPath (String): The path to the directory where the VBA components will be
'                         exported.
'   SourceWorkbook (Workbook): The workbook from which the VBA components will be exported.
'   Overwrite (Boolean): Optional boolean to specify if existing files should be
'                        overwritten (default = False).
'   filters (Dictionary): Optional filters to categorize files into folders based on patterns.
' Outputs:
'   None
'*******************************************************************************
Private Function ExportVbaProject(ProjectPath As String, SourceWorkbook As Workbook, Optional Overwrite As Boolean = False, Optional filters As Dictionary = Nothing) As Boolean
    On Error GoTo HandleGeneralError

    Dim objVbComp As VBComponent
    Dim exportDir As String
    Dim exportPath As String
    Dim folderName As String
    
    ' Remove any trailing path separator from ProjectPath if present
    If Right(ProjectPath, 1) = Application.PathSeparator Then
        ProjectPath = Left(ProjectPath, Len(ProjectPath) - 1)
    End If
    
    ' Construct the export directory using the workbook name
    exportDir = ProjectPath & Application.PathSeparator & Left(SourceWorkbook.name, InStrRev(SourceWorkbook.name, ".") - 1) & " VBA Project" & Application.PathSeparator
    
    ' Create the directory if it doesn't exist
    If Dir(exportDir, vbDirectory) = "" Then
        MkDir exportDir
    End If

    ' Loop through all VBA components in the specified workbook
    For Each objVbComp In SourceWorkbook.VBProject.VBComponents
        
        ' Determine the folder name based on the component type or custom filters
        folderName = GetFolderNameForComponent(objVbComp, filters)
        exportPath = exportDir & folderName & Application.PathSeparator
        
        ' Create the sub-folder if it doesn't exist
        If Dir(exportPath, vbDirectory) = "" Then
            MkDir exportPath
        End If
        
        ' Ensure the component name is valid for a file path
        exportPath = exportPath & SanitizeFilename(objVbComp.name)
        
        ' Determine the type of component and set the appropriate extension
        Select Case objVbComp.Type
            Case vbext_ct_StdModule
                exportPath = exportPath & ".bas"
            Case vbext_ct_Document, vbext_ct_ClassModule
                exportPath = exportPath & ".cls"
            Case vbext_ct_MSForm
                exportPath = exportPath & ".frm"
            Case Else
                ' For unknown types, log and continue
                exportPath = exportPath & ".txt"
                Debug.Print "Unknown component type for: " & objVbComp.name
        End Select

        ' Check if the file exists and either overwrite or skip based on the Overwrite parameter
        If Dir(exportPath) = "" Or Overwrite Then
            On Error GoTo HandleExportError
            objVbComp.Export exportPath
            On Error GoTo 0
        Else
            Debug.Print "File " & exportPath & " already exists and was skipped."
        End If
    Next objVbComp
    
    ExportVbaProject = True
    Exit Function

HandleGeneralError:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    ExportVbaProject = False
    Exit Function

HandleExportError:
    MsgBox "Error exporting component: " & objVbComp.name & vbCrLf & "Error: " & Err.Description, vbCritical
    Resume Next
End Function

'*******************************************************************************
' Function: GetFolderNameForComponent
' Purpose: Determines the folder name for a VBA component based on its type or custom filters.
' Inputs:
'   objVbComp (VBComponent): The VBA component.
'   filters (Dictionary): Optional filters to categorize components into folders.
' Outputs:
'   String: The folder name where the component should be stored.
'*******************************************************************************
Private Function GetFolderNameForComponent(objVbComp As VBComponent, filters As Dictionary) As String
    Dim pattern As Variant
    Dim folderName As String
    
    ' Check custom filters first
    If Not filters Is Nothing Then
        For Each pattern In filters.Keys
            ' Filters is a dictionary with a pattern as key and folder name as value
            If objVbComp.name Like pattern Then
                GetFolderNameForComponent = filters(pattern)
                Exit Function
            End If
        Next pattern
    End If
    
    ' Default folder names based on component type
    Select Case objVbComp.Type
        Case vbext_ct_ClassModule
            folderName = "Class"
        Case vbext_ct_MSForm
            folderName = "Forms"
        Case vbext_ct_Document
            folderName = "Local"
        Case vbext_ct_StdModule
            folderName = "Modules"
        Case Else
            folderName = "Other"
    End Select
    
    GetFolderNameForComponent = folderName
End Function

'*******************************************************************************
' Function: SanitizeFilename
' Purpose: Ensures that a filename is valid by removing invalid characters.
' Inputs:
'   str (String): The original string to be sanitized for use as a filename.
' Outputs:
'   String: The sanitized string with invalid filename characters removed.
'*******************************************************************************
Private Function SanitizeFilename(str As String) As String
    Dim invalidChars As String
    Dim i As Integer
    invalidChars = "/\:*?""<>|"
    For i = 1 To Len(invalidChars)
        str = Replace(str, Mid(invalidChars, i, 1), "_")
    Next i
    SanitizeFilename = str
End Function

'*******************************************************************************
' Test function to export the VBA project of the active workbook
' This is a simple example to demonstrate the usage of the ExportVbaProject function.
' It adds custom filters to group components based on name patterns.
'*******************************************************************************
Private Sub ExportThisWorkbook()
    Dim filters As Scripting.Dictionary
    Set filters = CreateObject("Scripting.Dictionary")
    
    ' Example: Adding custom filters for specific components
    filters.Add "*Helpers", "Helpers"
    filters.Add "Globals*", "Globals"
    
    If Not ExportVbaProject(ThisWorkbook.Path, ThisWorkbook, True, filters) Then
        MsgBox "Export failed.", vbCritical
    Else
        MsgBox "Export successful.", vbInformation
    End If
End Sub
