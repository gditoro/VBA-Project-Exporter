Attribute VB_Name = "ProjectExporter"
'@Folder("Exporter")
Option Explicit

'*******************************************************************************
' Name: ProjectExporter
' Kind: Module
' Author: Giovanni Di Toro
' Date: 02/10/2024
' Purpose: Exports all VBA components from the specified workbook to a directory.
'          Supports automatic sub-folder creation and custom filtering for components.
'          Enhanced with ValidationHelpers for security and robust validation.
' Dependencies: Microsoft Visual Basic for Applications Extensibility 5.3
'               Microsoft Scripting Runtime (for Dictionary and FileSystemObject)
'               VBScript.RegExp (for regex pattern matching)
'               ValidationHelpers.bas (for enhanced security validation)
'               ErrorHandler.bas (for logging functions)
'*******************************************************************************

' Add these constants at the top of the module
Private Const ERR_INVALID_PATH As Long = vbObjectError + 513
Private Const ERR_INVALID_WORKBOOK As Long = vbObjectError + 514
Private Const ERR_EXPORT_FAILED As Long = vbObjectError + 515
Private Const MAX_PATH_LENGTH As Long = 260

' Add these folder-related constants
Private Const FOLDER_PATTERN As String = "@Folder"
Private Const ROOT_FOLDER As String = "Root"
Private Const DEFAULT_FOLDER As String = ""

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
Private Function ExportVbaProject(ProjectPath As String, SourceWorkbook As Workbook,_
                                  Optional Overwrite As Boolean = False,_
                                  Optional filters As Dictionary = Nothing,_
                                  Optional useFolderComment As Boolean = True,_
                                  Optional ignoreRootComment As Boolean = True) As Boolean
    On Error GoTo HandleGeneralError

    ' Enhanced input validation using ValidationHelpers
    Dim pathValidation As ValidationResult
    Set pathValidation = ValidationHelpers.ValidateStringLength("Project Path", ProjectPath, 1, MAX_PATH_LENGTH)
    If Not pathValidation.IsValid Then
        Err.Raise ERR_INVALID_PATH, "ExportVbaProject", "Project path validation failed: " & pathValidation.GetErrorsAsString()
    End If

    ' Additional security validation for path
    If Not ValidationHelpers.ValidateFilePath(ProjectPath) Then
        Err.Raise ERR_INVALID_PATH, "ExportVbaProject", "Project path contains security issues or invalid characters"
    End If

    If SourceWorkbook Is Nothing Then
        Err.Raise ERR_INVALID_WORKBOOK, "ExportVbaProject", "Source workbook cannot be Nothing"
    End If

    ' Validate workbook name
    Dim workbookValidation As ValidationResult
    Set workbookValidation = ValidationHelpers.ValidateRequired("Workbook Name", SourceWorkbook.Name)
    If Not workbookValidation.IsValid Then
        Err.Raise ERR_INVALID_WORKBOOK, "ExportVbaProject", "Workbook name validation failed: " & workbookValidation.GetErrorsAsString()
    End If

    If Not SourceWorkbook.VBProject.Protection = vbext_pp_none Then
        Err.Raise ERR_EXPORT_FAILED, "ExportVbaProject", "VBA Project is protected"
    End If

    ' Create backup of existing files if overwriting
    If Overwrite Then
        CreateBackup ProjectPath, SourceWorkbook.name
    End If

    Dim objVbComp As VBComponent
    Dim exportDir As String
    Dim exportPath As String
    Dim folderName As String

    ' Remove any trailing path separator from ProjectPath if present
    If Right(ProjectPath, 1) = Application.PathSeparator Then
        ProjectPath = Left(ProjectPath, Len(ProjectPath) - 1)
    End If

    ' Construct the export directory using the workbook name
    Dim workbookBaseName As String
    workbookBaseName = Left(SourceWorkbook.Name, InStrRev(SourceWorkbook.Name, ".") - 1)

    ' Validate workbook base name
    Dim baseNameValidation As ValidationResult
    Set baseNameValidation = ValidationHelpers.ValidateStringLength("Workbook Base Name", workbookBaseName, 1, 100)
    If Not baseNameValidation.IsValid Then
        ErrorHandler.LogWarning MODULE_NAME, "ExportVbaProject", "Workbook name validation failed: " & baseNameValidation.GetErrorsAsString()
        workbookBaseName = "VBAProject"  ' Fallback name
    End If

    exportDir = ProjectPath & Application.PathSeparator & workbookBaseName & " VBA Project" & Application.PathSeparator

    ' Improved folder creation with error handling
    If Not CreateExportDirectory(exportDir) Then
        Err.Raise ERR_EXPORT_FAILED, "ExportVbaProject", "Failed to create export directory"
    End If

    ' Loop through all VBA components in the specified workbook
    For Each objVbComp In SourceWorkbook.VBProject.VBComponents

        ' Validate component name using ValidationHelpers
        If Not ValidateComponentName(objVbComp.Name) Then
            ErrorHandler.LogWarning MODULE_NAME, "ExportVbaProject", "Skipping component with invalid name: " & objVbComp.Name
            GoTo NextComponent
        End If

        ' Determine the folder name based on the @Folder comment or custom filters
        folderName = GetFolderFromComment(objVbComp, ignoreRootComment)
        If folderName = "" Or useFolderComment = False Then
            folderName = GetFolderNameForComponent(objVbComp, filters)
        End If

        ' Handle root folder case
        If folderName = workbookBaseName Then
            folderName = ""
        End If

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

NextComponent:
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
' Function: GetFolderFromComment
' Purpose: Extracts the folder path from the @Folder comment in the component.
' Inputs:
'   objVbComp (VBComponent): The VBA component.
'   ignoreFirstPart (Boolean): Whether to ignore the first part of the @Folder path.
' Outputs:
'   String: The folder path specified in the @Folder comment, or an empty string if not found.
'*******************************************************************************
Private Function GetFolderFromComment(objVbComp As VBComponent,_
                                      Optional ignoreFirstPart As Boolean = False) As String
    On Error GoTo ErrorHandler

    Dim codeLines As Variant
    Dim line As Variant
    Dim folderPath As String
    Dim i As Integer
    folderPath = ""

    If objVbComp.CodeModule.CountOfLines < 1 Then
        ErrorHandler.LogError MODULE_NAME, "GetFolderFromComment", "Module has no code"
        Exit Function
    End If

    ' Read the code lines of the component
    codeLines = Split(objVbComp.CodeModule.Lines(1, objVbComp.CodeModule.CountOfLines), vbCrLf)

    ' Look for the @Folder comment before any non-comment code
    For i = LBound(codeLines) To UBound(codeLines)
        line = Trim(codeLines(i))
        If Left(line, 1) <> "'" And line <> "" Then
            Exit For
        End If
        If InStr(line, "@Folder") > 0 Then
            folderPath = ExtractFolderPath(line)
            Exit For
        End If
    Next i

    ' Ignore the first part of the path if specified
    If ignoreFirstPart And InStr(folderPath, ".") > 0 Then
        folderPath = Mid(folderPath, InStr(folderPath, ".") + 1)
    End If

    ' Validate folder path before returning using ValidationHelpers for enhanced security
    Dim folderValidation As ValidationResult
    Set folderValidation = ValidationHelpers.ValidateStringLength("Folder Path", folderPath, 0, MAX_PATH_LENGTH)
    If Not folderValidation.IsValid Or (Len(folderPath) > 0 And Not ValidationHelpers.ValidateFilePath(folderPath)) Then
        ErrorHandler.LogError MODULE_NAME, "GetFolderFromComment", "Invalid folder path: " & folderPath & " - " & folderValidation.GetErrorsAsString()
        folderPath = ""
    End If

    GetFolderFromComment = folderPath
    Exit Function

ErrorHandler:
    ErrorHandler.LogError MODULE_NAME, "GetFolderFromComment", "Error: " & Err.Description
    GetFolderFromComment = ""
End Function

Private Function ExtractFolderPath(ByVal line As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = True
        .MultiLine = False
        .IgnoreCase = True

        ' Match @Folder with various formats:
        ' @Folder("path"), @Folder(path), @Folder "path", @Folder /path, @Folder path, etc.
        .pattern = "(@Folder)[\s]*(\([""]?|[""])?(([\./])?([^\)""\./]+))+([""]?\)?)[\n]?$"
    End With

    Dim matches As Object
    Set matches = regex.Execute(line)

    If matches.count = 0 Then Exit Function

    ' Get the captured path
    line = matches(0).SubMatches(0)

    ' Remove leading/trailing whitespace
    line = Trim$(line)

    ' Remove leading dot or slash
    If Left$(line, 1) = "." Or Left$(line, 1) = "/" Then
        line = Mid$(line, 2)
    End If

    ' Convert separators to forward slashes
    line = Replace(Replace(line, ".", "/"), "\", "/")

    ' Clean up multiple slashes
    With regex
        .pattern = "/{2,}"
        line = .Replace(line, "/")
    End With

    ' Remove trailing slash
    If Right$(line, 1) = "/" Then
        line = Left$(line, Len(line) - 1)
    End If

    ExtractFolderPath = line
End Function

'*******************************************************************************
' Function: CreateExportDirectory
' Purpose: Creates the export directory with proper error handling
' Inputs: path (String): The path to create
' Outputs: Boolean: True if successful, False otherwise
'*******************************************************************************
Private Function CreateExportDirectory(path As String) As Boolean
    On Error Resume Next
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    CreateExportDirectory = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
' Function: CreateBackup
' Purpose: Creates a backup of existing exported files
' Inputs:
'   ProjectPath (String): The base path
'   WorkbookName (String): Name of the workbook
'*******************************************************************************
Private Sub CreateBackup(ProjectPath As String, WorkbookName As String)
    Dim backupFolder As String
    backupFolder = ProjectPath & Application.PathSeparator & "Backup_" & Format(Now, "yyyymmdd_hhnnss")

    On Error Resume Next
    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If

    ' Copy existing files to backup
    FileCopy ProjectPath & Application.PathSeparator & "*.*", backupFolder & Application.PathSeparator
    On Error GoTo 0
End Sub

'*******************************************************************************
' Function: ValidateComponentName
' Purpose: Validates a component name for export using ValidationHelpers
' Inputs: name (String): The component name to validate
' Outputs: Boolean: True if valid, False otherwise
'*******************************************************************************
Private Function ValidateComponentName(name As String) As Boolean
    ' Use ValidationHelpers for comprehensive validation
    Dim nameValidation As ValidationResult
    Set nameValidation = ValidationHelpers.ValidateStringLength("Component Name", name, 1, 31)

    ValidateComponentName = nameValidation.IsValid

    ' Log validation errors if any
    If Not nameValidation.IsValid Then
        ErrorHandler.LogWarning MODULE_NAME, "ValidateComponentName", "Component name validation failed: " & nameValidation.GetErrorsAsString()
    End If
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

    If Not ExportVbaProject(ThisWorkbook.path, ThisWorkbook, True, filters) Then
        MsgBox "Export failed.", vbCritical
    Else
        MsgBox "Export successful.", vbInformation
    End If
End Sub
