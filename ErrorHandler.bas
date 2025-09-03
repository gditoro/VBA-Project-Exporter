Attribute VB_Name = "ErrorHandler"
'@Folder "Faturamento.Helpers"
Option Explicit

'*******************************************************************************
' Module: ErrorHandler
' Author: Giovanni Di Toro
' Date: 21-11-2024
' Purpose: Provides error handling and logging functionality
' Dependencies:
'   - FileSystem (FilesSystemHelpers.bas)
'   - CPropertyValidation (CPropertyValidation.cls)
' Notes: This module provides error handling and logging functionality for VBA projects.
' Revision History:
'   - 21-11-2024: Initial version with basic error handling and logging functionality.
'   - 15-12-2024: Added support for batch logging and performance optimizations.
'   - 10-01-2025: Added support for log rotation and file size management.
'   - 15-02-2025: Added support for log level filtering and improved error handling.
'   - 15-03-2025: Added support for log file compression and improved performance.
'   - 26-03-2025: Added support for batch logging and performance optimizations.
'   - 26-08-2025: Finalized error handling and logging module with comprehensive features.
'   - 27-08-2025: Moved enums and types to separate module for better organization.
'*******************************************************************************

'*******************************************************************************
' Module Constants
'*******************************************************************************
Private Const MODULE_NAME As String = "ErrorHandler"

'***************************************************************************
' Logging Configuration
'***************************************************************************
Private Const LOG_ENABLED As Boolean = True
Private Const LOG_LEVEL As Long = LvlDebug ' Configurable logging level (use Long for enum values)
Private Const LOG_TO_FILE As Boolean = True
Private Const LOG_TO_IMMEDIATE As Boolean = True
Private Const INCLUDE_TIMESTAMP As Boolean = True
Private Const INCLUDE_STACK_TRACE As Boolean = True

'***************************************************************************
' Log File Configuration
'***************************************************************************
Private Const FILE_PATH As String = "\logs\"
Private Const LOG_FILE_PREFIX As String = "app_log_"
Private Const MAX_LOG_SIZE_MB As Long = 10
Private Const MAX_LOG_FILES As Long = 5
Private Const DATE_FORMAT As String = "dd-mm-yyyy hh:nn:ss"

'***************************************************************************
' Performance Configuration
'***************************************************************************
Private Const ENABLE_BATCH_LOGGING As Boolean = True
Private Const MAX_BATCH_SIZE As Long = 100
Private Const BATCH_FLUSH_INTERVAL_SEC As Long = 30

'***************************************************************************
' Error Handling Configuration
'***************************************************************************
Private Const MAX_RETRY_ATTEMPTS As Long = 3
Private Const RETRY_DELAY_MS As Long = 100

'***************************************************************************
' Debug Configuration
'***************************************************************************
Private Const DEBUG_MODE As Boolean = True
Private Const MAX_ERROR_COUNT As Long = 1000
Private Const MAX_STACK_DEPTH As Long = 20
Private Const ERR_CALL_STACK As Long = vbObjectError + 513

'******************************************
' Module Variables
'******************************************
Private m_State As ErrorState
Private m_Batch As TErrorBatch
Private m_ErrorCount As Long
Private m_IsInitialized As Boolean

'******************************************
' Win32 API Declarations
'******************************************
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function StackWalk Lib "dbghelp" ( _
        ByVal MachineType As Long, _
        ByVal ProcessId As Long, _
        ByVal ThreadId As Long, _
        ByRef StackFrame As Any, _
        ByRef ContextRecord As Any, _
        ByVal ReadMemoryRoutine As Long, _
        ByVal FunctionTableAccessRoutine As Long, _
        ByVal GetModuleBaseRoutine As Long, _
        ByVal TranslateAddress As Long) As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function StackWalk Lib "dbghelp" ( _
        ByVal MachineType As Long, _
        ByVal ProcessId As Long, _
        ByVal ThreadId As Long, _
        ByRef StackFrame As Any, _
        ByRef ContextRecord As Any, _
        ByVal ReadMemoryRoutine As Long, _
        ByVal FunctionTableAccessRoutine As Long, _
        ByVal GetModuleBaseRoutine As Long, _
        ByVal TranslateAddress As Long) As Long
#End If

'******************************************
' Public Interface
'******************************************

'*******************************************************************************
' Function: LogError
' Purpose: Logs an error message to the log file
' Inputs:
'   ModuleName <String> - The name of the module where the error occurred
'   MethodName <String> - The name of the method where the error occurred
'   ErrorMsg <String> - The error message
' Notes: This function should be called from the error handler of a method to log
'        the error message.
'*******************************************************************************
Public Sub LogError(ByVal ModuleName As String, ByVal MethodName As String, ByVal ErrorMsg As String)
    ProcessLogRequest LvlError, ModuleName, MethodName, ErrorMsg, True
End Sub

'*******************************************************************************
' Function: LogWarning
' Purpose: Logs a warning message to the log file
' Inputs:
'   ModuleName <String> - The name of the module where the warning occurred
'   MethodName <String> - The name of the method where the warning occurred
'   Msg <String> - The warning message
' Notes: This function should be called from the error handler of a method to log
'        the warning message.
'*******************************************************************************
Public Sub LogWarning(ByVal ModuleName As String, ByVal MethodName As String, ByVal Msg As String)
    ProcessLogRequest LvlWarning, ModuleName, MethodName, Msg, False
End Sub

'*******************************************************************************
' Function: LogInfo
' Purpose: Logs an informational message to the log file
' Inputs:
'   ModuleName <String> - The name of the module where the info message occurred
'   MethodName <String> - The name of the method where the info message occurred
'   Msg <String> - The info message
' Notes: This function should be called from the error handler of a method to log
'        the info message.
'*******************************************************************************
Public Sub LogInfo(ByVal ModuleName As String, ByVal MethodName As String, ByVal Msg As String)
    If DEBUG_MODE Then ProcessLogRequest LvlInfo, ModuleName, MethodName, Msg, False
End Sub

'*******************************************************************************
' Function: HandleError
' Purpose: Handles an error by logging it and optionally showing a message box
' Inputs:
'   Source <String> - The name of the module where the error occurred
'   Method <String> - The name of the method where the error occurred
'   ShowUser <Boolean> - Whether to show a message box to the user
' Notes: This function should be called from the error handler of a method to log
'        the error message and optionally show a message box to the user.
'*******************************************************************************
Public Sub HandleError(ByVal Source As String, _
                      ByVal Method As String, _
                      Optional ByVal ShowUser As Boolean = False)
    On Error Resume Next

    If Not ValidateParams(Source, Method, Err.Description) Then Exit Sub

    Dim Msg As String
    Msg = FormatErrorMessage(Source, Method, Err.Description)
    LogError Source, Method, Msg

    If ShowUser Then ShowErrorMessage Msg
End Sub

'*******************************************************************************
' Function: AssertNotNothing
' Purpose: Asserts that an object is not Nothing
' Inputs:
'   obj <Object> - The object to check
'   objectName <String> - The name of the object
' Notes: This function should be used to validate that an object is not Nothing
'        before using it in a method.
'*******************************************************************************
Public Sub AssertNotNothing(ByVal obj As Object, ByVal objectName As String)
    If obj Is Nothing Then
        Err.Raise 91, "Assert", objectName & " cannot be Nothing"
    End If
End Sub

'******************************************
' Private Implementation
'******************************************

'*******************************************************************************
' Function: InitializeModule
' Purpose: Initializes the error handler module
' Notes: This function should be called before using any other functions in the module
'*******************************************************************************
Private Sub InitializeModule()
    If m_IsInitialized Then Exit Sub

    On Error GoTo ErrorHandler

    CreateLogDirectory
    ReDim m_Batch.Entries(0 To 99)  ' Initialize with 100 slots
    m_Batch.EntryCount = 0
    m_Batch.LastFlush = Now
    m_ErrorCount = 0
    m_IsInitialized = True

    Exit Sub

ErrorHandler:
    Debug.Print "Failed to initialize ErrorHandler: " & Err.Description
End Sub

'*******************************************************************************
' Function: CreateLogDirectory
' Purpose: Creates the log directory if it does not exist
'*******************************************************************************
Private Sub CreateLogDirectory()
    Dim logPath As String
    logPath = ThisWorkbook.path & FILE_PATH

    If Not fsIsFolder(logPath) Then
        fsCreateFolder logPath
    End If
End Sub

'*******************************************************************************
' Function: ValidateParams
' Purpose: Validates the input parameters for the logging functions
' Inputs:
'   ModuleName <String> - The name of the module
'   MethodName <String> - The name of the method
'   Msg <string> - The log message
' Returns <Boolean>: True if the parameters are valid, False otherwise
'*******************************************************************************
Private Function ValidateParams(ModuleName As String, MethodName As String, Msg As String) As Boolean
    ValidateParams = False

    If ModuleName = vbNullString Or MethodName = vbNullString Then Exit Function
    If Msg = vbNullString Then Exit Function

    ValidateParams = True
End Function

'*******************************************************************************
' Function: FormatLogEntry
' Purpose: Formats a log entry for writing to the log file
' Inputs:
'  entry <TLogEntry> - The log entry to format
' Returns: The formatted log entry as a string
'*******************************************************************************
Private Function FormatLogEntry(entry As TLogEntry) As String
    Dim Msg As String

    With entry
        Msg = Format$(.Timestamp, DATE_FORMAT) & " " & _
            GetLevelPrefix(.Level) & " [" & _
            .ModuleName & "." & .MethodName & "] " & _
            .Message

        If .ErrorNumber <> 0 Then
            Msg = Msg & vbNewLine & _
                "Error: " & .ErrorNumber & vbNewLine & _
                "Source: " & .Source & vbNewLine & _
                "Stack: " & .stackTrace
        End If
    End With

    FormatLogEntry = Msg
End Function

'********************************************************************************
' Function: ProcessLogEntry
' Purpose: Processes a log entry by writing it to the log file or batching it
' Inputs:
'  entry <TLogEntry> - The log entry to process
'********************************************************************************
Private Sub ProcessLogEntry(ByRef entry As TLogEntry)
    If m_State = ErrorState.Processing Then Exit Sub
    m_State = ErrorState.Processing

    On Error GoTo ErrorHandler

    If ShouldBatchLog(entry.Level) Then
        AddToBatchSafely entry
    Else
        WriteLogSafely entry
    End If

CleanUp:
    m_State = ErrorState.Ready
    Exit Sub

ErrorHandler:
    DebugPrint MODULE_NAME, "ProcessLogEntry", Err.Description
    Resume CleanUp
End Sub

'********************************************************************************
' Function: WriteLogEntry
' Purpose: Writes a message to the log file
' Inputs:
'  Level <LogLevel> - The log level
'  ModuleName <String> - The name of the module
'  MethodName <String> - The name of the method
'  Msg <String> - The log message
'********************************************************************************
Private Sub WriteLogEntry(ByVal Level As LogLevel, _
                         ByVal ModuleName As String, _
                         ByVal MethodName As String, _
                         ByVal Msg As String)
    EnsureInitialized

    Dim entry As TLogEntry
    With entry
        .Level = Level
        .ModuleName = ModuleName
        .MethodName = MethodName
        .Message = Msg
        .Timestamp = Now
    End With

    ProcessLogEntry entry
End Sub

'********************************************************************************
' Function: WriteToLog
' Purpose: Writes a log entry to the log file
' Inputs:
'  entry <TLogEntry> - The log entry to write
'********************************************************************************
Private Sub WriteToLog(entry As TLogEntry)
    Dim logMsg As String
    logMsg = FormatLogEntry(entry)

    Debug.Print logMsg
    WriteToFile logMsg, entry.Level
End Sub

'********************************************************************************
' Function: GetLevelPrefix
' Purpose: Returns the log level prefix for a given log level
' Inputs:
'  Level - The log level (LogLevel)
' Returns <String>: The log level prefix as a string
' Note: The log level prefix is used to identify the log level in the log file
'********************************************************************************
Private Function GetLevelPrefix(Level As LogLevel) As String
    Select Case Level
        Case LvlError: GetLevelPrefix = "ERROR"
        Case LvlWarning: GetLevelPrefix = "WARN"
        Case LvlInfo: GetLevelPrefix = "INFO"
        Case LvlDebug: GetLevelPrefix = "DEBUG"
    End Select
End Function

'********************************************************************************
' Function: ValidateLogEntry
' Purpose: Validates a log entry
' Inputs:
'  entry <TLogEntry> - The log entry to validate
' Returns <Boolean>: True if the log entry is valid, False otherwise
' Note: A log entry is considered valid if it has a module name, method name, and message
'********************************************************************************
Private Function ValidateLogEntry(ByRef entry As TLogEntry) As Boolean
    On Error GoTo ErrorHandler

    ValidateLogEntry = False

    ' Check recursion
    If m_State = ErrorState.Processing Then
        DebugPrint MODULE_NAME, "ValidateLogEntry", "Recursive logging detected"
        Exit Function
    End If

    ' Check required fields
    If entry.ModuleName = vbNullString Then
        DebugPrint MODULE_NAME, "ValidateLogEntry", "Missing module name"
        Exit Function
    End If

    If entry.MethodName = vbNullString Then
        DebugPrint MODULE_NAME, "ValidateLogEntry", "Missing method name"
        Exit Function
    End If

    If entry.Message = vbNullString Then
        DebugPrint MODULE_NAME, "ValidateLogEntry", "Missing message"
        Exit Function
    End If

    ' Add retry limit check
    If entry.Retries >= 3 Then
        DebugPrint MODULE_NAME, "ValidateLogEntry", "Max retries exceeded"
        Exit Function
    End If

    ValidateLogEntry = True
    Exit Function

ErrorHandler:
    DebugPrint MODULE_NAME, "ValidateLogEntry", "Validation error: " & Err.Description
End Function
'********************************************************************************
' Function: WriteToFile
' Purpose: Writes a message to the log file
' Inputs:
'  Msg <String> - The message to write
'  Level <LogLevel> - The log level of the message
' Note: This function is used to write log messages to the log file and handle log rotation
'********************************************************************************
Private Sub WriteToFile(ByVal Msg As String, Level As LogLevel)
    On Error GoTo ErrorHandler

    EnsureInitialized
    If Not CanProcessLog() Then Exit Sub

    CheckErrorThreshold

    ' Prevent recursive logging
    If m_State = ErrorState.Processing Then Exit Sub
    m_State = ErrorState.Processing

    FlushBatchIfNeeded

    Dim logFile As String
    logFile = GetLogFilePath()

    If fsIsFile(logFile) Then
        If fsGetFileSize(logFile) > MAX_LOG_SIZE_MB Then
            RotateLogFile logFile
        End If
    End If

    WriteTextToFile logFile, Msg & vbNewLine

    m_State = ErrorState.Ready
    Exit Sub

ErrorHandler:
    DebugPrint MODULE_NAME, "WriteToFile", "Failed to write log: " & Err.Description
    m_State = ErrorState.Ready
End Sub

'********************************************************************************
' Function: GetLogFilePath
' Purpose: Returns the path of the log file for the current date
' Returns <String>: The path of the log file
'********************************************************************************
Private Function GetLogFilePath() As String
    GetLogFilePath = ThisWorkbook.path & FILE_PATH & _
                     Format$(Date, "ddmmyyyy") & ".log"
End Function

'********************************************************************************
' Function: RotateLogFile
' Purpose: Renames the log file to include the current time stamp
' Inputs:
'  filePath <String> - The path of the log file to rotate
' Note: This function is used to rotate the log file when it reaches the maximum size
'      or when a new log file is created for the current date
'********************************************************************************
Private Sub RotateLogFile(ByVal FilePath As String)
    On Error GoTo ErrorHandler

    Dim newPath As String
    newPath = FilePath & "." & Format$(Now, "hhnnss")

    If fsFileExists(FilePath) Then
        fsRenameFile FilePath, newPath
        ResetState
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Failed to rotate log file: " & Err.Description
End Sub

'********************************************************************************
' Function: AddToBatch
' Purpose: Adds a log entry to the batch for processing
' Inputs:
'  entry <TLogEntry> - The log entry to add to the batch
' Note: This function is used to batch log entries for more efficient writing to the log file
'********************************************************************************
Private Sub AddToBatch(entry As TLogEntry)
    ' Resize array if needed
    If m_Batch.EntryCount >= UBound(m_Batch.Entries) Then
        ReDim Preserve m_Batch.Entries(0 To UBound(m_Batch.Entries) + 50)
    End If

    ' Add entry to array
    m_Batch.Entries(m_Batch.EntryCount) = entry
    m_Batch.EntryCount = m_Batch.EntryCount + 1

    If m_Batch.EntryCount >= MAX_BATCH_SIZE Then
        FlushBatch
    End If
End Sub

'********************************************************************************
' Function: AddToBatchSafely
' Purpose: Adds a log entry to the batch for processing with retry logic
' Inputs:
'  entry <TLogEntry> - The log entry to add to the batch
'********************************************************************************
Private Sub AddToBatchSafely(ByRef entry As TLogEntry)
    If m_Batch.State = ErrorState.Processing Then Exit Sub
    m_Batch.State = ErrorState.Processing

    On Error GoTo RetryHandler
    AddToBatch entry
    m_Batch.State = ErrorState.Ready
    Exit Sub

RetryHandler:
    If m_Batch.RetryCount < 3 Then
        m_Batch.RetryCount = m_Batch.RetryCount + 1
        AddToBatchSafely entry
    Else
        DebugPrint MODULE_NAME, "AddToBatchSafely", "Max retries exceeded"
    End If
    m_Batch.State = ErrorState.Ready
End Sub

Private Sub WriteLogSafely(ByRef entry As TLogEntry)
    On Error GoTo RetryHandler

    WriteToLog entry
    Exit Sub

RetryHandler:
    If entry.Retries < 3 Then
        entry.Retries = entry.Retries + 1
        WriteLogSafely entry
    Else
        DebugPrint MODULE_NAME, "WriteLogSafely", "Failed after retries"
    End If
End Sub

'********************************************************************************
' Function: FlushBatch
' Purpose: Writes all log entries in the batch to the log file and clears the batch
' Note: This function is used to write all log entries in the batch to the log file
'      and clear the batch after writing
'********************************************************************************
Private Sub FlushBatch()
    If m_Batch.EntryCount = 0 Then Exit Sub

    Dim i As Long
    For i = 0 To m_Batch.EntryCount - 1
        WriteToLog m_Batch.Entries(i)
    Next i

    ' Reset the batch
    m_Batch.EntryCount = 0
    m_Batch.LastFlush = Now
End Sub

Private Sub EnsureInitialized()
    If Not m_IsInitialized Then InitializeModule
End Sub

'********************************************************************************
' Function: GetCallStack
' Purpose: Gets the current call stack
' Returns: String containing the formatted call stack
' Note: This is a VBA7-specific implementation using Win32 API
'********************************************************************************
Private Function GetCallStack() As String
    On Error GoTo ErrorHandler

    Dim stackTrace As String
    Dim caller As String
    Dim i As Long
    Dim depth As Long

    stackTrace = ""
    depth = 0

    ' Get call stack info using error raising technique
    Do
        depth = depth + 1
        If depth > MAX_STACK_DEPTH Then Exit Do

        caller = GetCallerInfo(depth)
        If caller = "" Then Exit Do

        stackTrace = stackTrace & IIf(stackTrace = "", "", vbNewLine) & _
                    String(depth * 2, " ") & caller
    Loop

    GetCallStack = stackTrace
    Exit Function

ErrorHandler:
    GetCallStack = "Failed to get call stack: " & Err.Description
End Function

'********************************************************************************
' Function: GetCallerInfo
' Purpose: Gets information about a specific caller in the stack
' Inputs: depth - How far up the call stack to look
' Returns: String containing caller information or empty if not found
'********************************************************************************
Private Function GetCallerInfo(ByVal depth As Long) As String
    On Error Resume Next

    ' Use error object to get call stack info
    Err.Raise ERR_CALL_STACK

    If Err.Number = ERR_CALL_STACK Then
        Dim stackInfo As String
        stackInfo = Err.Source

        ' Parse the call stack info
        If InStr(stackInfo, "Line") > 0 Then
            Dim parts() As String
            parts = Split(stackInfo, " ")
            If UBound(parts) >= 3 Then
                GetCallerInfo = "at " & parts(1) & " line " & parts(3)
            End If
        End If
    End If

    Err.Clear
End Function

'********************************************************************************
' Function: GetStackTrace
' Purpose: Returns the call stack as a string for debugging purposes (VBA6)
' Returns <String>: The call stack as a string
'********************************************************************************
Private Function GetStackTrace() As String
    #If VBA7 Then
        GetStackTrace = GetCallStack()
    #Else
        GetStackTrace = "Stack trace not available in VBA6"
    #End If
End Function

'********************************************************************************
' Function: FormatErrorMessage
' Purpose: Formats an error message with the source and method names
' Inputs:
'  Source <String> - The name of the source module
'  Method <String> - The name of the method
'  Description <String> - The error description
' Returns <String>: The formatted error message as a string
'********************************************************************************
Private Function FormatErrorMessage(ByVal Source As String, _
                                  ByVal Method As String, _
                                  ByVal Description As String) As String
    FormatErrorMessage = "Error in " & Source & "." & Method & vbNewLine & _
                        Description
End Function

'********************************************************************************
' Function: ShouldLogError
' Purpose: Determines whether an error should be logged based on the log level
' Inputs:
'  Level <LogLevel> - The log level of the error
'  ErrorNumber <Long> - The error number
' Returns <Boolean>: True if the error should be logged, False otherwise
'********************************************************************************
Private Function ShouldLogError(ByVal Level As LogLevel, ByVal ErrorNumber As Long) As Boolean
    ' Filter out certain error numbers or limit frequency
    Select Case ErrorNumber
        Case 424, 91 ' Object required errors
            ShouldLogError = (Level = LvlError)
        Case Else
            ShouldLogError = True
    End Select
End Function

'********************************************************************************
' Function: RotateLogs
' Purpose: Rotates log files by deleting the oldest files beyond the limit
' Note: This function is used to manage the number of log files by deleting the oldest
'      files when the limit is exceeded
'********************************************************************************
Private Sub RotateLogs()
    Dim logPath As String
    logPath = ThisWorkbook.path & FILE_PATH

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get log files sorted by date
    Dim files As Collection
    Set files = GetLogFiles(logPath)

    ' Delete oldest files beyond limit
    While files.count > MAX_LOG_FILES
        fso.DeleteFile files(1)
        files.Remove 1
    Wend
End Sub

'********************************************************************************
' Function: CheckErrorThreshold
' Purpose: Checks if the error count has exceeded the threshold and disables logging
'********************************************************************************
Private Sub CheckErrorThreshold()
    m_ErrorCount = m_ErrorCount + 1

    If m_ErrorCount > MAX_ERROR_COUNT Then
        LogWarning MODULE_NAME, "CheckErrorThreshold", "Error count exceeded. Logging disabled."
        ResetState
    End If
End Sub

'********************************************************************************
' Function: FlushBatchIfNeeded
' Purpose: Flushes the log batch if it exceeds the maximum size or time threshold
'********************************************************************************
Private Sub FlushBatchIfNeeded()
    If Not ShouldFlushBatch() Then Exit Sub
    FlushBatch
End Sub

'********************************************************************************
' Function: ShouldBatchLog
' Purpose: Determines whether a log entry should be batched based on the log level
' Inputs:
' Level - The log level of the entry
' Returns <Boolean>: True if the entry should be batched, False otherwise
' Note: Log entries with a log level of LvlInfo or higher are batched, because
'       they are less critical and can be written to the log file in batches
'********************************************************************************
Private Function ShouldBatchLog(Level As LogLevel) As Boolean
    ShouldBatchLog = (Level >= LvlInfo)
End Function

'********************************************************************************
' Function: ShouldFlushBatch
' Purpose: Determines whether the log batch should be flushed based on the size and time
' Returns: True if the batch should be flushed, False otherwise
'********************************************************************************
Private Function ShouldFlushBatch() As Boolean
    If m_Batch.EntryCount = 0 Then Exit Function

    ShouldFlushBatch = (m_Batch.EntryCount >= MAX_BATCH_SIZE) Or _
                       (DateDiff("s", m_Batch.LastFlush, Now) > 60)
End Function

'********************************************************************************
' Function: ResetState
' Purpose: Resets the error handler state and clears the log batch and error count
'********************************************************************************
Private Sub ResetState()
    m_State = ErrorState.Ready
    m_ErrorCount = 0
    m_Batch.State = ErrorState.Ready
    m_Batch.RetryCount = 0
    m_Batch.EntryCount = 0  ' Reset entry count instead of creating new collection
    m_Batch.LastFlush = Now
End Sub

'********************************************************************************
' Function: CanProcessLog
' Purpose: Determines wheter the log is enabled and the error count is within the limit
' Returns: True if the log can be processed, False otherwise
'********************************************************************************
Private Function CanProcessLog() As Boolean
    If Not LOG_ENABLED Then Exit Function
    EnsureInitialized
    If m_State = ErrorState.disabled Then Exit Function
    If m_ErrorCount > MAX_ERROR_COUNT Then Exit Function

    CanProcessLog = True
End Function

'********************************************************************************
' Function: ShowErrorMessage
' Purpose: Shows an error message (MsgBox) to the user
' Inputs:
'  msg <String> - The error message to show
'********************************************************************************
Private Sub ShowErrorMessage(ByVal Msg As String)
    If Not DEBUG_MODE Then Exit Sub
    On Error Resume Next
    MsgBox Msg, vbExclamation, "Error"
End Sub

'********************************************************************************
' Function: ProcessLogRequest
' Purpose: Processes a log request by creating a log entry and validating it
' Inputs:
'  Level <LogLevel> - The log level of the message
'  ModuleName <String> - The name of the module
'  MethodName <String> - The name of the method
'  Msg <String> - The log message
'  IsError <Boolean> - Whether the message is an error
'********************************************************************************
Private Sub ProcessLogRequest(ByVal Level As LogLevel, _
                            ByVal ModuleName As String, _
                            ByVal MethodName As String, _
                            ByVal Msg As String, _
                            ByVal IsError As Boolean)
    If Not CanProcessLog() Then Exit Sub

    Dim entry As TLogEntry
    InitializeLogEntry entry, Level, ModuleName, MethodName, Msg, IsError

    If ValidateLogEntry(entry) Then ProcessLogEntry entry
End Sub

Private Sub InitializeLogEntry(ByRef entry As TLogEntry, _
                             ByVal Level As LogLevel, _
                             ByVal ModuleName As String, _
                             ByVal MethodName As String, _
                             ByVal Msg As String, _
                             ByVal IsError As Boolean)
    With entry
        .Level = Level
        .ModuleName = ModuleName
        .MethodName = MethodName
        .Message = Msg
        .Timestamp = Now
        If IsError Then
            .ErrorNumber = Err.Number
            .Source = Err.Source
            .stackTrace = GetStackTrace()
        End If
        .Retries = 0
    End With
End Sub

'********************************************************************************
' Function: DebugPrint
' Purpose: Prints a debug message to the Immediate window without the need for conditional compilation
'          and avoiding recursion inside ErrorHandler.bas
' Inputs:
'   ModuleName <String> - The name of the module
'   MethodName <String> - The name of the method
'   Msg <String> - The debug message
'********************************************************************************
Private Sub DebugPrint(ByVal ModuleName As String, ByVal MethodName As String, ByVal Msg As String)
    If Not DEBUG_MODE Then Exit Sub
    Debug.Print Format$(Now, DATE_FORMAT) & " [DEBUG] " & ModuleName & "." & MethodName & ": " & Msg
End Sub

'********************************************************************************
' Function: WriteTextToFile
' Purpose: Writes text to a file, creating the file if it doesn't exist
' Inputs:
'  FilePath <String> - The full path to the file
'  Text <String> - The text to write to the file
' Note: Appends text to existing file or creates new file if it doesn't exist
'********************************************************************************
Private Sub WriteTextToFile(ByVal FilePath As String, ByVal Text As String)
    On Error GoTo ErrorHandler

    Const ForAppending As Long = 8
    Const TristateTrue As Long = -1  ' Unicode

    Dim fileNum As Long
    fileNum = FreeFile()

    ' Open file for appending in Unicode mode
    Open FilePath For Append As #fileNum
    Print #fileNum, Text
    Close #fileNum

    Exit Sub

ErrorHandler:
    ' Log error but don't raise to avoid recursion
    DebugPrint MODULE_NAME, "WriteTextToFile", _
              "Failed to write to file '" & FilePath & "': " & Err.Description

    ' Ensure file is closed
    On Error Resume Next
    Close #fileNum
End Sub

'********************************************************************************
' Function: GetLogFiles
' Purpose: Gets a collection of log files sorted by date
' Inputs: logPath - The path to search for log files
' Returns: Collection of log file paths sorted by date
'********************************************************************************
Private Function GetLogFiles(ByVal logPath As String) As Collection
    Set GetLogFiles = New Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(logPath) Then Exit Function

    Dim file As Object
    Dim files() As String
    ReDim files(0)

    For Each file In fso.GetFolder(logPath).files
        If Right$(file.name, 4) = ".log" Then
            ReDim Preserve files(UBound(files) + 1)
            files(UBound(files)) = file.path
        End If
    Next

    ' Sort files by date (newest first) using ArrayHelpers
    If UBound(files) > 0 Then
        On Error Resume Next
        arrSortFilesByDate files
        If Err.Number <> 0 Then
            ' Fall back to simple date comparison if ArrayHelpers is not available
            DebugPrint MODULE_NAME, "GetLogFiles", "ArrayHelpers not available: " & Err.Description
        End If
        On Error GoTo 0
    End If

    ' Add sorted files to collection
    Dim i As Long
    For i = LBound(files) To UBound(files)
        If files(i) <> "" Then GetLogFiles.Add files(i)
    Next i
End Function
