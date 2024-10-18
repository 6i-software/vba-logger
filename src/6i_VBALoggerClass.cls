VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBALoggerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ------------------------------------- '
'                                       '
'    VBA Logger                         '
'    Copyright © 2024, 6i software      '
'                                       '
' ------------------------------------- '
'
' ---
' version: 1.0.0
' last release date: 2024.10.17
' author: 'v20100v <vb20100bv@gmail.com>'
' GUID : 2435fa00-af72-4c60-b90f-608139a102c6
' repository: https://github.com/6i-software/vba-logger
' ---
'
' The `VBALogger` class module is a logging system designed to track and record events or
' messages within a VBA application. It allows logging of different types of messages with
' varying levels of severity and supports outputting these logs either to the VBA console or
' to log files.
'
'---


Option Explicit


'__________________________'
'                          '
'          Members         '
'__________________________'

Public Enum LogVerbosityLevel
    LevelError = -2    ' Runtime errors should be logged and monitored.
    LevelWarning = -1  ' Exceptional occurrences that are not errors.
    LevelLog = 0       ' Normal log.
    LevelNotice = 1    ' Significant events.
    LevelInfo = 2      ' Interesting events.
    LevelTrace = 3     ' For detailed debug information.
End Enum

Public Enum LogOutput
    Console = 1  ' Send message to VBA console (a.k.a. the immediate Window)
    File = 2     ' Send message to a logfile (persist in ./var/log folder of workbook)
    All = LogOutput.Console Or LogOutput.File
End Enum

Public LogVerbosityLevelSession As LogVerbosityLevel ' The level verbosity of log for the current session
Public LogOutputSession As LogOutput                 ' Define where log will be written (console, file, ...)
Public OptionShowSplashscreen As Boolean             ' Show a splashcreen when session logger is starting
Public splashscreen As String                        ' Customizes the content of splashcreen

Private pLogFileName As String                       ' The name of the logfile
Private pLogFilePath As String                       ' The path where the logfile is save
Private pLogFileFolder As String                     ' The folder where the logfile is save
Private pCounterLogEntries As Long                   ' Counts the number of log messages that have been logged
Private pInstantiationSource As String               ' Indicates how the class was instantiated (via factory or directly)
Private pLibraryVersion As String                    ' The version number of class module


'-------------------------------------------------------------------------------------
' Getter and setter of the logfile name
'-------------------------------------------------------------------------------------
Public Property Get LogFileName() As String
    LogFileName = pLogFileName
End Property
Public Property Let LogFileName(ByVal newLogFileName As String)
    If Not IsValidFileName(newLogFileName) Then
        Err.Raise vbObjectError + 1001, "Logger::SetLogFileName", "The new name for logfile (" & newLogFileName & ") contains invalid characters !"
    End If
    pLogFileName = newLogFileName
End Property
Private Function IsValidFileName(ByVal fileName As String) As Boolean
    Dim invalidChars As String
    Dim i As Integer
    
    ' Set unvalid characters for file name
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        If InStr(fileName, Mid(invalidChars, i, 1)) > 0 Then
            IsValidFileName = False
            Exit Function
        End If
    Next i
    
    IsValidFileName = True
End Function


'-------------------------------------------------------------------------------------
' Getter and setter of the logfile folder
'-------------------------------------------------------------------------------------
Public Property Get LogFileFolder() As String
    LogFileFolder = pLogFileFolder
End Property
Public Property Let LogFileFolder(ByVal newLogFileFolder As String)
    ' Try to create the new folder recursvily
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer

    On Error Resume Next
    pathParts = Split(newLogFileFolder, "\")
    currentPath = pathParts(0)
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & "\" & pathParts(i)
        If Dir(currentPath, vbDirectory) = "" Then
            MkDir currentPath
        End If
    Next i
    On Error GoTo 0

    ' Check final folder
    If Dir(newLogFileFolder, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1002, "Logger::SetLogFileFolder", "Failed to create the new folder (" & newLogFileFolder & ") !"
    End If

    pLogFileFolder = newLogFileFolder
    pLogFilePath = newLogFileFolder & "\" & pLogFileName
End Property


'-------------------------------------------------------------------------------------
' Getter of the logfile path
'-------------------------------------------------------------------------------------
Public Property Get LogFilePath() As String
    LogFilePath = pLogFilePath
End Property


'-------------------------------------------------------------------------------------
' Getter of the version
'-------------------------------------------------------------------------------------
Public Property Get Version() As String
    Version = pLibraryVersion
End Property


'_____________________________'
'                             '
'          Initialize         '
'_____________________________'

'-------------------------------------------------------------------------------------
' Setting up default values when an instance is created.
'
' Class_Initialize is a default event procedure that is triggered when the class is
' initialized, meaning when the constructor is called by using the keyword 'New'
' (e.g. Set myCustomLogger = New VBALogger).
'-------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    LogVerbosityLevelSession = LogVerbosityLevel.LevelTrace
    LogOutputSession = LogOutput.Console
        
    ' Splashscreen is disabled by default
    OptionShowSplashscreen = False
    splashscreen = ""
    
    ' By default, we consider this is a direct instanciation (and not factory approach)
    pInstantiationSource = "Direct"
    
    ' Set the version number
    pLibraryVersion = "1.0.0"
End Sub


Public Sub Initialize( _
    Optional ByVal paramSplashscreen As String = "", _
    Optional ByVal paramLogOutputSession As LogOutput = LogOutput.Console, _
    Optional ByVal paramLogFileName As String = "", _
    Optional ByVal paramLogFileFolder As String = "" _
)
    If paramSplashscreen <> "" Then
        OptionShowSplashscreen = True
        splashscreen = paramSplashscreen
    End If
    
    ' Change log output for current session
    LogOutputSession = paramLogOutputSession
    If (paramLogOutputSession = LogOutput.All) Or (paramLogOutputSession = LogOutput.File) Then
        If paramLogFileName = "" Then
            LogFileName = "logfile_" & format(Now, "yyyy-mm-dd") & ".log"
        Else
            LogFileName = paramLogFileName
        End If

        If paramLogFileFolder = "" Then
            LogFileFolder = ThisWorkbook.Path & "\var\log"
        Else
            LogFileFolder = paramLogFileFolder
        End If
    End If
    
    If OptionShowSplashscreen Then
        Log (splashscreen)
    End If
End Sub


Public Sub MarkInstanciationFromFactroy()
    pInstantiationSource = "Factory"
End Sub




'_________________________________'
'                                 '
'          Public methods         '
'_________________________________'

Public Sub Error(ByVal message As String)
    Call PrepareLogEntry(LogVerbosityLevel.LevelError, message)
End Sub


Public Sub Warning(ByVal message As String)
    Call PrepareLogEntry(LogVerbosityLevel.LevelWarning, message)
End Sub


Public Sub Notice(ByVal message As String)
    Call PrepareLogEntry(LogVerbosityLevel.LevelNotice, message)
End Sub


Public Sub Info(ByVal message As String)
    Call PrepareLogEntry(LogVerbosityLevel.LevelInfo, message)
End Sub


Public Sub Trace(ByVal message As String, Optional ByVal context As Variant)
    Call PrepareLogEntry(LogVerbosityLevel.LevelTrace, message, context)
End Sub


Public Sub Log(ByVal message As String, Optional ByVal withNewLine As Boolean = True)
    ' Increments the number of log messages that have been logged
    pCounterLogEntries = pCounterLogEntries + 1
        
    Call WriteLogEntry(message, withNewLine)
End Sub


Public Function ToString() As String
    Dim output As String
    output = vbCrLf
    output = output & "Debug VBALogger instance" & vbCrLf
    output = output & "------------------------" & vbCrLf
    output = output & "VBALoggerClass version: " & Version() & vbCrLf
    output = output & "Log verbosity level session: '" & GetLogVerbosityLevelHumanReadable(LogVerbosityLevelSession) & "'" & vbCrLf
    output = output & "Log output session: '" & GetLogOutputHumanReadable(LogOutputSession) & "'" & vbCrLf
    output = output & "Log file name: '" & LogFileName & "'" & vbCrLf
    output = output & "Log file path: '" & pLogFilePath & "'" & vbCrLf
    output = output & "Log file folder: '" & LogFileFolder & "'" & vbCrLf
    output = output & "Option show splashscreen: " & IIf(OptionShowSplashscreen, "True", "False") & vbCrLf
    output = output & "Total messages logged: " & pCounterLogEntries & vbCrLf
    output = output & "Object Type: '" & TypeName(Me) & "'" & vbCrLf
    output = output & "Instantiation source: '" & pInstantiationSource & "'" & vbCrLf
    output = output & "Memory Address: " & VarPtr(Me) & vbCrLf
    output = output & "------------------------" & vbCrLf
    
    ToString = output
End Function




'__________________________________'
'                                  '
'          Private methods         '
'__________________________________'

Private Sub PrepareLogEntry(ByVal paramLogVerbosityLevel As LogVerbosityLevel, ByVal paramMessage As String, Optional ByVal context As Variant)
    If paramLogVerbosityLevel <= LogVerbosityLevelSession Then
        ' Prepare the align level (prefix of log entry)
        Dim level As String
        level = "[" & GetLogVerbosityLevelHumanReadable(paramLogVerbosityLevel) & "]"
        Dim alignedLevel As String
        alignedLevel = level & Space(9 - Len(level))
        
        ' Prepare the log entry from given message and level verbosity
        Dim logEntry As String
        logEntry = alignedLevel & " | " & format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & paramMessage
        
        ' Concatenate context informations if given
        If Not IsMissing(context) Then
            Dim contextInfo As String
            contextInfo = GetContextInfo(context)
            logEntry = logEntry & " | context=" & contextInfo & ""
        End If
        
        ' Increments the number of log messages that have been logged
        pCounterLogEntries = pCounterLogEntries + 1
        
        Call WriteLogEntry(logEntry)
    End If
End Sub


Private Sub WriteLogEntry(ByVal paramLogEntry As String, Optional ByVal withNewLine As Boolean = True)
    Dim output As String
    
    Select Case LogOutputSession
        Case LogOutput.Console
            If withNewLine = True Then
                Debug.Print paramLogEntry
            Else
                Debug.Print paramLogEntry;
            End If
        Case LogOutput.File
            WriteIntoLogFile paramLogEntry, withNewLine
        Case LogOutput.All
            If withNewLine = True Then
                Debug.Print paramLogEntry
            Else
                Debug.Print paramLogEntry;
            End If
            WriteIntoLogFile paramLogEntry, withNewLine
    End Select
End Sub


Private Sub WriteIntoLogFile(ByVal paramLogEntry As String, ByVal withNewLine As Boolean)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open pLogFilePath For Append As fileNum
    If withNewLine Then
        Print #fileNum, paramLogEntry ' Adds a new line
    Else
        Print #fileNum, paramLogEntry; ' No new line
    End If
    Close fileNum
End Sub


Private Function GetLogVerbosityLevelHumanReadable(ByVal paramLogVerbosityLevel As LogVerbosityLevel) As String
    Select Case paramLogVerbosityLevel
        Case LogVerbosityLevel.LevelError: GetLogVerbosityLevelHumanReadable = "ERROR"
        Case LogVerbosityLevel.LevelWarning: GetLogVerbosityLevelHumanReadable = "WARNING"
        Case LogVerbosityLevel.LevelLog: GetLogVerbosityLevelHumanReadable = "LOG"
        Case LogVerbosityLevel.LevelNotice: GetLogVerbosityLevelHumanReadable = "NOTICE"
        Case LogVerbosityLevel.LevelInfo: GetLogVerbosityLevelHumanReadable = "INFO"
        Case LogVerbosityLevel.LevelTrace: GetLogVerbosityLevelHumanReadable = "TRACE"
        Case Else
            GetLogVerbosityLevelHumanReadable = "Unknow"
            Err.Raise vbObjectError + 1000, "VBALogger::GetLogVerbosityHumanReadable", "Use unvalid parameter LogVerbosityLevel ! (" & paramLogVerbosityLevel & ")."
    End Select
End Function


Private Function GetLogOutputHumanReadable(ByVal paramLogOutput As LogOutput) As String
    Select Case paramLogOutput
        Case LogOutput.Console: GetLogOutputHumanReadable = "Console"
        Case LogOutput.File: GetLogOutputHumanReadable = "File"
        Case LogOutput.All: GetLogOutputHumanReadable = "All"
        Case Else
            GetLogOutputHumanReadable = "Unknown"
            Err.Raise vbObjectError + 1000, "VBALogger::GetLogOutputHumanReadable", "Use unvalid parameter LogOutput ! (" & paramLogOutput & ")."
    End Select
End Function


Private Function GetLogVerbosityLevelFromString(ByVal logVerbosityLevelInString As String) As LogVerbosityLevel
    Select Case UCase(logVerbosityLevelInString)
        Case "ERROR": GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelError
        Case "WARNING": GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelWarning
        Case "LOG": GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelLog
        Case "NOTICE": GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelNotice
        Case "INFO": GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelInfo
        Case "TRACE": GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelTrace
        Case Else
            GetLogVerbosityLevelFromString = LogVerbosityLevel.LevelWarning ' Default value for unvalid string log level
            Err.Raise vbObjectError + 1000, "VBALogger::GetLogVerbosityLevelFromString", "Use invalid string to get a LogVerbosityLevel ! (" & logLevelStr & ")."
    End Select
End Function


Private Function GetLogOutputFromString(ByVal logOutputStr As String) As LogOutput
    Select Case UCase(logOutputStr)
        Case "CONSOLE"
            GetLogOutputFromString = LogOutput.Console
        Case "FILE"
            GetLogOutputFromString = LogOutput.File
        Case "ALL"
            GetLogOutputFromString = LogOutput.All
        Case Else
            GetLogOutputFromString = OutputConsole ' Default value for unvalid string log output
            Err.Raise vbObjectError + 1000, "VBALogger::GetLogOutputFromString", "Use invalid string to get a LogOutput ! (" & logOutputStr & ")."
    End Select
End Function


Private Function GetContextInfo(ByVal context As Variant) As String
    Dim output As String
    Dim i As Long
    
    ' Vérifier le type de context et générer le bon format
    If IsNull(context) Then
        output = "Null"
    ElseIf varType(context) = vbBoolean Then
        output = CStr(context)
    ElseIf varType(context) = vbByte Or varType(context) = vbInteger Or _
           varType(context) = vbLong Or varType(context) = vbSingle Or _
           varType(context) = vbDouble Or varType(context) = vbCurrency Then
        output = CStr(context)
    ElseIf varType(context) = vbString Then
        output = """" & context & """"
    ElseIf (varType(context) = vbArray Or IsArray(context)) Then
        output = "["
        For i = LBound(context) To UBound(context)
            output = output & """" & CStr(context(i)) & """, "
        Next i
        If Right(output, 2) = ", " Then
            output = Left(output, Len(output) - 2)
        End If
        output = output & "]"
    ElseIf varType(context) = vbObject Then
        If context Is Nothing Then
            output = "Nothing"
        ElseIf TypeOf context Is collection Then
            output = "Collection(#" & context.Count & ")"
            If context.Count > 0 Then
                output = output & vbCrLf & CollectionToJson(context)
            End If
        Else
            output = "Object: "
        End If
    Else
        output = "Unsupported type: " & TypeName(context)
    End If
    
    GetContextInfo = output
End Function


Private Function CollectionToJson(ByVal coll As collection) As String
    Dim json As String
    Dim item As Variant
    Dim isFirst As Boolean
    
    json = "["
    isFirst = True
    
    For Each item In coll
        If Not isFirst Then
            json = json & ", "
        End If
        
        If IsObject(item) Then
            json = json & ObjectToJSON(item)
        Else
            json = json & """" & CStr(item) & """"
        End If

        isFirst = False
    Next item
    json = json & "]"
    
    CollectionToJson = json
End Function


Private Function ObjectToJSON(ByVal obj As Object) As String
    Dim output As String
    Dim key As Variant
    Dim isFirst As Boolean
    
    output = "{"
    isFirst = True

    If TypeName(obj) = "Dictionary" Then
        For Each key In obj.Keys
            If Not isFirst Then output = output & ", "
            output = output & """" & key & """: """ & obj(key) & """"
            isFirst = False
        Next key
    Else
        On Error Resume Next
        For Each key In obj
            If Not isFirst Then output = output & ", "
            output = output & """" & key.Name & """: """ & key.value & """"
            isFirst = False
        Next key
        On Error GoTo 0
    End If
    
    output = output & "}"
    
    If output = "{}" Then
        output = "null"
    End If
    
    ObjectToJSON = output
End Function
