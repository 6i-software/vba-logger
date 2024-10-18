VBA Logger
==========


## API Documentation

### Members

| Member                     | Type     | Description                                                                                                                                                                                                                                                          |
|----------------------------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `LogVerbosityLevel`        | Enum     | - `LevelError`: Logs runtime errors.<br/>- `LevelWarning`: Logs warnings.<br/> - `LevelLog`: Normal log level.<br/> - `LevelNotice`: Logs significant events.<br/> - `LevelInfo`: Logs interesting events.<br/> - `LevelTrace`: Logs detailed debugging information. |
| `LogOutput`                | Enum     | - `Console`: Output logs to the VBA Immediate Window. <br/> - `File`: Write logs to a file.<br/> - `All`: Log to both the console and a file.                                                                                                                        |
| `LogVerbosityLevelSession` | Property | The current session verbosity level.                                                                                                                                                                                                                                 |
| `LogOutputSession`         | Property | Where logs are output (console, file, or both).                                                                                                                                                                                                                      |
| `LogFileName`              | Property | Name of the log file.                                                                                                                                                                                                                                                |
| `LogFileFolder`            | Property | Folder where the log file is stored.                                                                                                                                                                                                                                 |
| `LogFilePath`              | Property | Full path where the log file is stored.                                                                                                                                                                                                                              |
| `OptionShowSplashscreen`   | Property | Enable or disable a splashscreen message to display at the session start.                                                                                                                                                                                            |
| `SplashscreenContent`      | Property | Content of the splash screenmessage.                                                                                                                                                                                                                                 |



### Methods public

#### Public Sub Initialize

```vba
Public Sub Initialize( _
    Optional ByVal paramSplashscreen As String = "", _
    Optional ByVal paramLogOutputSession As LogOutput = LogOutput.Console, _
    Optional ByVal paramLogFileName As String = "", _
    Optional ByVal paramLogFileFolder As String = "" _
)
```
**Description**: Initializes the logger with optional parameters such as splash screen, log output type, log file name, and log file folder.

**Parameters**:
 - `paramSplashscreen` (String): Custom message to display as a splash screen.
 - `paramLogOutputSession` (LogOutput): Determines where to log output (Console, File, All).
 - `paramLogFileName` (String): Name of the log file.
 - `paramLogFileFolder` (String): Path to the folder where the log file will be saved.

**Example**:

```vba
Sub Create_and_configure_VBALogger()
    ' Create a new instance of VBALogger and configure settings.
    Dim Logger As VBALoggerClass
    Set Logger = New VBALoggerClass
    
    ' Prepapre a custom splashscreen
    Dim splashscreen As String
    splashscreen = splashscreen & "-----------------------" & vbCrLf
    splashscreen = splashscreen & "    Start VBALogger    " & vbCrLf
    splashscreen = splashscreen & "-----------------------" & vbCrLf    

    ' Initialize logger with custom configuration:
    Call Logger.Initialize( _ 
        splashscreen, _ 
        LogOutput.All, _ 
        "CustomLogfile.log", 
        ThisWorkbook.Path & "\another\log\folder" _
    )

    ' Log messages
    Logger.Info "This is an informational message."
    Logger.Trace "This is a debug message, useful for troubleshooting."
    Logger.Log "This is a normal log message."
End Sub

' <<< Output results >>>
'
' -----------------------
'     Start VBALogger    
' -----------------------
'
' [INFO]    | 2024-10-17 11:38:01 | This is an informational message.
' [TRACE]   | 2024-10-17 11:38:01 | This is a debug message, useful for troubleshooting.
' This is a normal log message.
```

---

#### Public Sub Error

```vba
Public Sub Error(ByVal message As String)
```

**Description**: Logs an error message.

**Parameters**:
- `message` (String): The error message to log.

**Example**:

```vba
Logger.Error "An unexpected error has occurred."
```

---

#### Public Sub Warning

```vba
Public Sub Warning(ByVal message As String)
```

**Description**: Logs a warning message.

**Parameters**:
- `message` (String): The warning message to log.

**Example**:

```vba
Logger.Warning "This is a warning about potential issues."
```

---

#### Public Sub Notice

```vba
Public Sub Notice(ByVal message As String)
```

**Description**: Logs a notice message.

**Parameters**:
- `message` (String): The notice message to log.

**Example**:

```vba
Logger.Notice "This is a notice for your attention."
```

---

#### Public Sub Info

```vba
Public Sub Info(ByVal message As String)
```

**Description**: Logs an informational message.

**Parameters**:
- `message` (String): The informational message to log.

**Example**:

```vba
Logger.Info "The process completed successfully."
```

---

#### Public Sub Trace

```vba
Public Sub Trace(ByVal message As String, Optional ByVal context As Variant)
```

**Description**: Logs a trace message for detailed debugging.

**Parameters**:
- `message` (String): The trace message to log.
- `context` (Variant): Optional context information related to the message.

**Example**:

```vba
Sub trace_log_entries_with_context_values()
    Dim Logger As VBALoggerClass
    Set Logger = New VBALoggerClass

    ' Testing with simple type context
    Logger.Trace "Testing with a nothing context.", Nothing
    Logger.Trace "Testing with a null context.", Null
    Logger.Trace "Testing with a boolean context (True).", True
    Logger.Trace "Testing with a boolean context (False).", False
    Logger.Trace "Testing with a number context (Integer).", 42
    Logger.Trace "Testing with a number context (Double).", 3.14
    Logger.Trace "Testing with a string context.", "Hello, World!"
    
    ' Testing with an array context
    Dim testArray(0 To 2) As Variant
    testArray(0) = "The first"
    testArray(1) = 2
    testArray(2) = 3.14
    Logger.Trace "Testing with an array context.", testArray
    
    ' Testing with a simple collection
    Dim testCollection1 As New collection
    testCollection1.Add 1
    testCollection1.Add 2
    testCollection1.Add 3.14
    Logger.Trace "Testing with a simple collection context.", testCollection1
    
    ' Testing with a hybrid collection
    Dim testCollection2 As New collection
    testCollection2.Add "Item 1"
    testCollection2.Add 2
    testCollection2.Add 3.14
    Dim myObject As Object
    Set myObject = CreateObject("Scripting.Dictionary") ' Utiliser un dictionnaire comme exemple d'objet
    myObject.Add "Name", "Bob"
    myObject.Add "Age", 24
    myObject.Add "Value", 3.14
    testCollection2.Add myObject
    Logger.Trace "Testing with a hybrid collection context.", testCollection2
End Sub

' <<< Output results >>>
'
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a nothing context. | context=Nothing
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a null context. | context=Null
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a boolean context (True). | context=Vrai
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a boolean context (False). | context=Faux
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a number context (Integer). | context=42
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a number context (Double). | context=3,14
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a string context. | context="Hello, World!"
' [TRACE]   | 2024-10-17 13:20:36 | Testing with an array context. | context=["The first", "2", "3,14"]
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a simple collection context. | context=Collection(#3)
' ["1", "2", "3,14"]
' [TRACE]   | 2024-10-17 13:20:36 | Testing with a hybrid collection context. | context=Collection(#4)
' ["Item 1", "2", "3,14", {"Name": "Bob", "Age": "24", "Value": "3,14"}]
```

---

#### Public Sub Log

```vba
Public Sub Log(ByVal message As String, Optional ByVal withNewLine As Boolean = True)
```

**Description**: Logs a generic message.

**Parameters**:
- `message` (String): The message to log.
- `withNewLine` (Boolean): Indicates whether to add a new line after the message (default is True).

**Example**:

```vba
Sub test_newline()
    Dim Logger As VBALoggerClass
    Set Logger = New VBALoggerClass
    
    Logger.Log "A log message."
    
    Logger.Log "This is ", False
    Logger.Log "a log ", False
    Logger.Log "message ", False
    Logger.Log "output on the same line."
    
    Logger.Log "Antoher message !"
End Sub

' <<< Output results >>>
'
' A log message.
' This is a log message output on the same line.
' Antoher message !
```

---

#### Public Function ToString

```vba
Public Function ToString() As String
```

**Description**: Returns a string representation of the logger instance, including current settings.

**Returns**: A formatted string containing representation fo the logger instance with key property values for debugging.

**Example**:

```vba
Dim loggerDetails As String
loggerDetails = Logger.ToString()

Debug.Print loggerDetails

' <<< Output results >>>
'
' VBALogger Instance
' ------------------
' Log verbosity level session: WARNING
' Log output session: Console
' Log file name: 
' Log file path: 
' Log file folder: 
' Option show splashscreen: Faux
' Total messages logged: 6
' Object Type: VBALoggerClass
' Memory Address: 2611909723304
' ------------------
```

---


### Methods private

#### Private Sub PrepareLogEntry

```vba
Private Sub PrepareLogEntry(ByVal paramLogVerbosityLevel As LogVerbosityLevel, ByVal paramMessage As String, Optional ByVal context As Variant)
```
**Description**: Prepares a log entry with the specified verbosity level and message.

**Parameters**:
- `paramLogVerbosityLevel` (LogVerbosityLevel): The verbosity level of the log entry.
- `paramMessage` (String): The message to log.
- `context` (Variant): Optional context information.

---

#### Private Sub WriteLogEntry

```vba
Private Sub WriteLogEntry(ByVal paramLogEntry As String, Optional ByVal withNewLine As Boolean = True)
```
**Description**: Writes a log entry to the specified output (console or file).

**Parameters**:
- `paramLogEntry` (String): The log entry to write.
- `withNewLine` (Boolean): Indicates whether to add a new line (default is True).

---

#### Private Sub WriteIntoLogFile

```vba
Private Sub WriteIntoLogFile(ByVal paramLogEntry As String, ByVal withNewLine As Boolean)
```
**Description**: Writes a log entry to the log file.

**Parameters**:
- `paramLogEntry` (String): The log entry to write to the file.
- `withNewLine` (Boolean): Indicates whether to add a new line.

---

#### Private Function GetLogVerbosityLevelHumanReadable

```vba
Private Function GetLogVerbosityLevelHumanReadable(ByVal paramLogVerbosityLevel As LogVerbosityLevel) As String
```
**Description**: Returns a human-readable string representation of the specified verbosity level.

**Parameters**:
- `paramLogVerbosityLevel` (LogVerbosityLevel): The verbosity level to convert.

**Returns**: A string representation of the verbosity level.

---

#### Private Function GetLogOutputHumanReadable

```vba
Private Function GetLogOutputHumanReadable(ByVal paramLogOutput As LogOutput) As String
```
**Description**: Returns a human-readable string representation of the specified log output type.

**Parameters**:
- `paramLogOutput` (LogOutput): The log output type to convert.

**Returns**: A string representation of the log output type.

---

#### Private Function GetLogVerbosityLevelFromString

```vba
Private Function GetLogVerbosityLevelFromString(ByVal logVerbosityLevelInString As String) As LogVerbosityLevel
```
**Description**: Converts a string representation of a verbosity level to its corresponding enum value.

**Parameters**:
- `logVerbosityLevelInString` (String): The string representation of the verbosity level.

**Returns**: The corresponding LogVerbosityLevel enum value.

---

#### Private Function GetLogOutputFromString

```vba
Private Function GetLogOutputFromString(ByVal logOutputStr As String) As LogOutput
```
**Description**: Converts a string representation of a log output type to its corresponding enum value.

**Parameters**:
- `logOutputStr` (String): The string representation of the log output type.

**Returns**: The corresponding LogOutput enum value.

---

#### Private Function GetContextInfo

```vba
Private Function GetContextInfo(ByVal context As Variant) As String
```
**Description**: Generates a string representation of the context information provided.

**Parameters**:
- `context` (Variant): The context information to format.

**Returns**: A string representation of the context information.

---

#### Private Function CollectionToJson

```vba
Private Function CollectionToJson(ByVal coll As collection) As String
```
**Description**: Converts a VBA collection to a JSON string representation.

**Parameters**:
- `coll` (Collection): The collection to convert.

**Returns**: A JSON string representation of the collection.

---

#### Private Function ObjectToJSON

```vba
Private Function ObjectToJSON(ByVal obj As Object) As String
```
**Description**: Converts a VBA object to a JSON string representation.

**Parameters**:
- `obj` (Object): The object to convert.

**Returns**: A JSON string representation of the object.

---