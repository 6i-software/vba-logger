VBA Logger
==========

> Welcome to the VBA Logger documentation ! This page will provide you with an introduction to this logger, including installation instructions and usage guidelines.


## Getting started

### Prerequisites

Before getting started with **VBA Logger**, ensure the following prerequisites are met:

1. **Excel version** : You should have Microsoft Excel installed on your system, preferably Excel 2010 or newer, as older versions may not fully support this module.


2. **VBA enabled in your workbook**: Your workbook should support VBA macros, such as `.xlsm` (macro-enabled workbook) or `.xlsb` (binary workbook with macros). And make sure VBA macros are enabled in your Excel settings to execute the tests.

   > - Go to Trust Center Settings *File* > *Options* > *Trust Center* > *Trust Center Settings*.
   > - Under Macro Settings, select Disable all macros with notification (recommended) or Enable all macros (for full access).



### Installation 

#### Download VBA Logger

First, obtain the VBA Logger setup file (not yet available), clone the project or download archive zip from the repository. You should get:

 - `6i_VBALogger.xlam`: This is the Excel Macro-Enabled Add-In file of VBA Logger that developers can load through Excel’s settings. Once referenced, it grants access to logger that can be utilized across all VBA projects.


 - `VBALoggerClass.cls`: Developers can import this class module directly into their existing VBA projects. 


#### Install "VBA Logger" by importing the class module

Note that in this scenario, the **VBALoggerClass** has been set to `private` and imported as a class module into the VBA project of the developer who intends to use this logger. With this installation, the class can only be accessed within the project where it is defined or imported, ensuring encapsulation and preventing external code from directly interacting with it. To utilize the logging functionality, the developer must instantiate the class within their own procedures, allowing for controlled access to its methods and properties.

1. Launch Excel and open your workbook (typically `.xlsm` or `.xlsb`) where you wish to add logger. Access the Visual Basic for Application Editor by pressing <kbd>Alt + F11</kbd>.

   > If you don’t already have the **Developer** tab in Excel:
   > - Go to *File* > *Options*.
   > - In the *Excel Options* window, select *Customize Ribbon*.
   > - Check the box for *Developer* and click *OK*.


2. In the VBA Editor, import the class module **VBALoggerClass.cls**
   
   > - Go to *File* > *Import File* 
   > - Select the **VBALoggerClass.cls** class modules to import into your VBA project.
   > - Ensure that the class module visibility is set to `Private`.
   > 
   >   ![import_VBALoggerClass.png](assets%2Fimport_VBALoggerClass.png)

   That's it ! You are now ready to use **VBALoggerClass**. 


#### Install "VBA Logger" as a XLAM reference

But you have a second scenario available, when the **VBALoggerClass** is configured as `PublicNotCreatable`. In this case, the developer use a factory method to obtain an instance of the logger.


1. Launch Excel and open your workbook (typically `.xlsm` or `.xlsb`) where you wish to add logger. Access the Visual Basic for Application Editor by pressing <kbd>Alt + F11</kbd> and click in the VBA editor menu on *Tools* > *References*.


2. Add the `6i_VBALogger.xlam` file as a Reference in your VBA Project

   ![reference_VBALogger_xlam.png](assets%2Freference_VBALogger_xlam.png)

   You should see this reference in the VBA editor:

   ![see_reference_VBALogger_xlam.png](assets%2Fsee_reference_VBALogger_xlam.png)

   That's it ! You are now ready to use the **VBA Logger** by its factory method.


## Usages

### Verbosity levels

VBALogger allows logging at different verbosity levels to suit different priorities. The levels are defined in the `LogVerbosityLevelSession` property according to the `LogVerbosityLevel` enumeration, ranging from critical errors to trace (a.k.a. detailed debug) information:

| **Level** | **Value** | **Description**                                      |
|-----------|-----------|------------------------------------------------------|
| Error     | -2        | Logs runtime errors.                                 |
| Warning   | -1        | Logs exceptional occurrences that are not errors.    |
| Log       | 0         | Normal logs for general purposes.                    |
| Notice    | 1         | Logs significant events.                             |
| Info      | 2         | Logs interesting events, useful for understanding.   |
| Trace     | 3         | Logs detailed debug information for troubleshooting. |


### Where logs are written

By default, log entries are shown only in VBA console (Excel's immediate Window). The log output can be configured based on the desired destination using the `LogOutputSession` property with the values of `LogOutput` enumeration. Logs can be sent to the VBA console, to a file, or to both:

| **Output option** | **Value** | **Description**                               |
|-------------------|-----------|-----------------------------------------------|
| Console           | 1         | Logs are sent to the VBA immediate window.    |
| File              | 2         | Logs are written to a file.                   |
| All               | 3         | Logs are sent both to the console and a file. |

When you opt to log output to a file, the logs are saved by default in the `./var/log/` directory of the workbook path, and by using the pattern filename `logfile_2024-10-17.log`.


### Logging a message

To log a message, you need to create a logger as a new instance of `VBALoggerClass`, and this depends on the type of installation you have chosen: either by importing the class module or by referencing the VBALogger XLAM.


#### Instantiate a logger when installing by importing the VBALoggerClass module

If you install "**VBA Logger**" by importing the class module in your VBA project, so you can use its constructor directly, like in this below example. Note that in this scenario, the class module **VBALoggerClass** is set to `private` and can only be accessed within the project where it is defined or imported, ensuring encapsulation and preventing external code from directly interacting with it.

```vba
Sub Use_class_to_create_default_VBALogger() 
    Dim Logger As VBALoggerClass

   ' Create a new instance of VBALogger with default settings (output to console).
    Set Logger = New VBALoggerClass

    ' Log a normal message
    Logger.Log "This is a normal log message."

    ' Log messages at different verbosity levels
    Logger.Error "This message indicates an error."
    Logger.Warning "This is a warning message."
    Logger.Notice "This is a significant event notice."
    Logger.Info "This is an informational message."
    Logger.Trace "This is a debug message, useful for troubleshooting."
    
    ' Adjust verbosity level to Warning and log another message
    Logger.LogVerbosityLevelSession = LevelWarning
    Logger.Info "This message won't be logged due to the verbosity settings."

    ' Debug VBALogger instance
    Debug.Print Logger.ToString       
End Sub

' <<< Output results >>>
'
' [ERROR]   | 2024-10-17 11:33:20 | This message indicates an error. e
' [WARNING] | 2024-10-17 11:33:20 | This is a warning message.
' [NOTICE]  | 2024-10-17 11:33:20 | This is a significant event notice.
' [INFO]    | 2024-10-17 11:33:20 | This is an informational message.
' [TRACE]   | 2024-10-17 11:33:20 | This is a debug message, useful for    troubleshooting.
' This is a normal log message.
'
' Debug VBALogger instance
' ------------------------
' VBALoggerClass version: 1.0.0
' Log verbosity level session: 'WARNING'
' Log output session: 'Console'
' Log file name: ''
' Log file path: ''
' Log file folder: ''
' Option show splashscreen: False
' Total messages logged: 6
' Object Type: 'VBALoggerClass'
' Instantiation source: 'Direct'
' Memory Address: 1627419241072
' ------------------------
```

You can see result in the VBA console.

![output_result_in_VBA_console.png](assets%2Foutput_result_in_VBA_console.png)

> **Note**: By default, log entries are shown only in VBA console (Excel's immediate Window). If it is not visible go to the menu and select *View* > *Immediate Window*. Alternatively, you can press <kbd>Ctrl + G</kbd> to quickly open the VBA console.



#### Instanciate in case of installation by importing the class module VBALoggerClass

If you install "**VBA Logger**" as a XLAM reference in your VBA project, you can not use its constructor directly. In this case, the developer use a factory method to obtain an instance of the logger because the **VBALoggerClass** is configured as `PublicNotCreatable`. 

This is accomplished by calling the Create method from the factory of VBALogger module, as shown in the example below:

```vba
Sub Use_factory_to_create_default_VBALogger()
   ' Use reference to fetch the type of VBALoggerClass   
   Dim Logger As VBALogger.VBALoggerClass
   
   ' Create a new instance of VBALogger by its factory
   Set Logger = VBALogger.Factory.Create

    ' Log a normal message
    Logger.Log "This is a normal log message."
   
    ' Log messages at different verbosity levels
    Logger.Error "This message indicates an error."
    Logger.Warning "This is a warning message."
    Logger.Notice "This is a significant event notice."
    Logger.Info "This is an informational message."
    Logger.Trace "This is a debug message, useful for troubleshooting."
       
    ' Adjust verbosity level to Warning and log another message
    Logger.LogVerbosityLevelSession = LevelWarning
    Logger.Info "This message won't be logged due to the verbosity settings."
    
    ' Debug VBALogger instance
    Debug.Print Logger.ToString
End Sub

' <<< Output results >>>
'
' [ERROR]   | 2024-10-17 11:43:30 | This message indicates an error. e
' [WARNING] | 2024-10-17 11:43:30 | This is a warning message.
' [NOTICE]  | 2024-10-17 11:43:30 | This is a significant event notice.
' [INFO]    | 2024-10-17 11:43:30 | This is an informational message.
' [TRACE]   | 2024-10-17 11:43:30 | This is a debug message, useful for    troubleshooting.
' This is a normal log message.
'
' Debug VBALogger instance
' ------------------------
' VBALoggerClass version: 1.0.0
' Log verbosity level session: 'WARNING'
' Log output session: 'Console'
' Log file name: ''
' Log file path: ''
' Log file folder: ''
' Option show splashscreen: False
' Total messages logged: 6
' Object Type: 'VBALoggerClass'
' Instantiation source: 'Factory'
' Memory Address: 2697303585320
' ------------------------
```

You can see result in the VBA console.
![output_result_in_VBA_console_from_factory.png](assets%2Foutput_result_in_VBA_console_from_factory.png)

> **Note**: At the end of this example, we use the `ToString` method, which provides a representation of the logger instance along with key property values for debugging purposes. You will also notice that the instantiation source correctly shows the value '*Factory*'.

Here we go !



### Configuration the logger

The following procedure demonstrates how to set up a custom configuration for the VBA Logger. This configuration enables a variety of functionalities that enhance the logging experience.

- Add a Splashscreen.
- Configure log output to both VBA console and file.
- Change the name of the log file.
- Change the path ot the log storage. Folders are automatically creating, if they do not exist.

```vba
Sub Custom_configuration_VBALogger()
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
        "CustomLogfile.log", _
        ThisWorkbook.Path & "\another\log\folder" _
    )

    ' Log messages
    Logger.Info "This is an informational message."
    Logger.Trace "This is a debug message, useful for troubleshooting."
    Logger.Log "This is a normal log message."
    
    Debug.Print Logger.ToString
End Sub
```


### Debug instance of VBALoggerClass

VBALoggerClass implements a `ToString` method, in order to output a string representation of the logger instance with key property values for debugging.

```vba
Sub Test_ToString_VBALogger()
    ' Create a new instance of VBALogger and configure settings.
    Dim Logger As VBALogger.VBALoggerClass
    Set Logger = VBALogger.Factory.Create

    ' Log messages
    Logger.Info "This is an informational message."
    Logger.Trace "This is a debug message, useful for troubleshooting."
    Logger.Log "This is a normal log message."
    
    ' ToString instance
    Debug.Print Logger.ToString
End Sub
```
