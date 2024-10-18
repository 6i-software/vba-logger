VBA Logger
==========

[![Support me by offer me a coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-donate-informational.svg)](https://www.buymeacoffee.com/vincent.blain)


> A simple VBA logging utility for tracking messages with customizable log levels and output options. 


## Overview

**VBA Logger** is a logging system designed to track and record events or messages within a VBA (Visual Basic for Applications) application. It supports logging messages with different levels of severity and can output logs either to the VBA console or to a specific log file. It supports also allows for customization, such as adjusting log verbosity, customizing the log output location, and more.

![VBALogger-output-all.png](./documentation%2Fassets%2FVBALogger-output-all.png)
<small>*VBA Logger - Simultaneously output to the VBA Console (also known as Excel's Immediate Window) and to a log file.*</small>

**Main features:**

- **Verbosity levels**: Control the granularity of logs by adjusting the verbosity level, from critical errors to detailed debug information. The current verbosity level is set using the `LogVerbosityLevelSession` property, which corresponds to the values of the `LogVerbosityLevel` enumeration.

  | **Level** | **Description**                                      |
  |-----------|------------------------------------------------------|
  | Error     | Logs runtime errors.                                 |
  | Warning   | Logs exceptional occurrences that are not errors.    |
  | Log       | Normal logs for general purposes.                    |
  | Notice    | Logs significant events.                             |
  | Info      | Logs interesting events, useful for understanding.   |
  | Trace     | Logs detailed debug information for troubleshooting. |


- **Each log entry is prefixed with the type of message** (*e.g.*, `[ERROR]`, `[WARNING]`, etc.), except for general "Log" entries, which have no prefix. This helps in quickly identifying the type and severity of the log entry.


- **Configure log output**: By default, logs are displayed only in the VBA console (Excel's immediate window). You can change their destination using the `LogOutputSession` property according to the values of the `LogOutput` enumeration. 

  | **Log output** | **Description**                               |
  |----------------|-----------------------------------------------|
  | Console        | Logs are sent to the VBA immediate window.    |
  | File           | Logs are written to a file.                   |
  | All            | Logs are sent both to the console and a file. |  

  Logs can be sent to the VBA console, a file, or both. The logger also allows you to choose a name and folder for the log file. If the folder (and its subfolders) does not exist, the logger will create it automatically.


- **Contextual logging to simplify debugging of complex variables**: At the trace level, you can add extra context to your logs, such as collections or objects. This approach improves debugging and issue tracking by providing deeper insights into variable values.


- **Add a customizable splashscreen**: A customizable splashscreen can be shown at the beginning of a logging session, providing a visual confirmation that logging is initialized. This can be turned off as needed.



## Documentation

Please refer to the documentation for details on how to use VBA Logger and its features.

- [Getting started with VBA Logger](documentation%2F1%20-%20Getting%20started.md)
- [API documentation](documentation%2F2%20-%20API%20documentation.md)



## Installation

VBA Logger supports several installation methods :

- by importing the class module VBALoggerClass in your VBA project. 
    
    > In this case, the class module `VBALoggerClass` is set to `Private` and can only be accessed within the project where it is defined or imported, ensuring encapsulation and preventing external code from directly interacting with it. To utilize the logger functionality, the developer must instantiate the class within their own procedures, allowing for controlled access to its methods and properties.
  

- by reference the XLAM VBALogger in your VBA project.

    > If you want to use the logger into multilples VBA project, you should use the VBA Logger XLAM. In this case, the class module `VBALoggerClass` is set to `PublicNotCreatable` and the developer use a factory method `VBALogger.Factory.Create` to obtain an instance of the logger.


- *by using the installer (Windows setup) - ***not yet, working in progress****

    > Use this solution in order to make easy installation and deployment of VBA Logger XLAM.
  
You can find the detailed instructions for each installation method in the documentation of [VBALogger Installation](documentation%2F1%20-%20Getting%20started.md#installation). You should consult it for the simplest setup experience.



## About

### Support

**VBA Logger** is free and open source under the [MIT License](./LICENSE), but if you want to support me, you can [offer me a coffee here](https://www.buymeacoffee.com/vincent.blain) or by scanning this QR code.

<img alt="Buy me a coffee ?" src="./documentation%2Fassets%2Fv20100v_buy-me-a-coffee_qrcode.png" width="300" height="300" />


### Contributing

Bug reports, reports a typo in documentation, comments, pull-request & Github stars are always welcome !


### Releases

- VBA Logger v1.0.0 - 2024.10.17


### License

Release under [MIT License](./LICENSE),<br/>
Copyright (c) 2024 by 2o1oo vb20100bv@gmail.com

