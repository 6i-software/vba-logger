VBA Logger
==========

[![Support me by offer me a coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-donate-informational.svg)](https://www.buymeacoffee.com/vincent.blain)


> A simple VBA logging utility for tracking messages with customizable log levels and output options. 


## Overview

**VBA Logger** is a logging system designed to track and record events or messages within a VBA (Visual Basic for Applications) application. It supports logging messages with different levels of severity and can output logs either to the VBA console or to a specific log file. It supports also allows for customization, such as adjusting log verbosity, customizing the log output location, and more.

![VBALogger-output-all.png](./documentation%2Fassets%2FVBALogger-output-all.png)
<small>*VBA Logger - Simultaneously output to the VBA Console (also known as Excel's Immediate Window) and to a log file.*</small>

**Main features:**

- **Verbosity levels**: Control the granularity of logs by adjusting the verbosity level, from critical errors to detailed debug information, according to the `LogVerbosityLevel` enumeration.


- **Each log entry is prefixed with the type of message** (*e.g.*, `[ERROR]`, `[WARNING]`, etc.), except for general "Log" entries, which have no prefix. This helps in quickly identifying the type and severity of the log entry.


- **Configure log output**: You can choose to send logs to the VBA console (Excel's immediate Windows), a file, or both. And also defines the log file's name and folder. The logger allows you to specify a custom log file name and storage folder. If the folder does not exist, the logger will create it recursively, ensuring that the file can be saved without manual folder creation.


- **Contextual logging to simplify debugging of complex variables**: At the trace level, you can add extra context to your logs, such as collections or objects. This approach improves debugging and issue tracking by providing deeper insights into variable values.


- **Add a customizable splashscreen**: A customizable splashscreen can be shown at the beginning of a logging session, providing a visual confirmation that logging is initialized. This can be turned off as needed.


### Verbosity levels

VBALogger allows logging at different verbosity levels to suit different priorities. The levels are defined in the `LogVerbosityLevelSession` property according to the `LogVerbosityLevel` enumeration, ranging from critical errors to detailed debug information:

| **Level** | **Description**                                      |
|-----------|------------------------------------------------------|
| Error     | Logs runtime errors.                                 |
| Warning   | Logs exceptional occurrences that are not errors.    |
| Log       | Normal logs for general purposes.                    |
| Notice    | Logs significant events.                             |
| Info      | Logs interesting events, useful for understanding.   |
| Trace     | Logs detailed debug information for troubleshooting. |


### Where logs are written

By default, log entries are shown only in VBA console (Excel's immediate Window). The log output can be configured based on the desired destination using the `LogOutputSession` property with the values of `LogOutput` enumeration. Logs can be sent to the VBA console, to a file, or to both:

| **Output option** | **Description**                               |
|-------------------|-----------------------------------------------|
| Console           | Logs are sent to the VBA immediate window.    |
| File              | Logs are written to a file.                   |
| All               | Logs are sent both to the console and a file. |

When you opt to log output to a file, the logs are saved by default in the `./var/log/` directory of the workbook path, and by using the pattern filename `logfile_2024-10-17.log`, but you can also change it.


## Documentation

Please refer to the [documentation](./documentation%2F1%20-%20Getting%20started.md) for details on how to install and use VBA Logger.


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

