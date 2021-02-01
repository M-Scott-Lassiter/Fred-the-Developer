# Fred the Developer Design Document

## Overview
This document details the custom Excel Workbook properties that Fred uses to make his developer tools work.

## Properties
These are stored under Excel's built in `ThisWorkbook.CustomDocumentProperties`. By storing values here, you can use them in your other modules to guide behavior while debugging (such as controlling data logging output).

If the Type column refers to a [custom Fred enumeration](#enumerations), then the property is of type `msoPropertyTypeNumber`. Microsoft defines the available property types on their [Office VBA documentation](https://docs.microsoft.com/en-us/office/vba/api/office.msodocproperties).

| Field Name | Document Property Type | Description | Default Value |
| :--------- | :----: | :---------- | :----: |
| `DebugLogging` | *msoPropertyTypeNumber* | Of type [*fredDebugLogMode*](#freddebuglogmode). If not set to `Disabled`, then during run time any debugging lines will write to the Immediate Window, an external file path as specified in `DebugExternalFileLoggingPath`, or both | `Disabled` |
| `DebugExternalFilePath` | *msoPropertyTypeString* | If `DebugLogging` set to an external mode, debug lines will write to the text file log specified within this parameter | `ThisWorkbook.Path` |

## Enumerations

### fredDebugLogMode

This enum is contained within the `fredDeveloper_Logging` class.

| Enum Name | Value | Description |
| :-------- | :---: | :---------- |
| `Disabled` 			|0| Fred will ignore calls to output a debugging line |
| `ToImmediateOnly`	  	|1| Debug lines will print to the Immediate Window only |
| `ToExternalOnly`	  	|2| Debug lines will print to an external file specified by [`DebugExternalFilePath`](#properties)|
| `ToImmediateAndExternal`	|3| Debug lines will print to both the Immediate Window *and* the external file |

## Class Structure

- fredDeveloper
  - fredDeveloper_Assert
    - AreEqual
    - AreNotEqual
    - IsTrue
    - IsFalse
    - IsGreater
    - IsGreaterOrEqual
    - IsLess
    - IsLessOrEqual
    - Inconclusive
    - CountSuccess
    - CountFailure
    - CountInconclusive
    - Report
  - fredDeveloper_ClientSettings
    - Add
    - Clear
    - ClientSettingExists
    - Delete
  - fredDeveloper_ModuleManager
    - CountAll
    - CountClasses
    - CountForms
    - CountModules
    - Directory
    - Export
    - Import
  - RestoreDefaultSetting
  - Log
  - LoggingMode
  - LoggingFilePath
  - Tic
  - Toc

