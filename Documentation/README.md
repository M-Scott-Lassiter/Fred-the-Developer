# Fred the Developer Design Document

## How This Documentation is Organized

Fred has several different types of documentaiton.

- For detailed usage guidance on each of Fred's classes and methods, see the [Reference Guide](ReferenceGuide.md).
- (IN PROGRESS) For example usage, see Tutorials.
- The rest of this document details the custom Excel Workbook properties that Fred uses to make his developer tools work.

## Properties
These are stored under Excel's built in `ThisWorkbook.CustomDocumentProperties`. By storing values here, the set values persist not only between VBA run calls, but remain even if you close and reopen your workbook (assuming you saved, of course). Fred uses these behind the scenes when doing tasks such as logging debugging output.

If the Type column refers to a [custom Fred enumeration](#enumerations), then the property is of type `msoPropertyTypeNumber`. Microsoft defines the available property types on their [Office VBA documentation](https://docs.microsoft.com/en-us/office/vba/api/office.msodocproperties).

| Field Name | Document Property Type | Description | Default Value |
| :--------- | :----: | :---------- | :----: |
| `DebugLogging` | *msoPropertyTypeNumber* | Of type [*fredDebugLogMode*](#freddebuglogmode). If not set to `Disabled`, then during run time any debugging lines will write to the Immediate Window, an external file path as specified in `DebugExternalFileLoggingPath`, or both | `Disabled` |
| `DebugExternalFilePath` | *msoPropertyTypeString* | If `DebugLogging` set to an external mode, debug lines will write to the text file log specified within this parameter | `ThisWorkbook.Path` |

## Enumerations

### fredDebugLogMode

This enum is contained within the `fredDeveloper` class.

| Enum Name | Value | Description |
| :-------- | :---: | :---------- |
| `Disabled` 			|0| Fred will ignore calls to output a debugging line |
| `ToImmediateOnly`	  	|1| Debug lines will print to the Immediate Window only |
| `ToExternalOnly`	  	|2| Debug lines will print to an external file specified by [`DebugExternalFilePath`](#properties)|
| `ToImmediateAndExternal`	|3| Debug lines will print to both the Immediate Window *and* the external file |

## Class Structure

This is a high level overview of Fred's class and method structure. For more detailed information, see the [Reference Guide](ReferenceGuide.md).

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

