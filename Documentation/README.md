# Fred the Developer Design Document

Fred consists of four different classes, each written 100% in VBA and completely independent of the others. That reduces bloat in projects because you can take only what you need and leave the rest.

This is an overview of Fred's class and method structure. Click on one of the links below to go to more detailed use information.

## fredDeveloper_Assert
- AreEqual
- AreNotEqual
- Failed
- Inconclusive
- IsTrue
- IsFalse
- IsGreater
- IsGreaterOrEqual
- IsLess
- IsLessOrEqual
- IsInconclusive
- Successful
- Report

## fredDeveloper_Events
- Log
- LoggingMode
- LoggingFilePath
- RestoreDefaultSetting
- Tic
- Toc

## fredDeveloper_ClientSettings
- Add
- Clear
- Count
- Delete
- Exists
- Report

## fredDeveloper_ModuleManager
- CountAll
- CountClasses
- CountForms
- CountModules
- Directory
- Export
- Import


# Properties
These are stored under Excel's built in `ThisWorkbook.CustomDocumentProperties` when using the `fredDeveloper_Events` class. By storing values here, the set values persist not only between VBA run calls, but remain even if you close and reopen your workbook (assuming you saved, of course). Fred uses these behind the scenes when doing tasks such as logging debugging output.

If the Type column refers to a [custom Fred enumeration](#enumerations), then the property is of type `msoPropertyTypeNumber`. Microsoft defines the available property types on their [Office VBA documentation](https://docs.microsoft.com/en-us/office/vba/api/office.msodocproperties).

| Field Name | Document Property Type | Description | Default Value |
| :--------- | :----: | :---------- | :----: |
| `DebugLogging` | *msoPropertyTypeNumber* | Of type [*fredDebugLogMode*](#freddebuglogmode). If not set to `Disabled`, then during run time any debugging lines will write to the Immediate Window, an external file path as specified in `DebugExternalFileLoggingPath`, or both | `Disabled` |
| `DebugExternalFilePath` | *msoPropertyTypeString* | If `DebugLogging` set to an external mode, debug lines will write to the text file log specified within this parameter | `ThisWorkbook.Path` |

# Enumerations


## fredDebugLogMode

This enum is contained within the `fredDeveloper` class.

| Enum Name | Value | Description |
| :-------- | :---: | :---------- |
| `Disabled` 			|0| Fred will ignore calls to output a debugging line |
| `ToImmediateOnly`	  	|1| Debug lines will print to the Immediate Window only |
| `ToExternalOnly`	  	|2| Debug lines will print to an external file specified by [`DebugExternalFilePath`](#properties)|
| `ToImmediateAndExternal`	|3| Debug lines will print to both the Immediate Window *and* the external file |


## fredEventSettings

This enum is contained within the `fredDeveloper` class. This is used exclusively by the `RestoreDefaultSetting` method to limit the script to only the values in CustomDocumentProperties from `fredDeveloper_Events`. No others can be affected.

| Enum Name | Value | Description |
| :-------- | :---: | :---------- |
| `DebugLogging` 		|0| |
| `DebugExternalFilePath`  	|1| |

