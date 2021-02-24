# Fred the Developer Design Document

Fred consists of four different classes, each written 100% in VBA and completely independent of the others. That reduces bloat in projects because you can take only what you need and leave the rest.

This is an overview of Fred's class and method structure. Click on one of the links below to go to more detailed use information.

**NOTE: Documentation is still under construction**

## fredDeveloper_Assert
- [Failed](Assert.md#failed)
- [Inconclusive](Assert.md#inconclusive)
- [IsEqual](Assert.md#isequal)
- [IsFalse](Assert.md#isfalse)
- [IsGreater](Assert.md#isgreater)
- [IsGreaterOrEqual](Assert.md#isgreaterorequal)
- [IsInconclusive](Assert.md#isinconclusive)
- [IsLess](Assert.md#isless)
- [IsLessOrEqual](Assert.md#islessorequal)
- [IsNotEqual](Assert.md#isequal)
- [IsTrue](Assert.md#istrue)
- [Successful](Assert.md#successful)
- [Report](Assert.md#report)

## fredDeveloper_ClientSettings
- [Add](ClientSettings.md#add)
- [Clear](ClientSettings.md#clear)
- [Count](ClientSettings.md#count)
- [Delete](ClientSettings.md#delete)
- [Exists](ClientSettings.md#exists)
- [Report](ClientSettings.md#report)

## fredDeveloper_Events
- [Log](Events.md#log)
- [LoggingFilePath](Events.md#loggingfilepath)
- [LoggingMode](Events.md#loggingmode)
- [RestoreDefaultSetting](Events.md#restoredefaultsetting)
- [Tic](Events.md#tic)
- [Toc](Events.md#toc)

## fredDeveloper_ModuleManager
- [CountAll](ModuleManager.md#countall)
- [CountClasses](ModuleManager.md#countclasses)
- [CountForms](ModuleManager.md#countforms)
- [CountModules](ModuleManager.md#countmodules)
- [Directory](ModuleManager.md#directory)
- [Export](ModuleManager.md#export)
- [Import](ModuleManager.md#import)


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

