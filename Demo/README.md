# Fred the Developer Demo

This directory houses the demo [Demo.xlsm](Demo.xlsm) macro-enabled excel workbook to demo the features of the Fred-the-Developer ecosystem of VBA class modules.

## Overview

[Demo.xlsm](Demo.xlsm) contains all four of the [original class modules](../Source%20Code), renamed for simplicity, and an additional module [modDemo](src/modDemo.bas) to run the examples displaying the functionality for each class module. See the [src](src) folder (generated using the [Module Manager Class Module](src/clsModuleManager.cls)):

- [src/](src/):
  - Class Modules:
    - [clsAssert.cls](src/clsAssert.cls): Assertion Class Module
    - [clsClientSettings.cls](src/clsClientSettings.cls): Client Settings Class Module
    - [clsEvents.cls](src/clsEvents.cls): Events Class Module
    - [clsModuleManager.cls](src/clsModuleManager.cls): Module Manager (Import/Export) Class Module
  - Demo Module:
    - [modDemo.bas](src/modDemo.bas): Module showcasing each class module with a respective demonstration sub-procedure.

## Demos

### DemoUnitTesting

```visual basic
Sub DemoUnitTesting()
    'Show an example use case for Test Driven Development
    
    Dim Assert As New clsAssert
    
    Assert.IsEqual 1, 1, "Num1IsEqualToNum1"
    Assert.IsEqual 1, "1", "String1EqualsNum1"         'Should fail, does not implicitly type convert
    Assert.IsNotEqual 1, 2, "Num1NotEqualToNum2"
    Assert.IsNotEqual 1, "1", "Num1NotEqualToString1"  'Should pass
    
    Assert.IsFalse 1 = 2, "FalseComparisonTest"
    Assert.IsTrue 1 = 1, "TrueComparisonTest"
    
    Assert.IsGreater 5, 1, "Num5IsGreaterThanNum1"
    Assert.IsGreaterOrEqual 6, 6, "Num6IsGreaterThanOrEqualToNum6"
    Assert.IsLess 4, 9, "Num5IsLessThanNum1"
    Assert.IsLessOrEqual 8, 8, "Num8IsLessThanOrEqualToNum8"
    
    Assert.IsInconclusive "ThisTestIsInconclusive"      'Flagged as inconlcusive to remind us to come back later
    
    Debug.Print Assert.Report
    
End Sub
```

### DemoClientSettings

```visual basic
Sub DemoClientSettings()

  'Show how to clear and add client settings, then run a report to easily show you what settings are available
    
    Dim ClientSettings As New clsClientSettings
    Dim cdp As Object
    
    Set cdp = ThisWorkbook.CustomDocumentProperties
    
    ClientSettings.Clear
    Debug.Print "Settings: " & ClientSettings.Count
    
    ClientSettings.Add "MyBooleanSetting", msoPropertyTypeBoolean, True
    ClientSettings.Add "MyDateSetting", msoPropertyTypeDate, Now
    ClientSettings.Add "MyNumberSetting", msoPropertyTypeNumber, 1245
    ClientSettings.Add "MyStringSetting", msoPropertyTypeString, "Test string"
    ClientSettings.Add "MyFloatSetting", msoPropertyTypeFloat, 3.14159
    
    Debug.Print "MyFloatSetting Exists: " & ClientSettings.Exists("MyFloatSetting") & ", and its value is " & cdp("MyFloatSetting")
    Debug.Print "Settings: " & ClientSettings.Count
    
    ClientSettings.Report
End Sub
```

### DemoEvents

```visual basic
Sub DemoEvents()

   Dim Events As New clsEvents
    Dim i As Long
       
    Events.LoggingFilePath = ThisWorkbook.Path & "\Example Debug Log.txt"
    Events.LoggingMode = ToImmediateAndExternal
    
    Events.Tic
    Do While i < 10000000
        i = i + 1
    Loop
    
    Events.Log ("10,000,000 Iteration Loop complete. Run time: " & Events.Toc & " seconds.")
End Sub
```

### DemoModuleManager

```visual basic
Sub DemoModuleManager()

'Show how to export the four class modules in Fred to a directory on your computer
    
    Dim ModuleManager As New clsModuleManager
    
    With ModuleManager
        .Directory = ThisWorkbook.Path & "\src"
        .Export "modDemo"
        .Export "clsAssert"
        .Export "clsClientSettings"
        .Export "clsEvents"
        .Export "clsModuleManager"
    End With
    
End Sub

````

