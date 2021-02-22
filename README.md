# Fred the Developer

Fred enhances your projects with a standalone VBA development framework that provides logging, debugging, source control, and Test Driven Development tools written in 100% VBA.


## About the Project

I built this project to help VBA developers

1) Output debugging or status lines to an external log file with a single line of code
2) Add or remove Custom Document Properties to enable persistent variables between VBA run time executions or Workbook Close/Open events without writing to spreadsheets
3) Write unit tests using various `Assert` methods
4) Export and Import code modules, forms, or classes from within another VBA script

Why 100% in VBA? Because my entire professional career has revolved around scant computer resources. The one ubiquitous constant was the Microsoft Office suite. If you are using a computer with Windows, chances are you have Office. While fully featured tools such as [RubberDuck VBA](https://rubberduckvba.com/) provides robust capabilities, we don't all have the luxury to install add-ins without system administrators permission.

Because I built Fred from 100% VBA class modules, you can simply add these modules to any macro enabled workbook, document, presentation, or more to flesh out your own robustly developed stand alone product.


## What does FRED stand for?
Fredrick... not everything has to be an acronym.


## Requirements

- Microsoft Office 2007 or newer (Not tested for earlier versions)
- A macro enabled file

You don't need to make any additional references. Fred makes references to the "Microsoft Scripting Runtime" and "Microsoft Visual Basic For Applications Extensibility 5.3" libraries. However, both are late bound to prevent end users from having to enable the references on each machine they use.

**Note:** the functionality in the `ModuleManager` class requires trusted access to the VBA project object model. See the Getting Started section below for a walkthrough to enable this. Note that your system administrator may have disabled this feature on your machine, and he or she is right to be cautious about that one. You can do a lot of damage to the object model if you don't know what you're doing in there. Therefore, the `ModuleManager` functions gracefully fail and do nothing if access is denied.


## Getting Started

Fred is contained to four independent VBA class modules. That means if you only want the logging functionality, you don't have to bother with any of the other ones and bloat your project. Take only what you need.

To use it in your project, save the modules from the [Source Code](/source code) to your machine, then use one of the following methods to add them in the IDE.

- Select all of the class modules in the file explorer, then drag and drop into the Project window on the IDE
- Save the Fred Demonstration Workbook (note: workbook and link not yet available), then open both it and your project. In the IDE, drag each class file into your project
- From the IDE, use File -> Import File to manually import each one into your project
- From your project, create a blank class module and then copy/paste the code from this website like a monster


# Example Use

The following example cases reside in normal code modules and demonstrates how you might use Fred as a developer. Learn how to use all of the functions in the [documentation](/Documentation)!


## Assert

The `fredDeveloper_Assert` class provides a framework for enabling unit testing. For example:

```VBA
Sub FredDemonstrationUnitTesting()
    'Show an example use case for Test Driven Development
    
    Dim Assert As New fredDeveloper_Assert
    
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

Running this script gives the following output in the Immediate window:

> Test battery complete. 2 issues detected.
>
> Successful: 9
>
> Failed: 1 (String1EqualsNum1)
>
> Inconclusive: 1 (ThisTestIsInconclusive)

If we now know which tests are causing the problems so we can go track them down and work through bug patching.


## ClientSettings

The `fredDeveloper_ClientSettings` class provides an easy interface to the Workbook's `CustomDocumentProperties`. Variables saved there persist between VBA runtimes and save/close events on the workbook. Instead of saving data to a cell in a hidden worksheet, you can use these properties to make it easier to maintain your code as well as obfuscate it from end users more effectively.

```VBA
Sub FredDemonstrationClientSettings()
    'Show how to clear and add client settings, then run a report to easily show you what settings are available
    
    Dim ClientSettings As New fredDeveloper_ClientSettings
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

If we check the Immediate window, we find:

> Settings: 0
>
> MyFloatSetting Exists: True, and its value is 3.14159
>
> Settings: 5
>
> CustomDocumentProperties List (Index|Name|Value)
>
> 1|MyBooleanSetting|True
>
> 2|MyDateSetting|2/21/2021 7:42:42 AM
>
> 3|MyNumberSetting|1245
>
> 4|MyStringSetting|Test string
>
> 5|MyFloatSetting|3.14159


## Events

The `fredDeveloper_Events` class gives you logging functionality. Outputting to the Immediate Window is easy, but it slows down your code execution, and the end user certainly won't ever see it. Using the `Log` function takes care of that. You can write in log events for debugging into your code at key points. Set `LoggingMode = Disabled` for your production run. If you're having bugs, you can go back and redirect that output to the Immediate window, an external file, or both to help you follow what's going on.

Alternatively, you can leave the external writing on and use strategically placed `Log` commands in your scripts to keep a record of updates. In the example below, we also make use of the `Tic` and `Toc` functions to track how long it took to do an obscenely large number of calculations.

```VBA
Sub FredDemonstrationEvents()
    'Show how to log a debugging message with the time it took to perform 10,000,000 additions.
    
    Dim Events As New fredDeveloper_Events
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

Running this will produce this in the Immediate window:

> 10,000,000 Iteration Loop complete. Run time: 0.119140625 seconds.

And in the external log file:

> Debug Log file for Fred-the-Developer 0.1.0.xlsm
>
> Logging powered by Fred the Developer (https://github.com/M-Scott-Lassiter/Fred-the-Developer) under the MIT license, Copyright (c) 2021.
>
> Log created within Microsoft Excel by MSL on 02/21/2021 07:56:34.
>
>
> 02/21/2021 07:57:17|MSL|10,000,000 Iteration Loop complete. Run time: 0.119140625 seconds.

Every subsequent run will leave another log entry you can review later.


## ModuleManager

The `fredDeveloper_ModuleManager` class requires trusted access to the VBA module. If you have not granted that access (or can't), don't worry, VAB won't barrage you and your end users with run time errors because every property and method in this class will gracefully fail. If you do have this functionality though, you can use it to export and import class modules. This is particularly useful for git based source control. Although you won't be able to track git commits and changes real time in the IDE, git handles the exported code (all of which are just text files) like a champ. All of the files in [Source Code](/source code) were exported and tracked using a script like this:

```VBA
Sub FredDemonstrationModuleManager()
    'Show how to export the four class modules in Fred to a directory on your computer
    
    Dim ModuleManager As New fredDeveloper_ModuleManager
    
    With ModuleManager
        .Directory = ThisWorkbook.Path & "\Source Code"

        .Export "fredDeveloper_Assert"
        .Export "fredDeveloper_ClientSettings"
        .Export "fredDeveloper_Events"
        .Export "fredDeveloper_ModuleManager"
    End With

End Sub
```

# Project Roadmap

This project is in beta.

The next major addition is the [Assert class](/../../projects/1) which will take it to [version 1.0](/../../milestone/1).

See the open [issues](/../../issues) to see what work is coming up. The [Bug Fixing](/../../projects/2) board tracks known bug status.


## Contributing

This has been a solo development project so far, but I welcome any contributions or feedback. To fix an issue,

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m "Add some AmazingFeature"`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

If you just want to flag it for me to handle, then just open a new issue.


## License

Distributed under the MIT License. See [LICENSE](/../../blob/main/LICENSE) for more information.


## Contact

Reach me on [LinkedIn](https://www.linkedin.com/in/mscottlassiter/) or [Twitter](https://twitter.com/MScottLassiter).