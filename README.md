# Fred the Developer
A standalone VBA development framework to enable better debugging, source control, and Test Driven Development in Excel VBA.

## About the Project

I built this project to help VBA developers

1) Output debugging or status lines to an external log file with a single line of code
2) Write unit tests using various `Assert` methods
3) Export and Import code modules, forms, or classes from within another VBA script
4) Add or remove Custom Document Properties to enable persistent variables between VBA executions or Workbook Close/Open events without writing to spreadsheets

My entire professional career has revolved around scant computer resources. The one ubiquitous constant? The Microsoft Office suite. While a fully featured tool such as [RubberDuck VBA](https://rubberduckvba.com/) provides robust capabilities, it is a C# add-in that you likely can't install without your system administrator's permission. 

Because I built Fred from 100% VBA class modules, you can simply add these modules to any macro enabled workbook, document, presentation, or more to flesh out your own robustly developed stand alone product.


## What does FRED stand for?
Fredrick... not everything has to be an acronym.


## Requirements

- Microsoft Office 2007 or newer
- A macro enabled file

You don't need to make any additional references. Fred makes references to the "Microsoft Scripting Runtime" and "Microsoft Visual Basic For Applications Extensibility 5.3" libraries. However, both are late bound to prevent end users from having to enable the references on each machine they use.

**Note:** the functionality in the `ModuleManager` class requires trusted access to the VBA project object model. See the Getting Started section below for a walkthrough to enable this. Note that your system administrator may have disabled this feature on your machine, and he or she is right to be cautious about that one. You can do a lot of damage to the object model if you don't know what you're doing in there. Therefore, the `ModuleManager` functions gracefully fail and do nothing if access is denied.


## Getting Started

Fred is completely contained within four VBA class modules. If you want to use Fred, then I assume you are familiar enough with VBA to access the Microsoft VBA IDE.

To use it in your project, either

- Save the modules from the [Source Code](/source code) then manually import each one into your project
- Open the Fred Demonstration Workbook (note: workbook and link not yet available) and drag each one into your project
- From your project, create a blank class module and then copy/paste the code from this website like a monster


## Example Use

This sub resides in a normal code module and demonstrates how you might use Fred as a developer. It exports three code modules to the "Source Code" folder in the workbook's path then logs the results to "FredTheDeveloper_DebugLog.txt" as well as the Immediate window in the VBA IDE. The `dev.Tic` and `dev.Toc` functions work together to measure how long it took to run that code segment.

```VBA
Sub ExportCodeModulesWithFred()

    Dim dev As New fredDeveloper
    
    dev.Tic
    dev.LoggingFilePath = ThisWorkbook.Path & "\FredTheDeveloper_DebugLog.txt"
    dev.LoggingMode = ToImmediateAndExternal
    
    With dev.ModuleManager
        .Directory = ThisWorkbook.Path & "\Source Code"
        .Export "fredDeveloper"
        .Export "fredDeveloper_ClientSettings"
        .Export "fredDeveloper_ModuleManager"
    End With
    
    dev.Log ("All code modules exported. Run time: " & dev.Toc & " seconds.")
    
End Sub
```

Learn how to use all of the functions in the [documentation](/Documentation/ReferenceGuide.md)!


## Project Roadmap

This project is in Beta phase. 

The next major addition is the [Assert class](/projects/1) which will take it to [version 1.0](/milestone/1).

See the open [issues](/issues) to see what work is coming up. The [Bug Fixing](/projects/2) board tracks known bug status.


### Contributing

This has been a solo project so far, but I welcome any contributions or feedback. To fix an issue,

1. Fork the Project
2. Create your Feature Branch (git checkout -b feature/AmazingFeature)
3. Commit your Changes (git commit -m 'Add some AmazingFeature')
4. Push to the Branch (git push origin feature/AmazingFeature)
5. Open a Pull Request


## License

Distributed under the MIT License. See [LICENSE](/blob/main/LICENSE) for more information.


## Contact

Reach me on Twitter [@mscottlassiter](https://twitter.com/MScottLassiter) or on [LinkedIn](https://www.linkedin.com/in/mscottlassiter/)

Project Link: https://github.com/your_username/repo_name.
