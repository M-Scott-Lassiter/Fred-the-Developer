VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModuleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ExternalCodeDirectory As String
Private VBProject As Object
Private fso As Object   'FileScriptingObject

Private Sub Class_Initialize()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Not having trust access settings right will throw a 429 error. This class will thus be useless.
    On Error Resume Next
    Set VBProject = ThisWorkbook.VBProject
    If Err <> 0 Then Set VBProject = Nothing
    On Error GoTo 0
    
    ExternalCodeDirectory = ""   'Default until told otherwise
    
End Sub

Public Sub Export(Optional ModuleList As String)
    'Export a Form, Module, or Class Module using Excel's built in export tools to the folder specified by this class's Directory property.
    'If ModuleList is left empty, this sub will export all modules. Otherwise, ModuleList takes a comma separated list
    '   as an input, and will output only those modules by name.
    'Requires user to set "Trust access to the VBA vbproject object model" in the Macro Security settings. If not enabled,
    '   this sub will gracefully exit without doing anything.
    
    Dim ModuleArray() As String
    Dim ExportFlag As Boolean
    Dim FileExtension As String
    Dim MsgResponse As Integer
    Dim i As Integer, j As Integer
    
    If VBProject Is Nothing Then Exit Sub   'Code will otherwise fail because appropriate permissions not granted by Excel
    
    If Not fso.FolderExists(ExternalCodeDirectory) Then
        MsgResponse = MsgBox("No valid export filepath set.", vbCritical, "Export Folder Doesn't Exist")
        Exit Sub
    End If
    
    'If the optional ModuleList parameter was used, format the input string into a well trimmed array
    If ModuleList <> "" Then
        ModuleArray = Split(ModuleList, ",")
        For i = 0 To UBound(ModuleArray)
            ModuleArray(i) = Application.WorksheetFunction.Trim(ModuleArray(i))
        Next i
    End If
    
    'If the ModuleList was empty, export all modules.
    'If it wasn't, loop through each module and determine if it was one of the specified ones then export that.
    For i = 1 To VBProject.VBComponents.Count
        Select Case VBProject.VBComponents(i).Type
            Case 1  'vbext_ct_StdModule
                FileExtension = ".bas"
                If ModuleList = "" Then
                    ExportFlag = True
                Else
                    For j = 0 To UBound(ModuleArray)
                        If VBProject.VBComponents(i).Name = ModuleArray(j) Then
                            ExportFlag = True
                            Exit For
                        End If
                    Next j
                End If
            Case 2  'vbext_ct_ClassModule
                FileExtension = ".cls"
                If ModuleList = "" Then
                    ExportFlag = True
                Else
                    For j = 0 To UBound(ModuleArray)
                        If VBProject.VBComponents(i).Name = ModuleArray(j) Then
                            ExportFlag = True
                            Exit For
                        End If
                    Next j
                End If
            Case 3  'vbext_ct_MSForm
                FileExtension = ".frm"
                If ModuleList = "" Then
                    ExportFlag = True
                Else
                    For j = 0 To UBound(ModuleArray)
                        If VBProject.VBComponents(i).Name = ModuleArray(j) Then
                            ExportFlag = True
                            Exit For
                        End If
                    Next j
                End If
        End Select
        
        If ExportFlag Then VBProject.VBComponents(i).Export ExternalCodeDirectory & "\" & VBProject.VBComponents(i).Name & FileExtension
        ExportFlag = False
    Next i

    
End Sub

Public Sub Import(Optional ModuleList As String)
    'Import a Form, Module, or Class Module using Excel's built in import tools from the folder specified by this class's Directory property.
    'If ModuleList is left empty, this sub will import all modules. Otherwise, ModuleList takes a comma separated list
    '   as an input, and will import and overwrite only those modules by name.
    'Requires user to set "Trust access to the VBA vbproject object model" in the Macro Security settings. If not enabled,
    '   this sub will gracefully exit without doing anything.
    
    Dim fsofile As Object
    Dim ModuleArray() As String
    Dim ImportFlag As Boolean
    Dim FileExtension As String
    Dim i As Integer, j As Integer
    
    If VBProject Is Nothing Then Exit Sub   'Code will otherwise fail because appropriate permissions not granted by Excel
    
    'If the optional ModuleList parameter was used, format the input string into a well trimmed array
    If ModuleList <> "" Then
        ModuleArray = Split(ModuleList, ",")
        For i = 0 To UBound(ModuleArray)
            ModuleArray(i) = Application.WorksheetFunction.Trim(ModuleArray(i))
        Next i
    End If
    
    
    ''''NOTE: ADD LINES THAT PREVENT IMPORTING OVER THE CLASSES THAT ARE IN USE DURING THIS FUNCTION!!!!!!!!!!
    
    For Each fsofile In fso.GetFolder(ExternalCodeDirectory).Files
        If fso.GetExtensionName(fsofile.Name) = "bas" Or fso.GetExtensionName(fsofile.Name) = "cls" Or fso.GetExtensionName(fsofile.Name) = "frm" Then
            If ModuleList = "" Then
                ImportFlag = True
            Else
                For j = 0 To UBound(ModuleArray)
                    If fso.GetBaseName(fsofile.Path) = ModuleArray(j) Then
                        ImportFlag = True
                        Exit For
                    End If
                Next j
            End If
            
            If ImportFlag Then
                On Error Resume Next    'Remove this module if it exists to prevent conflicts. If it doesn't exist, the next line will throw an error
                VBProject.VBComponents.Remove VBProject.VBComponents(fso.GetBaseName(fsofile.Path))
                On Error GoTo 0
                
                VBProject.VBComponents.Import fsofile.Path
            End If
            ImportFlag = False
        End If
    Next fsofile
    
End Sub
    
Public Property Let Directory(FilePath As String)
    'This property ensures that there is always a valid file path directory to use when exporting/importing code modules, or else it is an empty string.
    
    Dim MsgResponse As Integer
    
    'Ensure the user entered the filepath appropriately formatted.
    Do While Right(FilePath, 1) = "\"
        FilePath = Left(FilePath, Len(FilePath) - 1)
    Loop
    
    If Not fso.FolderExists(FilePath) Then
        MsgResponse = MsgBox("This file path does not yet exist. Press 'Ok' to create the folder" & vbCrLf & vbCrLf & FilePath, vbOKCancel, "Folder Doesn't Exist")
        If MsgResponse = vbOK Then
            fso.CreateFolder (FilePath)
        Else
            FilePath = ""
        End If
        
    End If
    
    ExternalCodeDirectory = FilePath
    
End Property

Public Property Get Directory() As String

    Directory = ExternalCodeDirectory
    
End Property

Public Property Get CountAll() As String
    'Returns the number of user forms, standard modules, and class modules in this project.
    
    Dim i As Integer, ModuleCount As Integer
    
    If VBProject Is Nothing Then
        ModuleCount = 0
    Else
        For i = 1 To VBProject.VBComponents.Count
            With VBProject.VBComponents(i)
                'vbext_ct_StdModule = 1; 'vbext_ct_ClassModule = 2; vbext_ct_MSForm = 3
                If .Type = 1 Or .Type = 2 Or .Type = 3 Then ModuleCount = ModuleCount + 1
            End With
        Next i
    End If
    
    CountAll = ModuleCount
    
End Property

Public Property Get CountModules() As String
    'Returns the number of standard modules in this project.
    
    Dim i As Integer, ModuleCount As Integer
    
    If VBProject Is Nothing Then
        ModuleCount = 0
    Else
        For i = 1 To VBProject.VBComponents.Count
            With VBProject.VBComponents(i)
                'vbext_ct_StdModule = 1; 'vbext_ct_ClassModule = 2; vbext_ct_MSForm = 3
                If .Type = 1 Then ModuleCount = ModuleCount + 1
            End With
        Next i
    End If
    
    CountModules = ModuleCount
    
End Property

Public Property Get CountClasses() As String
    'Returns the number of class modules in this project.
    
    Dim i As Integer, ModuleCount As Integer
    
    If VBProject Is Nothing Then
        ModuleCount = 0
    Else
        For i = 1 To VBProject.VBComponents.Count
            With VBProject.VBComponents(i)
                'vbext_ct_StdModule = 1; 'vbext_ct_ClassModule = 2; vbext_ct_MSForm = 3
                If .Type = 2 Then ModuleCount = ModuleCount + 1
            End With
        Next i
    End If
    
    CountClasses = ModuleCount
    
End Property

Public Property Get CountForms() As String
    'Returns the number of user forms in this project.
    
    Dim i As Integer, ModuleCount As Integer
    
    If VBProject Is Nothing Then
        ModuleCount = 0
    Else
        For i = 1 To VBProject.VBComponents.Count
            With VBProject.VBComponents(i)
                'vbext_ct_StdModule = 1; 'vbext_ct_ClassModule = 2; vbext_ct_MSForm = 3
                If .Type = 3 Then ModuleCount = ModuleCount + 1
            End With
        Next i
    End If
    
    CountForms = ModuleCount
    
End Property

