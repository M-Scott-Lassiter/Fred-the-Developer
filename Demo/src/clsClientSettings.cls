VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClientSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cdp As Object   'Custom Document Properties, declared in Initialization

Private Sub Class_Initialize()

    Set cdp = ThisWorkbook.CustomDocumentProperties
    
End Sub

Public Sub Add(SettingName As String, SettingType As MsoDocProperties, DefaultValue As Variant)
    'Inserts a custom setting into ThisWorkbook.CustomDocumentProperties.

    On Error Resume Next
    cdp.Add Name:=SettingName, _
         LinkToContent:=False, _
         Type:=SettingType, _
         Value:=DefaultValue
    On Error GoTo 0
    
End Sub

Public Sub Delete(SettingName As String)
    'Deletes the selected setting in ThisWorkbook.CustomDocumentProperties (if it exists) without any further warning.
    
    On Error Resume Next
    cdp(SettingName).Delete
    On Error GoTo 0
    
End Sub

Public Sub Clear()
    'Deletes ALL client settings in ThisWorkbook.CustomDocumentProperties without further warning.
    
    Dim prop As Object
    
    For Each prop In cdp
        prop.Delete
    Next prop
    
End Sub

Public Property Get Count() As Integer
    'Deletes ALL client settings in ThisWorkbook.CustomDocumentProperties without further warning.
    
    Dim prop As Object
    
    Count = 0
    For Each prop In cdp
        Count = Count + 1
    Next prop
    
End Property

Public Property Get Exists(ClientSetting As String) As Boolean
    'Attempts to set a variant equal to the setting. If it works, the setting obviously exists. If there's an error, then it doesn't.
    
    Dim SettingValue As Variant
    
    On Error Resume Next
    SettingValue = cdp(ClientSetting)
    If Err <> 0 Then Exists = False Else Exists = True
    On Error GoTo 0
    
End Property

Public Sub Report()

    Dim prop As Object
    Dim i As Integer
    
    Debug.Print "CustomDocumentProperties List (Index|Name|Value)"
    i = 0
    For Each prop In cdp
        i = i + 1
        Debug.Print i & "|" & prop.Name & "|" & prop.Value
    Next prop
    
    
End Sub


