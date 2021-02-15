VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fredDeveloper_ClientSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Version 0.1.0                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MIT License                                                                    '
'                                                                               '
'Copyright (c) 2021 M. Scott Lassiter                                           '
'                                                                               '
'Permission is hereby granted, free of charge, to any person obtaining a copy   '
'of this software and associated documentation files (the "Software"), to deal  '
'in the Software without restriction, including without limitation the rights   '
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell      '
'copies of the Software, and to permit persons to whom the Software is          '
'furnished to do so, subject to the following conditions:                       '
'                                                                               '
'The above copyright notice and this permission notice shall be included in all '
'copies or substantial portions of the Software.                                '
'                                                                               '
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR     '
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,       '
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE    '
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER         '
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,  '
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE  '
'SOFTWARE.                                                                      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Dim i As Integer
    
    i = 0
    For Each prop In cdp
        i = i + 1
    Next prop
    
End Property

Public Property Get ClientSettingExists(ClientSetting As String) As Boolean
    'Attempts to set a variant equal to the setting. If it works, the setting obviously exists. If there's an error, then it doesn't.
    
    Dim SettingValue As Variant
    
    On Error Resume Next
    SettingValue = cdp(ClientSetting)
    If Err <> 0 Then ClientSettingExists = False Else ClientSettingExists = True
    On Error GoTo 0
    
End Property


