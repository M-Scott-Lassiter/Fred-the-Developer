VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fredDeveloper_Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MIT License                                                                    '
'                                                                               '
'Copyright (c) 2021 M. Scott Lassiter                                            '
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

Private NumSuccessful As Long
Private NumFailed As Long
Private NumInconclusive As Long


Private Sub Class_Initialize()
    
    NumSuccessful = 0
    NumFailed = 0
    NumInconclusive = 0
    
End Sub


Public Sub AreEqual(Value1 As Variant, Value2 As Variant)
    'Does comparisons based on VBA's implicit type conversions. For example, AreEqual("1",1) would be true.
    
    On Error GoTo TestInconclusivePoint
    
    If Value1 = Value2 Then
        NumSuccessful = NumSuccessful + 1
    Else
        NumFailed = NumFailed + 1
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    NumInconclusive = NumInconclusive + 1
    
End Sub

Public Property Get Failed() As Long
    'Number of successful tests that have been run since class was initialized/reset
    
    Successful = NumFailed
    
End Property

Public Property Get Inconclusive() As Long
    'Number of successful tests that have been run since class was initialized/reset
    
    Successful = NumInconclusive
    
End Property

Public Property Get Successful() As Long
    'Number of successful tests that have been run since class was initialized/reset
    
    Successful = NumSuccessful
    
End Property
