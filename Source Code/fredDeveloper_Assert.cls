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

Public Enum TestType    'Used for test battery reporting functionality, particularly with the TestsInError function
    [_First]
    TestFailed = 0
    TestInconclusive
    [_Last]
End Enum

Private numSuccessful As Long
Private numFailed As Long
Private numInconclusive As Long
Private failedDict As Object
Private inconclusiveDict As Object

Private Sub Class_Initialize()
    
    numSuccessful = 0
    numFailed = 0
    numInconclusive = 0
    
    'Use Scripting Dictionaries to track the names of tests that fail and are inconclusive. Not concerned about tests that passed.
    Set failedDict = CreateObject("Scripting.Dictionary")
    Set inconclusiveDict = CreateObject("Scripting.Dictionary")

End Sub


Public Sub AreEqual(Value1 As Variant, Value2 As Variant, TestName As String)
    'Does comparisons based on VBA's implicit type conversions. For example, AreEqual("1",1) would be successful.
    
    On Error GoTo TestInconclusivePoint
    
    If Value1 = Value2 Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Sub AreNotEqual(Value1 As Variant, Value2 As Variant, TestName As String)
    'Does comparisons based on VBA's implicit type conversions. For example, AreNotEqual("1",2) would be successful.
    
    On Error GoTo TestInconclusivePoint
    
    If Value1 <> Value2 Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Sub IsTrue(Value1 As Variant, TestName As String)
    
    On Error GoTo TestInconclusivePoint
    
    If Value1 = True Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Sub IsFalse(Value1 As Variant, TestName As String)
    
    On Error GoTo TestInconclusivePoint
    
    If Value1 = False Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Sub IsGreater(GreaterValue As Variant, LesserValue As Variant, TestName As String)
    
    On Error GoTo TestInconclusivePoint
    
    If GreaterValue > LesserValue Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Sub IsGreaterOrEqual(GreaterValue As Variant, LesserValue As Variant, TestName As String)
    
    On Error GoTo TestInconclusivePoint
    
    If GreaterValue >= LesserValue Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Sub IsLess(LesserValue As Variant, GreaterValue As Variant, TestName As String)
    
    On Error GoTo TestInconclusivePoint
    
    If LesserValue < GreaterValue Then
        numSuccessful = numSuccessful + 1
    Else
        numFailed = numFailed + 1
        failedDict.Add TestName, False
    End If
    
    On Error GoTo 0
    Exit Sub
    
TestInconclusivePoint:
    On Error GoTo 0
    numInconclusive = numInconclusive + 1
    inconclusiveDict.Add TestName, False
    
End Sub

Public Property Get Report() As String
    'Returns a string, possibly on multiple lines, detailing how the current test battery went.
    
    If numSuccessful + numFailed + numInconclusive = 0 Then
        Report = "No tests run."
        
    ElseIf numSuccessful > 0 And numFailed + numInconclusive = 0 Then   'All tests were successful
        Report = "Test battery complete. Ran " & numSuccessful
        If numSuccessful = 1 Then Report = Report & " test." Else Report = Report & " tests."
        
    Else
        'At least one test has been run, but something wasn't right. Go into detail about it.
        Report = "Test battery complete. " & numFailed + numInconclusive & " issues detected." & vbCrLf
        Report = Report & "Successful: " & numSuccessful & vbCrLf
        
        'Report the failed tests
        Report = Report & "Failed: " & numFailed & vbCrLf
        If numFailed > 0 Then Report = Report & "Failed Tests: " & TestsInError(TestFailed) & vbCrLf
        
        'Report the inconclusive tests
        Report = Report & "Inconclusive: " & numInconclusive & vbCrLf
        If numInconclusive > 0 Then Report = Report & "Inconclusive Tests: " & TestsInError(TestInconclusive)
        
    End If
    
End Property

Private Function TestsInError(TypeToReport As TestType) As String
    'This function returns a comma separated list of all recorded failed or inconclusive tests. Helper function to the property Report.
    
    Dim dict As Object
    Dim dictKey As Variant
    
    If TypeToReport = TestFailed Then
        Set dict = failedDict
    ElseIf TypeToReport = TestInconclusive Then
        Set dict = inconclusiveDict
    End If
    
    For Each dictKey In dict
        TestsInError = TestsInError & dictKey & ", "
    Next dictKey
    
    If Len(TestsInError) <= 2 Then Exit Function    'If the next line tries to run with a string length <2, it will cause an error
    TestsInError = Left(TestsInError, Len(TestsInError) - 2)  'Remove the trailing ", "
    
End Function

Public Property Get Failed() As Long
    'Number of successful tests that have been run since class was initialized/reset
    
    Failed = numFailed
    
End Property

Public Property Get Inconclusive() As Long
    'Number of successful tests that have been run since class was initialized/reset
    
    Inconclusive = numInconclusive
    
End Property

Public Property Get Successful() As Long
    'Number of successful tests that have been run since class was initialized/reset
    
    Successful = numSuccessful
    
End Property
