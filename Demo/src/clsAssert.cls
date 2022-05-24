VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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


Public Sub IsEqual(Value1 As Variant, Value2 As Variant, TestName As String)
    
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

Public Sub IsNotEqual(Value1 As Variant, Value2 As Variant, TestName As String)
    
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

Public Sub IsLessOrEqual(LesserValue As Variant, GreaterValue As Variant, TestName As String)
    
    On Error GoTo TestInconclusivePoint
    
    If LesserValue <= GreaterValue Then
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

Public Sub IsInconclusive(TestName As String)
    'Used to immediately flag a test as inconclusive without further ado. Useful when initially constructing tests as a reminder that they are unfinished.
    
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
        Report = Report & "Failed: " & numFailed
        If numFailed > 0 Then Report = Report & " (" & FailedTestList & ")"
        Report = Report & vbCrLf
        
        'Report the inconclusive tests
        Report = Report & "Inconclusive: " & numInconclusive
        If numInconclusive > 0 Then Report = Report & " (" & InconclusiveTestList & ")"
        
    End If
    
End Property

Private Property Get FailedTestList() As String
    'This function returns a comma separated list of all recorded failed tests. If there weren't any, it returns "None"
    
    Dim dictKey As Variant
    
    If failedDict.Count = 0 Then
        FailedTestList = "None"
    Else
        For Each dictKey In failedDict
            FailedTestList = FailedTestList & dictKey & ", "
        Next dictKey
    
        FailedTestList = Left(FailedTestList, Len(FailedTestList) - 2)  'Remove the trailing ", "
    End If
    
End Property

Private Property Get InconclusiveTestList() As String
    'This function returns a comma separated list of all recorded inconclusive tests. If there weren't any, it returns "None"
    
    Dim dictKey As Variant
    
    If inconclusiveDict.Count = 0 Then
        InconclusiveTestList = "None"
    Else
        For Each dictKey In inconclusiveDict
            InconclusiveTestList = InconclusiveTestList & dictKey & ", "
        Next dictKey
    
        InconclusiveTestList = Left(InconclusiveTestList, Len(InconclusiveTestList) - 2)  'Remove the trailing ", "
    End If
    
End Property

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

