Attribute VB_Name = "MRootTests"
Option Explicit

Public Type TRoot
    Value  As Double
    Result As Double
End Type
Public Type TRootTest
    StartTime As Double
    EndTime   As Double
    Count     As Long
    Test()    As TRoot
End Type
'Private m_n As Long
'Public m_Roots() As TRoot

Public Function New_RootTest(ByVal n As Long) As TRootTest
    With New_RootTest
        .Count = n
        ReDim .Test(0 To n - 1)
    End With
End Function

Public Sub RootTest_InitRandomNumbers(this As TRootTest)
    Dim i As Long
    With this
        Randomize Timer
        For i = 0 To .Count - 1
            With .Test(i)
                .Value = Rnd() * 12345678901.2345
            End With
        Next
    End With
End Sub

Public Function RootTest_Clone(this As TRootTest) As TRootTest
    With RootTest_Clone
        .Count = this.Count
        .Test = this.Test
    End With
End Function

Public Sub RootTest_ClearResults(this As TRootTest)
    Dim i As Long
    With this
        For i = 0 To .Count - 1
            With .Test(i)
                .Result = 0#
            End With
        Next
    End With
End Sub

Public Function RootTest_ResultsAreEqual(this As TRootTest, other As TRootTest) As Boolean
    Dim i As Long
    With this
        If .Count <> other.Count Then Exit Function
        For i = 0 To .Count - 1
            With .Test(i)
                Dim d As Double: d = Abs(.Result - other.Test(i).Result)
                If d > 0.000000001 Then
                    Debug.Print CStr(i) & " " & .Result & " " & other.Test(i).Result
                    Exit Function 'no not equal
                End If
            End With
        Next
    End With
    RootTest_ResultsAreEqual = True
End Function

Public Function RootTest_GetTime(this As TRootTest) As Double
    With this
        RootTest_GetTime = (.EndTime - .StartTime) * 1000
    End With
End Function

Public Function RootTest_ToStr(this As TRootTest, funcname As String) As String
    With this
        RootTest_ToStr = "Testing " & .Count & " times " & funcname & " : " & RootTest_GetTime(this) & " ms"
    End With
End Function
