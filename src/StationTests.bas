Attribute VB_Name = "StationTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Alignment.Dimensioning")

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As AssertClass
    Private Fakes As FakesProvider
#End If

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New AssertClass
        Set Fakes = New FakesProvider
    #End If
    
    #If DebugMode Then
        Debug.Print "*************************************************"
        Debug.Print "Begin StationTests"
        Debug.Print "*************************************************"
    #End If
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    #If DebugMode Then
        Debug.Print "*************************************************"
        Debug.Print "End StationTests"
        Debug.Print "*************************************************"
    #End If
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Arithmetic")
Public Sub TestAddStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim RightStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_STATION_VALUE As Double = 150#
    
    'Act:
    Set LeftStation = New Station
    Set RightStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    RightStation.Value = RIGHT_STATION_VALUE
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestAddStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RightStation Output String: " & RightStation.ToString
    #End If
    
    LeftStation.AddStation OtherStation:=RightStation
    
    #If DebugMode Then
        Debug.Print "Final Output String: ", LeftStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE + RIGHT_STATION_VALUE), LeftStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output")
Public Sub TestToString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestStation As Station
    Set TestStation = New Station
    Const STATION_NUMBER As Double = 12345.67
    Const STATION_STRING As String = "123+45.67"

    'Act:
    TestStation.Value = STATION_NUMBER
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestToString()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Input Value: ", , STATION_NUMBER
        Debug.Print "Output String: ", TestStation.ToString
    #End If

    'Assert:
    Assert.AreEqual STATION_STRING, TestStation.ToString
    
'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output")
Public Sub TestToStringZeroStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestStation As Station
    Set TestStation = New Station
    Const STATION_NUMBER As Double = 0
    Const STATION_STRING As String = "0+00.00"

    'Act:
    TestStation.Value = STATION_NUMBER
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestToStringZeroStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Input Value: ", , STATION_NUMBER
        Debug.Print "Output String: ", TestStation.ToString
    #End If

    'Assert:
    Assert.AreEqual STATION_STRING, TestStation.ToString

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestAddValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_VALUE As Double = 150#
    
    'Act:
    Set LeftStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestAddValue()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RIGHT_VALUE: ", , RIGHT_VALUE
    #End If
    
    LeftStation.AddValue Value:=RIGHT_VALUE
    
    #If DebugMode Then
        Debug.Print "Final Output String: ", LeftStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE + RIGHT_VALUE), LeftStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestSubtractStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim RightStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_STATION_VALUE As Double = 50#
    
    'Act:
    Set LeftStation = New Station
    Set RightStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    RightStation.Value = RIGHT_STATION_VALUE
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSubtractStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RightStation Output String: " & RightStation.ToString
    #End If
    
    LeftStation.SubtractStation OtherStation:=RightStation
    
    #If DebugMode Then
        Debug.Print "Final Output String: ", LeftStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE - RIGHT_STATION_VALUE), LeftStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestSubtractValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_VALUE As Double = 50#
    
    'Act:
    Set LeftStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSubtractValue()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RIGHT_VALUE: ", , RIGHT_VALUE
    #End If
    
    LeftStation.SubtractValue Value:=RIGHT_VALUE
    
    #If DebugMode Then
        Debug.Print "Final Output String: ", LeftStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE - RIGHT_VALUE), LeftStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Equals")
Public Sub TestEqualsTrueCondition()
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_VALUE As Double = 12345.67
    Dim StationOne As Station
    Set StationOne = New Station
    
    Dim StationTwo As Station
    Set StationTwo = New Station

    'Act:
    StationOne.Value = STATION_VALUE
    StationTwo.Value = STATION_VALUE
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestEqualsTrueCondition()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Station One: ", StationOne.Value
        Debug.Print "Station Two: ", StationTwo.Value
    #End If

    'Assert:
    Assert.IsTrue StationOne.Equals(StationTwo)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Equals")
Public Sub TestEqualsFalseCondition()
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_VALUE As Double = 12345.67
    Dim StationOne As Station
    Set StationOne = New Station
    
    Dim StationTwo As Station
    Set StationTwo = New Station

    'Act:
    StationOne.Value = STATION_VALUE
    StationTwo.Value = STATION_VALUE + 1
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestEqualsFalseCondition()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Station One: ", StationOne.Value
        Debug.Print "Station Two: ", StationTwo.Value
    #End If

    'Assert:
    Assert.IsFalse StationOne.Equals(StationTwo)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


