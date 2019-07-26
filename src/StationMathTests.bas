Attribute VB_Name = "StationMathTests"
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
        Debug.Print "Begin StationMathTests"
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
        Debug.Print "End StationMathTests"
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
Public Sub TestAddStations()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim RightStation As Station
    Dim NewStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_STATION_VALUE As Double = 150#
    
    'Act:
    Set LeftStation = New Station
    Set RightStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    RightStation.Value = RIGHT_STATION_VALUE
        
    Set NewStation = StationMath.AddStations(LeftStation, RightStation)
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestAddStations()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RightStation Output String: " & RightStation.ToString
        Debug.Print "NewStation Output String: ", NewStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE + RIGHT_STATION_VALUE), NewStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestAddValueToStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim NewStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_VALUE As Double = 150#
    
    'Act:
    Set LeftStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
        
    Set NewStation = StationMath.AddValueToStation(LeftStation, RIGHT_VALUE)
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestAddValueToStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RIGHT_VALUE Output String: ", RIGHT_VALUE
        Debug.Print "NewStation Output String: ", NewStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE + RIGHT_VALUE), NewStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestDifferenceInStations()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim RightStation As Station
    Dim Difference As Double
    Const LEFT_STATION_VALUE As Double = 200#
    Const RIGHT_STATION_VALUE As Double = 50#
    Const Expected As Double = 1.5
    
    'Act:
    Set LeftStation = New Station
    Set RightStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    RightStation.Value = RIGHT_STATION_VALUE
        
    Difference = StationMath.DifferenceInStations(LeftStation, RightStation)
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestDifferenceInStations()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RightStation Output String: " & RightStation.ToString
        Debug.Print "Difference: ", , Str$(Difference)
    #End If

    'Assert:
    Assert.AreEqual Expected, Difference

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestSubtractStations()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim RightStation As Station
    Dim NewStation As Station
    Const LEFT_STATION_VALUE As Double = 250#
    Const RIGHT_STATION_VALUE As Double = 150#
    
    'Act:
    Set LeftStation = New Station
    Set RightStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
    RightStation.Value = RIGHT_STATION_VALUE
        
    Set NewStation = StationMath.SubtractStations(LeftStation, RightStation)
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSubtractStations()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RightStation Output String: " & RightStation.ToString
        Debug.Print "NewStation Output String: ", NewStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE - RIGHT_STATION_VALUE), NewStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Arithmetic")
Public Sub TestSubtractValueFromStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim LeftStation As Station
    Dim NewStation As Station
    Const LEFT_STATION_VALUE As Double = 100#
    Const RIGHT_VALUE As Double = 50#
    
    'Act:
    Set LeftStation = New Station
    
    LeftStation.Value = LEFT_STATION_VALUE
        
    Set NewStation = StationMath.SubtractValueFromStation(LeftStation, RIGHT_VALUE)
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSubtractValueFromStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "LeftStation Output String: ", LeftStation.ToString
        Debug.Print "RIGHT_VALUE Output String: ", RIGHT_VALUE
        Debug.Print "NewStation Output String: ", NewStation.ToString
    #End If

    'Assert:
    Assert.AreEqual (LEFT_STATION_VALUE - RIGHT_VALUE), NewStation.Value

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


