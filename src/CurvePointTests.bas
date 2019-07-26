Attribute VB_Name = "CurvePointTests"
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
        Debug.Print "Begin CurvePointTests"
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
        Debug.Print "End CurvePointTests"
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

'@TestMethod("Property")
Public Sub TestGetStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Station As Station
    Set Station = New Station
    Station.Value = 12345.67
    
    Dim Elevation As Double
    Elevation = 100#
    
    'Act:
    Dim Point As CurvePoint
    Set Point = New CurvePoint
    Point.SetCurvePoint Station, Elevation
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestGetStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "TypeOf Station: ", TypeOf Point.Station Is Station
    #End If
    
    'Assert:
    Assert.IsTrue TypeOf Point.Station Is Station

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestGetElevation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Station As Station
    Set Station = New Station
    Station.Value = 12345.67
    
    Dim Elevation As Double
    Elevation = 100#
    
    'Act:
    Dim Point As CurvePoint
    Set Point = New CurvePoint
    Point.SetCurvePoint Station, Elevation
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestGetElevation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Input Elevation: ", Elevation
        Debug.Print "Output Elevation: ", Point.Elevation
    #End If
    
    'Assert:
    Assert.AreEqual Elevation, Point.Elevation

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestSlopeToPositiveSlope()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = 0.0325
    
    Const StationOneValue As Double = 10000#
    Const StationOneElevation As Double = 100
    
    Dim StationOne As Station
    Set StationOne = New Station
    StationOne.Value = StationOneValue
    
    Const StationTwoValue As Double = 10100#
    Const StationTwoElevation As Double = 103.25
    
    Dim StationTwo As Station
    Set StationTwo = New Station
    StationTwo.Value = StationTwoValue

    'Act:
    Dim CurvePointOne As CurvePoint
    Set CurvePointOne = New CurvePoint
    CurvePointOne.SetCurvePoint StationOne, StationOneElevation
    
    Dim CurvePointTwo As CurvePoint
    Set CurvePointTwo = New CurvePoint
    CurvePointTwo.SetCurvePoint StationTwo, StationTwoElevation
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSlopeToPositiveSlope()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Slope: ", Expected
        Debug.Print "Actual SLope: ", CurvePointOne.SlopeTo(CurvePointTwo)
    #End If

    'Assert:
    Assert.AreEqual Expected, CurvePointOne.SlopeTo(CurvePointTwo)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestSlopeToNegativeSlope()
    On Error GoTo TestFail
    
    'Arrange:
    Const Expected As Double = -0.0325
    
    Const StationOneValue As Double = 10000#
    Const StationOneElevation As Double = 103.25
    
    Dim StationOne As Station
    Set StationOne = New Station
    StationOne.Value = StationOneValue
    
    Const StationTwoValue As Double = 10100#
    Const StationTwoElevation As Double = 100#
    
    Dim StationTwo As Station
    Set StationTwo = New Station
    StationTwo.Value = StationTwoValue

    'Act:
    Dim CurvePointOne As CurvePoint
    Set CurvePointOne = New CurvePoint
    CurvePointOne.SetCurvePoint StationOne, StationOneElevation
    
    Dim CurvePointTwo As CurvePoint
    Set CurvePointTwo = New CurvePoint
    CurvePointTwo.SetCurvePoint StationTwo, StationTwoElevation
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSlopeToNegativeSlope()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Slope: ", Expected
        Debug.Print "Actual SLope: ", CurvePointOne.SlopeTo(CurvePointTwo)
    #End If

    'Assert:
    Assert.AreEqual Expected, CurvePointOne.SlopeTo(CurvePointTwo)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
