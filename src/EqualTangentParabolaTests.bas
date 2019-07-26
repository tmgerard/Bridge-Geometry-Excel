Attribute VB_Name = "EqualTangentParabolaTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Alignment.Vertical")

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As AssertClass
    Private Fakes As FakesProvider
#End If

' Values for Vertical Curve from CERM Example 79.5
Const exGradeIn As Double = 0.01
Const exGradeOut As Double = -0.0175
Const exPVISta As Double = 3500#
Const exPVIElev As Double = 549.2
Const exLength As Double = 400#
Const exBeginSta As Double = 3300#
Const exBeginElev As Double = 547.2
Const exEndSta As Double = 3700#
Const exEndElev As Double = 545.7
Const exRateOfGradeChange As Double = -0.6875
Const exChangeInGradient As Double = 2.75
Const exMiddleOrdinate As Double = 1.375
Const exToTurningPoint As Double = 145.454545454545
Const exElevAtSta34 As Double = 547.86
Const exSlopeAtSta24 As Double = 0.3125

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
        Debug.Print "Begin EqualTangentParabolaTests"
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
        Debug.Print "End EqualTangentParabolaTests"
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
Public Sub TestCreateCurveByPointsGetBeginStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestCreateCurveByPointsGetBeginStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Station Value: ", exBeginSta
        Debug.Print "Actual Station Value: " & VerticalCurve.BeginStationValue
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exBeginSta, VerticalCurve.BeginStationValue)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestCreateCurveByPointsGetBeginElevation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestCreateCurveByPointsGetBeginElevation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Station Value: ", exBeginElev
        Debug.Print "Actual Station Value: " & VerticalCurve.BeginElevationValue
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exBeginElev, VerticalCurve.BeginElevationValue)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestCreateCurveByPointsGetEndStation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestCreateCurveByPointsGetEndStation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Station Value: ", exEndSta
        Debug.Print "Actual Station Value: " & VerticalCurve.EndStationValue
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exEndSta, VerticalCurve.EndStationValue)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestCreateCurveByPointsGetEndElevation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestCreateCurveByPointsGetEndElevation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Station Value: ", exEndElev
        Debug.Print "Actual Station Value: " & VerticalCurve.EndElevationValue
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exBeginElev, VerticalCurve.BeginElevationValue)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestCreateByCurvePointsGetGradeIn()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVCSta As Station
    Set PVCSta = New Station
    PVCSta.Value = exBeginSta
    
    Dim PVC As CurvePoint
    Set PVC = New CurvePoint
    PVC.SetCurvePoint PVCSta, exBeginElev
    
    Dim PVISta As Station
    Set PVISta = New Station
    PVISta.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVISta, exPVIElev
    
    Dim PVTSta As Station
    Set PVTSta = New Station
    PVTSta.Value = exEndSta
    
    Dim PVT As CurvePoint
    Set PVT = New CurvePoint
    PVT.SetCurvePoint PVTSta, exEndElev
    
    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    VerticalCurve.CreateByCurvePoints PVC, PVI.Elevation, PVT
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestCreateByCurvePointsGetGradeIn()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Grade In Value: ", exGradeIn
        Debug.Print "Actual Grade In Value: ", VerticalCurve.GradeIn
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exGradeIn, VerticalCurve.GradeIn)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestCreateByCurvePointsGetGradeOut()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVCSta As Station
    Set PVCSta = New Station
    PVCSta.Value = exBeginSta
    
    Dim PVC As CurvePoint
    Set PVC = New CurvePoint
    PVC.SetCurvePoint PVCSta, exBeginElev
    
    Dim PVISta As Station
    Set PVISta = New Station
    PVISta.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVISta, exPVIElev
    
    Dim PVTSta As Station
    Set PVTSta = New Station
    PVTSta.Value = exEndSta
    
    Dim PVT As CurvePoint
    Set PVT = New CurvePoint
    PVT.SetCurvePoint PVTSta, exEndElev
    
    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    VerticalCurve.CreateByCurvePoints PVC, PVI.Elevation, PVT
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestCreateByCurvePointsGetGradeOut()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Grade Out Value: ", exGradeOut
        Debug.Print "Actual Grade Out Value: ", VerticalCurve.GradeOut
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exGradeOut, VerticalCurve.GradeOut)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestRateOfGradeChange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestRateOfGradeChange()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Rate of Grade Change: " & exRateOfGradeChange
        Debug.Print "Actual Rate of Grade Change: " & VerticalCurve.RateOfGradeChangePerStation
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exRateOfGradeChange, VerticalCurve.RateOfGradeChangePerStation)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestLength()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestLength()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Length: " & exLength
        Debug.Print "Actual Length: " & VerticalCurve.Length
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exLength, VerticalCurve.Length)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestChangeInGradient()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestMiddleOrdinate()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Change in Gradient: " & exChangeInGradient
        Debug.Print "Actual Change in Gradient: " & VerticalCurve.ChangeInGradient
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exChangeInGradient, VerticalCurve.ChangeInGradient)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestMiddleOrdinate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestMiddleOrdinate()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Middle Ordinate Distance: " & exMiddleOrdinate
        Debug.Print "Actual Middle Ordinate Distance: " & VerticalCurve.MiddleOrdinate
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exMiddleOrdinate, VerticalCurve.MiddleOrdinate)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestToTurningPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestToTurningPoint()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected To Turning Point: " & exToTurningPoint
        Debug.Print "Actual Distance To Turning Point: " & VerticalCurve.ToTurningPoint
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exToTurningPoint, VerticalCurve.ToTurningPoint)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestElevationAt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev
    
    Dim TestStation As Station
    Set TestStation = New Station
    TestStation.Value = 3400#

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestElevationAt()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Elevation: " & exElevAtSta34
        Debug.Print "Actual Elevation: " & VerticalCurve.ElevationAt(TestStation)
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exElevAtSta34, VerticalCurve.ElevationAt(TestStation), 2)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Public Sub TestElevationAt_StationBeforeCurve()
    Const ExpectedError As Long = AlignmentElementError.InvalidStationLimits
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev
    
    Dim TestStation As Station
    Set TestStation = New Station
    TestStation.Value = 3000#

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    Dim Elevation As Double
    Elevation = VerticalCurve.ElevationAt(TestStation)
    
Assert:
    Assert.Fail "Expected error was not raised."
    
'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:

    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestElevationAt_StationBeforeCurve()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Error: " & AlignmentElementError.InvalidStationLimits
        Debug.Print "Actual Error: " & Err.Number
    #End If
    
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestElevationAt_StationAfterCurve()
    Const ExpectedError As Long = AlignmentElementError.InvalidStationLimits
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev
    
    Dim TestStation As Station
    Set TestStation = New Station
    TestStation.Value = 3800#

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    Dim Elevation As Double
    Elevation = VerticalCurve.ElevationAt(TestStation)

Assert:
    Assert.Fail "Expected error was not raised."
    
'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestElevationAt_StationBeforeCurve()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Error: " & AlignmentElementError.InvalidStationLimits
        Debug.Print "Actual Error: " & Err.Number
    #End If
    
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Calculation")
Public Sub TestSlopeAt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev
    
    Dim TestStation As Station
    Set TestStation = New Station
    TestStation.Value = 3400#

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSlopeAt()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Slope: " & exSlopeAtSta24
        Debug.Print "Actual Slope: " & VerticalCurve.SlopeAt(TestStation)
    #End If

    'Assert:
    Assert.IsTrue DoubleCompare.CompareDoubleRound(exSlopeAtSta24, VerticalCurve.SlopeAt(TestStation), 4)

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Public Sub TestSlopeAt_StationBeforeCurve()
    Const ExpectedError As Long = AlignmentElementError.InvalidStationLimits
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev
    
    Dim TestStation As Station
    Set TestStation = New Station
    TestStation.Value = 3000#

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    Dim Slope As Double
    Slope = VerticalCurve.SlopeAt(TestStation)
    
Assert:
    Assert.Fail "Expected error was not raised."
    
'@Ignore LineLabelNotUsed
TestExit:
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSlopeAt_StationBeforeCurve()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Error: " & AlignmentElementError.InvalidStationLimits
        Debug.Print "Actual Error: " & Err.Number
    #End If
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestSlopeAt_StationAfterCurve()
    Const ExpectedError As Long = AlignmentElementError.InvalidStationLimits
    On Error GoTo TestFail
    
    'Arrange:
    Dim PVIStation As Station
    Set PVIStation = New Station
    PVIStation.Value = exPVISta
    
    Dim PVI As CurvePoint
    Set PVI = New CurvePoint
    PVI.SetCurvePoint PVIStation, exPVIElev
    
    Dim TestStation As Station
    Set TestStation = New Station
    TestStation.Value = 3800#

    'Act:
    Dim VerticalCurve As EqualTangentParabola
    Set VerticalCurve = New EqualTangentParabola
    
    VerticalCurve.CreateByTangentIntersection PVI, exLength, exGradeIn, exGradeOut
    
    Dim Slope As Double
    Slope = VerticalCurve.SlopeAt(TestStation)

Assert:
    Assert.Fail "Expected error was not raised."
    
'@Ignore LineLabelNotUsed
TestExit:
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestSlopeAt_StationAfterCurve()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Error: " & AlignmentElementError.InvalidStationLimits
        Debug.Print "Actual Error: " & Err.Number
    #End If
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

