Attribute VB_Name = "LinearGradeTests"
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

Private Const BeginStationValue As String = "100+00.00"
Private Const BeginStationElevation As Double = 100#
Private Const EndStationValue As String = "101+00.00"
Private Const EndStationElevation As Double = 105#
Private Const MidStationValue As String = "100+50.00"
Private Const MidStationElevation As Double = 102.5
Private Const LinearGrade As Double = 0.05

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
        Debug.Print "Begin LinearGradeTests"
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
        Debug.Print "End LinearGradeTests"
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

'@TestMethod("Creation")
Public Sub TestCreate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade

    'Assert:
    Assert.IsTrue Not IsNull(Grade.BeginStationValue)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestGetBeginStationElevation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade

    'Assert:
    Assert.AreEqual BeginStationElevation, Grade.BeginStationElevation

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Property")
Public Sub TestGetBeginStationValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestGetBeginStationValue()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Station Value: ", BeginStationValue
        Debug.Print "Actual Station Value: " & Grade.BeginStationValue
    #End If

    'Assert:
    Assert.AreEqual BeginStation.Value, Grade.BeginStationValue

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetEndStationElevation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestGetEndStationElevation()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Elevation: ", EndStationElevation
        Debug.Print "Actual Elevation: " & Grade.EndStationElevation
    #End If

    'Assert:
    Assert.AreEqual EndStationElevation, Grade.EndStationElevation

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetEndStationValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestGetEndStationValue()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Station Value: ", EndStationValue
        Debug.Print "Actual Station Value: " & Grade.EndStationValue
    #End If

    'Assert:
    Assert.AreEqual EndStation.Value, Grade.EndStationValue

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestElevationAt()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation
    
    Dim MidStation As Station
    Set MidStation = New Station
    MidStation.Value = StationStringParser.ToDouble(MidStationValue)

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestElevationAt()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Elevation: ", MidStationElevation
        Debug.Print "Actual Elevation: " & Grade.ElevationAt(MidStation)
    #End If

    'Assert:
    Assert.AreEqual MidStationElevation, Grade.ElevationAt(MidStation)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Calculation")
Public Sub TestGrade()
    On Error GoTo TestFail
    
    'Arrange:
    Dim BeginStation As Station
    Set BeginStation = New Station
    BeginStation.Value = StationStringParser.ToDouble(BeginStationValue)
    
    Dim BeginGrade As CurvePoint
    Set BeginGrade = New CurvePoint
    BeginGrade.SetCurvePoint BeginStation, BeginStationElevation
    
    Dim EndStation As Station
    Set EndStation = New Station
    EndStation.Value = StationStringParser.ToDouble(EndStationValue)
    
    Dim EndGrade As CurvePoint
    Set EndGrade = New CurvePoint
    EndGrade.SetCurvePoint EndStation, EndStationElevation

    'Act:
    Dim Grade As LinearGrade
    Set Grade = New LinearGrade
    Grade.Create BeginGrade, EndGrade
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestGrade()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Expected Grade: ", LinearGrade
        Debug.Print "Actual Elevation: " & Grade.Grade
    #End If

    'Assert:
    Assert.AreEqual LinearGrade, Grade.Grade

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


