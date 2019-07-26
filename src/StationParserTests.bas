Attribute VB_Name = "StationParserTests"
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
        Debug.Print "Begin StationParserTests"
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
        Debug.Print "End StationParserTests"
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

'@TestMethod("Output")
Public Sub TestStationStringToDouble()
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_STRING As String = "123+45.67"
    Const Expected As Double = 12345.67

    'Act:
    Dim Actual As Double
    Actual = StationStringParser.ToDouble(STATION_STRING)
    
    #If DebugMode Then
        Debug.Print "-------------------------------------------------"
        Debug.Print "TestToString()"
        Debug.Print "-------------------------------------------------"
        Debug.Print "Station Input String: ", STATION_STRING
        Debug.Print "Station Output Value: ", Actual
    #End If

    'Assert:
    Assert.AreEqual Expected, Actual

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Expected Error")
Public Sub TestStationStringToDoubleTooManyDelimeters()
    Const ExpectedError As Long = StationParserError.InvalidStationFormat
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_STRING As String = "12+3+45.67"

    'Act:
    Dim Actual As Double
    Actual = StationStringParser.ToDouble(STATION_STRING)

Assert:
    Assert.Fail "Expected error was not raised."

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        #If DebugMode Then
            Debug.Print "-------------------------------------------------"
            Debug.Print "TestStationStringToDoubleTooManyDelimeters()"
            Debug.Print "-------------------------------------------------"
            Debug.Print "Station Input String: ", STATION_STRING
            Debug.Print "Expected Error: ", ExpectedError
            Debug.Print "Error Number: ", Err.Number
        #End If
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestStationStringToDoubleNoDelimeter()
    Const ExpectedError As Long = StationParserError.InvalidStationFormat
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_STRING As String = "12345.67"

    'Act:
    Dim Actual As Double
    Actual = StationStringParser.ToDouble(STATION_STRING)

Assert:
    Assert.Fail "Expected error was not raised."

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        #If DebugMode Then
            Debug.Print "-------------------------------------------------"
            Debug.Print "TestStationStringToDoubleNoDelimeter()"
            Debug.Print "-------------------------------------------------"
            Debug.Print "Station Input String: ", STATION_STRING
            Debug.Print "Expected Error: ", ExpectedError
            Debug.Print "Error Number: ", Err.Number
        #End If
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestStationStringToDoubleBadDelimeter()
    Const ExpectedError As Long = StationParserError.InvalidStationFormat
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_STRING As String = "123-45.67"

    'Act:
    Dim Actual As Double
    Actual = StationStringParser.ToDouble(STATION_STRING)

Assert:
    Assert.Fail "Expected error was not raised."

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        #If DebugMode Then
            Debug.Print "-------------------------------------------------"
            Debug.Print "TestStationStringToDoubleBadDelimeter()"
            Debug.Print "-------------------------------------------------"
            Debug.Print "Station Input String: ", STATION_STRING
            Debug.Print "Expected Error: ", ExpectedError
            Debug.Print "Error Number: ", Err.Number
        #End If
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestStationStringNonNumericStation()
    Const ExpectedError As Long = StationParserError.NonNumericStation
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_STRING As String = "1A3+45.67"

    'Act:
    Dim Actual As Double
    Actual = StationStringParser.ToDouble(STATION_STRING)

Assert:
    Assert.Fail "Expected error was not raised."

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        #If DebugMode Then
            Debug.Print "-------------------------------------------------"
            Debug.Print "TestStationStringNonNumericStation()"
            Debug.Print "-------------------------------------------------"
            Debug.Print "Station Input String: ", STATION_STRING
            Debug.Print "Expected Error: ", ExpectedError
            Debug.Print "Error Number: ", Err.Number
        #End If
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Expected Error")
Public Sub TestStationStringNonNumericFeet()
    Const ExpectedError As Long = StationParserError.NonNumericStation
    On Error GoTo TestFail
    
    'Arrange:
    Const STATION_STRING As String = "123+4A.67"

    'Act:
    Dim Actual As Double
    Actual = StationStringParser.ToDouble(STATION_STRING)

Assert:
    Assert.Fail "Expected error was not raised."

'@Ignore LineLabelNotUsed
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        #If DebugMode Then
            Debug.Print "-------------------------------------------------"
            Debug.Print "TestStationStringNonNumericStation()"
            Debug.Print "-------------------------------------------------"
            Debug.Print "Station Input String: ", STATION_STRING
            Debug.Print "Expected Error: ", ExpectedError
            Debug.Print "Error Number: ", Err.Number
        #End If
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


