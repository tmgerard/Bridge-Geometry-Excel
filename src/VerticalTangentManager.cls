VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VerticalTangentManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("BridgeGeometry.Alignment.Vertical")

Private Type TVerticalTangentManager
    pointsOfIntersection As Collection
    linearGrades As Collection
End Type

Private This As TVerticalTangentManager
Private Const FirstItemIndex As Long = 1

Private Sub Class_Initialize()
    With This
        Set .pointsOfIntersection = New Collection
        Set .linearGrades = New Collection
    End With
End Sub

Private Sub Class_Terminate()
    With This
        Set .pointsOfIntersection = Nothing
        Set .linearGrades = Nothing
    End With
End Sub

Public Property Get GetPI(ByVal index As Long)
    GetPI = This.pointsOfIntersection(index)
End Property

Public Property Get GetLinearGrade(ByVal index As Long)
    GetLinearGrade = This.linearGrades(index)
End Property

Public Sub AddPI(ByRef piPoint As CurvePoint)
    
    Dim checkPI As CurvePoint
    
    If This.pointsOfIntersection.Count = 0 Then
        
        This.pointsOfIntersection.Add Item:=piPoint
    
    ElseIf This.pointsOfIntersection.Count = 1 Then
        
        
        Set checkPI = This.pointsOfIntersection.Item(FirstItemIndex)
        
        If piPoint.Station.Value < checkPI.Station.Value Then
            This.pointsOfIntersection.Add Item:=piPoint, Before:=FirstItemIndex
        Else
            This.pointsOfIntersection.Add Item:=piPoint
        End If
        
        
    ElseIf This.pointsOfIntersection.Count > 1 Then
    
        Dim index As Long
        For index = FirstItemIndex To This.pointsOfIntersection.Count
            Set checkPI = This.pointsOfIntersection.Item(index)
            ' add PI to correct location
        Next index
    
    End If
    
End Sub


