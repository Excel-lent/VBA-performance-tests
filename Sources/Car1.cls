VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Car1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public name As String
Public SeatCount As Long
Public DoorCount As Long
Public Distance As Long

Public Function Create(ByVal sName As String, iSeatCount As Long, iDoorCount As Long) As Car1
  Set Create = New Car1
  Call Create.protInit(sName, iSeatCount, iDoorCount)
End Function
Friend Sub protInit(ByVal sName As String, iSeatCount As Long, iDoorCount As Long)
  name = sName
  SeatCount = iSeatCount
  DoorCount = iDoorCount
End Sub

Public Sub Tick()
  Distance = Distance + 1
End Sub
