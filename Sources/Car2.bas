Attribute VB_Name = "Car2"
Public Type CarData
  name As String
  SeatCount As Long
  DoorCount As Long
  Distance As Long
End Type

Public Function Car_Create(ByVal sName As String, iSeatCount As Long, iDoorCount As Long) As CarData
  Car_Create.name = sName
  Car_Create.SeatCount = iSeatCount
  Car_Create.DoorCount = iDoorCount
End Function

Public Sub Car_Tick(ByRef data As CarData)
  data.Distance = data.Distance + 1
End Sub
