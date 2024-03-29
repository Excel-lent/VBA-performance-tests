VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'This class has been designed to meet your performance testing and optimisation needs. stdPerformance uses the `Sentry Object` design pattern
'which allows for cleaner more maintainable code.
'Functions implemented on the class
'CONSTRUCTORS
'    [X] Create   - With Cache
'    [X] init       #PROTECTED
'    [X] Measure  - Create a performance measuring Sentry Object
'    [X] Optimise - Create a object which toggles runtime options for optimisation. Currently sets: `ScreenUpdating`, `EnableEvents` and `XLCalculation`.
'                   This is intended to be application agnostic.
'
'STATIC PROPERTIES
'    [X] Get MeasureKeys - Get an array of procs/blocks which have been measured
'    [x] Get Measurement(sProcOrBlock) - Get the average time it took to execute a block.
'    [ ] Get MeasuresStr
'    [ ] Get MeasuresHtml
'
'STATIC methods
'    [x] MeasuresClear() - Clear the performance stack.
'
'OUT-OF-SCOPE
'    * Anything performance related which is specific, should realistically be honed to a specific class for that thing.
'
'EXAMPLES
'# 1 - Usage of Optimser
'
'   'Disable numerous options for performance
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'   With stdPerformance.Optimiser()
'     Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4135
'   End With
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'
'   'Disable everything BUT Calculation
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'   With stdPerformance.Optimiser(Calculation:=xlCalculation.xlCalculationAutomatic)
'     Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'   End With
'   Debug.Print Application.ScreenUpdating, Application.EnableEvents, Application.Calculation         'false,false,-4105
'
'# 2 - measuring performance:
'
'   With stdPerformance.measure("#1 Select then set")
'     For i = 1 to C_MAX
'       cells(1,1).select
'       selection.value = "hello"
'     Next
'   End With
'
'   With stdPerformance.measure("#2 Set directly")
'     For i = 1 to C_MAX
'       cells(1,1).value = "hello"
'     next
'   End With
'
'Declares for performance counters
#If Mac Then
   #If MAC_OFFICE_VERSION >= 15 Then
      Private Declare Function GetTickCount Lib "/Applications/Microsoft Excel.app/Contents/Frameworks/MicrosoftOffice.framework/MicrosoftOffice" () As Long
   #Else
      Private Declare Function GetTickCount Lib "Applications:Microsoft Office 2011:Office:MicrosoftOffice.framework:MicrosoftOffice" () As Long
   #End If
#Else ' Win32 or Win64
   #If VBA7 Then
      Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
   #Else
      Private Declare Function GetTickCount Lib "kernel32" () As Long
   #End If
#End If

'Enum for sentry object type
Public Enum EPerfObjType
  iMeasure = 1
  iOptimiser = 2
End Enum

'The instance object type.
Private pObjType As EPerfObjType

'iOptimiser Fields...
Private pEnableEvents As Boolean
Private pScreenUpdating As Boolean
Private pCalculation As Long

'iMeasure definitions
Private pStartTime As Long
Private pKeyName As String
Private pDivider As Double

'Measurement storage
Private Type FakeDictItem
  key As String
  val As Variant
End Type
Private FakeDict() As FakeDictItem

'Create
'@constructor
'@param {EPerfObjType} - Type of performance object to create. iMeasure - used for measuring performance, iOptimiser - used for optimising performance
'@param {Variant()} - Additional params supplied as array.
'@returns {stdPerformance<EPerfObjType>} - Object termination has special behaviour. See Measure and Optimise methods for further details.
Public Function Create(ByVal objType As EPerfObjType, ByVal params As Variant) As stdPerformance
  Set Create = New stdPerformance
  Call Create.Init(objType, params)
End Function

'Init
'PROTECTED - Don't call this method unless you know what you are doing.
'Initialises the class
'@protected
'@param {EPerfObjType} - Type of performance object to create. iMeasure - used for measuring performance, iOptimiser - used for optimising performance
'@param {Variant()} - Additional params supplied as array.
Public Sub Init(ByVal objType As EPerfObjType, ByVal params As Variant)
  pObjType = objType
  Select Case objType
    Case iMeasure
      pKeyName = params(0)
      pDivider = params(1)
      pStartTime = GetTickCount()
    Case iOptimiser
      'Store vals
      pScreenUpdating = Application.ScreenUpdating
      pEnableEvents = Application.EnableEvents
      
      'Set vals
      Application.ScreenUpdating = params(0)
      Application.EnableEvents = params(1)
      
      'Different options for different applications
      Select Case Application.name
        Case "Microsoft Excel"
          pCalculation = Application.Calculation
          Application.Calculation = params(2)
      End Select

  End Select
End Sub

'Measure
'@constructor
'@param {String} - Name of method or block to measure
'@returns {stdPerformance<iMeasure>} - Object which upon termination, adds measurement of block to global cache
'@usage
'  ```vb
'  With stdPerformance.Measure("Hello world")
'    For i = 1 to 1000
'      Debug.print "Hello world"
'    next
'  End With
'  ```
Public Function Measure(ByVal sProc As String, Optional ByVal nCount As Double = 1) As stdPerformance
  Set Measure = stdPerformance.Create(iMeasure, Array(sProc, nCount))
End Function

'Optimise
'@constructor
'@param {Boolean} - ScreenUpdating set value
'@param {Boolean} - EnableEvents set value
'@returns {stdPerformance<iOptimiser>} - Object termination has special behaviour. See Measure and Optimise methods for further details.
'@note Calculation is defined as long instead of xlCalculation so the function continues to work without compile error in Word, Powerpoint etc.
'@usage
'  ```vb
'  With stdPerformance.Optimise
'    'some heavy code here
'  End With
'  ```
Public Function Optimise(Optional ByVal ScreenUpdating As Boolean = False, Optional ByVal EnableEvents As Boolean = False, Optional ByVal Calculation As Long = -4135) As stdPerformance
  Set Optimise = stdPerformance.Create(iOptimiser, Array(ScreenUpdating, EnableEvents, Calculation))
End Function


'Measurement
'@param {String} - Name of measurement to get
'@returns {Double} - Average measurement time
Public Property Get Measurement(ByVal sKey As String) As Double
  If Me Is stdPerformance Then
    Dim v: v = getDictItem(sKey)
    If TypeName(v) = "Variant()" Then
        Measurement = getDictItem(sKey)(0)
    Else
        Measurement = Empty
    End If
  End If
End Function

'AddMeasurement
'If a time is added that was previously also added then the average of the times is calculated.
'@param {String} - Name of measurement to add to global cache
'@param {Double} - time to add to global cache
'@param {Double} - number of operations (divisor)
Public Sub AddMeasurement(ByVal sKey As String, ByVal time As Double, ByVal nCount As Double)
  If Me Is stdPerformance Then
    Debug.Print sKey & ": " & time & " ms" & iif(nCount > 1, " (" & (1000 * time / nCount) & chr(181) & "s per operation)", "")
    Dim ind As Long: ind = getDictIndex(sKey)
    If ind = -1 Then
      Call setDictItem(sKey, Array(time, 1))
    Else
      Dim vItem: vItem = getDictItem(sKey)
      Dim average As Long: average = vItem(0)
      Dim count As Long: count = vItem(1) + 1
      average = average + (time - average) / count
      Call setDictItem(sKey, Array(average, count))
    End If
  End If
End Sub

'MeasuresClear
'Clears all procedures/blocks and times that have been measured
Public Sub MeasuresClear()
  ReDim FakeDict(0 To 0)
End Sub

'MeasuresKeys
'@returns {string()} - Array containing the procedures or blocks that have been measured.
Public Property Get MeasuresKeys() As String()
  If Me Is stdPerformance Then
    If UBound(FakeDict) = 0 Then
      MeasuresKeys = Split("")
    Else
      'Define return array
      Dim sOut() As String
      ReDim Preserve sOut(0 To UBound(FakeDict) - 1)

      'Fill keys array
      Dim i As Long
      For i = 0 To UBound(FakeDict) - 1
        sOut(i) = FakeDict(i).key
      Next

      'return data
      MeasuresKeys = sOut
    End If
  End If
End Property

'Used by static class only
'@constructor
Private Sub Class_Initialize()
  If Me Is stdPerformance Then
    ReDim FakeDict(0 To 0)
  End If
End Sub

'Used by instance objects only
'@destructor
Private Sub Class_Terminate()
  If Not Me Is stdPerformance Then
    Select Case pObjType
      Case iMeasure
        Dim pEndTime As Long: pEndTime = GetTickCount()
        Call stdPerformance.AddMeasurement(pKeyName, Abs(pEndTime - pStartTime), pDivider)
      Case iOptimiser
        'Store vals
        Application.ScreenUpdating = pScreenUpdating
        Application.EnableEvents = pEnableEvents
        
        'Different options for different applications
        Select Case Application.name
          Case "Microsoft Excel"
            Application.Calculation = pCalculation
        End Select
    End Select
  End If
End Sub




'FakeDict Helpers
'==========================================================================================================================================
'NOTE: These functions are completely unoptimised and are largely in use for the purpose of making this class multi-platform friendly.
'These will be unlikely to be optimised given that this is largely a debugging library.

'getDictIndex
'Returns the index where a particular key is stored
'@param {string} - Key to find in dictionary
'@returns {long} = Index of key in dictionary
Private Function getDictIndex(ByVal key As String) As Long
  On Error GoTo ErrorOccurred
    Dim i As Long
    For i = 0 To UBound(FakeDict)
      If FakeDict(i).key = key Then
        getDictIndex = i
        Exit Function
      End If
    Next
  On Error GoTo 0
ErrorOccurred:
  getDictIndex = -1
End Function

'setDictItem
'Set an item within a dictionary to a particular value
'@param {string} - Key to find in dictionary
'@param {variant} - Value to set dictionary too
'@param {optional long} - Index of item. If not given getDictIndex() is used
Private Sub setDictItem(ByVal key As String, ByVal v As Variant, Optional ByVal ind As Long = -1)
  'get index of item in fake dict
  If ind = -1 Then ind = getDictIndex(key)
  
  'If item not exist, add it
  If ind = -1 Then
    ind = getUB(FakeDict)
    FakeDict(ind).key = key
    ReDim Preserve FakeDict(0 To ind + 1)
  End If

  'Assign value to index
  If isObject(v) Then
    Set FakeDict(ind).val = v
  Else
    FakeDict(ind).val = v
  End If
End Sub

'getUB
'gets the upper bound of an array, if the array is uninitialised return -1
'@param {ByRef FakeDictItem()} Array of dict items
'@returns {Long} - Upper bound of array OR -1 if not initialised
Private Function getUB(ByRef items() As FakeDictItem) As Long
    On Error GoTo ErrorOccurred
        getUB = UBound(items)
        Exit Function
ErrorOccurred:
    getUB = -1
End Function

'getDictIndex
'Returns the item paired with some key
'@param {string} - Key to find in dictionary
'@returns {variant} = Item stored at key
Private Function getDictItem(ByVal key As String) As Variant
  Dim ind As Long: ind = getDictIndex(key)
  If ind <> -1 Then
    If isObject(FakeDict(ind).val) Then
      Set getDictItem = FakeDict(ind).val
    Else
      getDictItem = FakeDict(ind).val
    End If
  Else
    getDictItem = Empty
  End If
End Function
