Attribute VB_Name = "PerformanceTest"
Option Explicit

' https://sancarn.github.io/vba-articles/performance-tips.html

#If VBA7 Then
    Public Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#Else
    Public Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
#End If

Const strWsPerformance As String = "Performance"
Const strWsMeasurements As String = "Measurements"
Const strWsTmp1 As String = "tmp1"
Const strWsTmp2 As String = "tmp2"
Const strHiddenSheet As String = "HiddenSheet"

Const SlowComputer As Long = 10
Const DesiredExecutionTime_ms As Double = 2000

Private wsPerformance As Worksheet
Private wsMeasurements As Worksheet
Private wsTmp1 As Worksheet
Private wsTmp2 As Worksheet
Private wsHiddenSheet As Worksheet
Private Scaling As Double

Sub InitWorksheets()
    Set wsPerformance = ThisWorkbook.Worksheets(strWsPerformance)
    Set wsMeasurements = ThisWorkbook.Worksheets(strWsMeasurements)
    
    With ThisWorkbook
        Set wsTmp1 = .Sheets.Add(After:=.Sheets(.Sheets.count))
        wsTmp1.name = strWsTmp1
        Set wsTmp2 = .Sheets.Add(After:=.Sheets(.Sheets.count))
        wsTmp2.name = strWsTmp2
        Set wsHiddenSheet = .Sheets.Add(After:=.Sheets(.Sheets.count))
        wsHiddenSheet.name = strHiddenSheet
    End With
    
    wsMeasurements.Range("D2:D30").ClearContents
    wsMeasurements.Range("F2:F30").ClearContents
    wsMeasurements.Range("H2:H30").ClearContents
    wsMeasurements.Range("J2:J30").ClearContents
    wsMeasurements.Range("L2:L30").ClearContents
End Sub

Sub PerformanceTest()
    Dim time1_ms As Double
    Dim time2_ms As Double
    Dim ratio_prc As Double
    
    InitWorksheets
    
    wsTmp1.Activate
    
    wsMeasurements.Cells(2, 4) = S1_1(1000)
    wsMeasurements.Cells(2, 6) = S1_2(1000)
    
    ' First measurement is used to calculate the scaling factor to make the running time of measurements on fast and slow computers the same (at least to try to make it the same).
    Scaling = 1
    
    wsMeasurements.Cells(3, 4) = S2_1(CLng(90 * Scaling))
    wsMeasurements.Cells(3, 6) = S2_2(CLng(90 * Scaling))
    wsMeasurements.Cells(4, 4) = S3_1(CLng(17500 * Scaling))
    wsMeasurements.Cells(4, 6) = S3_2(CLng(17500 * Scaling))
    wsMeasurements.Cells(5, 4) = S4a1(CLng(50 * Scaling))
    wsMeasurements.Cells(5, 6) = S4a2(CLng(50 * Scaling))
    wsMeasurements.Cells(5, 8) = S4a3(CLng(50 * Scaling))
    wsMeasurements.Cells(6, 4) = S4b1(CLng(800 * Scaling))
    wsMeasurements.Cells(6, 6) = S4b2(CLng(800 * Scaling))
    wsMeasurements.Cells(6, 8) = S4b3(CLng(800 * Scaling))
    wsMeasurements.Cells(7, 4) = S4c1(CLng(400 * Scaling))
    wsMeasurements.Cells(7, 6) = S4c2(CLng(400 * Scaling))
    wsMeasurements.Cells(8, 4) = S4d1(CLng(10000000 * Scaling))
    wsMeasurements.Cells(8, 6) = S4d2(CLng(10000000 * Scaling))
    wsMeasurements.Cells(8, 8) = S4d3(CLng(2000 * Scaling))
    wsMeasurements.Cells(8, 10) = S4d4(CLng(2000 * Scaling))
    wsMeasurements.Cells(9, 4) = S4e1(CLng(3000000 * Scaling))
    wsMeasurements.Cells(9, 6) = S4e2(CLng(3000000 * Scaling))
    wsMeasurements.Cells(10, 4) = S4f1(CLng(10000000 * Scaling))
    wsMeasurements.Cells(10, 6) = S4f2(CLng(10000000 * Scaling))
    
    ' Option "PrintCommunication" was introduced in Office 2010:
    If CInt(Replace(Application.Version, ".", Application.DecimalSeparator)) > 12 Then
        wsMeasurements.Cells(11, 4) = S4g1(CLng(10000000 * Scaling))
        wsMeasurements.Cells(11, 6) = S4g2(CLng(10000000 * Scaling))
    End If
    
    wsMeasurements.Cells(12, 4) = S5_1(CLng(1500000 * Scaling))
    wsMeasurements.Cells(12, 6) = S5_2(CLng(1500000 * Scaling))
    wsMeasurements.Cells(13, 4) = S6_1(CLng(1000000 * Scaling))
    wsMeasurements.Cells(13, 6) = S6_2(CLng(1000000 * Scaling))
    wsMeasurements.Cells(14, 4) = S7_1(CLng(1500 * Scaling))
    wsMeasurements.Cells(14, 6) = S7_2(CLng(1500 * Scaling))
    
    wsMeasurements.Cells(15, 4) = S8_1(CLng(500000 * Scaling))
    wsMeasurements.Cells(15, 6) = S8_2(CLng(500000 * Scaling))
    wsMeasurements.Cells(15, 8) = S8_3(CLng(500000 * Scaling))
    wsMeasurements.Cells(15, 10) = S8_4(CLng(500000 * Scaling))
    
    wsMeasurements.Cells(16, 4) = S9_1(CLng(40000000 * Scaling))
    wsMeasurements.Cells(16, 6) = S9_2(CLng(40000000 * Scaling))
    wsMeasurements.Cells(17, 4) = S10_1(CLng(1000 * Scaling))
    wsMeasurements.Cells(17, 6) = S10_2(CLng(1000 * Scaling))
    
    wsMeasurements.Cells(18, 4) = S11a1(CLng(200000 * Scaling))
    wsMeasurements.Cells(18, 6) = S11a2(CLng(200000 * Scaling))
    wsMeasurements.Cells(19, 4) = S11b1(CLng(100000 * Scaling))
    wsMeasurements.Cells(19, 6) = S11b2(CLng(100000 * Scaling))
    wsMeasurements.Cells(19, 8) = S11b3(CLng(100000 * Scaling))
    wsMeasurements.Cells(19, 10) = S11b4(CLng(100000 * Scaling))
    wsMeasurements.Cells(20, 4) = S11c1(CLng(1000000 * Scaling))
    wsMeasurements.Cells(20, 6) = S11c2(CLng(1000000 * Scaling))
    wsMeasurements.Cells(21, 4) = S12_1(CLng(4000000 * Scaling))
    wsMeasurements.Cells(21, 6) = S12_2(CLng(4000000 * Scaling))
    
    wsMeasurements.Cells(22, 4) = D1_1(CLng(250000 * Scaling))
    wsMeasurements.Cells(22, 6) = D1_2(CLng(250000 * Scaling))
    wsMeasurements.Cells(22, 8) = D1_3(CLng(250000 * Scaling))
    wsMeasurements.Cells(22, 10) = D1_4(CLng(250000 * Scaling))
    wsMeasurements.Cells(22, 12) = D1_5(CLng(250000 * Scaling))
    
    wsMeasurements.Cells(23, 4) = ST_Concatenation(CLng(140 * Sqr(Scaling)))
    wsMeasurements.Cells(23, 6) = ST_Join(CLng(140 * Sqr(Scaling)))
    
    ' No scaling applied to file writing functions because it is a property of processor, not a hard disc.
    wsMeasurements.Cells(24, 4) = FW1(CLng(2000))
    wsMeasurements.Cells(24, 6) = FW2(CLng(2000))
    wsMeasurements.Cells(24, 8) = FW3(CLng(2000))
    
    Application.DisplayAlerts = False
    wsTmp1.Delete
    wsTmp2.Delete
    wsHiddenSheet.Delete
    Application.DisplayAlerts = True
    
    wsPerformance.Activate
End Sub

Function S1_1(C_MAX As Long) As Double
    Dim i As Long
    
    With stdPerformance.Measure("S1 #1 Select and set", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Select
            Selection.Value = "hello"
        Next
    End With
    
    S1_1 = stdPerformance.Measurement("S1 #1 Select and set")
End Function

Function S1_2(C_MAX As Long) As Double
    Dim i As Long
    
    With stdPerformance.Measure("S1 #2 Set directly", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Value = "hello"
        Next
    End With
    
    S1_2 = stdPerformance.Measurement("S1 #2 Set directly")
End Function

Function S2_1(C_MAX As Long) As Double
    Dim i As Long
    
    With stdPerformance.Measure("S2 #1 Cut and paste", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Cut
            wsTmp1.Cells(1, 2).Select
            ActiveSheet.Paste
        Next
    End With
    
    S2_1 = stdPerformance.Measurement("S2 #1 Cut and paste")
End Function

Function S2_2(C_MAX As Long) As Double
    Dim i As Long
    
    With stdPerformance.Measure("S2 #2 Set directly + clear", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 2).Value = wsTmp1.Cells(1, 1).Value
            wsTmp1.Cells(1, 1).Clear
        Next
    End With
    
    S2_2 = stdPerformance.Measurement("S2 #2 Set directly + clear")
End Function

Function S3_1(C_MAX As Long) As Double
    Dim i As Long, v As Variant, r As Range
    
    With stdPerformance.Measure("S3 #1 Looping through individual cells setting values", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(i, 1).Value2 = Rnd()
        Next
    End With
    
    S3_1 = stdPerformance.Measurement("S3 #1 Looping through individual cells setting values")
End Function

Function S3_2(C_MAX As Long) As Double
    Dim i As Long, v As Variant, r As Range
    
    wsTmp1.Activate
    With stdPerformance.Measure("S3 #2 Exporting array in bulk, set values, Import array in bulk", C_MAX)
        'GetRange
        Set r = wsTmp1.Range("A1").Resize(C_MAX, 10)
        
        'Values of Range --> Array
        v = r.Value2
        
        'Modify array
        For i = 1 To C_MAX  'Using absolute just to be clear no extra work is done, but you'd usually use ubound(v,1)
            v(i, 1) = Rnd()
        Next
        
        'Values of array  -->  Range
        r.Value2 = v
    End With
    
    S3_2 = stdPerformance.Measurement("S3 #2 Exporting array in bulk, set values, Import array in bulk")
End Function

Function S4a1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.ScreenUpdating
    wsTmp1.Range("E1:E1000").Formula = "=RandBetween(1,10)"
    Application.ScreenUpdating = True
    With stdPerformance.Measure("S4a #1 Looping through individual cells setting values", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Value = Empty
        Next
    End With
    Application.ScreenUpdating = prevState
    
    S4a1 = stdPerformance.Measurement("S4a #1 Looping through individual cells setting values")
End Function

Function S4a2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.ScreenUpdating
    wsTmp1.Range("E1:E1000").Formula = "=RandBetween(1,10)"
    With stdPerformance.Measure("S4a #2 w/ ScreenUpdating within loop", C_MAX)
        For i = 1 To C_MAX
            Application.ScreenUpdating = False
            wsTmp1.Cells(1, 1).Value = Empty
            Application.ScreenUpdating = True
        Next
    End With
    Application.ScreenUpdating = prevState
    
    S4a2 = stdPerformance.Measurement("S4a #2 w/ ScreenUpdating within loop")
End Function

Function S4a3(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.ScreenUpdating
    wsTmp1.Range("E1:E1000").Formula = "=RandBetween(1,10)"
    Application.ScreenUpdating = True
    With stdPerformance.Measure("S4a #3 w/ ScreenUpdating", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Value = Empty
        Next
    End With
    Application.ScreenUpdating = prevState
    wsTmp1.Range("E1:E1000").ClearContents
    
    S4a3 = stdPerformance.Measurement("S4a #3 w/ ScreenUpdating")
End Function

Function S4b1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.EnableEvents
    Application.EnableEvents = True
    With stdPerformance.Measure("S4 #1 Looping through individual cells setting values", C_MAX)
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Value = Empty
        Next
    End With
    Application.EnableEvents = prevState
    
    S4b1 = stdPerformance.Measurement("S4 #1 Looping through individual cells setting values")
End Function

Function S4b2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.EnableEvents
    With stdPerformance.Measure("S4b #2 w/ EnableEvents", C_MAX)
        Application.EnableEvents = False
        For i = 1 To C_MAX
            wsTmp1.Cells(1, 1).Value = Empty
        Next
        Application.EnableEvents = True
    End With
    Application.EnableEvents = prevState
    
    S4b2 = stdPerformance.Measurement("S4b #2 w/ EnableEvents")
End Function

Function S4b3(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.EnableEvents
    With stdPerformance.Measure("S4b #3 w/ EnableEvents within loop", C_MAX)
        For i = 1 To C_MAX
            Application.EnableEvents = False
            wsTmp1.Cells(1, 1).Value = Empty
            Application.EnableEvents = True
        Next
    End With
    Application.EnableEvents = prevState
    
    S4b3 = stdPerformance.Measurement("S4b #3 w/ EnableEvents within loop")
End Function

Function S4c1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Long
    Dim rCell As Range: Set rCell = wsTmp1.Range("A1")
    
    prevState = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    wsTmp1.Range("E1:E1000").Formula = "=RandBetween(1,10)"
    With stdPerformance.Measure("S4c #1 Calculation = xlCalculationAutomatic", C_MAX)
        For i = 1 To C_MAX
            rCell.Formula = "=1"
        Next
    End With
    Application.Calculation = prevState
    
    S4c1 = stdPerformance.Measurement("S4c #1 Calculation = xlCalculationAutomatic")
End Function

Function S4c2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Long
    Dim rCell As Range: Set rCell = wsTmp1.Range("A1")
    
    prevState = Application.Calculation
    wsTmp1.Range("E1:E1000").Formula = "=RandBetween(1,10)"
    With stdPerformance.Measure("S4c #2 Calculation = xlCalculationManual", C_MAX)
        Application.Calculation = xlCalculationManual
        For i = 1 To C_MAX
            rCell.Formula = "=1"
        Next
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
    End With
    Application.Calculation = prevState
    wsTmp1.Range("E1:E1000").ClearContents
    
    S4c2 = stdPerformance.Measurement("S4c #2 Calculation = xlCalculationManual")
End Function

Function S4d1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("S4d #1 Don't change StatusBar", C_MAX)
            For i = 1 To C_MAX
                v = ""
            Next
        End With
    End With
    Application.DisplayStatusBar = prevState
    
    S4d1 = stdPerformance.Measurement("S4d #1 Don't change StatusBar")
End Function

Function S4d2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("S4d #2 Change StatusBar setting", C_MAX)
            Application.DisplayStatusBar = False
            For i = 1 To C_MAX
                v = ""
            Next
            Application.DisplayStatusBar = True
        End With
    End With
    Application.DisplayStatusBar = prevState
    
    S4d2 = stdPerformance.Measurement("S4d #2 Change StatusBar setting")
End Function

Function S4d3(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.DisplayStatusBar
    With stdPerformance.Measure("S4d #3 Don't change StatusBar", C_MAX)
        For i = 1 To C_MAX
            Application.StatusBar = i
        Next
    End With
    Application.DisplayStatusBar = prevState
    
    S4d3 = stdPerformance.Measurement("S4d #3 Don't change StatusBar")
End Function

Function S4d4(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.DisplayStatusBar
    With stdPerformance.Measure("S4d #4 Change StatusBar setting", C_MAX)
        Application.DisplayStatusBar = False
        For i = 1 To C_MAX
            Application.StatusBar = i
        Next
        Application.DisplayStatusBar = True
    End With
    Application.DisplayStatusBar = prevState
    
    S4d4 = stdPerformance.Measurement("S4d #4 Change StatusBar setting")
End Function

Function S4e1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = wsPerformance.DisplayPageBreaks
    wsPerformance.DisplayPageBreaks = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("#1 w/o change DisplayPageBreaks setting", C_MAX)
            For i = 1 To C_MAX
                v = ""
            Next
        End With
    End With
    wsPerformance.DisplayPageBreaks = prevState
    
    S4e1 = stdPerformance.Measurement("#1 w/o change DisplayPageBreaks setting")
End Function

Function S4e2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = wsPerformance.DisplayPageBreaks
    wsPerformance.DisplayPageBreaks = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("#2 w/ change DisplayPageBreaks setting", C_MAX)
            wsPerformance.DisplayPageBreaks = False
            For i = 1 To C_MAX
                v = ""
            Next
            wsPerformance.DisplayPageBreaks = True
        End With
    End With
    wsPerformance.DisplayPageBreaks = prevState
    
    S4e2 = stdPerformance.Measurement("#2 w/ change DisplayPageBreaks setting")
End Function

Function S4f1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.EnableAnimations
    Application.EnableAnimations = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("#1 w/o change EnableAnimations setting", C_MAX)
            For i = 1 To C_MAX
                v = ""
            Next
        End With
    End With
    Application.EnableAnimations = prevState
    
    S4f1 = stdPerformance.Measurement("#1 w/o change EnableAnimations setting")
End Function

Function S4f2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.EnableAnimations
    Application.EnableAnimations = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("#2 w/ change EnableAnimations setting", C_MAX)
            Application.EnableAnimations = False
            For i = 1 To C_MAX
                v = ""
            Next
            Application.EnableAnimations = True
        End With
    End With
    Application.EnableAnimations = prevState
    
    S4f2 = stdPerformance.Measurement("#2 w/ change EnableAnimations setting")
End Function

Function S4g1(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.PrintCommunication
    Application.PrintCommunication = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("#1 w/o change PrintCommunication setting", C_MAX)
            For i = 1 To C_MAX
                v = ""
            Next
        End With
    End With
    Application.PrintCommunication = prevState
    
    S4g1 = stdPerformance.Measurement("#1 w/o change PrintCommunication setting")
End Function

Function S4g2(C_MAX As Long) As Double
    Dim i As Long
    Dim prevState As Boolean
    
    prevState = Application.PrintCommunication
    Application.PrintCommunication = True
    With stdPerformance.Optimise   'Disable screen updating and application events
        Dim v
        With stdPerformance.Measure("#2 w/ change PrintCommunication setting", C_MAX)
            Application.PrintCommunication = False
            For i = 1 To C_MAX
                v = ""
            Next
            Application.PrintCommunication = True
        End With
    End With
    Application.PrintCommunication = prevState
    
    S4g2 = stdPerformance.Measurement("#2 w/ change PrintCommunication setting")
End Function

Function S5_1(C_MAX As Long) As Double
    Dim o As New oClass
    Dim i As Long
    Dim v As Boolean
    
    With stdPerformance.Measure("D10, No with block", C_MAX)
        For i = 1 To C_MAX
            v = o.Self.Self.Self.Self.Self.Self.Self.Self.Self.data
        Next
    End With
    
    S5_1 = stdPerformance.Measurement("D10, No with block")
End Function

Function S5_2(C_MAX As Long) As Double
    Dim o As New oClass
    Dim i As Long
    Dim v As Boolean
    
    With stdPerformance.Measure("D10, With block", C_MAX)
        With o.Self.Self.Self.Self.Self.Self.Self.Self.Self
            For i = 1 To C_MAX
                v = .data
            Next
        End With
    End With
    
    S5_2 = stdPerformance.Measurement("D10, With block")
End Function

Function S6_1(C_MAX As Long) As Double
    Dim i As Long, r1 As Object, r2 As VBScript_RegExp_55.RegExp
    
    Set r1 = CreateObject("VBScript.Regexp")
    With stdPerformance.Measure("#B-1 Late bound calls", C_MAX)
        For i = 1 To C_MAX
            r1.Pattern = "something"
        Next
    End With
    
    S6_1 = stdPerformance.Measurement("#B-1 Late bound calls")
End Function

Function S6_2(C_MAX As Long) As Double
    Dim i As Long, r1 As Object, r2 As VBScript_RegExp_55.RegExp
    
    Set r2 = New VBScript_RegExp_55.RegExp
    With stdPerformance.Measure("#B-2 Early bound calls", C_MAX)
        For i = 1 To C_MAX
            r2.Pattern = "something"
        Next
    End With
    
    S6_2 = stdPerformance.Measurement("#B-2 Early bound calls")
End Function

Function S7_1(C_MAX As Long) As Double
    Dim v() As String, i As Long
    
    v = Split(Space(1000), " ")
    With stdPerformance.Measure("#1 `ByVal`", C_MAX)
        For i = 1 To C_MAX
            Call testByVal(v)
        Next
    End With
    
    S7_1 = stdPerformance.Measurement("#1 `ByVal`")
End Function

Function S7_2(C_MAX As Long) As Double
    Dim v() As String, i As Long
    
    v = Split(Space(1000), " ")
    With stdPerformance.Measure("#2 `ByRef`", C_MAX)
        For i = 1 To C_MAX
            Call testByRef(v)
        Next
    End With
    
    S7_2 = stdPerformance.Measurement("#2 `ByRef`")
End Function

Sub testByVal(ByVal v)
    wsTmp1.Cells(1, 1).Value2 = Rnd()
End Sub

Sub testByRef(ByRef v)
    wsTmp1.Cells(1, 1).Value2 = Rnd()
End Sub

Function S8_1(C_MAX As Long) As Double
    Dim i As Long
    Dim c1 As New Car1
    
    With stdPerformance.Measure("A-#1 Object creation (Class)", C_MAX)
        For i = 1 To C_MAX
            Set c1 = c1.Create("hello", 10, 2)
'            Set c1 = Car1.Create("hello", 10, 2)
        Next
    End With
    
    S8_1 = stdPerformance.Measurement("A-#1 Object creation (Class)")
End Function

Function S8_2(C_MAX As Long) As Double
    Dim i As Long
    Dim c2 As Car2.CarData
    
    With stdPerformance.Measure("A-#2 Object creation (Module)", C_MAX)
        For i = 1 To C_MAX
            c2 = Car2.Car_Create("hello", 10, 2)
        Next
    End With
    
    S8_2 = stdPerformance.Measurement("A-#2 Object creation (Module)")
End Function

Function S8_3(C_MAX As Long) As Double
    Dim i As Long
    Dim c1 As New Car1
    
    'Objects for instance tests
    Set c1 = c1.Create("hello", 10, 2)
    
    'Test calling public methods speeds
    With stdPerformance.Measure("B-#1 Object method calls (Class)", C_MAX)
        For i = 1 To C_MAX
            Call c1.Tick
        Next
    End With
    
    S8_3 = stdPerformance.Measurement("B-#1 Object method calls (Class)")
End Function

Function S8_4(C_MAX As Long) As Double
    Dim i As Long
    Dim c2 As Car2.CarData
    
    'Objects for instance tests
    c2 = Car2.Car_Create("hello", 10, 2)
    
    'Test calling public methods speeds
    With stdPerformance.Measure("B-#2 Object method calls (Module)", C_MAX)
        For i = 1 To C_MAX
            Call Car2.Car_Tick(c2)
        Next
    End With
    
    S8_4 = stdPerformance.Measurement("B-#2 Object method calls (Module)")
End Function

Function S9_1(C_MAX As Long) As Double
    Dim i As Long
    
    With stdPerformance.Measure("#1 - Variant", C_MAX)
        Dim v() As Variant
        ReDim v(1 To C_MAX)
        For i = 1 To C_MAX
            v(i) = i
        Next
    End With
    
    S9_1 = stdPerformance.Measurement("#1 - Variant")
End Function

Function S9_2(C_MAX As Long) As Double
    Dim i As Long
    
    With stdPerformance.Measure("#2 - Type", C_MAX)
        Dim l() As Long
        ReDim l(1 To C_MAX)
        For i = 1 To C_MAX
            l(i) = i
        Next
    End With
    
    S9_2 = stdPerformance.Measurement("#2 - Type")
End Function

Function S10_1(C_MAX As Long) As Double
    Dim i As Long, Rng As Range
    
    wsPerformance.Activate
    
    wsTmp1.Range("A1:X" & C_MAX).Value = "Some cool data here"
    
    With stdPerformance.Measure("#1 Delete rows 1 by 1", C_MAX)
        For i = C_MAX To 1 Step -1
            'Delete only even rows
            If i Mod 2 = 0 Then
                wsTmp1.Rows(i).Delete
            End If
        Next
    End With
    
    S10_1 = stdPerformance.Measurement("#1 Delete rows 1 by 1")
End Function

Function S10_2(C_MAX As Long) As Double
    Dim i As Long, Rng As Range
    
    wsPerformance.Activate
    
    wsTmp1.Range("A1:X" & C_MAX).Value = "Some cool data here"
    
    With stdPerformance.Measure("#2 Delete all rows in a single operation less branching", C_MAX)
        Set Rng = wsTmp1.Cells(Rows.count, 1)
        For i = 1 To C_MAX
            If i Mod 2 = 0 Then
                Set Rng = Application.Union(Rng, wsTmp1.Rows(i))
            End If
        Next i
        Set Rng = Application.Intersect(Rng, wsTmp1.Range("1:" & C_MAX))
        Rng.Delete
    End With
    
    S10_2 = stdPerformance.Measurement("#2 Delete all rows in a single operation less branching")
End Function

Function S11a1(C_MAX As Long) As Double
    'Initialisation. Initialise a sheet containing data and an output sheet containing headers.
    Dim arr As Variant: arr = getArray(C_MAX)
    Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
    wsTmp1.Range("A1").Resize(iNumRow, iNumCol).Value = arr
    wsTmp2.UsedRange.Clear
    
    With stdPerformance.Optimise()
        'Copy data from sheet into an array,
        'loop through rows, move filtered rows to top of array,
        'only return required size of array
        With stdPerformance.Measure("#1 Use of array")
            Dim v: v = wsTmp1.Range("A1").CurrentRegion.Value
            Dim iRowLen As Long: iRowLen = UBound(v, 1)
            Dim iColLen As Long: iColLen = UBound(v, 2)
            
            Dim i As Long, j As Long, iRet As Long
            iRet = 1
            For i = 2 To iRowLen
                If v(i, 2) = "A" Then
                    iRet = iRet + 1
                    If iRet < i Then
                        For j = 1 To iColLen
                            v(iRet, j) = v(i, j)
                        Next
                    End If
                End If
            Next
            
            wsTmp2.Range("A1").Resize(iRet, iColLen).Value = v
        End With
    End With
    
    S11a1 = stdPerformance.Measurement("#1 Use of array")
End Function

Function S11a2(C_MAX As Long) As Double
    'Initialisation. Initialise a sheet containing data and an output sheet containing headers.
    Dim arr As Variant: arr = getArray(C_MAX)
    Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
    wsTmp1.Range("A1").Resize(iNumRow, iNumCol).Value = arr
    wsTmp2.UsedRange.Clear
    
    With stdPerformance.Optimise()
        'Use advanced filter, copy result and paste to new location. Use range.currentRegion.value to obtain result
        wsTmp2.UsedRange.Clear
        With stdPerformance.Measure("#2 Advanced filter and copy")
            'Choose headers
            wsTmp2.Range("A1:B1").Value = Array("ID", "Key")
            
            'Choose filter
            wsHiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
            'filter and copy data
            With wsTmp1.Range("A1").CurrentRegion
                Call .AdvancedFilter(xlFilterCopy, wsHiddenSheet.Range("A1:A2"), wsTmp2.Range("A1:B1"))
            End With
            
            'Cleanup
            wsHiddenSheet.UsedRange.Clear
        End With
    End With
    
    S11a2 = stdPerformance.Measurement("#2 Advanced filter and copy")
End Function

Function S11b1(C_MAX As Long) As Double
    'Obtain test data
    Dim arr As Variant: arr = getArray(C_MAX)
    
    'Some of these tests may take some time, so optimise to ensure these don't have an impact
    With stdPerformance.Optimise()
        'Use advanced filter, copy result and paste to new location. Use range.currentRegion.value to obtain result
        Dim vResult
        With stdPerformance.Measure("#1 Advanced filter and copy - array result")
            'Get data dimensions
            Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
            Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
            
            'Create filters data
            wsHiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
            With wsHiddenSheet.Range("A4").Resize(iNumRow, iNumCol)
                'Dump data to sheet
                .Value = arr
                
                'Call advanced filter
                Call .AdvancedFilter(xlFilterInPlace, wsHiddenSheet.Range("A1:A2"))
                
                'Get result
                .Resize(, 1).Copy wsHiddenSheet.Range("D4")
                vResult = wsHiddenSheet.Range("D4").CurrentRegion.Value
            End With
            
            'Cleanup
            wsHiddenSheet.ShowAllData
            wsHiddenSheet.UsedRange.Clear
        End With
    End With
    
    S11b1 = stdPerformance.Measurement("#1 Advanced filter and copy - array result")
End Function

Function S11b2(C_MAX As Long) As Double
    'Obtain test data
    Dim arr As Variant: arr = getArray(C_MAX)
    
    'Some of these tests may take some time, so optimise to ensure these don't have an impact
    With stdPerformance.Optimise()
        'Use advanced filter, extract results by looping over the range areas
        Dim vResult2() As Variant
        With stdPerformance.Measure("#2 Advanced filter and areas loop - array result")
            'get dimensions
            Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
            Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
            
            'Store capacity for at least iNumRow items in result
            ReDim vResult2(1 To iNumRow)
            
            'Create filters data
            wsHiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
            With wsHiddenSheet.Range("A4").Resize(iNumRow, iNumCol)
                'Set data and call filter
                .Value = arr
                Call .AdvancedFilter(xlFilterInPlace, wsHiddenSheet.Range("A1:A2"))
                
                'Loop over all visible cells and dump data to array
                Dim rArea As Range, vArea As Variant, i As Long, iRes As Long
                iRes = 0
                For Each rArea In .Resize(, 1).SpecialCells(xlCellTypeVisible).Areas
                    vArea = rArea.Value
                    If rArea.CountLarge = 1 Then
                        iRes = iRes + 1
                        vResult2(iRes) = vArea
                    Else
                        For i = 1 To UBound(vArea, 1)
                            iRes = iRes + 1
                            vResult2(iRes) = vArea(i, 1)
                        Next
                    End If
                Next
                
                'Trim size of array to total number of inserted elements
                ReDim Preserve vResult2(1 To iRes)
            End With
            
            'Cleanup
            wsHiddenSheet.ShowAllData
            wsHiddenSheet.UsedRange.Clear
        End With
    End With
    
    S11b2 = stdPerformance.Measurement("#2 Advanced filter and areas loop - array result")
End Function

Function S11b3(C_MAX As Long) As Double
    'Obtain test data
    Dim arr As Variant: arr = getArray(C_MAX)
    
    'Some of these tests may take some time, so optimise to ensure these don't have an impact
    With stdPerformance.Optimise()
        'Use a VBA filter
        Dim vResult3() As Variant, iRes As Long, i As Long
        With stdPerformance.Measure("#3 Array filter - array result")
            'Get total row count
            Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
            
            'Make result at least the same number of rows as the base data (We can't return more data than rows in our source data)
            ReDim vResult3(1 To iNumRow)
            
            'Loop over rows, filter condition and assign to result
            For i = 1 To iNumRow
                If arr(i, 2) = "A" Then
                    iRes = iRes + 1
                    vResult3(iRes) = arr(i, 1)
                End If
            Next
            
            'Trim array to total result size
            ReDim Preserve vResult3(1 To iRes)
        End With
    End With
    
    S11b3 = stdPerformance.Measurement("#3 Array filter - array result")
End Function

Function S11b4(C_MAX As Long) As Double
    'Obtain test data
    Dim arr As Variant: arr = getArray(C_MAX)
    
    'Some of these tests may take some time, so optimise to ensure these don't have an impact
    With stdPerformance.Optimise()
        'Use a VBA filter - return a collection
        'This algorithm is much the same as the above, however we simply add results to a collection instead of to an array. Collections are generally a fast way to have dynamic sizing data.
        Dim cResult As Collection, i As Long
        With stdPerformance.Measure("#4 Array filter - collection result")
            Set cResult = New Collection
            Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
            For i = 1 To iNumRow
                If arr(i, 2) = "A" Then
                    cResult.Add arr(i, 1)
                End If
            Next
        End With
    End With
    
    S11b4 = stdPerformance.Measurement("#4 Array filter - collection result")
End Function

Function S11c1(C_MAX As Long) As Double
    'Initialisation
    Dim arr As Variant: arr = getArray(C_MAX)
    Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
    wsTmp1.Range("A1").Resize(iNumRow, iNumCol).Value = arr
    
    With stdPerformance.Optimise()
        'Use advanced filter, copy result and paste to new location. Use range.currentRegion.value to obtain result
        Dim vResult1
        With stdPerformance.Measure("#1 Advanced filter and copy")
            'Choose output headers
            wsHiddenSheet.Range("A4:B4").Value = Array("ID", "Key")
            
            'Choose filters
            wsHiddenSheet.Range("A1:A2").Value = Application.Transpose(Array("Key", "A"))
            
            'Call advanced filter
            With wsTmp1.Range("A1").CurrentRegion
                Call .AdvancedFilter(xlFilterCopy, wsHiddenSheet.Range("A1:A2"), wsHiddenSheet.Range("A4:B4"))
            End With
            
            'Obtain results
            vResult1 = wsHiddenSheet.Range("A4").CurrentRegion.Value
            
            'Cleanup
            wsHiddenSheet.UsedRange.Clear
        End With
    End With
    
    S11c1 = stdPerformance.Measurement("#1 Advanced filter and copy")
End Function

Function S11c2(C_MAX As Long) As Double
    'Initialisation
    Dim arr As Variant: arr = getArray(C_MAX)
    Dim iNumRow As Long: iNumRow = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim iNumCol As Long: iNumCol = UBound(arr, 2) - LBound(arr, 2) + 1
    wsTmp1.Range("A1").Resize(iNumRow, iNumCol).Value = arr
    
    With stdPerformance.Optimise()
        'Array
        'Copy data from sheet into an array,
        'loop through rows, move filtered rows to top of array,
        'in this scenario remove value from any row which isn't required.
        Dim vResult2
        With stdPerformance.Measure("#2 Use of array")
            vResult2 = wsTmp1.Range("A1").CurrentRegion.Value
            Dim iRowLen As Long: iRowLen = UBound(vResult2, 1)
            Dim iColLen As Long: iColLen = UBound(vResult2, 2)
            
            Dim i As Long, j As Long, iRet As Long
            iRet = 1
            For i = 2 To iRowLen
                If vResult2(i, 2) = "A" Then
                    iRet = iRet + 1
                    If iRet < i Then
                        For j = 1 To iColLen
                            vResult2(iRet, j) = vResult2(i, j)
                            vResult2(i, j) = Empty
                        Next
                    End If
                Else
                    vResult2(i, 1) = Empty
                    vResult2(i, 2) = Empty
                End If
            Next
        End With
    End With
    
    S11c2 = stdPerformance.Measurement("#2 Use of array")
End Function

'Obtain an array of data to C_MAX size.
'@param {Long} Max length of data
'@returns {Variant(nx2)} Returns an array of data 2 columns wide and n rows deep.
Public Function getArray(C_MAX As Long) As Variant
    Dim arr() As Variant
    ReDim arr(1 To C_MAX, 1 To 2)
    
    arr(1, 1) = "ID"
    arr(1, 2) = "Key"
    Dim i As Long
    For i = 2 To C_MAX
        'ID
        arr(i, 1) = i
        Select Case True
            Case i Mod 17 = 0: arr(i, 2) = "A"
            Case i Mod 13 = 0: arr(i, 2) = "B"
            Case i Mod 11 = 0: arr(i, 2) = "C"
            Case i Mod 7 = 0: arr(i, 2) = "D"
            Case Else
                arr(i, 2) = "E"
        End Select
    Next
    getArray = arr
End Function

Function S12_1(C_MAX As Long) As Double
    Dim v1, v2, i As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("#1 VBA", C_MAX)
            For i = 1 To C_MAX
                Call VariantCopyVBA(v1, v2)
            Next
        End With
    End With
    
    S12_1 = stdPerformance.Measurement("#1 VBA")
End Function

Function S12_2(C_MAX As Long) As Double
    Dim v1, v2, i As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("#2 DLL", C_MAX)
            For i = 1 To C_MAX
                Call VariantCopy(v1, v2)
            Next
        End With
    End With
    
    S12_2 = stdPerformance.Measurement("#2 DLL")
End Function

Public Sub VariantCopyVBA(ByRef v1, ByVal v2)
    If isObject(v2) Then
        Set v1 = v2
    Else
        v1 = v2
    End If
End Sub

Function D1_1(C_MAX As Long) As Double
    'Arrays vs dictionary
    Dim arr: arr = getArray(C_MAX)
    Dim arrLookup: arrLookup = getLookupArray()
    Dim dictLookup: Set dictLookup = getLookupDict()
    Dim dictLookup2 As Dictionary: Set dictLookup2 = dictLookup
    Dim i As Long, j As Long, iVal As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("D1 #1 Lookup in array - Naïve approach", C_MAX)
            For i = 2 To C_MAX
                'Lookup key in arrLookup
                For j = 1 To 5
                    If arr(i, 2) = arrLookup(j, 1) Then
                        iVal = arrLookup(j, 2)
                        Exit For
                    End If
                Next
            Next
        End With
    End With
    
    D1_1 = stdPerformance.Measurement("D1 #1 Lookup in array - Naïve approach")
End Function

Function D1_2(C_MAX As Long) As Double
    'Arrays vs dictionary
    Dim arr: arr = getArray(C_MAX)
    Dim arrLookup: arrLookup = getLookupArray()
    Dim dictLookup: Set dictLookup = getLookupDict()
    Dim dictLookup2 As Dictionary: Set dictLookup2 = dictLookup
    Dim i As Long, j As Long, iVal As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("D1 #2 Lookup in dictionary - late binding", C_MAX)
            For i = 2 To C_MAX
                'Lookup key in dict
                iVal = dictLookup(arr(i, 2))
            Next
        End With
    End With
    
    D1_2 = stdPerformance.Measurement("D1 #2 Lookup in dictionary - late binding")
End Function

Function D1_3(C_MAX As Long) As Double
    'Arrays vs dictionary
    Dim arr: arr = getArray(C_MAX)
    Dim arrLookup: arrLookup = getLookupArray()
    Dim dictLookup: Set dictLookup = getLookupDict()
    Dim dictLookup2 As Dictionary: Set dictLookup2 = dictLookup
    Dim i As Long, j As Long, iVal As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("D1 #3 Lookup in dictionary - early binding", C_MAX)
            For i = 2 To C_MAX
                'Lookup key in dict
                iVal = dictLookup2(arr(i, 2))
            Next
        End With
    End With
    
    D1_3 = stdPerformance.Measurement("D1 #3 Lookup in dictionary - early binding")
End Function

Function D1_4(C_MAX As Long) As Double
    'Arrays vs dictionary
    Dim arr: arr = getArray(C_MAX)
    Dim arrLookup: arrLookup = getLookupArray()
    Dim dictLookup: Set dictLookup = getLookupDict()
    Dim dictLookup2 As Dictionary: Set dictLookup2 = dictLookup
    Dim i As Long, j As Long, iVal As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("D1 #4 Generate through logic", C_MAX)
            For i = 2 To C_MAX
                'Generate value from key
                iVal = getLookupFromCalc(arr(i, 2))
            Next
        End With
    End With
    
    D1_4 = stdPerformance.Measurement("D1 #4 Generate through logic")
End Function

Function D1_5(C_MAX As Long) As Double
    'Arrays vs dictionary
    Dim arr: arr = getArray(C_MAX)
    Dim arrLookup: arrLookup = getLookupArray()
    Dim dictLookup: Set dictLookup = getLookupDict()
    Dim dictLookup2 As Dictionary: Set dictLookup2 = dictLookup
    Dim i As Long, j As Long, iVal As Long
    
    With stdPerformance.Optimise
        With stdPerformance.Measure("D1 #5 Generate through logic direct", C_MAX)
            For i = 2 To C_MAX
                'Generate value from key
                Dim iChar As Long: iChar = Asc(arr(i, 2)) - 64
                iVal = iChar * 10
                If iChar = 5 Then iVal = 99
            Next
        End With
    End With
    
    D1_5 = stdPerformance.Measurement("D1 #5 Generate through logic direct")
End Function

'Obtain dictionary to lookup key,value pairs
Public Function getLookupDict() As Object
    Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
    o("A") = 10: o("B") = 20: o("C") = 30: o("D") = 40: o("E") = 99
    Set getLookupDict = o
End Function

'Obtain array to lookup key,value pairs
Public Function getLookupArray() As Variant
    Dim arr()
    ReDim arr(1 To 5, 1 To 2)
    arr(1, 1) = "A": arr(1, 2) = 10
    arr(2, 1) = "B": arr(2, 2) = 20
    arr(3, 1) = "C": arr(3, 2) = 30
    arr(4, 1) = "D": arr(4, 2) = 40
    arr(5, 1) = "E": arr(5, 2) = 99
    getLookupArray = arr
End Function

'Obtain value from key
Public Function getLookupFromCalc(ByVal key As String) As Long
    Dim iChar As Long: iChar = Asc(key) - 64
    getLookupFromCalc = iChar * 10
    If iChar = 5 Then getLookupFromCalc = 99
End Function

' Creates concatenation vs join graph.
Sub ConcatenationVsJoinGraph()
    Dim time1_ms As Double
    Dim time2_ms As Double
    Dim ratio_prc As Double
    Dim wsStrings As Worksheet
    Dim i As Long
    Dim j As Long
    
    Set wsPerformance = ThisWorkbook.Worksheets(strWsPerformance)
    Set wsStrings = ThisWorkbook.Worksheets("Concatenation vs Join")
    
    For j = 1 To 10
        For i = 100 To 180 Step 10
            wsStrings.Cells(i / 10 - 8, 2) = i
            wsStrings.Cells(i / 10 - 8, 4) = ST_Concatenation(i)
            wsStrings.Cells(i / 10 - 8, 5) = ST_Join(i)
        Next i
    Next j
End Sub

' Strings: Concatenation.
Public Function ST_Concatenation(C_MAX As Long) As Double
    Dim i&, j&, stringContent$
    Dim notBig2DArray() As Variant
    ReDim notBig2DArray(1 To C_MAX, 1 To C_MAX)
    
    For i = 1 To C_MAX
        For j = 1 To C_MAX
            notBig2DArray(i, j) = Rnd
        Next j
    Next i
    
    stringContent = ""
    With stdPerformance.Measure("ST #1 Concatenation " & C_MAX, C_MAX ^ 2)
        For i = 1 To C_MAX
            For j = 1 To C_MAX
                stringContent = stringContent & notBig2DArray(i, j) & ", "
            Next j
            stringContent = stringContent & vbCrLf
        Next i
    End With
    
    ST_Concatenation = stdPerformance.Measurement("ST #1 Concatenation " & C_MAX)
End Function

' Strings: Join.
Public Function ST_Join(C_MAX As Long) As Double
    Dim i&, j&, stringContent$
    Dim notBig2DArray() As Variant
    ReDim notBig2DArray(1 To C_MAX, 1 To C_MAX)
    
    For i = 1 To C_MAX
        For j = 1 To C_MAX
            notBig2DArray(i, j) = Rnd
        Next j
    Next i
    
    stringContent = ""
    Dim tmpStrings() As Variant
    ReDim tmpStrings(1 To C_MAX)
    Dim tmpDoubles() As Variant
    ReDim tmpDoubles(1 To C_MAX)
    
    With stdPerformance.Measure("ST #2 Join " & C_MAX, C_MAX ^ 2)
        For i = 1 To C_MAX
            For j = 1 To C_MAX
                tmpDoubles(j) = notBig2DArray(i, j)
            Next j
            tmpStrings(i) = Join(tmpDoubles, ", ")
        Next i
        stringContent = Join(tmpStrings, vbCrLf)
    End With
    
    ST_Join = stdPerformance.Measurement("ST #2 Join " & C_MAX)
End Function

' Creates random array and returns it as a string.
Private Function RandomArray2String(C_MAX As Long)
    Dim i&, j&, stringContent$
    Dim notBig2DArray() As Variant
    ReDim notBig2DArray(1 To C_MAX, 1 To C_MAX)
    
    For i = 1 To C_MAX
        For j = 1 To C_MAX
            notBig2DArray(i, j) = Rnd
        Next j
    Next i
    
    RandomArray2String = ""
    Dim tmpStrings() As Variant
    ReDim tmpStrings(1 To C_MAX)
    Dim tmpDoubles() As Variant
    ReDim tmpDoubles(1 To C_MAX)
    With stdPerformance.Measure("#2 Join " & C_MAX, C_MAX ^ 2)
        For i = 1 To C_MAX
            For j = 1 To C_MAX
                tmpDoubles(j) = notBig2DArray(i, j)
            Next j
            tmpStrings(i) = Join(tmpDoubles, ", ")
        Next i
        RandomArray2String = Join(tmpStrings, vbCrLf)
    End With
End Function

' Files: Standard writing.
' C_MAX As Long = 2000 corresponds to file size of 44.6 Mb.
Function FW1(C_MAX As Long) As Double
    Const FileName = "anylongnametonotoverwriteyourfiles"
    Dim strFileContent$
    Dim iFile As Integer
    
    strFileContent = RandomArray2String(C_MAX)
    
    ' standard writing function
    With stdPerformance.Measure("FW #1 Standard writing function", C_MAX)
        iFile = FreeFile
        Open Application.ActiveWorkbook.Path & "\" & FileName For Output As #iFile
        Print #iFile, strFileContent
        Close #iFile
    End With
    Kill Application.ActiveWorkbook.Path & "\" & FileName
    
    FW1 = stdPerformance.Measurement("FW #1 Standard writing function")
End Function

' Files: FSO.
' C_MAX As Long = 2000 corresponds to file size of 44.6 Mb.
Function FW2(C_MAX As Long) As Double
    Const FileName = "anylongnametonotoverwriteyourfiles"
    Dim strFileContent$
    Dim fso As Object
    Dim oFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strFileContent = RandomArray2String(C_MAX)
    
    ' FSO
    With stdPerformance.Measure("FW #2 FSO", C_MAX)
        Set oFile = fso.CreateTextFile(Application.ActiveWorkbook.Path & "\" & FileName)
        oFile.WriteLine strFileContent
        oFile.Close
        Set fso = Nothing
        Set oFile = Nothing
    End With
    Kill Application.ActiveWorkbook.Path & "\" & FileName
    
    FW2 = stdPerformance.Measurement("FW #2 FSO")
End Function

' Files: Binary writing.
' C_MAX As Long = 2000 corresponds to file size of 44.6 Mb.
Function FW3(C_MAX As Long) As Double
    Const FileName = "anylongnametonotoverwriteyourfiles"
    Dim strFileContent$
    Dim iFile As Integer
    
    strFileContent = RandomArray2String(C_MAX)
    
    With stdPerformance.Measure("FW #3 Binary writing function", C_MAX)
        iFile = FreeFile
        Open Application.ActiveWorkbook.Path & "\" & FileName For Binary As #iFile
        Put #iFile, , strFileContent
        Close #iFile
    End With
    Kill Application.ActiveWorkbook.Path & "\" & FileName
    
    FW3 = stdPerformance.Measurement("FW #3 Binary writing function")
End Function
