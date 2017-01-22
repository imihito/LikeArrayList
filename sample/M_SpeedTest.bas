Attribute VB_Name = "M_SpeedTest"
Option Explicit

Private StartTime As Single

Sub SpeedTest()
    Dim AList As Object
    Set AList = CreateObject("System.Collections.ArrayList")
    
    Dim LAList As LikeArrayList
    Set LAList = New LikeArrayList
    
    Dim i As Long
    Const MaxLoop As Long = 100000
    
    Start
    AList.Capacity = MaxLoop
    For i = 1 To MaxLoop
        AList.Add i
    Next i
    Dump "ArrayListAdd"
    
    
    'Debug.Print LAList.ArrayType   'Empty
    Dim tmp() As Long
    Call LAList.InitInternalArray(tmp)
    'Debug.Print LAList.ArrayType   'Long
    
    Start
    For i = 1 To MaxLoop
        LAList.AddVal i
    Next i
    Dump "LAListAddVal"
    
    LAList.Clear (True)
    Start
    For i = 1 To MaxLoop
        LAList.Add i
    Next i
    Dump "LAListAdd"
    
    Dim buf As Variant
    Start
    For i = 0 To AList.Count - 1
        buf = AList.Item(i)
    Next i
    Dump "ArrayListItem"
    
    Start
    For i = LAList.ArrayLBound To LAList.ArrayUBound
        buf = LAList.Item(i)
    Next i
    Dump "LAListItem"
    
    Start
    For i = LAList.ArrayLBound To LAList.ArrayUBound
        buf = LAList.ItemAsValue(i)
    Next i
    Dump "LAListItemVal"
    
    Stop
End Sub

Private Sub Start()
    StartTime = VBA.Timer
End Sub

Private Function Lap() As Single
    Lap = VBA.Timer - StartTime
End Function

Private Sub Dump(Msg As String)
    Debug.Print Msg, VBA.Format$(Lap, "0.000s")
End Sub
