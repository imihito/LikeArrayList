VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LikeArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum InternalClsErr
    iceOutOfIndex = 9
    iceTypeMissMatch = 13
    iceUnimplemented = vbObjectError
    
End Enum

Private clsArray As Variant
Private clsInitialized As Boolean
Private clsBase As Long

Private clsCapacity As Long
Private clsArrayIndex As Long

'内部の配列の初期化
Public Sub InitInternalArray(ByRef BaseArray As Variant, Optional TypeOnly As Boolean = True)
    If Not VBA.IsArray(BaseArray) Then
        Call Err.Raise(iceTypeMissMatch)
    End If
    
    Let clsArray = BaseArray
    Let clsInitialized = True
    If TypeOnly Then
        Call Me.Clear(False)
        Exit Sub
    End If
    
    On Error GoTo NotoInitArray
        Let clsBase = LBound(BaseArray)
        Let clsCapacity = UBound(BaseArray)
        Let clsArrayIndex = clsCapacity
    On Error GoTo 0
Exit Sub

NotoInitArray:
    On Error GoTo 0
    Let clsBase = 1
    Let clsCapacity = 0
    Let clsArrayIndex = 0
Exit Sub

End Sub

'基本的にLbound(clsArray)と同一。
Property Get ArrayLBound() As Long
    Let ArrayLBound = clsBase
End Property

'紛らわしいですが、clsArrayで使用している最大Index。
Property Get ArrayUBound() As Long
    Let ArrayUBound = clsArrayIndex
End Property


'基本的にUbound(clsArray)と同一。
Public Property Get Capacity() As Long
    Let Capacity = clsCapacity
End Property

Public Property Let Capacity(ByVal Size As Long)
    Call Extend(Size)
    Let clsCapacity = Size
End Property


'中の要素数と同等
Public Property Get Count() As Long
    Let Count = clsArrayIndex + 1 - clsBase
End Property


'Itemは値型かオブジェクト型かを自動判定。
'ItemAs～は判定をしない（その分パフォーマンス改善）。
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Call AssignItem(clsArray(Index), Item)
End Property


Public Property Let Item(ByVal Index As Long, ByVal Value As Variant)
    Let clsArray(Index) = Value
End Property

Public Property Get ItemAsValue(ByVal Index As Long) As Variant
    Let ItemAsValue = clsArray(Index)
End Property


Public Property Set Item(ByVal Index As Long, ByVal Value As Variant)
    Set clsArray(Index) = Value
End Property

Public Property Get ItemAsObject(ByVal Index As Long) As Object
    Set ItemAsObject = clsArray(Index)
End Property



'値型かオブジェクト型かを自動判定。
Public Sub Add(ByRef Item As Variant)
    Call Extend(clsArrayIndex + 1)
    Call AssignItem(Item, clsArray(clsArrayIndex + 1))
    Let clsArrayIndex = clsArrayIndex + 1   '追加失敗に備え、追加が終わってからインクリメント。
End Sub

'値型か、オブジェクト型か決まっている場合は以下のメソッド。
Public Sub AddValue(ByRef Item As Variant)
    Call Extend(clsArrayIndex + 1)
    Let clsArray(clsArrayIndex + 1) = Item
    Let clsArrayIndex = clsArrayIndex + 1
End Sub

Public Sub AddObject(ByRef Item As Object)
    Call Extend(clsArrayIndex + 1)
    Set clsArray(clsArrayIndex + 1) = Item
    Let clsArrayIndex = clsArrayIndex + 1
End Sub


'値型かオブジェクト型かを自動判定。
Public Sub AddRange(ByRef List As Variant)
    Call Extend(clsArrayIndex + ListLength(List))
    
    Dim i As Long
    Dim tVar As Variant
    
    For Each tVar In List
        i = i + 1
        Call AssignItem(tVar, clsArray(clsArrayIndex + i))
    Next tVar
    
    Let clsArrayIndex = clsArrayIndex + i
End Sub

Private Function ListLength(List As Variant) As Long
    If VBA.IsArray(List) Then
        Let ListLength = UBound(List) - LBound(List) + 1
    Else
        Let ListLength = List.Count
    End If
End Function

Public Sub AddValueRange(ByRef List As Variant)
    Call Extend(clsArrayIndex + ListLength(List))
    
    Dim i As Long
    Dim tVar As Variant
    
    For Each tVar In List
        i = i + 1
        Let clsArray(clsArrayIndex + i) = tVar
    Next tVar
    
    Let clsArrayIndex = clsArrayIndex + i
End Sub

Public Sub AddObjectRange(ByRef List As Variant)
    Call Extend(clsArrayIndex + ListLength(List))
    
    Dim i As Long
    Dim tVar As Variant
    
    For Each tVar In List
        i = i + 1
        Set clsArray(clsArrayIndex + i) = tVar
    Next tVar
    
    Let clsArrayIndex = clsArrayIndex + i
End Sub


Public Function BinarySearch(ByRef Item As Variant) As Long
    Call ClsErrRaise(iceUnimplemented, "BinarySearch")
End Function


Public Sub Clear(Optional Force As Boolean = False)
    If Force Then
        clsArray = Empty
        clsInitialized = False
    Else
        Erase clsArray  '配列の型情報は維持される。
    End If
    Call Reset
End Sub


Public Function Clone() As LikeArrayList
    Dim tmp As LikeArrayList:    Set tmp = New LikeArrayList
    Call tmp.InitInternalArray(Me.ToArray)
    tmp.Capacity = Me.Capacity
    Set Clone = tmp
End Function


Public Function Contains(ByRef Item As Variant) As Boolean
    Call ClsErrRaise(iceUnimplemented, "Contains")
    On Error GoTo Fail
        If VBA.IsObject(Item) Then
            
        ElseIf VBA.IsArray(Item) Then
            
        Else
            
        End If
    On Error GoTo 0
Exit Function

Fail:
End Function


Public Sub CopyTo(ByRef CopiedArray As Variant)
    Dim i As Long
    For i = clsBase To clsArrayIndex
        Call AssignItem(clsArray(i), CopiedArray(i))
    Next i
End Sub


Public Function IndexOf(ByRef Item As Variant) As Long
    Call ClsErrRaise(iceUnimplemented, "IndexOf")
End Function

Public Sub Insert(Index As Long, ByRef Item As Variant)
    Call ClsErrRaise(iceUnimplemented, "Insert")
End Sub


Public Function ToArray() As Variant
    Dim tmp As Variant
    If clsInitialized Then
        Let tmp = clsArray
        ReDim Preserve tmp(clsBase To clsArrayIndex)
    Else
        Let tmp = Array()
    End If
    Let ToArray = tmp
End Function

Public Function ToString() As String
    Let ToString = VBA.TypeName(Me) & "(Of " & VBA.TypeName(clsArray) & ")"
End Function

Public Function ArrayType() As String
    Dim tName As String:    tName = VBA.TypeName(clsArray)
    Select Case tName
        Case VBA.TypeName(Empty)
            Let ArrayType = tName
        Case Else
            Let ArrayType = VBA.Left$(tName, VBA.Len(tName) - 2)
    End Select
End Function


Public Sub TrimToSize()
    If clsInitialized Then
        ReDim Preserve clsArray(clsBase To clsArrayIndex)
    End If
End Sub

Private Sub Extend(RequireIndex As Long)
    If RequireIndex <= clsCapacity Then Exit Sub
    Dim nextUbound As Long
    If RequireIndex > clsCapacity * 2 Then
        nextUbound = RequireIndex
    Else
        nextUbound = clsCapacity * 2
    End If
    If clsInitialized Then
        ReDim Preserve clsArray(clsBase To nextUbound)
    Else
        ReDim clsArray(clsBase To nextUbound)
    End If
    Let clsCapacity = nextUbound
End Sub

Private Sub AssignItem(ByRef InputItem As Variant, ByRef OutputItem As Variant)
    If VBA.IsObject(InputItem) Then
        Set OutputItem = InputItem
    Else
        Let OutputItem = InputItem
    End If
End Sub


Private Sub Reset()
    Let clsBase = 1
    Let clsArrayIndex = 0
    Let clsCapacity = 0
End Sub

Private Sub ClsErrRaise(ErrNo As InternalClsErr, SourceProc As String)
    Select Case ErrNo
        Case Is >= 0
            Call Err.Raise(ErrNo, SourceProc)
        Case InternalClsErr.iceUnimplemented
            Call Err.Raise(ErrNo, SourceProc, "未実装です。" & vbCrLf & "This method is Unimplemented")
    End Select
End Sub

Private Sub Class_Initialize()
    Call Reset
    Let clsInitialized = False
End Sub

Private Sub Class_Terminate()
    Erase clsArray
End Sub
