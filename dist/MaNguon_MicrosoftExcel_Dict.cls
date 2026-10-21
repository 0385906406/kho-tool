Attribute VB_Name = "Dict"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'==========================================================================
' Dict: thay the Scripting.Dictionary tren Excel cho Mac (Mac khong co scrrun.dll)
' Ho tro: d(k) doc/ghi, .Add, .Exists, .Keys, .Items, .Count, .Remove, .RemoveAll
' Khoa phan biet hoa/thuong va phan biet so / chu giong Scripting.Dictionary.
' Tren Windows tool dung Scripting.Dictionary that (nhanh hon), lop nay khong duoc goi.
'==========================================================================
Option Explicit

Private mKeys() As Variant
Private mItems() As Variant
Private mCount As Long
Private mIdx As Collection

Private Sub Class_Initialize()
    Set mIdx = New Collection
    ReDim mKeys(0 To 15): ReDim mItems(0 To 15)
End Sub

' Collection so khoa khong phan biet hoa/thuong -> doi khoa thanh chuoi ma hex
Private Function KeyStr(ByRef k As Variant) As String
    Dim s As String, i As Long, n As Long, out As String
    Select Case VarType(k)
        Case vbString
            s = k: n = Len(s)
            out = String$(n * 4 + 1, "0")
            Mid$(out, 1, 1) = "s"
            For i = 1 To n
                Mid$(out, i * 4 - 2, 4) = Right$("000" & Hex$(AscW(Mid$(s, i, 1)) And &HFFFF&), 4)
            Next
            KeyStr = out
        Case vbEmpty
            KeyStr = "e"
        Case vbBoolean
            KeyStr = "b" & CStr(CLng(k))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, 20
            KeyStr = "n" & CStr(CDbl(k))
        Case vbDate
            KeyStr = "d" & CStr(CDbl(k))
        Case Else
            KeyStr = "x" & CStr(k)
    End Select
End Function

Private Function IndexOf(ByRef k As Variant) As Long
    On Error GoTo NotFound
    IndexOf = mIdx(KeyStr(k))
    Exit Function
NotFound:
    IndexOf = -1
End Function

Private Sub Grow()
    If mCount > UBound(mKeys) Then
        ReDim Preserve mKeys(0 To UBound(mKeys) * 2 + 1)
        ReDim Preserve mItems(0 To UBound(mItems) * 2 + 1)
    End If
End Sub

Private Sub AddNew(ByRef k As Variant, ByRef v As Variant)
    Grow
    mIdx.Add mCount, KeyStr(k)
    mKeys(mCount) = k
    If IsObject(v) Then Set mItems(mCount) = v Else mItems(mCount) = v
    mCount = mCount + 1
End Sub

Public Sub Add(ByVal Key As Variant, ByVal Item As Variant)
    If IndexOf(Key) >= 0 Then Err.Raise 457, "Dict", "This key is already associated with an element of this collection"
    AddNew Key, Item
End Sub

Public Function Exists(ByVal Key As Variant) As Boolean
    Exists = (IndexOf(Key) >= 0)
End Function

Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim i As Long, e As Variant
    i = IndexOf(Key)
    If i < 0 Then AddNew Key, e: i = mCount - 1   ' giong Scripting.Dictionary: doc khoa chua co -> them gia tri rong
    If IsObject(mItems(i)) Then Set Item = mItems(i) Else Item = mItems(i)
End Property

Public Property Let Item(ByVal Key As Variant, ByVal v As Variant)
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = IndexOf(Key)
    If i < 0 Then AddNew Key, v Else mItems(i) = v
End Property

Public Property Set Item(ByVal Key As Variant, ByVal v As Variant)
Attribute Item.VB_UserMemId = 0
    Dim i As Long
    i = IndexOf(Key)
    If i < 0 Then AddNew Key, v Else Set mItems(i) = v
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function Keys() As Variant
    Dim a() As Variant, i As Long
    If mCount = 0 Then Keys = Array(): Exit Function
    ReDim a(0 To mCount - 1)
    For i = 0 To mCount - 1: a(i) = mKeys(i): Next
    Keys = a
End Function

Public Function Items() As Variant
    Dim a() As Variant, i As Long
    If mCount = 0 Then Items = Array(): Exit Function
    ReDim a(0 To mCount - 1)
    For i = 0 To mCount - 1
        If IsObject(mItems(i)) Then Set a(i) = mItems(i) Else a(i) = mItems(i)
    Next
    Items = a
End Function

Public Sub Remove(ByVal Key As Variant)
    Dim i As Long, j As Long, ks() As Variant, vs() As Variant, n As Long
    i = IndexOf(Key)
    If i < 0 Then Err.Raise 32811, "Dict", "Element not found"
    n = mCount: ks = mKeys: vs = mItems
    RemoveAll
    For j = 0 To n - 1
        If j <> i Then AddNew ks(j), vs(j)
    Next
End Sub

Public Sub RemoveAll()
    Set mIdx = New Collection
    ReDim mKeys(0 To 15): ReDim mItems(0 To 15)
    mCount = 0
End Sub

Public Property Get CompareMode() As Long
    CompareMode = 0
End Property

Public Property Let CompareMode(ByVal v As Long)
End Property
