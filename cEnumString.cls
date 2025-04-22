VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements WinDevLibImpl.IEnumString
Private sItems() As String
Private nItems As Long
Private nCur As Long
Private iescnt As Long
 
Private Sub IEnumString_Next(ByVal celt As Long, rgelt As LongPtr, pceltFetched As Long)
Debug.Print "cEnumString_Next"
Dim lpString As LongPtr
Dim i As Long
Dim celtFetched As Long
If rgelt = 0 Then
    Err.ReturnHResult = E_POINTER
    Exit Sub
End If

For i = 0 To (celt - 1)
    If nCur = nItems Then Exit For
    lpString = CoTaskMemAlloc(LenB(sItems(nCur) & vbNullChar))
    If lpString = 0 Then Err.ReturnHResult = S_FALSE: Exit Sub
    
    CopyMemory ByVal lpString, ByVal StrPtr(sItems(nCur)), LenB(sItems(nCur) & vbNullChar)
    CopyMemory ByVal UnsignedAdd(VarPtr(rgelt), i * LenB(Of LongPtr)), lpString, LenB(Of LongPtr)
    
    nCur = nCur + 1
    celtFetched = celtFetched + 1
Next i
 If pceltFetched Then
     pceltFetched = celtFetched
 End If
 If i <> celt Then Err.ReturnHResult = S_FALSE
'  Debug.Print "IES_Next retval=" & Err.ReturnHResult

End Sub

Private Sub IEnumString_Skip(ByVal celt As Long)
If nCur + celt <= nItems Then
    nCur = nCur + celt
    Err.ReturnHResult = S_OK
Else
    Err.ReturnHResult = S_FALSE
End If
End Sub
Private Sub IEnumString_Reset()
StringCountReset
End Sub
Private Sub IEnumString_Clone(ppenum As IEnumString)
Err.ReturnHResult = E_NOTIMPL
End Sub
 
Public Sub SetACStringList(sz() As String)
Dim i As Long
ReDim sItems(UBound(sz))
For i = 0 To UBound(sz)
    sItems(i) = sz(i)
Next i
nItems = UBound(sItems) + 1
End Sub
Public Sub ClearACStringList()
ReDim sItems(0)
nItems = 0
End Sub
Public Sub StringCountReset()
nCur = 0
End Sub
