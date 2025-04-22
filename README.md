# cAutoComplete
### Custom Autocomplete Class for twinBASIC

This is a port of my original VB6 project: [[VB6] Using IAutoComplete / IAutoComplete2 including autocomplete with custom lists ](https://www.vbforums.com/showthread.php?836015-VB6-Using-IAutoComplete-IAutoComplete2-including-autocomplete-with-custom-lists)

It replaces oleexp.tlb with WinDevLib/WinDevLibImpl, which with tB allows greatly simplifying things by eliminating v-table redirects. It's been updated to also be x64 compatible.

>[!NOTE]
>The twinproj is in a zip since it exceeds GitHub's 25MB limit; tB still includes all files and duplicates references, so it includes WinDevLib (13MB) and WinDevLibImpl which has its own copy of WinDevLib as its base.
----

![jgjDdBD](https://github.com/user-attachments/assets/894a28ba-c1e6-4165-bb4e-fbe498aadfb0)

## `IAutoComplete` / `IAutoComplete2` / `IEnumString`

`SHAutocomplete` has many well known limitations, the biggest being if you want to supply your own list to use with it. I was very impressed with Krool's work on this interface, and not wanting to include a whole other TLB set out to do it with oleexp, and now, bring it to twinBASIC/WinDevLib as Krool has done with his.

IAutoCompleteDropdown is used to provide the status of the dropdown autosuggest list. The .DropdownStatus method reports whether it's down, and the text of an item if an item in the list is selected. In the sample project, this is run on an automatically updated timer enabled in the 'basic filesystem' routine. It also exposes the .ResetEnumerator call to update the dropdown list while it's open.

Here's what the code looks like:


cAutoComplete.cls

```vba
Option Explicit

Private pACO As AutoComplete
Private pACL As ACListISF
Private pACL2 As IACList2
Private pACLH As ACLHistory
Private pACLMRU As ACLMRU
Private pACM As ACLMulti
Private pObjMgr As IObjMgr
Private pDD As IAutoCompleteDropDown
Private pUnk As IUnknownUnrestricted
Private m_hWnd As LongPtr
Private pCust As cEnumString
 
Private Sub Class_Initialize()
Set pACO = New AutoComplete
End Sub

Public Sub AC_Filesys(hWnd As LongPtr, lOpt As AUTOCOMPLETEOPTIONS)
Set pACL = New ACListISF
pACO.Init hWnd, pACL, "", ""
pACO.SetOptions lOpt
pACO.Enable 1
m_hWnd = hWnd
End Sub
Public Sub AC_Disable()
pACO.Enable 0
End Sub
Public Sub AC_Enable()
pACO.Enable 1
End Sub
Public Sub AC_Custom(hWnd As LongPtr, sTerms() As String, lOpt As AUTOCOMPLETEOPTIONS)
Set pCust = New cEnumString
pCust.SetACStringList sTerms
pACO.Init hWnd, pCust, "", ""
pACO.SetOptions lOpt
pACO.Enable 1
m_hWnd = hWnd
End Sub
Public Sub UpdateCustomTerms(sTerms() As String)
If (pCust Is Nothing) = False Then
    pCust.SetACStringList sTerms
End If
End Sub
Public Sub AC_ACList2(hWnd As LongPtr, lOpt As AUTOCOMPLETEOPTIONS, lOpt2 As AUTOCOMPLETELISTOPTIONS)
Set pACL = New ACListISF
Set pACL2 = pACL
If (pACL2 Is Nothing) = False Then
    pACL2.SetOptions lOpt2
    pACO.Init hWnd, pACL2, "", ""
    pACO.SetOptions lOpt
    pACO.Enable 1
    m_hWnd = hWnd
Else
    Debug.Print "Failed to create IACList2"
End If
End Sub
Public Sub AC_History(hWnd As LongPtr, lOpt As AUTOCOMPLETEOPTIONS)
Set pACLH = New ACLHistory
pACO.Init hWnd, pACLH, "", ""
pACO.SetOptions lOpt
pACO.Enable 1
m_hWnd = hWnd

End Sub
Public Sub AC_MRU(hWnd As LongPtr, lOpt As AUTOCOMPLETEOPTIONS)
Set pACLMRU = New ACLMRU
pACO.Init hWnd, pACLMRU, "", ""
pACO.SetOptions lOpt
pACO.Enable 1
m_hWnd = hWnd

End Sub

Public Sub AC_Multi(hWnd As LongPtr, lOpt As AUTOCOMPLETEOPTIONS, lFSOpts As AUTOCOMPLETELISTOPTIONS, bFileSys As Boolean, bHistory As Boolean, bMRU As Boolean, bCustom As Boolean, Optional vStringArrayForCustom As Variant)

   On Error GoTo e0

Set pACM = New ACLMulti
Set pObjMgr = pACM

If bFileSys Then
    Set pACL = New ACListISF
    Set pACL2 = pACL
    pACL2.SetOptions lFSOpts
    pObjMgr.Append pACL2
End If
If bMRU Then
    Set pACLMRU = New ACLMRU
    pObjMgr.Append pACLMRU
End If
If bHistory Then
    Set pACLH = New ACLHistory
    pObjMgr.Append pACLH
End If
If bCustom Then
    Dim i As Long
    Dim sTerms() As String
    ReDim sTerms(UBound(vStringArrayForCustom))
    For i = 0 To UBound(vStringArrayForCustom)
        sTerms(i) = vStringArrayForCustom(i)
    Next i
    Set pCust = New cEnumString
    pCust.SetACStringList sTerms
    pObjMgr.Append pCust
End If

pACO.Init hWnd, pObjMgr, "", ""
pACO.SetOptions lOpt
pACO.Enable 1
m_hWnd = hWnd
   On Error GoTo 0
   Exit Sub

e0:

    Debug.Print "cAutocomplete.AC_Multi.Error->" & Err.Description & " (" & Err.Number & ")"

End Sub

Public Function DropdownStatus(lpStatus As Long, sText As String) As Long
If pDD Is Nothing Then
    Set pDD = pACO
End If
Dim lp As LongPtr

pDD.GetDropdownStatus lpStatus, lp
SysReAllocStringW VarPtr(sText), lp
CoTaskMemFree lp

End Function
Public Sub ResetEnum()
If pDD Is Nothing Then
    Set pDD = pACO
End If
pDD.ResetEnumerator
End Sub
```

Implementing IEnumString's functions:

```vba
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
 Debug.Print "IES_Next retval=" & Err.ReturnHResult

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
```

For the complete code, see the attached project.

### Requirements 
-Windows Development Library for twinBASIC (WinDevLib)\
-WinDevLib for Implements (should be after WinDevLib in priority list)

### Thanks 
Krool's project mentioned above is what inspired me to do this, and I borrowed a few techniques from his project, especially for IEnumString.
