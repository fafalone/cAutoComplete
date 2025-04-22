VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------
'
'cAutoComplete v0.4 (twinBASIC/x64)
'by fafalone
'
'Provides simplified access to IAutocomplete and related interfaces
'
'Requirements:
' -Windows Development Library for twinBASIC (WinDevLib)
' -WinDevLib for Implements
' -cEnumString.cls (included)
'
'==History==
'
'Version 0.4 (twinBASIC/x64)
'-Converted to use WinDevLib and twinBASIC, and support x64
'-Bug fix in cEnumString: 
'   CoTaskMemAlloc(LenB(sItems(nCur)) & vbNullChar) should be
'   CoTaskMemAlloc(LenB(sItems(nCur) & vbNullChar))
'  That crashed in x64, not sure how it didn't in 32bit mode too.
'
'Version 0.4
'-Added sub to update the custom terms list without a full reset
'
'Version 0.3
'-Updated to use oleexp 4.0 or higher
'
'Version 0.2
'-Now supports multiple custom lists at the same time in a project
'--------------------------------------------------------------------


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
