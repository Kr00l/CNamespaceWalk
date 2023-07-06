Attribute VB_Name = "NamespaceWalkBase"
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As LongPtr
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As Long
#End If
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = &H0
Private VTableINamespaceWalkCB(0 To 6) As LongPtr
Private VTableINamespaceWalkCB2(0 To 7) As LongPtr

Private Function ShadowINamespaceWalkCB(ByVal Ptr As LongPtr) As INamespaceWalkCB
Dim ObjectPointer As LongPtr, TempObj As Object
CopyMemory ObjectPointer, ByVal UnsignedAdd(Ptr, PTR_SIZE * 2), PTR_SIZE
CopyMemory TempObj, ObjectPointer, PTR_SIZE
Set ShadowINamespaceWalkCB = TempObj
CopyMemory TempObj, NULL_PTR, PTR_SIZE
End Function

Private Function ShadowINSWCBObject(ByVal Ptr As LongPtr, ByVal LpIShellFolder As LongPtr, ByVal LpIDList As LongPtr) As INSWCBObject
Dim ObjectPointer As LongPtr, TempObj As Object
CopyMemory ObjectPointer, ByVal UnsignedAdd(Ptr, PTR_SIZE * 3), PTR_SIZE
CopyMemory TempObj, ObjectPointer, PTR_SIZE
Set ShadowINSWCBObject = TempObj
CopyMemory TempObj, NULL_PTR, PTR_SIZE
Dim VarPointer As LongPtr
CopyMemory VarPointer, ByVal UnsignedAdd(Ptr, PTR_SIZE * 4), PTR_SIZE
CopyMemory ByVal VarPointer, LpIShellFolder, PTR_SIZE
CopyMemory VarPointer, ByVal UnsignedAdd(Ptr, PTR_SIZE * 5), PTR_SIZE
CopyMemory ByVal VarPointer, LpIDList, PTR_SIZE
End Function

#If VBA7 Then
Public Function INamespaceWalkCBPtr(ByVal This As INamespaceWalkCB, ByVal Object As INSWCBObject, ByRef LpIShellFolder As LongPtr, ByRef LpIDList As LongPtr) As LongPtr
#Else
Public Function INamespaceWalkCBPtr(ByVal This As INamespaceWalkCB, ByVal Object As INSWCBObject, ByRef LpIShellFolder As Long, ByRef LpIDList As Long) As Long
#End If
Dim VTableData(0 To 5) As LongPtr
VTableData(0) = GetVTableINamespaceWalkCB()
VTableData(1) = 0 ' RefCount is uninstantiated
VTableData(2) = ObjPtr(This)
VTableData(3) = ObjPtr(Object)
VTableData(4) = VarPtr(LpIShellFolder)
VTableData(5) = VarPtr(LpIDList)
Dim hMem As LongPtr
hMem = CoTaskMemAlloc(PTR_SIZE * 6)
If hMem <> NULL_PTR Then
    CopyMemory ByVal hMem, VTableData(0), PTR_SIZE * 6
    INamespaceWalkCBPtr = hMem
End If
End Function

Private Function GetVTableINamespaceWalkCB() As LongPtr
If VTableINamespaceWalkCB(0) = 0 Then
    VTableINamespaceWalkCB(0) = ProcPtr(AddressOf INamespaceWalkCB_QueryInterface)
    VTableINamespaceWalkCB(1) = ProcPtr(AddressOf INamespaceWalkCB_AddRef)
    VTableINamespaceWalkCB(2) = ProcPtr(AddressOf INamespaceWalkCB_Release)
    VTableINamespaceWalkCB(3) = ProcPtr(AddressOf INamespaceWalkCB_FoundItem)
    VTableINamespaceWalkCB(4) = ProcPtr(AddressOf INamespaceWalkCB_EnterFolder)
    VTableINamespaceWalkCB(5) = ProcPtr(AddressOf INamespaceWalkCB_LeaveFolder)
    VTableINamespaceWalkCB(6) = ProcPtr(AddressOf INamespaceWalkCB_InitializeProgressDialog)
End If
GetVTableINamespaceWalkCB = VarPtr(VTableINamespaceWalkCB(0))
End Function

Private Function INamespaceWalkCB2Ptr(ByVal Ptr As LongPtr) As LongPtr
Dim VTableData(0 To 5) As LongPtr
VTableData(0) = GetVTableINamespaceWalkCB2()
VTableData(1) = 0 ' RefCount is uninstantiated
CopyMemory VTableData(2), ByVal UnsignedAdd(Ptr, PTR_SIZE * 2), PTR_SIZE
CopyMemory VTableData(3), ByVal UnsignedAdd(Ptr, PTR_SIZE * 3), PTR_SIZE
CopyMemory VTableData(4), ByVal UnsignedAdd(Ptr, PTR_SIZE * 4), PTR_SIZE
CopyMemory VTableData(5), ByVal UnsignedAdd(Ptr, PTR_SIZE * 5), PTR_SIZE
Dim hMem As LongPtr
hMem = CoTaskMemAlloc(PTR_SIZE * 6)
If hMem <> NULL_PTR Then
    CopyMemory ByVal hMem, VTableData(0), PTR_SIZE * 6
    INamespaceWalkCB2Ptr = hMem
End If
End Function

Private Function GetVTableINamespaceWalkCB2() As LongPtr
If VTableINamespaceWalkCB2(0) = 0 Then
    VTableINamespaceWalkCB2(0) = ProcPtr(AddressOf INamespaceWalkCB_QueryInterface)
    VTableINamespaceWalkCB2(1) = ProcPtr(AddressOf INamespaceWalkCB_AddRef)
    VTableINamespaceWalkCB2(2) = ProcPtr(AddressOf INamespaceWalkCB_Release)
    VTableINamespaceWalkCB2(3) = ProcPtr(AddressOf INamespaceWalkCB_FoundItem)
    VTableINamespaceWalkCB2(4) = ProcPtr(AddressOf INamespaceWalkCB_EnterFolder)
    VTableINamespaceWalkCB2(5) = ProcPtr(AddressOf INamespaceWalkCB_LeaveFolder)
    VTableINamespaceWalkCB2(6) = ProcPtr(AddressOf INamespaceWalkCB_InitializeProgressDialog)
    VTableINamespaceWalkCB2(7) = ProcPtr(AddressOf INamespaceWalkCB2_WalkComplete)
End If
GetVTableINamespaceWalkCB2 = VarPtr(VTableINamespaceWalkCB2(0))
End Function

Private Function INamespaceWalkCB_QueryInterface(ByVal Ptr As LongPtr, ByRef IID As CLSID, ByRef pvObj As LongPtr) As Long
If VarPtr(pvObj) = NULL_PTR Then
    INamespaceWalkCB_QueryInterface = E_POINTER
    Exit Function
End If
' IID_INamespaceWalkCB = {D92995F8-CF5E-4A76-BF59-EAD39EA2B97E}
' IID_INamespaceWalkCB2 = {7AC7492B-C38E-438A-87DB-68737844FF70}
If IID.Data1 = &HD92995F8 And IID.Data2 = &HCF5E And IID.Data3 = &H4A76 Then
    If IID.Data4(0) = &HBF And IID.Data4(1) = &H59 And IID.Data4(2) = &HEA And IID.Data4(3) = &HD3 _
    And IID.Data4(4) = &H9E And IID.Data4(5) = &HA2 And IID.Data4(6) = &HB9 And IID.Data4(7) = &H7E Then
        pvObj = Ptr
        INamespaceWalkCB_AddRef Ptr
        INamespaceWalkCB_QueryInterface = S_OK
    Else
        INamespaceWalkCB_QueryInterface = E_NOINTERFACE
    End If
ElseIf IID.Data1 = &H7AC7492B And IID.Data2 = &HC38E And IID.Data3 = &H438A Then
    If IID.Data4(0) = &H87 And IID.Data4(1) = &HDB And IID.Data4(2) = &H68 And IID.Data4(3) = &H73 _
    And IID.Data4(4) = &H78 And IID.Data4(5) = &H44 And IID.Data4(6) = &HFF And IID.Data4(7) = &H70 Then
        pvObj = INamespaceWalkCB2Ptr(Ptr)
        INamespaceWalkCB_AddRef pvObj
        INamespaceWalkCB_QueryInterface = S_OK
    Else
        INamespaceWalkCB_QueryInterface = E_NOINTERFACE
    End If
Else
    INamespaceWalkCB_QueryInterface = E_NOINTERFACE
End If
End Function

Private Function INamespaceWalkCB_AddRef(ByVal Ptr As LongPtr) As Long
CopyMemory INamespaceWalkCB_AddRef, ByVal UnsignedAdd(Ptr, PTR_SIZE), 4
INamespaceWalkCB_AddRef = INamespaceWalkCB_AddRef + 1
CopyMemory ByVal UnsignedAdd(Ptr, PTR_SIZE), INamespaceWalkCB_AddRef, 4
End Function

Private Function INamespaceWalkCB_Release(ByVal Ptr As LongPtr) As Long
CopyMemory INamespaceWalkCB_Release, ByVal UnsignedAdd(Ptr, PTR_SIZE), 4
INamespaceWalkCB_Release = INamespaceWalkCB_Release - 1
CopyMemory ByVal UnsignedAdd(Ptr, PTR_SIZE), INamespaceWalkCB_Release, 4
If INamespaceWalkCB_Release = 0 Then CoTaskMemFree Ptr
End Function

Private Function INamespaceWalkCB_FoundItem(ByVal Ptr As LongPtr, ByVal LpIShellFolder As LongPtr, ByVal LpIDList As LongPtr) As Long
ShadowINamespaceWalkCB(Ptr).FoundItem ShadowINSWCBObject(Ptr, LpIShellFolder, LpIDList)
INamespaceWalkCB_FoundItem = S_OK
End Function

Private Function INamespaceWalkCB_EnterFolder(ByVal Ptr As LongPtr, ByVal LpIShellFolder As LongPtr, ByVal LpIDList As LongPtr) As Long
INamespaceWalkCB_EnterFolder = S_OK
ShadowINamespaceWalkCB(Ptr).EnterFolder ShadowINSWCBObject(Ptr, LpIShellFolder, LpIDList), INamespaceWalkCB_EnterFolder
End Function

Private Function INamespaceWalkCB_LeaveFolder(ByVal Ptr As LongPtr, ByVal LpIShellFolder As LongPtr, ByVal LpIDList As LongPtr) As Long
ShadowINamespaceWalkCB(Ptr).LeaveFolder ShadowINSWCBObject(Ptr, LpIShellFolder, LpIDList)
INamespaceWalkCB_LeaveFolder = S_OK
End Function

Private Function INamespaceWalkCB_InitializeProgressDialog(ByVal Ptr As LongPtr, ByRef lpszTitle As LongPtr, ByRef lpszCancel As LongPtr) As Long
Dim DialogTitle As String
ShadowINamespaceWalkCB(Ptr).InitializeProgressDialog DialogTitle
If StrPtr(DialogTitle) <> NULL_PTR Then
    DialogTitle = DialogTitle & vbNullChar
    Dim hMem As LongPtr
    hMem = CoTaskMemAlloc(LenB(DialogTitle))
    If hMem <> NULL_PTR Then
        CopyMemory ByVal hMem, ByVal StrPtr(DialogTitle), LenB(DialogTitle)
        ' [out] LPWSTR *ppszTitle
        ' The interface itself will take care to CoTaskMemFree the string later on automatically.
        lpszTitle = hMem
    End If
End If
INamespaceWalkCB_InitializeProgressDialog = S_OK
End Function

Private Function INamespaceWalkCB2_WalkComplete(ByVal Ptr As LongPtr, ByVal HResult As Long) As Long
ShadowINamespaceWalkCB(Ptr).WalkComplete HResult
INamespaceWalkCB2_WalkComplete = S_OK
End Function

Private Function ProcPtr(ByVal Address As LongPtr) As LongPtr
ProcPtr = Address
End Function

Private Function UnsignedAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr
#If Win64 Then
UnsignedAdd = ((Start Xor &H8000000000000000^) + Incr) Xor &H8000000000000000^
#Else
UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
#End If
End Function
