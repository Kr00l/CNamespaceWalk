VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNamespaceWalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
#If False Then
Private NSWDefault, NSWNoneImpliesAll, NSWOneImpliesAll, NSWDontTraverseLinks, NSWDontAccumulateResult, NSWTraverseStreamJunctions, NSWFileSystemOnly, NSWShowProgress, NSWDontResolveLinks, NSWAccumulateFolders, NSWDontSort, NSWUseTransferMedium, NSWDontTraverseStreamJunctions, NSWAnyImpliesAll
Private NSWCBResultContinue, NSWCBResultSkip, NSWCBResultCancel
#End If
Private Const NSWF_DEFAULT As Long = &H0
Private Const NSWF_NONE_IMPLIES_ALL As Long = &H1
Private Const NSWF_ONE_IMPLIES_ALL As Long = &H2
Private Const NSWF_DONT_TRAVERSE_LINKS As Long = &H4
Private Const NSWF_DONT_ACCUMULATE_RESULT As Long = &H8
Private Const NSWF_TRAVERSE_STREAM_JUNCTIONS As Long = &H10
Private Const NSWF_FILESYSTEM_ONLY As Long = &H20
Private Const NSWF_SHOW_PROGRESS As Long = &H40
Private Const NSWF_FLAG_VIEWORDER As Long = &H80 ' Only for IShellView
Private Const NSWF_IGNORE_AUTOPLAY_HIDA As Long = &H100 ' Only for IDataObject
Private Const NSWF_ASYNC As Long = &H200 ' Only if IAgileObject is implemented
Private Const NSWF_DONT_RESOLVE_LINKS As Long = &H400
Private Const NSWF_ACCUMULATE_FOLDERS As Long = &H800
Private Const NSWF_DONT_SORT As Long = &H1000
Private Const NSWF_USE_TRANSFER_MEDIUM As Long = &H2000
Private Const NSWF_DONT_TRAVERSE_STREAM_JUNCTIONS As Long = &H4000
Private Const NSWF_ANY_IMPLIES_ALL As Long = &H8000&
Public Enum NSWFlags
NSWDefault = NSWF_DEFAULT
NSWNoneImpliesAll = NSWF_NONE_IMPLIES_ALL
NSWOneImpliesAll = NSWF_ONE_IMPLIES_ALL
NSWDontTraverseLinks = NSWF_DONT_TRAVERSE_LINKS
NSWDontAccumulateResult = NSWF_DONT_ACCUMULATE_RESULT
NSWTraverseStreamJunctions = NSWF_TRAVERSE_STREAM_JUNCTIONS
NSWFileSystemOnly = NSWF_FILESYSTEM_ONLY
NSWShowProgress = NSWF_SHOW_PROGRESS
NSWDontResolveLinks = NSWF_DONT_RESOLVE_LINKS
NSWAccumulateFolders = NSWF_ACCUMULATE_FOLDERS
NSWDontSort = NSWF_DONT_SORT
NSWUseTransferMedium = NSWF_USE_TRANSFER_MEDIUM
NSWDontTraverseStreamJunctions = NSWF_DONT_TRAVERSE_STREAM_JUNCTIONS
NSWAnyImpliesAll = NSWF_ANY_IMPLIES_ALL
End Enum
Private Const S_OK As Long = &H0
Private Const S_FALSE As Long = &H1
Private Const ERROR_CANCELLED As Long = &H800704C7
Public Enum NSWCBResultConstants
NSWCBResultContinue = S_OK
NSWCBResultSkip = S_FALSE
NSWCBResultCancel = ERROR_CANCELLED
End Enum
Private Type CLSID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Const MAX_PATH As Long = 260
Private Type STRRET
uType As Long
lpOleStr As LongPtr
uOffset As Long
cString(0 To ((MAX_PATH * 2) - 1)) As Byte
End Type
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal lpvInstance As LongPtr, ByVal oVft As LongPtr, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As LongPtr, ByVal prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As LongPtr, ByRef pCLSID As Any) As Long
Private Declare PtrSafe Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As Any) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function SHGetFolderLocation Lib "shell32" (ByVal hWndOwner As LongPtr, ByVal nFolder As Long, ByVal hToken As LongPtr, ByVal dwReserved As Long, ByRef LpIDList As LongPtr) As Long
Private Declare PtrSafe Function ILCreateFromPath Lib "shell32" (ByVal lpszPath As LongPtr) As LongPtr
Private Declare PtrSafe Function SHGetDesktopFolder Lib "shell32" (ByRef ppshf As IUnknown) As Long
Private Declare PtrSafe Function SHBindToObject Lib "shell32" (ByVal lpsf As LongPtr, ByVal LpIDList As LongPtr, ByVal lpbc As LongPtr, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare PtrSafe Function StrRetToBuf Lib "shlwapi.dll" Alias "StrRetToBufW" (ByRef pStrRet As STRRET, ByVal LpIDList As LongPtr, ByVal lpszBuf As LongPtr, ByVal cchBuf As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal lpvInstance As Long, ByVal oVft As Long, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByRef pvargResult As Variant) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As Any) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByRef LpIDList As Long) As Long
Private Declare Function ILCreateFromPath Lib "shell32" (ByVal lpszPath As Long) As Long
Private Declare Function SHGetDesktopFolder Lib "shell32" (ByRef ppshf As IUnknown) As Long
Private Declare Function SHBindToObject Lib "shell32" (ByVal lpsf As Long, ByVal LpIDList As Long, ByVal lpbc As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare Function StrRetToBuf Lib "shlwapi.dll" Alias "StrRetToBufW" (ByRef pStrRet As STRRET, ByVal LpIDList As Long, ByVal lpszBuf As Long, ByVal cchBuf As Long) As Long
#End If
Private Const CLSID_NamespaceWalker As String = "{72eb61e0-8672-4303-9175-f2e4c68b2e7c}"
Private Const IID_INamespaceWalk As String = "{57CED8A7-3F4A-432C-9350-30F24483F74F}"
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Const CC_STDCALL As Long = 4
Implements INSWCBObject
Private NamespaceWalkINamespaceWalk As IUnknown
Private NamespaceWalkIShellFolder As IUnknown
Private NamespaceWalkIDLRoot As LongPtr
Private NamespaceWalkSHBindToObject As Integer
Private NamespaceWalkLpIShellFolder As LongPtr
Private NamespaceWalkLpIDList As LongPtr
Private PropKey As String
Private PropRootFolder As Variant
Private PropFlags As NSWFlags
Private PropLevels As Long

Private Sub Class_Initialize()
Set NamespaceWalkINamespaceWalk = CreateINamespaceWalk()
End Sub

Private Sub Class_Terminate()
If NamespaceWalkIDLRoot <> NULL_PTR Then
    CoTaskMemFree NamespaceWalkIDLRoot
    NamespaceWalkIDLRoot = NULL_PTR
End If
End Sub

Private Property Get INSWCBObject_LpIShellFolder() As LongPtr
INSWCBObject_LpIShellFolder = NamespaceWalkLpIShellFolder
End Property

Private Property Get INSWCBObject_LpIDList() As LongPtr
INSWCBObject_LpIDList = NamespaceWalkLpIDList
End Property

Private Function INSWCBObject_GetDisplayNameOf() As String
Const SHGDN_FORPARSING As Long = &H8000&
Dim uFlags As Long, pStrRet As STRRET
uFlags = SHGDN_FORPARSING
' IShellFolder::GetDisplayNameOf
VTableCall vbLong, NamespaceWalkLpIShellFolder, 12, NamespaceWalkLpIDList, uFlags, VarPtr(pStrRet)
Dim Buffer As String
Buffer = String(MAX_PATH, vbNullChar) & vbNullChar
StrRetToBuf pStrRet, NamespaceWalkLpIDList, StrPtr(Buffer), MAX_PATH
INSWCBObject_GetDisplayNameOf = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End Function

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

Public Property Get RootFolder() As Variant
Attribute RootFolder.VB_Description = "Returns/sets the root folder from which to begin the namespace walk."
Attribute RootFolder.VB_UserMemId = 0
RootFolder = PropRootFolder
End Property

Public Property Let RootFolder(ByVal Value As Variant)
Select Case VarType(Value)
    Case vbEmpty, vbLong, vbInteger, vbByte, vbString, vbDouble, vbSingle
        PropRootFolder = Value
    Case Else
        Err.Raise 380
End Select
If NamespaceWalkIDLRoot <> NULL_PTR Then
    CoTaskMemFree NamespaceWalkIDLRoot
    NamespaceWalkIDLRoot = NULL_PTR
End If
Set NamespaceWalkIShellFolder = Nothing
Select Case VarType(PropRootFolder)
    Case vbEmpty
        NamespaceWalkIDLRoot = NULL_PTR
    Case vbLong, vbInteger, vbByte
        SHGetFolderLocation NULL_PTR, PropRootFolder, 0, 0, NamespaceWalkIDLRoot
    Case vbString
        NamespaceWalkIDLRoot = ILCreateFromPath(StrPtr(Left$(PropRootFolder, MAX_PATH)))
    Case vbDouble, vbSingle
        SHGetFolderLocation NULL_PTR, CLng(PropRootFolder), 0, 0, NamespaceWalkIDLRoot
End Select
' IID_IShellFolder = {000214E6-0000-0000-C000-000000000046}
Dim IID As CLSID
With IID
.Data1 = &H214E6
.Data4(0) = &HC0
.Data4(7) = &H46
End With
If NamespaceWalkSHBindToObject = 0 Then
    Dim hLib As LongPtr
    hLib = LoadLibrary(StrPtr("shell32.dll"))
    If hLib <> NULL_PTR Then
        If GetProcAddress(hLib, "SHBindToObject") <> NULL_PTR Then
            NamespaceWalkSHBindToObject = 1
        Else
            NamespaceWalkSHBindToObject = -1
        End If
        FreeLibrary hLib
        hLib = NULL_PTR
    End If
End If
If NamespaceWalkSHBindToObject > -1 Then
    ' Requires shell32.dll version 6.0 or higher.
    SHBindToObject NULL_PTR, NamespaceWalkIDLRoot, NULL_PTR, IID, NamespaceWalkIShellFolder
Else
    Dim IShellFolderDesktop As IUnknown
    SHGetDesktopFolder IShellFolderDesktop
    ' IShellFolder::BindToObject
    VTableCall vbLong, ObjPtr(IShellFolderDesktop), 6, NamespaceWalkIDLRoot, NULL_PTR, VarPtr(IID), VarPtr(NamespaceWalkIShellFolder)
End If
End Property

Public Property Get Flags() As NSWFlags
Attribute Flags.VB_Description = "Returns/sets the options for a namespace walk."
Flags = PropFlags
End Property

Public Property Let Flags(ByVal Value As NSWFlags)
PropFlags = Value
End Property

Public Property Get Levels() As Long
Attribute Levels.VB_Description = "Returns/sets the maximum depth to descend through the namespace hierarchy."
Levels = PropLevels
End Property

Public Property Let Levels(ByVal Value As Long)
If Value < 0 Then Err.Raise 380
PropLevels = Value
End Property

Public Function Walk(Optional ByVal Callback As INamespaceWalkCB) As Boolean
Attribute Walk.VB_Description = "Initiates a recursive walk of the namespace from the specified root folder."
If Not NamespaceWalkINamespaceWalk Is Nothing And Not NamespaceWalkIShellFolder Is Nothing Then
    Dim CBPtr As LongPtr
    If Not Callback Is Nothing Then CBPtr = INamespaceWalkCBPtr(Callback, Me, NamespaceWalkLpIShellFolder, NamespaceWalkLpIDList)
    ' INamespaceWalk::Walk
    Walk = CBool(VTableCall(vbLong, ObjPtr(NamespaceWalkINamespaceWalk), 4, ObjPtr(NamespaceWalkIShellFolder), PropFlags, PropLevels, CBPtr) = S_OK)
End If
End Function

Public Function GetIDArrayResult() As CNSWIDArrayResult
Attribute GetIDArrayResult.VB_Description = "Gets an array of PIDL objects found during a namespace walk."
If Not NamespaceWalkINamespaceWalk Is Nothing Then
    Set GetIDArrayResult = New CNSWIDArrayResult
    ' INamespaceWalk::GetIDArrayResult
    Dim cItems As Long, LpIDListArray As LongPtr
    If VTableCall(vbLong, ObjPtr(NamespaceWalkINamespaceWalk), 5, VarPtr(cItems), VarPtr(LpIDListArray)) = S_OK Then
        GetIDArrayResult.FInit cItems, LpIDListArray
    Else
        GetIDArrayResult.FInit 0, NULL_PTR
    End If
End If
End Function

Private Function CreateINamespaceWalk() As IUnknown
Dim CLSID As CLSID, IID As CLSID
On Error Resume Next
CLSIDFromString StrPtr(CLSID_NamespaceWalker), CLSID
CLSIDFromString StrPtr(IID_INamespaceWalk), IID
CoCreateInstance CLSID, NULL_PTR, CLSCTX_INPROC_SERVER, IID, CreateINamespaceWalk
End Function

Private Function VTableCall(ByVal RetType As VbVarType, ByVal InterfacePointer As LongPtr, ByVal Entry As Long, ParamArray ArgList() As Variant) As Variant
Debug.Assert Not (Entry < 1 Or InterfacePointer = NULL_PTR)
Dim VarArgList As Variant, HResult As Long
VarArgList = ArgList
If UBound(VarArgList) > -1 Then
    Dim i As Long, ArrVarType() As Integer, ArrVarPtr() As LongPtr
    ReDim ArrVarType(LBound(VarArgList) To UBound(VarArgList)) As Integer
    ReDim ArrVarPtr(LBound(VarArgList) To UBound(VarArgList)) ' As LongPtr
    For i = LBound(VarArgList) To UBound(VarArgList)
        ArrVarType(i) = VarType(VarArgList(i))
        ArrVarPtr(i) = VarPtr(VarArgList(i))
    Next i
    HResult = DispCallFunc(InterfacePointer, (Entry - 1) * PTR_SIZE, CC_STDCALL, RetType, i, VarPtr(ArrVarType(0)), VarPtr(ArrVarPtr(0)), VTableCall)
Else
    HResult = DispCallFunc(InterfacePointer, (Entry - 1) * PTR_SIZE, CC_STDCALL, RetType, 0, NULL_PTR, NULL_PTR, VTableCall)
End If
SetLastError HResult ' S_OK will clear the last error code, if any.
End Function

Private Function UnsignedAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr

#If Win64 Then

UnsignedAdd = Start + Incr

#Else

UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000

#End If

End Function
