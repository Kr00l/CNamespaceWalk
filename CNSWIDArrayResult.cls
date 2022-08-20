VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNSWIDArrayResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
Private Const MAX_PATH As Long = 260
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal LpIDList As LongPtr, ByVal lpszPath As LongPtr) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" (ByVal LpIDList As Long, ByVal lpszPath As Long) As Long
#End If
Private PropCount As Long
Private PropList() As LongPtr

Private Sub Class_Terminate()
Dim i As Long
For i = 0 To (PropCount - 1)
    CoTaskMemFree PropList(i)
Next i
Erase PropList()
End Sub

Friend Sub FInit(ByVal cItems As Long, ByVal LpIDListArray As LongPtr)
PropCount = cItems
If LpIDListArray <> NULL_PTR Then
    Dim i As Long
    ReDim PropList(0 To (cItems - 1)) ' As LongPtr
    For i = 0 To (cItems - 1)
        CopyMemory PropList(i), ByVal UnsignedAdd(LpIDListArray, i * PTR_SIZE), PTR_SIZE
    Next i
    CoTaskMemFree LpIDListArray
End If
End Sub

Public Property Get Object() As Object
Attribute Object.VB_Description = "Returns the own instance."
Set Object = Me
End Property

#If VBA7 Then

Public Property Get LpIDList(ByVal Index As Long) As LongPtr
Attribute LpIDList.VB_Description = "Returns a pointer to a PIDL object given its index."
Attribute LpIDList.VB_UserMemId = 0

#Else

Public Property Get LpIDList(ByVal Index As Long) As Long
Attribute LpIDList.VB_Description = "Returns a pointer to a PIDL object given its index."
Attribute LpIDList.VB_UserMemId = 0

#End If

If Index > PropCount Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
LpIDList = PropList(Index - 1)
End Property

Public Function GetPathFromIDList(ByVal Index As Long) As String
Attribute GetPathFromIDList.VB_Description = "Converts a PIDL object to a file system path given its index."
If Index > PropCount Or Index <= 0 Then Err.Raise Number:=35600, Description:="Index out of bounds"
Dim Buffer As String
Buffer = String(MAX_PATH, vbNullChar) & vbNullChar
If SHGetPathFromIDList(PropList(Index - 1), StrPtr(Buffer)) <> 0 Then GetPathFromIDList = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of PIDL objects."
Count = PropCount
End Property

Private Function UnsignedAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr

#If Win64 Then

UnsignedAdd = Start + Incr

#Else

UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000

#End If

End Function
