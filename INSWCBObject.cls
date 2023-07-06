VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INSWCBObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
Public Property Get LpIShellFolder() As LongPtr
Attribute LpIShellFolder.VB_Description = "Interface method to return a pointer to an IShellFolder object."
#Else
Public Property Get LpIShellFolder() As Long
Attribute LpIShellFolder.VB_Description = "Interface method to return a pointer to an IShellFolder object."
#End If
End Property

#If VBA7 Then
Public Property Get LpIDList() As LongPtr
Attribute LpIDList.VB_Description = "Interface method to return a pointer to a PIDL object."
#Else
Public Property Get LpIDList() As Long
Attribute LpIDList.VB_Description = "Interface method to return a pointer to a PIDL object."
#End If
End Property

Public Function GetDisplayNameOf() As String
Attribute GetDisplayNameOf.VB_Description = "Support IShellFolder::GetDisplayNameOf"
End Function
