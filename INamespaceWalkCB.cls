VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INamespaceWalkCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FoundItem(ByVal Item As INSWCBObject)
Attribute FoundItem.VB_Description = "Interface method when an object is found during a namespace walk."
End Sub

Public Sub EnterFolder(ByVal Folder As INSWCBObject, ByRef Result As NSWCBResultConstants)
Attribute EnterFolder.VB_Description = "Interface method when a folder is about to be entered during a namespace walk."
End Sub

Public Sub LeaveFolder(ByVal Folder As INSWCBObject)
Attribute LeaveFolder.VB_Description = "Interface method after a namespace walk through a folder."
End Sub

Public Sub InitializeProgressDialog(ByRef DialogTitle As String)
Attribute InitializeProgressDialog.VB_Description = "Interface method to initializes the caption of the progress dialog box displayed during a namespace walk."
End Sub

Public Sub WalkComplete(ByVal HResult As Long)
Attribute WalkComplete.VB_Description = "Interface method when a namespace walk has been completed or canceled. Use this method to perform any necessary cleanup."
End Sub
