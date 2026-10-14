Attribute VB_Name = "modDragForm"
Option Explicit

Public DragHandlers As Collection

Public Sub MakeDraggable(ByVal TargetControl As Object, ByVal ParentUserForm As Object)
    If DragHandlers Is Nothing Then
        Set DragHandlers = New Collection
    End If
    
    Dim newHandler As clsDragHandler
    Set newHandler = New clsDragHandler
    
    newHandler.Hook TargetControl, ParentUserForm
    DragHandlers.Add newHandler
End Sub
