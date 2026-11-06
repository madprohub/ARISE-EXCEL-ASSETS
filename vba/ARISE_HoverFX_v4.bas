Attribute VB_Name = "ARISE_HoverFX_v4"
Option Explicit
Private mItems As Collection
Private mCursor As StdPicture
Private mActive As ARISE_HoverItem_v4
Public Sub StartHover(ByVal frm As Object, Optional ByVal cursorPath As String = "")
    Dim c As Object, item As ARISE_HoverItem_v4, kind As String, label As String
    StopHover
    Set mItems = New Collection
    If Len(cursorPath) > 0 Then
        On Error Resume Next
        If Len(Dir$(cursorPath)) > 0 Then Set mCursor = LoadPicture(cursorPath)
        Err.Clear
        On Error GoTo 0
    End If
    For Each c In frm.Controls
        kind = ""
        If TypeName(c) = "CommandButton" Then
            label = UCase$(Trim$(CStr(c.Caption) & " " & CStr(c.Name)))
            If InStr(label, "ACTIVATE!") > 0 Or InStr(label, "OPEN SELECTED") > 0 Or InStr(label, "SHOW PRODUCT") > 0 Then
                kind = "OPEN"
            ElseIf InStr(label, "CLEAR FILTER!") > 0 Or InStr(label, "CLEAR FILTERS") > 0 Or InStr(label, "RESET") > 0 Then
                kind = "CLEAR"
            ElseIf InStr(label, "EXIT>>") > 0 Or InStr(label, "CLOSE") > 0 Then
                kind = "CLOSE"
            ElseIf InStr(label, "WORKING SHEET NAVIGATION") > 0 Or InStr(label, "GO TO WORKING SHEET") > 0 Then
                kind = "WORK"
            End If
            If Len(kind) > 0 Then
                Set item = New ARISE_HoverItem_v4
                item.AttachButton c, kind, mCursor
                mItems.Add item
            End If
        ElseIf TypeName(c) = "ListBox" Then
            If c.Height > 90 Then
                Set item = New ARISE_HoverItem_v4
                item.AttachList c, mCursor
                mItems.Add item
            End If
        End If
    Next c
    If Not mCursor Is Nothing Then
        On Error Resume Next
        Set frm.MouseIcon = mCursor
        frm.MousePointer = 99
        On Error GoTo 0
    End If
End Sub
Public Sub HighlightItem(ByVal item As ARISE_HoverItem_v4)
    If mItems Is Nothing Then Exit Sub
    If Not mActive Is Nothing Then
        If mActive Is item Then Exit Sub
        mActive.ResetVisual
    End If
    Set mActive = item
    item.SetHighlight
End Sub
Public Sub ClearHover()
    If Not mActive Is Nothing Then mActive.ResetVisual
    Set mActive = Nothing
End Sub
Public Sub StopHover()
    ClearHover
    Set mItems = Nothing
    Set mCursor = Nothing
End Sub
Public Sub ShowDetectedControls(ByVal frm As Object)
    Dim c As Object, msg As String
    For Each c In frm.Controls
        If TypeName(c) = "CommandButton" Or TypeName(c) = "ListBox" Then
            msg = msg & TypeName(c) & "  (Name)=" & c.Name
            If TypeName(c) = "CommandButton" Then msg = msg & "  Caption=" & c.Caption
            msg = msg & vbCrLf
        End If
    Next c
    MsgBox msg, vbInformation, "ARISE Control Names"
End Sub
