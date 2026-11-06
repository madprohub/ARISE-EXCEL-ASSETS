Attribute VB_Name = "MADPRO_Credits_v3"
Option Explicit
Public Sub ApplyMADPROSignature(ByVal targetForm As Object)
    Const OLD_LABEL As String = "lblMADPROCreatorCredit"
    Const NEW_LABEL As String = "lblMADPROPremiumCredit"
    Dim badge As Object, oldBadge As Object, formW As Single, formH As Single
    On Error GoTo Failed
    formW = CSng(targetForm.InsideWidth)
    formH = CSng(targetForm.InsideHeight)
    On Error Resume Next
    Set oldBadge = targetForm.Controls(OLD_LABEL)
    If Not oldBadge Is Nothing Then oldBadge.Visible = False
    Set badge = targetForm.Controls(NEW_LABEL)
    On Error GoTo Failed
    If badge Is Nothing Then Set badge = targetForm.Controls.Add("Forms.Label.1", NEW_LABEL, True)
    With badge
        .Caption = ChrW$(&HA9) & " 2026 MADPRO  //  CREATED BY MADPRO"
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = RGB(210, 242, 255)
        .BackStyle = 0
        .TextAlign = 3
        .WordWrap = False
        .AutoSize = False
        .Width = 226
        .Height = 13
        .Left = formW - .Width - 27
        .Top = formH - .Height - 20
        .Visible = True
        .ZOrder 0
    End With
    Exit Sub
Failed:
    MsgBox "MADPRO signature could not be displayed: " & Err.Description, vbExclamation, "MADPRO"
End Sub
Public Sub MADPRO_PreviewPremiumSignature()
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "frmSmartSearch" Then
            ApplyMADPROSignature frm
            Exit Sub
        End If
    Next frm
    MsgBox "Open frmSmartSearch first, or add the call in UserForm_Activate.", vbInformation, "MADPRO"
End Sub
