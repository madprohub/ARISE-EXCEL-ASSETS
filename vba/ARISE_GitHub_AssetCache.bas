Attribute VB_Name = "ARISE_GitHub_AssetCache"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal caller As LongPtr, ByVal url As String, ByVal filename As String, ByVal reserved As Long, ByVal callback As LongPtr) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal caller As Long, ByVal url As String, ByVal filename As String, ByVal reserved As Long, ByVal callback As Long) As Long
#End If
Private Const BASE_URL As String = "https://raw.githubusercontent.com/madprohub/ARISE-EXCEL-ASSETS/main/assets/premium/"
Public Function ARISE_CacheFolder() As String
    ARISE_CacheFolder = Environ$("LOCALAPPDATA") & "\MADPRO\ARISE\PREMIUM_NEON\"
End Function
Public Function ARISE_CacheComplete() As Boolean
    Dim i As Long, f As String
    For i = 0 To 17
        f = ARISE_CacheFolder() & "frame_" & Format$(i, "00") & ".jpg"
        If Dir$(f) = "" Then Exit Function
        If FileLen(f) < 1000 Then Exit Function
    Next i
    ARISE_CacheComplete = True
End Function
Public Sub ARISE_DownloadPremiumFrames()
    Dim p As String, i As Long, f As String, tmp As String, rc As Long
    On Error GoTo Failed
    p = Environ$("LOCALAPPDATA") & "\MADPRO"
    If Dir$(p, vbDirectory) = "" Then MkDir p
    p = p & "\ARISE"
    If Dir$(p, vbDirectory) = "" Then MkDir p
    p = p & "\PREMIUM_NEON"
    If Dir$(p, vbDirectory) = "" Then MkDir p
    For i = 0 To 17
        f = "frame_" & Format$(i, "00") & ".jpg"
        p = ARISE_CacheFolder() & f
        If Dir$(p) <> "" Then
            If FileLen(p) >= 1000 Then GoTo NextFrame
            Kill p
        End If
        tmp = p & ".download"
        If Dir$(tmp) <> "" Then Kill tmp
        rc = URLDownloadToFile(0, BASE_URL & f, tmp, 0, 0)
        If rc <> 0 Then Err.Raise vbObjectError + 901, , "Download failed: " & f
        If Dir$(tmp) = "" Then Err.Raise vbObjectError + 902, , "Missing: " & f
        If FileLen(tmp) < 1000 Then Err.Raise vbObjectError + 903, , "Invalid: " & f
        Name tmp As p
NextFrame:
    Next i
    MsgBox "Premium frames cached.", vbInformation
    Exit Sub
Failed:
    MsgBox "Download incomplete: " & Err.Description, vbExclamation
End Sub
