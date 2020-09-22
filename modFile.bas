Attribute VB_Name = "modFile"
Option Explicit

Function FileExist(ByVal Filepath As String) As Boolean

      Dim Filenum As Integer
    Filenum = FreeFile
    If Len(Filepath) > 2 Then
        On Error Resume Next
        Open Filepath For Input As #Filenum
        Close #Filenum
        FileExist = (Err = 0)
        On Error GoTo 0
    End If

End Function

Function FileToString(ByVal Filepath As String) As String

  Dim Filenum As Integer

    If FileExist(Filepath) Then
        Filenum = FreeFile
        Open Filepath For Binary As #Filenum
        FileToString = Space$(LOF(Filenum))
        Get #Filenum, , FileToString
        Close #Filenum
    End If

End Function

Sub FileSave(ByVal StringTxt As String, ByVal Filepath As String)

  Dim Filenum As Integer

    Filenum = FreeFile
    If FileExist(Filepath) Then KillFile Filepath
    Open Filepath For Binary As #Filenum
    Put #Filenum, , StringTxt
    Close #Filenum

End Sub

Public Sub FileAppend(ByVal StringTxt As String, ByVal Filepath As String)

    On Error Resume Next
    Dim Filenum As Integer
    Filenum = FreeFile
    Open Filepath For Append As #Filenum
    Print #Filenum, StringTxt
    Close #Filenum

End Sub

Public Sub KillFile(ByVal Filepath As String)

    On Error Resume Next
    SetAttr Filepath, vbNormal
    Kill Filepath

End Sub

Function FolderExist(ByVal FolderPath As String) As Boolean

    On Error Resume Next
    If Len(FolderPath) > 1 Then
        If Right$(FolderPath, 1) = "\" Then
            FolderPath = FolderPath & "nul"
        Else
            FolderPath = FolderPath & "\nul"
        End If
        FolderExist = (Dir$(FolderPath, vbDirectory) <> vbNullString)
    End If

End Function

Function StrToPic(ByVal Av As String) As IPictureDisp

    On Error Resume Next
    Dim Filepath As String

    Filepath = App.Path & "\temp.tmp"

    If Len(Av) = 0 Then Av = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)

    FileSave Av, Filepath
    Set StrToPic = LoadPicture(Filepath)

End Function

Sub MakeIniPretty(Optional ByVal strFilePath As String)

    On Error GoTo NoFile
    Dim strBuffer As String

    strBuffer = FileToString(strFilePath)

    If InStr(strBuffer, vbCrLf & "[") Then
        strBuffer = Replace$(strBuffer, vbCrLf & vbCrLf, vbCrLf)
        strBuffer = Replace$(strBuffer, vbCrLf & "[", vbCrLf & vbCrLf & "[")
        Kill strFilePath
        FileSave strBuffer, strFilePath
    End If

NoFile:

End Sub

Function FormatBytes(ByVal BytesCount As Double) As String

    If BytesCount >= 1024000000 Then
        FormatBytes = Format$(BytesCount / 1024000000, "###,###,###,###,##0.00") & " Gb"
    ElseIf BytesCount >= 1024000 Then    'Checks to see if its big enough to convert into MB
        FormatBytes = Format$(BytesCount / 1024000, "###,###,##0.00") & " Mb"
    ElseIf BytesCount >= 1024 Then    'Checks to see if file is big enough to convert into KB
        FormatBytes = Format$(BytesCount / 1024, "0.00") & " Kb"
    Else
        FormatBytes = CStr(BytesCount) & " Bytes"
    End If

End Function

Function TimeRemaining(ByVal StartTime As Date, ByVal CurrentPercent As Integer) As String

  Dim TimePassed As Long, MinutesLeft As Long

    If CurrentPercent > 0 Then
        TimePassed = DateDiff("s", StartTime, Time)
        TimePassed = (TimePassed / CurrentPercent) * (100& - CurrentPercent)
        MinutesLeft = TimePassed / 60&
        TimeRemaining = Format$(CLng(MinutesLeft / 60), "00") & ":" & Format$(CLng(MinutesLeft Mod 60), "00") & ":" & Format$(CLng(TimePassed Mod 60), "00")
    Else
        TimeRemaining = "00:00:00"
    End If

End Function

'calulates hours,minutes and seconds from number of milliseconds
Function CalculateTime(ByVal sngTotal As Single) As String

  Dim lhour As Single, lmin As Single, lsec As Single

    If sngTotal >= 3600000 Then
        lhour = sngTotal \ 3600000
        sngTotal = sngTotal - (lhour * 3600000)
        CalculateTime = CStr(lhour) & " Hr. "
    End If

    If sngTotal >= 60000 Then
        lmin = sngTotal \ 60000
        sngTotal = sngTotal - (lmin * 60000)
        CalculateTime = CalculateTime & CStr(lmin) & " Min. "
    End If

    If sngTotal >= 1000 Then
        lsec = sngTotal \ 1000
        CalculateTime = CalculateTime & CStr(lsec) & " Sec."
    Else
        CalculateTime = CalculateTime & ".0" & Left$(CStr(sngTotal), 1) & " Sec."
    End If

End Function

