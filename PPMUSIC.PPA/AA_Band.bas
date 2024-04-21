Attribute VB_Name = "AA_Band"
Option Explicit
Function GetDefaultBand(pStyle As Long) As String
    Dim strBand As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetDefaultBand(strBand, pStyle)
    iNull = InStr(1, strBand, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetDefaultBand = Left$(strBand, iLen)
End Function
Function GetFirstBand(pStyle As Long) As String
    Dim strBand As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetFirstBand(strBand, pStyle)
    iNull = InStr(1, strBand, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetFirstBand = Left$(strBand, iLen)
End Function
Function GetNextBand(pStyle As Long, strPrevBand As String) As String
    Dim strBand As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetNextBand(strBand, pStyle, strPrevBand)
    iNull = InStr(1, strBand, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetNextBand = Left$(strBand, iLen)
End Function
Sub InitBandList()
    Dim strBand As String
    Dim strDefaultBand As String
    Dim i As Integer
    Dim nCount As Integer
    DlgMT.ComboBand.Clear
    strDefaultBand = gstrTagBand
    If (Len(strDefaultBand) = 0) Or (gfUseTagState = False) Then
        strDefaultBand = GetDefaultBand(gpStyle)
    End If
    strBand = GetFirstBand(gpStyle)
    While Len(strBand) <> 0
        If DlgMT.ComboBand.ListCount = 0 Then
            DlgMT.ComboBand.AddItem strBand
        Else
            nCount = DlgMT.ComboBand.ListCount - 1
            For i = 0 To nCount
                If (StrComp(strBand, DlgMT.ComboBand.List(i)) = -1) Then
                    DlgMT.ComboBand.AddItem strBand, i
                    Exit For
                ElseIf i = nCount Then
                    DlgMT.ComboBand.AddItem strBand
                    Exit For
                End If
            Next i
        End If
        strBand = GetNextBand(gpStyle, strBand)
    Wend
    If DlgMT.ComboBand.ListCount > 0 Then
        nCount = DlgMT.ComboBand.ListCount - 1
        For i = 0 To nCount
            If (StrComp(strDefaultBand, DlgMT.ComboBand.List(i)) = 0) Then
                DlgMT.ComboBand.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub
Sub SetupNewBand()
    Dim strBand As String
    If gfStyleChange = True Then
        Exit Sub
    End If
    If DlgMT.ComboBand.ListIndex = -1 Then
        Exit Sub
    End If
    strBand = DlgMT.ComboBand.List(DlgMT.ComboBand.ListIndex)
    MT_SetBand gpStyle, strBand
    If DlgMT.BtnPlayStyle.Value = True Then
        MT_FlushMusicQueue
        QueueMusic
    End If
End Sub
