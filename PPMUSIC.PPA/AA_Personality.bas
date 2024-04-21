Attribute VB_Name = "AA_Personality"
Option Explicit
Function GetDefaultPersonality(pStyle As Long) As String
    Dim strPersonality As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetDefaultPersonality(strPersonality, pStyle)
    iNull = InStr(1, strPersonality, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetDefaultPersonality = Left$(strPersonality, iLen)
End Function
Function GetFirstPersonality(pStyle As Long) As String
    Dim strPersonality As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetFirstPersonality(strPersonality, pStyle)
    iNull = InStr(1, strPersonality, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetFirstPersonality = Left$(strPersonality, iLen)
End Function
Function GetNextPersonality(pStyle As Long, strPrevPersonality As String) As String
    Dim strPersonality As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetNextPersonality(strPersonality, pStyle, strPrevPersonality)
    iNull = InStr(1, strPersonality, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetNextPersonality = Left$(strPersonality, iLen)
End Function
Function GetPersonalityFilename(pStyle As Long, strPersonality As String) As String
    Dim strFilename As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetPersonalityFilename(strFilename, strPersonality)
    iNull = InStr(1, strFilename, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetPersonalityFilename = Left$(strFilename, iLen)
End Function
Sub InitPersonalityList()
    Dim strPersonality As String
    Dim strDefaultPersonality As String
    Dim i As Integer
    Dim nCount As Integer
    DlgMT.ComboPersonality.Clear
    strDefaultPersonality = gstrTagPersonality
    If (Len(strDefaultPersonality) = 0) Or (gfUseTagState = False) Then
        strDefaultPersonality = GetDefaultPersonality(gpStyle)
    End If
    strPersonality = GetFirstPersonality(gpStyle)
    While Len(strPersonality) <> 0
        If DlgMT.ComboPersonality.ListCount = 0 Then
            DlgMT.ComboPersonality.AddItem strPersonality
        Else
            nCount = DlgMT.ComboPersonality.ListCount - 1
            For i = 0 To nCount
                If (StrComp(strPersonality, DlgMT.ComboPersonality.List(i)) = -1) Then
                    DlgMT.ComboPersonality.AddItem strPersonality, i
                    Exit For
                ElseIf i = nCount Then
                    DlgMT.ComboPersonality.AddItem strPersonality
                    Exit For
                End If
            Next i
        End If
        strPersonality = GetNextPersonality(gpStyle, strPersonality)
    Wend
    If DlgMT.ComboPersonality.ListCount > 0 Then
        nCount = DlgMT.ComboPersonality.ListCount - 1
        For i = 0 To nCount
            If (StrComp(strDefaultPersonality, DlgMT.ComboPersonality.List(i)) = 0) Then
                DlgMT.ComboPersonality.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub
Sub SetupNewPersonality()
    Dim strPersonality As String
    Dim strBand As String
    If gfStyleChange = True Then
        Exit Sub
    End If
    If DlgMT.ComboPersonality.ListIndex = -1 Then
        Exit Sub
    End If
    strPersonality = DlgMT.ComboPersonality.List(DlgMT.ComboPersonality.ListIndex)
    If DlgMT.ComboBand.ListIndex = -1 Then
        strBand = ""
    Else
        strBand = DlgMT.ComboBand.List(DlgMT.ComboBand.ListIndex)
    End If
    If DlgMT.BtnPlayStyle.Value = True Then
        MT_SetPersonality gpStyle, strPersonality, strBand
    End If
End Sub
