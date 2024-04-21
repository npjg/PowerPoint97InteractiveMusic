Attribute VB_Name = "AA_Motif"
Option Explicit
Function GetFirstMotif(pStyle As Long) As String
    Dim strMotif As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetFirstMotif(strMotif, pStyle)
    iNull = InStr(1, strMotif, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetFirstMotif = Left$(strMotif, iLen)
End Function
Function GetNextMotif(pStyle As Long, strPrevMotif As String) As String
    Dim strMotif As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetNextMotif(strMotif, pStyle, strPrevMotif)
    iNull = InStr(1, strMotif, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetNextMotif = Left$(strMotif, iLen)
End Function
Sub InitMotifList()
    Dim strMotif As String
    Dim strDefaultMotif As String
    Dim i As Integer
    Dim nCount As Integer
    DlgMT.ComboMotif.Clear
    strDefaultMotif = gstrTagMotif
    If (Len(strDefaultMotif) = 0) Or (gfUseTagState = False) Then
        strDefaultMotif = ""
    End If
    strMotif = GetFirstMotif(gpStyle)
    While Len(strMotif) <> 0
        If DlgMT.ComboMotif.ListCount = 0 Then
            DlgMT.ComboMotif.AddItem strMotif
        Else
            nCount = DlgMT.ComboMotif.ListCount - 1
            For i = 0 To nCount
                If (StrComp(strMotif, DlgMT.ComboMotif.List(i)) = -1) Then
                    DlgMT.ComboMotif.AddItem strMotif, i
                    Exit For
                ElseIf i = nCount Then
                    DlgMT.ComboMotif.AddItem strMotif
                    Exit For
                End If
            Next i
        End If
        strMotif = GetNextMotif(gpStyle, strMotif)
    Wend
    If DlgMT.ComboMotif.ListCount > 0 Then
        DlgMT.ComboMotif.ListIndex = 0
        nCount = DlgMT.ComboMotif.ListCount - 1
        For i = 0 To nCount
            If (StrComp(strDefaultMotif, DlgMT.ComboMotif.List(i)) = 0) Then
                DlgMT.ComboMotif.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub
