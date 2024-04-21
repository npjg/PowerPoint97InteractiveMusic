Attribute VB_Name = "AA_Style"
Option Explicit
Function GetFirstStyle(strCategory As String) As String
    Dim strStyleGuid As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetFirstStyle(strStyleGuid, strCategory)
    iNull = InStr(1, strStyleGuid, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetFirstStyle = Left$(strStyleGuid, iLen)
End Function
Function GetNextStyle(strCategory As String, strPrevStyle As String) As String
    Dim strStyleGuid As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetNextStyle(strStyleGuid, strCategory, strPrevStyle)
    iNull = InStr(1, strStyleGuid, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetNextStyle = Left$(strStyleGuid, iLen)
End Function
Function GetStyleName(strCategory As String, strStyleGuid As String) As String
    Dim strStyleName As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetStyleName(strStyleName, strCategory, strStyleGuid)
    iNull = InStr(1, strStyleName, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetStyleName = Left$(strStyleName, iLen)
End Function
Function GetStyleFilename(strCategory As String, strStyleGuid As String) As String
    Dim strFilename As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetStyleFilename(strFilename, strCategory, strStyleGuid)
    iNull = InStr(1, strFilename, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetStyleFilename = Left$(strFilename, iLen)
End Function
Sub InitStyleList()
    Dim strCategory As String
    Dim strStyleGuid As String
    Dim strStyleName As String
    Dim i As Integer
    Dim nIndex As Integer
    If (DlgMT.ComboCategory.ListIndex = -1) Then
        Exit Sub
    End If
    i = 0
    nIndex = 0
    DlgMT.ComboStyle.Clear
    strCategory = DlgMT.ComboCategory.List(DlgMT.ComboCategory.ListIndex)
    strStyleGuid = GetFirstStyle(strCategory)
    While Len(strStyleGuid) <> 0
        strStyleName = GetStyleName(strCategory, strStyleGuid)
        DlgMT.ComboStyle.AddItem i
        DlgMT.ComboStyle.List(i, 0) = strStyleName
        DlgMT.ComboStyle.List(i, 1) = strStyleGuid
        If (StrComp(gstrTagStyleGuid, strStyleGuid) = 0) Then
            nIndex = i
        End If
        strStyleGuid = GetNextStyle(strCategory, strStyleGuid)
        i = i + 1
    Wend
    If DlgMT.ComboStyle.ListCount > 0 Then
        DlgMT.ComboStyle.ListIndex = nIndex
    End If
End Sub
Sub SetupNewStyle()
    Dim strCategory As String
    Dim strStyleGuid As String
    If DlgMT.ComboStyle.ListIndex = -1 Then
        Exit Sub
    End If
    If DlgMT.ComboCategory.ListIndex = -1 Then
        Exit Sub
    End If
    gfStyleChange = True
    strCategory = DlgMT.ComboCategory.List(DlgMT.ComboCategory.ListIndex)
    strStyleGuid = DlgMT.ComboStyle.Value
    gpStyle = MT_GetStylePtr(strCategory, strStyleGuid)
    InitPersonalityList
    InitBandList
    InitMotifList
    EnableControls
    If DlgMT.BtnPlayStyle.Value = True Then
        MT_SetStyle gpStyle
    End If
    gfStyleChange = False
End Sub
