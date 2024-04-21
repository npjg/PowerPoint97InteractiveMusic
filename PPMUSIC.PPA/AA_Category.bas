Attribute VB_Name = "AA_Category"
Option Explicit
Function GetFirstCategory() As String
    Dim strCategory As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetFirstCategory(strCategory)
    iNull = InStr(1, strCategory, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetFirstCategory = Left$(strCategory, iLen)
End Function
Function GetNextCategory(strPrevCategory As String) As String
    Dim strCategory As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetNextCategory(strCategory, strPrevCategory)
    iNull = InStr(1, strCategory, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetNextCategory = Left$(strCategory, iLen)
End Function
Sub InitCategoryList()
    Dim strCategory As String
    DlgMT.ComboCategory.Clear
    strCategory = GetFirstCategory
    While Len(strCategory) <> 0
        DlgMT.ComboCategory.AddItem strCategory
        strCategory = GetNextCategory(strCategory)
    Wend
End Sub
Sub SelectCategory()
    Dim i As Integer
    Dim nCount As Integer
    If DlgMT.ComboCategory.ListCount > 0 Then
        If Len(gstrTagCategory) <> 0 Then
            nCount = DlgMT.ComboCategory.ListCount - 1
            For i = 0 To nCount
                If (StrComp(gstrTagCategory, DlgMT.ComboCategory.List(i)) = 0) Then
                    DlgMT.ComboCategory.ListIndex = i
                    Exit For
                End If
            Next i
        Else
            DlgMT.ComboCategory.ListIndex = 0
        End If
    End If
End Sub
