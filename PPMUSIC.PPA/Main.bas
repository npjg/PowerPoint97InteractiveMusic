Attribute VB_Name = "Main"
Option Explicit
Sub Auto_Open()
    gfHaveAAEngine = False
    MenuSetup
End Sub
Sub Auto_Close()
    MenuCleanup
    If gfHaveAAEngine <> False Then
        MT_FreeMusicEngine
    End If
    gfHaveAAEngine = False
End Sub
Private Sub MenuSetup()
    Dim oSlideShowMenu As CommandBarControl
    Dim oSlideShowMenuItems As CommandBarControls
    Dim oItem As CommandBarControl
    Dim i, nIndex As Integer
    Dim s As String
    Dim strMusicTrackMenu As String
    Dim strNarrationMenu As String
    Dim fMenuExists As Boolean
    strMusicTrackMenu = GetResourceString(IDS_MUSICTRACK_MENU)
    strNarrationMenu = GetResourceString(IDS_NARRATION_MENU)
    Set oSlideShowMenu = Application.CommandBars("Menu Bar").Controls("Slide Show")
    Set oSlideShowMenuItems = oSlideShowMenu.Control.CommandBar.Controls
    fMenuExists = False
    nIndex = oSlideShowMenuItems.Count
    For i = 1 To oSlideShowMenuItems.Count
        s = oSlideShowMenuItems(i).Caption
        If s = strNarrationMenu Then
            nIndex = i
        End If
        If s = strMusicTrackMenu Then
            fMenuExists = True
        End If
    Next i
    If fMenuExists = False Then
        nIndex = nIndex + 1
        Set oItem = oSlideShowMenuItems.Add(Type:=msoControlButton, Id:=1, Before:=nIndex, Temporary:=True)
        oItem.Caption = strMusicTrackMenu
        oItem.OnAction = "OpenDlgMT"
    End If
End Sub
Private Sub MenuCleanup()
    Dim oSlideShowMenu As CommandBarControl
    Dim oSlideShowMenuItems As CommandBarControls
    Dim oItem As CommandBarControl
    Dim i, nIndex As Integer
    Dim s As String
    Dim strMusicTrackMenu As String
    Dim fMenuExists As Boolean
    strMusicTrackMenu = GetResourceString(IDS_MUSICTRACK_MENU)
    Set oSlideShowMenu = Application.CommandBars("Menu Bar").Controls("Slide Show")
    Set oSlideShowMenuItems = oSlideShowMenu.Control.CommandBar.Controls
    fMenuExists = False
    nIndex = oSlideShowMenuItems.Count
    For i = 1 To oSlideShowMenuItems.Count
        s = oSlideShowMenuItems(i).Caption
        If s = strMusicTrackMenu Then
            nIndex = i
            fMenuExists = True
        End If
    Next i
    If fMenuExists = True Then
        Set oItem = oSlideShowMenuItems(nIndex)
        oItem.Delete
    End If
End Sub
Function GetResourceString(iID As Integer) As String
    Dim sTemp As String * 256
    Dim iLen As Integer
    Dim iNull As Integer
    iLen = MT_GetResourceString(sTemp, iID)
    iNull = InStr(1, sTemp, Chr$(0))
    If (iNull > 0) And (iNull < iLen) Then
        iLen = iNull - 1
    End If
    GetResourceString = Left$(sTemp, iLen)
End Function
Sub TellUser(iID As Integer)
    Dim strMessage As String
    strMessage = GetResourceString(iID)
    MsgBox strMessage
End Sub
