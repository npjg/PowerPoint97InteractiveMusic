Attribute VB_Name = "Form"
Option Explicit
Sub OpenDlgMT()
    Dim oSlide As Slide
    Dim strLabel As String
    On Error GoTo ErrorTrap
    If gfHaveAAEngine = False Then
        gfHaveAAEngine = MT_LoadMusicEngine()
        If gfHaveAAEngine = False Then
            Exit Sub
        End If
    End If
    Set oSlide = GetCurrentSlide
    If oSlide Is Nothing Then
        Exit Sub
    Else
        Set oSlide = Nothing
    End If
    DlgMT.Hide
    DlgMT.MousePointer = fmMousePointerHourglass
    DlgMT.BtnPlayStyle.Value = False
    MT_StopMusic
    gfUseTagState = True
    gfStyleChange = False
    gfMultiSelect = IsMultiSelect()
    If gfMultiSelect = True Then
        gfMixedSelect = IsMixedSelect()
        strLabel = GetResourceString(DLG_PROMPT_MULTI)
    Else
        gfMixedSelect = False
        strLabel = GetResourceString(DLG_PROMPT)
    End If
    DlgMT.LabelTitle.Caption = strLabel
    GetCurrentTagState
    If Len(DlgMT.Caption) = 0 Then
        InitLabels
    End If
    If DlgMT.ComboCategory.ListCount <= 0 Then
        InitCategoryList
    End If
    InitRadioButtons
    SelectCategory
    InitCheckMotifButton
    gfUseTagState = False
    EnableControls
    MT_SetNotifySink
    DlgMT.MousePointer = fmMousePointerDefault
    DlgMT.Show
    Exit Sub
ErrorTrap:
    TellUser IDS_GENERIC_ERROR
    If gfHaveAAEngine <> False Then
        MT_StopMusic
        MT_RestoreNotifySink
        MT_MIDIOut FREE_MIDI_OUT
    End If
    DlgMT.Hide
End Sub
Private Sub InitLabels()
    Dim strLabel As String
    Dim strNewLabel As String
    Dim strAcc As String
    Dim iPos As Integer
    Dim iLen As Integer
    strLabel = GetResourceString(DLG_TITLE)
    DlgMT.Caption = strLabel
    strLabel = GetResourceString(DLG_RADIO_CONTINUE)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.RadioContinue.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.RadioContinue.Caption = strLabel
    strLabel = GetResourceString(DLG_RADIO_END)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.RadioEnd.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.RadioEnd.Caption = strLabel
    strLabel = GetResourceString(DLG_RADIO_BEGIN)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.RadioBegin.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.RadioBegin.Caption = strLabel
    strLabel = GetResourceString(DLG_STYLE_GROUP)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.LabelCategory.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.LabelCategory.Caption = strLabel
    strLabel = GetResourceString(DLG_STYLE)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.LabelStyle.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.LabelStyle.Caption = strLabel
    strLabel = GetResourceString(DLG_PERSONALITY)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.LabelPersonality.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.LabelPersonality.Caption = strLabel
    strLabel = GetResourceString(DLG_BAND)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.LabelBand.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.LabelBand.Caption = strLabel
    strLabel = GetResourceString(DLG_CHECK_MOTIF)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.CheckMotif.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.CheckMotif.Caption = strLabel
    strLabel = GetResourceString(DLG_SAMPLE_STYLE)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.BtnPlayStyle.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.BtnPlayStyle.Caption = strLabel
    strLabel = GetResourceString(DLG_SAMPLE_MOTIF)
    iLen = Len(strLabel)
    iPos = InStr(strLabel, "&")
    If (iPos > 0) And (iPos < iLen) Then
        strNewLabel = Left$(strLabel, iPos - 1)
        strNewLabel = strNewLabel & Right$(strLabel, iLen - iPos)
        strAcc = Mid$(strLabel, iPos + 1, 1)
        DlgMT.BtnPlayMotif.Accelerator = strAcc
        strLabel = strNewLabel
    End If
    DlgMT.BtnPlayMotif.Caption = strLabel
    strLabel = GetResourceString(DLG_OK)
    DlgMT.BtnOK.Caption = strLabel
    strLabel = GetResourceString(DLG_CANCEL)
    DlgMT.BtnCancel.Caption = strLabel
End Sub
Private Sub InitRadioButtons()
    If gfMultiSelect = True Then
        If gfMixedSelect = True Then
            DlgMT.RadioContinue.Value = False
        Else
            DlgMT.RadioContinue.Value = True
        End If
        DlgMT.RadioEnd.Value = False
        DlgMT.RadioBegin.Value = False
    Else
        If (StrComp(gstrTagMusic, AAMT_MUSIC_BEGIN) = 0) Then
            DlgMT.RadioContinue.Value = False
            DlgMT.RadioEnd.Value = False
            DlgMT.RadioBegin.Value = True
        ElseIf (StrComp(gstrTagMusic, AAMT_MUSIC_END) = 0) Then
            DlgMT.RadioContinue.Value = False
            DlgMT.RadioEnd.Value = True
            DlgMT.RadioBegin.Value = False
        Else
            DlgMT.RadioContinue.Value = True
            DlgMT.RadioEnd.Value = False
            DlgMT.RadioBegin.Value = False
        End If
    End If
End Sub
Private Sub InitCheckMotifButton()
    If Len(gstrTagMotif) = 0 Then
        DlgMT.CheckMotif.Value = False
    Else
        DlgMT.CheckMotif.Value = True
    End If
End Sub
Sub EnableControls()
    Dim fEnable As Boolean
    If gfUseTagState = True Then
        Exit Sub
    End If
    If gfMultiSelect = True Then
        DlgMT.RadioBegin.Enabled = False
        DlgMT.RadioEnd.Enabled = False
    Else
        DlgMT.RadioBegin.Enabled = True
        DlgMT.RadioEnd.Enabled = True
    End If
    If DlgMT.RadioBegin.Value = True Then
        fEnable = True
    Else
        fEnable = False
    End If
    DlgMT.ComboCategory.Enabled = fEnable
    DlgMT.ComboStyle.Enabled = fEnable
    DlgMT.ComboPersonality.Enabled = fEnable
    DlgMT.ComboBand.Enabled = fEnable
    DlgMT.CheckMotif.Enabled = fEnable
    DlgMT.ComboMotif.Enabled = fEnable
    DlgMT.BtnPlayStyle.Enabled = fEnable
    DlgMT.BtnPlayMotif.Enabled = fEnable
    If fEnable = True Then
        If DlgMT.ComboPersonality.ListCount = 0 Then
            DlgMT.ComboPersonality.Enabled = False
        End If
        If DlgMT.ComboBand.ListCount = 0 Then
            DlgMT.ComboBand.Enabled = False
        End If
        If DlgMT.ComboMotif.ListCount = 0 Then
            DlgMT.CheckMotif.Enabled = False
            DlgMT.ComboMotif.Enabled = False
            DlgMT.BtnPlayMotif.Enabled = False
        ElseIf DlgMT.CheckMotif.Value <> True Then
            DlgMT.ComboMotif.Enabled = False
            DlgMT.BtnPlayMotif.Enabled = False
        End If
    End If
    If DlgMT.RadioEnd.Value = False Then
        If DlgMT.RadioBegin.Value = False Then
            If DlgMT.RadioContinue.Value = False Then
                DlgMT.BtnOK.Enabled = False
            Else
                DlgMT.BtnOK.Enabled = True
            End If
        End If
    End If
End Sub
Sub DoCancel()
    MT_StopMusic
    MT_RestoreNotifySink
    MT_MIDIOut FREE_MIDI_OUT
    DlgMT.Hide
End Sub
Sub DoOK()
    Dim oCurrentSlide As Slide
    Dim i As Integer
    DlgMT.MousePointer = fmMousePointerHourglass
    DoEvents
    MT_StopMusic
    DlgMT.Hide
    If gfMultiSelect = True Then
        If DlgMT.RadioContinue.Value = True Then
            With ActiveWindow.Selection
                For i = 1 To .SlideRange.Count
                    Set oCurrentSlide = .SlideRange.Item(i)
                    DeleteAllTags oCurrentSlide
                Next i
            End With
        End If
    Else
        Set oCurrentSlide = GetCurrentSlide
        If oCurrentSlide Is Nothing = False Then
            DeleteAllTags oCurrentSlide
            If DlgMT.RadioBegin.Value = True Then
                WriteTagsBegin oCurrentSlide
            ElseIf DlgMT.RadioEnd.Value = True Then
                WriteTagsEnd oCurrentSlide
            End If
        End If
    End If
   MT_RestoreNotifySink
   MT_MIDIOut FREE_MIDI_OUT
   DlgMT.MousePointer = fmMousePointerDefault
End Sub
Sub QueueMusic()
    Dim strPersonality As String
    Dim strBand As String
    If gpStyle = 0 Then
        Exit Sub
    End If
    If (DlgMT.ComboPersonality.ListIndex = -1) Then
        strPersonality = ""
    Else
        strPersonality = DlgMT.ComboPersonality.List(DlgMT.ComboPersonality.ListIndex)
    End If
    If (DlgMT.ComboBand.ListIndex = -1) Then
        strBand = ""
    Else
        strBand = DlgMT.ComboBand.List(DlgMT.ComboBand.ListIndex)
    End If
    MT_QueueMusic gpStyle, strPersonality, strBand
End Sub
Sub StartMusic()
    Dim strPersonality As String
    Dim strBand As String
    If gpStyle = 0 Then
        Exit Sub
    End If
    If MT_MIDIOut(GET_MIDI_OUT) = False Then
        DlgMT.BtnPlayStyle.Value = False
        Exit Sub
    End If
    If (DlgMT.ComboPersonality.ListIndex = -1) Then
        strPersonality = ""
    Else
        strPersonality = DlgMT.ComboPersonality.List(DlgMT.ComboPersonality.ListIndex)
    End If
    If (DlgMT.ComboBand.ListIndex = -1) Then
        strBand = ""
    Else
        strBand = DlgMT.ComboBand.List(DlgMT.ComboBand.ListIndex)
    End If
    MT_StartMusic gpStyle, strPersonality, strBand
End Sub
Sub StopMusic()
    MT_StopMusic
End Sub
Sub PlayMotif()
    Dim strMotif As String
    If (DlgMT.ComboMotif.ListIndex = -1) Then
        Exit Sub
    End If
    If MT_MIDIOut(GET_MIDI_OUT) = False Then
        Exit Sub
    End If
    strMotif = DlgMT.ComboMotif.List(DlgMT.ComboMotif.ListIndex)
    If DlgMT.BtnPlayStyle.Value = False Then
        MT_QueueMotif gpStyle, strMotif
        DlgMT.BtnPlayStyle.Value = True
    Else
        MT_PlayMotif gpStyle, strMotif
    End If
End Sub
