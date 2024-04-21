Attribute VB_Name = "Tags_Slides"
Option Explicit
Function IsMultiSelect() As Boolean
    IsMultiSelect = False
    With ActiveWindow.View
        Select Case .Type
            Case ppViewOutline
                If ActiveWindow.Selection.SlideRange.Count > 1 Then
                    IsMultiSelect = True
                End If
            Case ppViewSlideSorter
                If ActiveWindow.Selection.Type <> ppSelectionNone Then
                    If ActiveWindow.Selection.SlideRange.Count > 1 Then
                        IsMultiSelect = True
                    End If
                End If
        End Select
    End With
End Function
Function IsMixedSelect() As Boolean
    Dim oSlide As Slide
    Dim i As Integer
    IsMixedSelect = False
    With ActiveWindow.Selection
        For i = 1 To .SlideRange.Count
            Set oSlide = .SlideRange.Item(i)
            If oSlide.Tags(TAG_MUSIC) = AAMT_MUSIC_BEGIN Then
                IsMixedSelect = True
                Exit For
            ElseIf oSlide.Tags(TAG_MUSIC) = AAMT_MUSIC_END Then
                IsMixedSelect = True
                Exit For
            End If
        Next i
    End With
End Function
Function GetCurrentSlide() As Slide
    Set GetCurrentSlide = Nothing
    If Presentations.Count = 0 Then
        TellUser IDS_NO_SLIDE
        Exit Function
    End If
    With ActiveWindow.View
        Select Case .Type
            Case ppViewHandoutMaster
                Set GetCurrentSlide = Nothing
            Case ppViewNotesMaster
                Set GetCurrentSlide = Nothing
            Case ppViewNotesPage
                Set GetCurrentSlide = ActiveWindow.View.Slide.Parent
            Case ppViewOutline
                Set GetCurrentSlide = ActiveWindow.Selection.SlideRange(1)
            Case ppViewSlide
                Set GetCurrentSlide = ActiveWindow.View.Slide
            Case ppViewSlideMaster
                Set GetCurrentSlide = Nothing
            Case ppViewSlideSorter
                If ActiveWindow.Selection.Type <> ppSelectionNone Then
                    Set GetCurrentSlide = ActiveWindow.Selection.SlideRange(1)
                Else
                    Set GetCurrentSlide = Nothing
                End If
            Case ppViewTitleMaster
                Set GetCurrentSlide = Nothing
        End Select
    End With
    If GetCurrentSlide Is Nothing Then
        TellUser IDS_NO_SLIDE
    End If
End Function
Sub GetCurrentTagState()
    Dim oCurrentSlide As Slide
    Dim oSlide As Slide
    Dim i As Integer
    gstrTagMusic = ""
    gstrTagCategory = ""
    gstrTagStyleGuid = ""
    gstrTagPersonality = ""
    gstrTagBand = ""
    gstrTagMotif = ""
    Set oCurrentSlide = GetCurrentSlide
    If oCurrentSlide Is Nothing = False Then
        If oCurrentSlide.Tags(TAG_MUSIC) = AAMT_MUSIC_BEGIN Then
            gstrTagMusic = AAMT_MUSIC_BEGIN
            gstrTagCategory = oCurrentSlide.Tags(TAG_CATEGORY)
            gstrTagStyleGuid = oCurrentSlide.Tags(TAG_STYLEGUID)
            gstrTagPersonality = oCurrentSlide.Tags(TAG_PERSONALITY)
            gstrTagBand = oCurrentSlide.Tags(TAG_BAND)
            gstrTagMotif = oCurrentSlide.Tags(TAG_MOTIF)
        Else
            i = oCurrentSlide.SlideIndex - 1
            While i > 0
                Set oSlide = ActivePresentation.Slides(i)
                If oSlide.Tags(TAG_MUSIC) = AAMT_MUSIC_BEGIN Then
                    gstrTagMusic = AAMT_MUSIC_BEGIN
                    gstrTagCategory = oSlide.Tags(TAG_CATEGORY)
                    gstrTagStyleGuid = oSlide.Tags(TAG_STYLEGUID)
                    gstrTagPersonality = oSlide.Tags(TAG_PERSONALITY)
                    gstrTagBand = oSlide.Tags(TAG_BAND)
                    gstrTagMotif = oSlide.Tags(TAG_MOTIF)
                    i = 0
                End If
                i = i - 1
            Wend
            If oCurrentSlide.Tags(TAG_MUSIC) = AAMT_MUSIC_END Then
                gstrTagMusic = AAMT_MUSIC_END
            Else
                If Len(gstrTagMusic) <> 0 Then
                    gstrTagMusic = AAMT_MUSIC_CONTINUE
                End If
            End If
        End If
    End If
    If Len(gstrTagMusic) = 0 Then
        gstrTagMusic = AAMT_MUSIC_BEGIN
    End If
End Sub
Sub DeleteAllTags(oSlide As Slide)
    With oSlide.Tags
        .Delete TAG_MUSIC
        .Delete TAG_CATEGORY
        .Delete TAG_STYLEBITS
        .Delete TAG_STYLEGUID
        .Delete TAG_PERSONALITYBITS
        .Delete TAG_PERSONALITY
        .Delete TAG_BAND
        .Delete TAG_MOTIF
    End With
End Sub
Sub WriteTagsBegin(oSlide As Slide)
    Dim strFilename As String
    Dim strCategory As String
    Dim strText As String
    strCategory = ""
    With oSlide.Tags
        .Add TAG_MUSIC, AAMT_MUSIC_BEGIN
        If DlgMT.ComboCategory.ListIndex <> -1 Then
            strCategory = DlgMT.ComboCategory.List(DlgMT.ComboCategory.ListIndex)
            .Add TAG_CATEGORY, strCategory
        End If
        If DlgMT.ComboStyle.ListIndex <> -1 Then
            strText = DlgMT.ComboStyle.List(DlgMT.ComboStyle.ListIndex, 1)
            .Add TAG_STYLEGUID, strText
            strFilename = GetStyleFilename(strCategory, strText)
            If Len(strFilename) <> 0 Then
                .AddBinary TAG_STYLEBITS, strFilename
            End If
        End If
        If DlgMT.ComboPersonality.ListIndex <> -1 Then
            strText = DlgMT.ComboPersonality.List(DlgMT.ComboPersonality.ListIndex)
            .Add TAG_PERSONALITY, strText
            strFilename = GetPersonalityFilename(gpStyle, strText)
            If Len(strFilename) <> 0 Then
                .AddBinary TAG_PERSONALITYBITS, strFilename
            End If
        End If
        If DlgMT.ComboBand.ListIndex <> -1 Then
            strText = DlgMT.ComboBand.List(DlgMT.ComboBand.ListIndex)
            .Add TAG_BAND, strText
        End If
        If DlgMT.CheckMotif.Value = True Then
            If DlgMT.ComboMotif.ListIndex <> -1 Then
                strText = DlgMT.ComboMotif.List(DlgMT.ComboMotif.ListIndex)
                .Add TAG_MOTIF, strText
            End If
        End If
    End With
End Sub
Sub WriteTagsEnd(oSlide As Slide)
    With oSlide.Tags
        .Add TAG_MUSIC, AAMT_MUSIC_END
    End With
End Sub
