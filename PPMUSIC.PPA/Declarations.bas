Attribute VB_Name = "Declarations"
Option Explicit
Public gstrTagMusic As String
Public gstrTagCategory As String
Public gstrTagStyleGuid As String
Public gstrTagPersonality As String
Public gstrTagBand As String
Public gstrTagMotif As String
Public gpStyle As Long
Public gfUseTagState As Boolean
Public gfStyleChange As Boolean
Public gfHaveAAEngine As Boolean
Public gfMultiSelect As Boolean
Public gfMixedSelect As Boolean
Declare Function MT_GetResourceString Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal nRsrcId As Integer) As Integer
Declare Function MT_LoadMusicEngine Lib "ppmusau.dll" () As Boolean
Declare Sub MT_FreeMusicEngine Lib "ppmusau.dll" ()
Declare Function MT_GetFirstCategory Lib "ppmusau.dll" (ByVal strBuffer As String) As Integer
Declare Function MT_GetNextCategory Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal strPrevCategory As String) As Integer
Declare Function MT_GetFirstStyle Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal strCategory As String) As Integer
Declare Function MT_GetNextStyle Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal strCategory As String, ByVal strPrevStyle As String) As Integer
Declare Function MT_GetStyleName Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal strCategory As String, ByVal strGuid As String) As Integer
Declare Function MT_GetStyleFilename Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal strCategory As String, ByVal strGuid As String) As Integer
Declare Function MT_GetStylePtr Lib "ppmusau.dll" (ByVal strCategory As String, ByVal strGuid As String) As Long
Declare Function MT_GetDefaultBand Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long) As Integer
Declare Function MT_GetFirstBand Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long) As Integer
Declare Function MT_GetNextBand Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long, ByVal strPrevBand As String) As Integer
Declare Function MT_GetFirstMotif Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long) As Integer
Declare Function MT_GetNextMotif Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long, ByVal strPrevMotif As String) As Integer
Declare Function MT_GetDefaultPersonality Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long) As Integer
Declare Function MT_GetFirstPersonality Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long) As Integer
Declare Function MT_GetNextPersonality Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal gpStyle As Long, ByVal strPrevPersonality As String) As Integer
Declare Function MT_GetPersonalityFilename Lib "ppmusau.dll" (ByVal strBuffer As String, ByVal strPersonality As String) As Integer
Declare Function MT_StartMusic Lib "ppmusau.dll" (ByVal gpStyle As Long, ByVal strPersonality As String, ByVal strBand As String) As Long
Declare Sub MT_StopMusic Lib "ppmusau.dll" ()
Declare Sub MT_PlayMotif Lib "ppmusau.dll" (ByVal gpStyle As Long, ByVal strMotif As String)
Declare Sub MT_QueueMotif Lib "ppmusau.dll" (ByVal gpStyle As Long, ByVal strMotif As String)
Declare Sub MT_SetBand Lib "ppmusau.dll" (ByVal gpStyle As Long, ByVal strBand As String)
Declare Sub MT_SetPersonality Lib "ppmusau.dll" (ByVal gpStyle As Long, ByVal strPersonality As String, ByVal strBand As String)
Declare Sub MT_SetStyle Lib "ppmusau.dll" (ByVal gpStyle As Long)
Declare Function MT_QueueMusic Lib "ppmusau.dll" (ByVal gpStyle As Long, ByVal strPersonality As String, ByVal strBand As String) As Long
Declare Sub MT_FlushMusicQueue Lib "ppmusau.dll" ()
Declare Sub MT_SetNotifySink Lib "ppmusau.dll" ()
Declare Sub MT_RestoreNotifySink Lib "ppmusau.dll" ()
Declare Function MT_MIDIOut Lib "ppmusau.dll" (ByVal fFlag As Long) As Boolean
Public Const AAMT_MUSIC_BEGIN As String = "1"
Public Const AAMT_MUSIC_END As String = "2"
Public Const AAMT_MUSIC_CONTINUE As String = "3"
Public Const TAG_MUSIC As String = "AAMT_Music"
Public Const TAG_CATEGORY As String = "AAMT_Group"
Public Const TAG_STYLEBITS As String = "AAMT_StyleBits"
Public Const TAG_STYLEGUID As String = "AAMT_Style"
Public Const TAG_PERSONALITYBITS As String = "AAMT_MoodBits"
Public Const TAG_PERSONALITY As String = "AAMT_Mood"
Public Const TAG_BAND As String = "AAMT_Band"
Public Const TAG_MOTIF As String = "AAMT_Motif"
Public Const GET_MIDI_OUT As Long = 1
Public Const FREE_MIDI_OUT As Long = 2
Public Const IDS_APP_NAME As Integer = 1
Public Const IDS_MUSICTRACK_MENU As Integer = 3
Public Const IDS_NARRATION_MENU As Integer = 4
Public Const IDS_NO_SLIDE As Integer = 9
Public Const DLG_TITLE As Integer = 11
Public Const DLG_PROMPT As Integer = 12
Public Const DLG_PROMPT_MULTI As Integer = 13
Public Const DLG_RADIO_CONTINUE As Integer = 14
Public Const DLG_RADIO_END As Integer = 15
Public Const DLG_RADIO_BEGIN  As Integer = 16
Public Const DLG_STYLE_GROUP As Integer = 17
Public Const DLG_STYLE  As Integer = 18
Public Const DLG_PERSONALITY As Integer = 19
Public Const DLG_BAND As Integer = 20
Public Const DLG_CHECK_MOTIF As Integer = 21
Public Const DLG_OK As Integer = 22
Public Const DLG_CANCEL As Integer = 23
Public Const DLG_SAMPLE_MOTIF  As Integer = 24
Public Const DLG_SAMPLE_STYLE As Integer = 25
Public Const IDS_GENERIC_ERROR As Integer = 27
