VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DlgMT 
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4695
   HelpContextID   =   14
   OleObjectBlob   =   "DlgMT.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "DlgMT"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub BtnCancel_Click()
    DoCancel
End Sub

Private Sub BtnOK_Click()
    DoOK
End Sub

Private Sub BtnPlayMotif_Click()
    PlayMotif
End Sub

Private Sub BtnPlayMotif_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    PlayMotif
End Sub

Private Sub BtnPlayStyle_Click()
    If BtnPlayStyle.Value = True Then
        StartMusic
    Else
        StopMusic
    End If
End Sub

Private Sub CheckMotif_Click()
    EnableControls
End Sub

Private Sub ComboBand_Change()
    SetupNewBand
End Sub

Private Sub ComboCategory_Change()
    DlgMT.MousePointer = fmMousePointerHourglass
    DoEvents
    InitStyleList
    DlgMT.MousePointer = fmMousePointerDefault
End Sub

Private Sub ComboPersonality_Change()
    SetupNewPersonality
End Sub

Private Sub ComboStyle_Change()
    DlgMT.MousePointer = fmMousePointerHourglass
    DoEvents
    SetupNewStyle
    DlgMT.MousePointer = fmMousePointerDefault
End Sub

Private Sub RadioBegin_Click()
    EnableControls
End Sub

Private Sub RadioContinue_Click()
    If BtnPlayStyle.Value = True Then
        BtnPlayStyle.Value = False
    End If
    EnableControls
End Sub

Private Sub RadioEnd_Click()
    If BtnPlayStyle.Value = True Then
        BtnPlayStyle.Value = False
    End If
    EnableControls
End Sub

Private Sub UserForm_Terminate()
    DoCancel
End Sub
