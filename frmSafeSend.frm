VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSafeSend
   Caption         =   "Confirm External Recipients"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   StartUpPosition =   1
   Begin MSForms.CheckBox chkSelectAll
      Height          =   240
      Left            =   120
      Top             =   120
      Width           =   6360
      Caption         =   "Select all"
   End
   Begin MSForms.CommandButton btnSend
      Height          =   360
      Left            =   4200
      Top             =   5520
      Width           =   1080
      Caption         =   "Send"
   End
   Begin MSForms.CommandButton btnCancel
      Height          =   360
      Left            =   5400
      Top             =   5520
      Width           =   1080
      Caption         =   "Cancel"
   End
End
Attribute VB_Name = "frmSafeSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' SafeSend Replacement - Confirmation Dialog
'==============================================================================
' This form is shown by SafeSendReplacement.bas when external recipients
' are detected. Users must tick every recipient and attachment checkbox
' (or use "Select all") before the Send button will work.
'==============================================================================
Option Explicit

Private m_confirmed As Boolean
Private m_recipients As Collection    ' "type|displayText" e.g. "TO|John <john@ext.com>"
Private m_attachments As Collection   ' "filename|sizeText" e.g. "report.pdf|1.2 MB"
Private m_subject As String
Private m_itemCheckboxes As Collection ' all dynamic checkboxes in the frame

' ---------------------------------------------------------------------------
'  PUBLIC INTERFACE
' ---------------------------------------------------------------------------

Public Property Get UserConfirmed() As Boolean
    UserConfirmed = m_confirmed
End Property

' Call this BEFORE calling .Show
Public Sub SetData(recipients As Collection, attachments As Collection, subject As String)
    Set m_recipients = recipients
    Set m_attachments = attachments
    m_subject = subject
End Sub

' Call this AFTER SetData, BEFORE .Show
Public Sub BuildUI()
    m_confirmed = False
    Set m_itemCheckboxes = New Collection

    ' -- Instruction label
    Dim lblHeader As MSForms.Label
    Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader")
    lblHeader.Left = 120
    lblHeader.Top = 60
    lblHeader.Width = 6360
    lblHeader.Height = 240
    lblHeader.Caption = "Please confirm that the following recipient should receive this message"
    lblHeader.WordWrap = True

    ' Position Select All below header
    chkSelectAll.Left = 120
    chkSelectAll.Top = 360
    chkSelectAll.Width = 6360
    chkSelectAll.Value = False

    ' -- Attachment warning label (red text, only if attachments exist)
    Dim yStart As Long
    yStart = 660

    If m_attachments.Count > 0 Then
        Dim lblAttach As MSForms.Label
        Set lblAttach = Me.Controls.Add("Forms.Label.1", "lblAttachWarn")
        lblAttach.Left = 120
        lblAttach.Top = yStart
        lblAttach.Width = 6360
        lblAttach.Height = 240
        lblAttach.Caption = "This email has file(s) attached and they should be confirmed below"
        lblAttach.ForeColor = RGB(255, 0, 0)
        lblAttach.WordWrap = True
        yStart = yStart + 300
    End If

    ' -- Scrollable frame for checkboxes
    Dim fra As MSForms.Frame
    Set fra = Me.Controls.Add("Forms.Frame.1", "fraItems")
    fra.Left = 60
    fra.Top = yStart
    fra.Width = 6480
    fra.Caption = ""
    fra.BorderStyle = 0  ' fmBorderStyleNone
    fra.SpecialEffect = 0  ' flat
    fra.ScrollBars = 2  ' fmScrollBarsVertical
    fra.KeepScrollBarsVisible = 0  ' fmScrollBarsNone (show only when needed)

    ' -- Add recipient checkboxes
    Dim yPos As Long: yPos = 6
    Dim i As Long
    Dim parts() As String
    Dim chk As MSForms.CheckBox
    Dim lbl As MSForms.Label

    For i = 1 To m_recipients.Count
        parts = Split(m_recipients(i), "|")

        ' Type label (TO: / CC: / BCC:)
        Set lbl = fra.Controls.Add("Forms.Label.1", "lblType" & i)
        lbl.Left = 6
        lbl.Top = yPos + 3
        lbl.Width = 420
        lbl.Height = 210
        lbl.Caption = parts(0) & ":"

        ' Checkbox with recipient
        Set chk = fra.Controls.Add("Forms.CheckBox.1", "chkRecip" & i)
        chk.Left = 420
        chk.Top = yPos
        chk.Width = 5940
        chk.Height = 210
        chk.Caption = parts(1)
        chk.Value = False
        m_itemCheckboxes.Add chk

        yPos = yPos + 270
    Next i

    ' -- Add attachment checkboxes
    If m_attachments.Count > 0 Then
        yPos = yPos + 60  ' small gap

        For i = 1 To m_attachments.Count
            parts = Split(m_attachments(i), "|")

            ' "Files:" label (only on first)
            Set lbl = fra.Controls.Add("Forms.Label.1", "lblFile" & i)
            lbl.Left = 6
            lbl.Top = yPos + 3
            lbl.Width = 420
            lbl.Height = 210
            If i = 1 Then
                lbl.Caption = "Files:"
            Else
                lbl.Caption = ""
            End If

            ' Checkbox with filename and size
            Set chk = fra.Controls.Add("Forms.CheckBox.1", "chkFile" & i)
            chk.Left = 420
            chk.Top = yPos
            chk.Width = 5940
            chk.Height = 210
            chk.Caption = parts(0) & "  " & parts(1)
            chk.Value = False
            m_itemCheckboxes.Add chk

            yPos = yPos + 270
        Next i
    End If

    ' -- Size the frame to fit content (scroll if too tall)
    Dim contentHeight As Long
    contentHeight = yPos + 12
    Dim maxFrameHeight As Long
    maxFrameHeight = 4200  ' max visible height before scrolling

    If contentHeight <= maxFrameHeight Then
        fra.Height = contentHeight
        fra.ScrollBars = 0  ' no scrollbars needed
    Else
        fra.Height = maxFrameHeight
        fra.ScrollHeight = contentHeight
    End If

    ' -- Reposition buttons below the frame
    Dim btnY As Long
    btnY = fra.Top + fra.Height + 120

    btnSend.Top = btnY
    btnCancel.Top = btnY

    ' -- Resize form to fit
    Dim formHeight As Long
    formHeight = btnY + btnSend.Height + 180
    Me.Height = formHeight + (Me.Height - Me.InsideHeight)  ' account for title bar
    Me.Width = 6600 + (Me.Width - Me.InsideWidth)
End Sub

' ---------------------------------------------------------------------------
'  EVENT HANDLERS
' ---------------------------------------------------------------------------

Private Sub chkSelectAll_Click()
    Dim chk As MSForms.CheckBox
    Dim i As Long
    For i = 1 To m_itemCheckboxes.Count
        Set chk = m_itemCheckboxes(i)
        chk.Value = chkSelectAll.Value
    Next i
End Sub

Private Sub btnSend_Click()
    ' Verify all checkboxes are ticked
    Dim chk As MSForms.CheckBox
    Dim allChecked As Boolean: allChecked = True
    Dim i As Long

    For i = 1 To m_itemCheckboxes.Count
        Set chk = m_itemCheckboxes(i)
        If chk.Value = False Then
            allChecked = False
            Exit For
        End If
    Next i

    If Not allChecked Then
        MsgBox "Please confirm all recipients and attachments before sending.", _
               vbExclamation, "Confirmation Required"
        Exit Sub
    End If

    m_confirmed = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    m_confirmed = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Treat X button as Cancel
    If CloseMode = 0 Then  ' vbFormControlMenu
        m_confirmed = False
        Cancel = 1
        Me.Hide
    End If
End Sub
