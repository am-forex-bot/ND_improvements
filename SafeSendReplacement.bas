Attribute VB_Name = "SafeSendReplacement"
'==============================================================================
' SafeSend Replacement - External Recipient Confirmation for Outlook
'==============================================================================
' PURPOSE:  Replaces VIPRE SafeSend with a lightweight VBA alternative.
'           Shows a checkbox-based confirmation dialog (matching SafeSend's UI)
'           for external recipients AND attachments.
'           Smart enough to only prompt ONCE per email - if another add-in
'           (e.g. NetDocuments ndMail) cancels the send after confirmation,
'           the user won't be prompted again on the retry.
'
' INSTALL:
'   1) Disable/remove the real SafeSend add-in
'   2) In Outlook: File -> Options -> Trust Center -> Trust Center Settings
'      -> Macro Settings -> tick "Trust access to the VBA project object model"
'   3) Alt+F11 -> File -> Import File... -> select this .bas
'   4) Alt+F8 -> run "InstallSafeSend"  (creates the form + event hook)
'   5) Restart Outlook. Done.
'
' CONFIG:   Edit the INTERNAL_DOMAINS constant below.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
'  CONFIGURATION
' ---------------------------------------------------------------------------
' Comma-separated list of your internal domains (case-insensitive).
' Any recipient NOT matching these domains triggers the confirmation dialog.
Private Const INTERNAL_DOMAINS As String = "wallace.co.uk,wallace.onmicrosoft.com"

' Set to True to also check BCC recipients
Private Const CHECK_BCC As Boolean = True

' Custom property name stamped on the mail item after confirmation.
' This prevents re-prompting if another add-in cancels the send.
Private Const CONFIRMED_FLAG As String = "_SafeSendConfirmed"

' ---------------------------------------------------------------------------
'  MAIN ENTRY POINT - called from ThisOutlookSession.Application_ItemSend
' ---------------------------------------------------------------------------
' Returns True to cancel the send, False to allow it.
Public Function SafeSendCheck(ByVal Item As Object) As Boolean
    On Error GoTo ErrHandler
    SafeSendCheck = False  ' default: allow send

    ' Only check mail items (not meeting requests, task requests, etc.)
    If TypeName(Item) <> "MailItem" Then Exit Function

    ' If we already confirmed this draft, skip the check
    If HasBeenConfirmed(Item) Then
        ClearConfirmedFlag Item
        Exit Function
    End If

    ' Collect external recipients
    Dim externals As Collection
    Set externals = GetExternalRecipients(Item)

    ' No external recipients - let it through
    If externals.Count = 0 Then Exit Function

    ' Collect attachments
    Dim attachments As Collection
    Set attachments = GetAttachmentList(Item)

    ' Show the confirmation form
    Dim frm As frmSafeSend
    Set frm = New frmSafeSend
    frm.SetData externals, attachments, Item.Subject
    frm.BuildUI
    frm.Show vbModal

    If frm.UserConfirmed Then
        ' User confirmed everything - stamp so we don't ask again
        StampConfirmedFlag Item
        SafeSendCheck = False  ' allow send
    Else
        ' User cancelled - block the send
        SafeSendCheck = True
    End If

    Unload frm
    Set frm = Nothing

    Exit Function
ErrHandler:
    ' On error, let the email through rather than blocking all sends
    SafeSendCheck = False
End Function

' ---------------------------------------------------------------------------
'  EXTERNAL RECIPIENT DETECTION
' ---------------------------------------------------------------------------

' Returns a Collection of strings in "TYPE|display" format.
' e.g. "TO|John Smith <john@external.com>"
Private Function GetExternalRecipients(ByVal Item As Object) As Collection
    Dim result As New Collection
    Dim domains() As String
    domains = Split(LCase$(INTERNAL_DOMAINS), ",")

    Dim i As Long
    For i = 0 To UBound(domains)
        domains(i) = Trim$(domains(i))
    Next i

    Dim recip As Object  ' Outlook.Recipient
    For Each recip In Item.Recipients
        ' Skip BCC if not configured to check
        If Not CHECK_BCC And recip.Type = 3 Then GoTo NextRecip  ' olBCC = 3

        Dim emailAddr As String
        emailAddr = GetSmtpAddress(recip)

        If Len(emailAddr) > 0 Then
            Dim recipDomain As String
            recipDomain = LCase$(Mid$(emailAddr, InStr(emailAddr, "@") + 1))

            ' Check against all internal domains
            Dim isInternal As Boolean: isInternal = False
            Dim d As Long
            For d = 0 To UBound(domains)
                If recipDomain = domains(d) Then
                    isInternal = True
                    Exit For
                End If
            Next d

            If Not isInternal Then
                Dim displayText As String
                If LCase$(recip.Name) <> LCase$(emailAddr) Then
                    displayText = recip.Name & " <" & emailAddr & ">"
                Else
                    displayText = emailAddr
                End If

                ' Format as "TYPE|display"
                Dim typeLabel As String
                Select Case recip.Type
                    Case 1: typeLabel = "TO"      ' olTo
                    Case 2: typeLabel = "CC"      ' olCC
                    Case 3: typeLabel = "BCC"     ' olBCC
                    Case Else: typeLabel = "TO"
                End Select

                result.Add typeLabel & "|" & displayText
            End If
        End If
NextRecip:
    Next recip

    Set GetExternalRecipients = result
End Function

' ---------------------------------------------------------------------------
'  ATTACHMENT LIST
' ---------------------------------------------------------------------------

' Returns a Collection of strings in "filename|sizeText" format.
' e.g. "report.pdf|1.2 MB"
Private Function GetAttachmentList(ByVal Item As Object) As Collection
    Dim result As New Collection

    On Error Resume Next
    Dim att As Object  ' Outlook.Attachment
    For Each att In Item.Attachments
        ' Skip hidden/inline attachments (embedded images etc.)
        ' Type 1 = olByValue (regular file attachment)
        If att.Type = 1 Then
            Dim sizeText As String
            Dim fileSize As Long
            fileSize = 0

            ' Try to get size (available in Outlook 2010+)
            fileSize = att.Size
            If Err.Number <> 0 Then
                Err.Clear
                sizeText = ""
            Else
                sizeText = FormatFileSize(fileSize)
            End If

            result.Add att.FileName & "|" & sizeText
        End If
    Next att
    On Error GoTo 0

    Set GetAttachmentList = result
End Function

' Formats bytes into a human-readable string.
Private Function FormatFileSize(ByVal bytes As Long) As String
    If bytes < 1024 Then
        FormatFileSize = bytes & " B"
    ElseIf bytes < 1048576 Then
        FormatFileSize = Format$(bytes / 1024, "#,##0") & " KB"
    Else
        FormatFileSize = Format$(bytes / 1048576, "#,##0.0") & " MB"
    End If
End Function

' ---------------------------------------------------------------------------
'  SMTP ADDRESS RESOLUTION
' ---------------------------------------------------------------------------

' Resolves the SMTP email address from a Recipient object.
' Handles both SMTP and Exchange (EX) address types.
Private Function GetSmtpAddress(recip As Object) As String
    On Error GoTo ErrHandler

    Dim pa As Object  ' PropertyAccessor
    Const PR_SMTP As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    ' Try resolving first
    If Not recip.Resolve Then
        ' Can't resolve - try the address as-is
        GetSmtpAddress = recip.Address
        Exit Function
    End If

    Dim addrEntry As Object
    Set addrEntry = recip.AddressEntry

    If addrEntry.Type = "SMTP" Then
        GetSmtpAddress = LCase$(addrEntry.Address)
    ElseIf addrEntry.Type = "EX" Then
        ' Exchange address - get SMTP via PropertyAccessor
        Set pa = recip.PropertyAccessor
        GetSmtpAddress = LCase$(pa.GetProperty(PR_SMTP))
    Else
        ' Other type - try GetExchangeUser
        On Error Resume Next
        Dim exUser As Object
        Set exUser = addrEntry.GetExchangeUser
        If Not exUser Is Nothing Then
            GetSmtpAddress = LCase$(exUser.PrimarySmtpAddress)
        Else
            GetSmtpAddress = LCase$(addrEntry.Address)
        End If
        On Error GoTo ErrHandler
    End If

    Exit Function
ErrHandler:
    ' Fallback: return raw address
    On Error Resume Next
    GetSmtpAddress = LCase$(recip.Address)
End Function

' ---------------------------------------------------------------------------
'  CONFIRMED FLAG (prevents double-prompting)
' ---------------------------------------------------------------------------

' Checks if this mail item has already been confirmed in this session.
Private Function HasBeenConfirmed(ByVal Item As Object) As Boolean
    On Error Resume Next
    Dim prop As Object
    Set prop = Item.UserProperties.Find(CONFIRMED_FLAG)
    HasBeenConfirmed = (Not prop Is Nothing)
    If HasBeenConfirmed Then
        HasBeenConfirmed = (prop.Value = True)
    End If
    On Error GoTo 0
End Function

' Stamps the mail item as confirmed.
Private Sub StampConfirmedFlag(ByVal Item As Object)
    On Error Resume Next
    Dim prop As Object
    Set prop = Item.UserProperties.Add(CONFIRMED_FLAG, 6, False)  ' olYesNo = 6
    prop.Value = True
    Item.Save
    On Error GoTo 0
End Sub

' Clears the confirmed flag (after successful send-through).
Private Sub ClearConfirmedFlag(ByVal Item As Object)
    On Error Resume Next
    Dim props As Object: Set props = Item.UserProperties
    Dim prop As Object: Set prop = props.Find(CONFIRMED_FLAG)
    If Not prop Is Nothing Then
        prop.Delete
    End If
    On Error GoTo 0
End Sub

' ===========================================================================
'  ONE-TIME INSTALLER - run this once via Alt+F8 -> InstallSafeSend
' ===========================================================================
' Creates the frmSafeSend UserForm, adds the 3 design-time controls,
' injects the code-behind, and wires up ThisOutlookSession.
' Requires: "Trust access to the VBA project object model" enabled.

Public Sub InstallSafeSend()
    On Error GoTo ErrHandler

    Dim proj As Object
    Set proj = Application.VBE.ActiveVBProject

    ' -- Check if form already exists
    Dim comp As Object
    Dim formExists As Boolean: formExists = False
    For Each comp In proj.VBComponents
        If comp.Name = "frmSafeSend" Then
            formExists = True
            Exit For
        End If
    Next comp

    If formExists Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("frmSafeSend already exists. Delete and recreate it?", _
                      vbYesNo + vbQuestion, "SafeSend Install")
        If ans = vbNo Then Exit Sub
        proj.VBComponents.Remove proj.VBComponents("frmSafeSend")
    End If

    ' -- Create the UserForm
    Dim frm As Object  ' VBComponent
    Set frm = proj.VBComponents.Add(3)  ' vbext_ct_MSForm = 3
    frm.Name = "frmSafeSend"

    ' Set form properties
    Dim designer As Object
    Set designer = frm.designer
    designer.Caption = "Confirm External Recipients"

    ' -- Add CheckBox: chkSelectAll
    Dim chkAll As Object
    Set chkAll = designer.Controls.Add("Forms.CheckBox.1", "chkSelectAll")
    chkAll.Caption = "Select all"
    chkAll.Left = 8
    chkAll.Top = 24
    chkAll.Width = 416
    chkAll.Height = 16

    ' -- Add CommandButton: btnSend
    Dim btnS As Object
    Set btnS = designer.Controls.Add("Forms.CommandButton.1", "btnSend")
    btnS.Caption = "Send"
    btnS.Left = 290
    btnS.Top = 200
    btnS.Width = 66
    btnS.Height = 24

    ' -- Add CommandButton: btnCancel
    Dim btnC As Object
    Set btnC = designer.Controls.Add("Forms.CommandButton.1", "btnCancel")
    btnC.Caption = "Cancel"
    btnC.Left = 362
    btnC.Top = 200
    btnC.Width = 66
    btnC.Height = 24

    ' -- Inject the code-behind
    Dim code As String
    code = FormCodeBehind()
    frm.CodeModule.DeleteLines 1, frm.CodeModule.CountOfLines
    frm.CodeModule.AddFromString code

    ' -- Wire up ThisOutlookSession if not already done
    Dim tos As Object
    Set tos = proj.VBComponents("ThisOutlookSession")
    Dim existing As String
    existing = tos.CodeModule.Lines(1, tos.CodeModule.CountOfLines)

    If InStr(1, existing, "SafeSendCheck", vbTextCompare) = 0 Then
        tos.CodeModule.AddFromString vbCrLf & _
            "Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)" & vbCrLf & _
            "    Cancel = SafeSendCheck(Item)" & vbCrLf & _
            "End Sub"
    End If

    MsgBox "SafeSend installed successfully!" & vbCrLf & vbCrLf & _
           "Please restart Outlook for the ItemSend hook to take effect.", _
           vbInformation, "SafeSend Install"
    Exit Sub

ErrHandler:
    If Err.Number = 6068 Or InStr(1, Err.Description, "programmatic access", vbTextCompare) > 0 Then
        MsgBox "Access denied." & vbCrLf & vbCrLf & _
               "Please enable: File -> Options -> Trust Center -> " & _
               "Trust Center Settings -> Macro Settings -> " & vbCrLf & _
               """Trust access to the VBA project object model""", _
               vbCritical, "SafeSend Install"
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, _
               vbCritical, "SafeSend Install"
    End If
End Sub

' Returns the complete code-behind for frmSafeSend as a string.
Private Function FormCodeBehind() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf & vbCrLf
    c = c & "Private m_confirmed As Boolean" & vbCrLf
    c = c & "Private m_recipients As Collection" & vbCrLf
    c = c & "Private m_attachments As Collection" & vbCrLf
    c = c & "Private m_subject As String" & vbCrLf
    c = c & "Private m_itemCheckboxes As Collection" & vbCrLf & vbCrLf

    ' -- UserConfirmed property
    c = c & "Public Property Get UserConfirmed() As Boolean" & vbCrLf
    c = c & "    UserConfirmed = m_confirmed" & vbCrLf
    c = c & "End Property" & vbCrLf & vbCrLf

    ' -- SetData
    c = c & "Public Sub SetData(recipients As Collection, attachments As Collection, subject As String)" & vbCrLf
    c = c & "    Set m_recipients = recipients" & vbCrLf
    c = c & "    Set m_attachments = attachments" & vbCrLf
    c = c & "    m_subject = subject" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' -- BuildUI
    c = c & "Public Sub BuildUI()" & vbCrLf
    c = c & "    m_confirmed = False" & vbCrLf
    c = c & "    Set m_itemCheckboxes = New Collection" & vbCrLf & vbCrLf
    c = c & "    Dim lblHeader As MSForms.Label" & vbCrLf
    c = c & "    Set lblHeader = Me.Controls.Add(""Forms.Label.1"", ""lblHeader"")" & vbCrLf
    c = c & "    lblHeader.Left = 8: lblHeader.Top = 4: lblHeader.Width = 416: lblHeader.Height = 16" & vbCrLf
    c = c & "    lblHeader.Caption = ""Please confirm that the following recipient(s) should receive this message""" & vbCrLf
    c = c & "    lblHeader.WordWrap = True" & vbCrLf & vbCrLf
    c = c & "    chkSelectAll.Left = 8: chkSelectAll.Top = 24: chkSelectAll.Width = 416" & vbCrLf
    c = c & "    chkSelectAll.Value = False" & vbCrLf & vbCrLf
    c = c & "    Dim yStart As Long: yStart = 44" & vbCrLf & vbCrLf
    c = c & "    If m_attachments.Count > 0 Then" & vbCrLf
    c = c & "        Dim lblAttach As MSForms.Label" & vbCrLf
    c = c & "        Set lblAttach = Me.Controls.Add(""Forms.Label.1"", ""lblAttachWarn"")" & vbCrLf
    c = c & "        lblAttach.Left = 8: lblAttach.Top = yStart: lblAttach.Width = 416: lblAttach.Height = 16" & vbCrLf
    c = c & "        lblAttach.Caption = ""This email has file(s) attached and they should be confirmed below""" & vbCrLf
    c = c & "        lblAttach.ForeColor = RGB(255, 0, 0)" & vbCrLf
    c = c & "        lblAttach.WordWrap = True" & vbCrLf
    c = c & "        yStart = yStart + 20" & vbCrLf
    c = c & "    End If" & vbCrLf & vbCrLf
    c = c & "    Dim fra As MSForms.Frame" & vbCrLf
    c = c & "    Set fra = Me.Controls.Add(""Forms.Frame.1"", ""fraItems"")" & vbCrLf
    c = c & "    fra.Left = 4: fra.Top = yStart: fra.Width = 432: fra.Caption = """"" & vbCrLf
    c = c & "    fra.BorderStyle = 0: fra.SpecialEffect = 0" & vbCrLf
    c = c & "    fra.ScrollBars = 2: fra.KeepScrollBarsVisible = 0" & vbCrLf & vbCrLf
    c = c & "    Dim yPos As Long: yPos = 4" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    Dim parts() As String" & vbCrLf
    c = c & "    Dim chk As MSForms.CheckBox" & vbCrLf
    c = c & "    Dim lbl As MSForms.Label" & vbCrLf & vbCrLf
    c = c & "    For i = 1 To m_recipients.Count" & vbCrLf
    c = c & "        parts = Split(m_recipients(i), ""|"")" & vbCrLf
    c = c & "        Set lbl = fra.Controls.Add(""Forms.Label.1"", ""lblType"" & i)" & vbCrLf
    c = c & "        lbl.Left = 4: lbl.Top = yPos + 2: lbl.Width = 28: lbl.Height = 14" & vbCrLf
    c = c & "        lbl.Caption = parts(0) & "":""" & vbCrLf
    c = c & "        Set chk = fra.Controls.Add(""Forms.CheckBox.1"", ""chkRecip"" & i)" & vbCrLf
    c = c & "        chk.Left = 32: chk.Top = yPos: chk.Width = 392: chk.Height = 14" & vbCrLf
    c = c & "        chk.Caption = parts(1): chk.Value = False" & vbCrLf
    c = c & "        m_itemCheckboxes.Add chk" & vbCrLf
    c = c & "        yPos = yPos + 18" & vbCrLf
    c = c & "    Next i" & vbCrLf & vbCrLf
    c = c & "    If m_attachments.Count > 0 Then" & vbCrLf
    c = c & "        yPos = yPos + 4" & vbCrLf
    c = c & "        For i = 1 To m_attachments.Count" & vbCrLf
    c = c & "            parts = Split(m_attachments(i), ""|"")" & vbCrLf
    c = c & "            Set lbl = fra.Controls.Add(""Forms.Label.1"", ""lblFile"" & i)" & vbCrLf
    c = c & "            lbl.Left = 4: lbl.Top = yPos + 2: lbl.Width = 28: lbl.Height = 14" & vbCrLf
    c = c & "            If i = 1 Then lbl.Caption = ""Files:"" Else lbl.Caption = """"" & vbCrLf
    c = c & "            Set chk = fra.Controls.Add(""Forms.CheckBox.1"", ""chkFile"" & i)" & vbCrLf
    c = c & "            chk.Left = 32: chk.Top = yPos: chk.Width = 392: chk.Height = 14" & vbCrLf
    c = c & "            chk.Caption = parts(0) & ""  "" & parts(1): chk.Value = False" & vbCrLf
    c = c & "            m_itemCheckboxes.Add chk" & vbCrLf
    c = c & "            yPos = yPos + 18" & vbCrLf
    c = c & "        Next i" & vbCrLf
    c = c & "    End If" & vbCrLf & vbCrLf
    c = c & "    Dim contentHeight As Long: contentHeight = yPos + 8" & vbCrLf
    c = c & "    Dim maxFrameHeight As Long: maxFrameHeight = 280" & vbCrLf
    c = c & "    If contentHeight <= maxFrameHeight Then" & vbCrLf
    c = c & "        fra.Height = contentHeight: fra.ScrollBars = 0" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        fra.Height = maxFrameHeight: fra.ScrollHeight = contentHeight" & vbCrLf
    c = c & "    End If" & vbCrLf & vbCrLf
    c = c & "    Dim btnY As Long: btnY = fra.Top + fra.Height + 8" & vbCrLf
    c = c & "    btnSend.Left = 290: btnSend.Top = btnY: btnSend.Width = 66: btnSend.Height = 24" & vbCrLf
    c = c & "    btnCancel.Left = 362: btnCancel.Top = btnY: btnCancel.Width = 66: btnCancel.Height = 24" & vbCrLf & vbCrLf
    c = c & "    Me.Caption = ""Confirm External Recipients""" & vbCrLf
    c = c & "    Dim formHeight As Long: formHeight = btnY + btnSend.Height + 12" & vbCrLf
    c = c & "    Me.Width = 450: Me.Height = formHeight + 30" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' -- Event handlers
    c = c & "Private Sub chkSelectAll_Click()" & vbCrLf
    c = c & "    Dim chk As MSForms.CheckBox: Dim i As Long" & vbCrLf
    c = c & "    For i = 1 To m_itemCheckboxes.Count" & vbCrLf
    c = c & "        Set chk = m_itemCheckboxes(i): chk.Value = chkSelectAll.Value" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub btnSend_Click()" & vbCrLf
    c = c & "    Dim chk As MSForms.CheckBox: Dim allChecked As Boolean: allChecked = True: Dim i As Long" & vbCrLf
    c = c & "    For i = 1 To m_itemCheckboxes.Count" & vbCrLf
    c = c & "        Set chk = m_itemCheckboxes(i)" & vbCrLf
    c = c & "        If chk.Value = False Then allChecked = False: Exit For" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    If Not allChecked Then" & vbCrLf
    c = c & "        MsgBox ""Please confirm all recipients and attachments before sending."", vbExclamation, ""Confirmation Required""" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    m_confirmed = True: Me.Hide" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub btnCancel_Click()" & vbCrLf
    c = c & "    m_confirmed = False: Me.Hide" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    c = c & "    If CloseMode = 0 Then m_confirmed = False: Cancel = 1: Me.Hide" & vbCrLf
    c = c & "End Sub" & vbCrLf

    FormCodeBehind = c
End Function
