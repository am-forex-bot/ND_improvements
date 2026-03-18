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
'   2) In Outlook: Alt+F11 -> File -> Import File...
'      -> import SafeSendReplacement.bas
'      -> import frmSafeSend.frm
'   3) Paste the following into ThisOutlookSession (double-click it):
'
'        Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'            Cancel = SafeSendCheck(Item)
'        End Sub
'
'   4) Restart Outlook. Done.
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
