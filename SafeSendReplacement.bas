Attribute VB_Name = "SafeSendReplacement"
'==============================================================================
' SafeSend Replacement - External Recipient Confirmation for Outlook
'==============================================================================
' PURPOSE:  Replaces VIPRE SafeSend with a lightweight VBA alternative.
'           Warns when sending to external recipients and lists attachments.
'           Only prompts ONCE per email - if another add-in (e.g. ndMail)
'           cancels the send after confirmation, you won't be asked again.
'
' INSTALL:
'   1) Disable/remove the real SafeSend add-in
'   2) In Outlook: Alt+F11 -> File -> Import File -> select this .bas
'   3) Double-click ThisOutlookSession and paste:
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
Private Const INTERNAL_DOMAINS As String = "wallace.co.uk,wallace.onmicrosoft.com"
Private Const CHECK_BCC As Boolean = True
Private Const CONFIRMED_FLAG As String = "_SafeSendConfirmed"

' ---------------------------------------------------------------------------
'  MAIN ENTRY POINT
' ---------------------------------------------------------------------------
Public Function SafeSendCheck(ByVal Item As Object) As Boolean
    On Error GoTo ErrHandler
    SafeSendCheck = False

    If TypeName(Item) <> "MailItem" Then Exit Function

    ' Already confirmed this draft? Skip.
    If HasBeenConfirmed(Item) Then
        ClearConfirmedFlag Item
        Exit Function
    End If

    ' Find external recipients
    Dim externals As Collection
    Set externals = GetExternalRecipients(Item)
    If externals.Count = 0 Then Exit Function

    ' Build the warning message
    Dim msg As String
    msg = "EXTERNAL RECIPIENTS:" & vbCrLf & vbCrLf

    Dim entry As Variant
    For Each entry In externals
        Dim parts() As String
        parts = Split(CStr(entry), "|")
        msg = msg & "  " & parts(0) & ":  " & parts(1) & vbCrLf
    Next entry

    ' List attachments if any
    Dim attachments As Collection
    Set attachments = GetAttachmentList(Item)

    If attachments.Count > 0 Then
        msg = msg & vbCrLf & "ATTACHMENTS:" & vbCrLf & vbCrLf
        Dim att As Variant
        For Each att In attachments
            parts = Split(CStr(att), "|")
            If Len(parts(1)) > 0 Then
                msg = msg & "  " & parts(0) & "  (" & parts(1) & ")" & vbCrLf
            Else
                msg = msg & "  " & parts(0) & vbCrLf
            End If
        Next att
    End If

    msg = msg & vbCrLf & "Subject: " & Item.Subject & vbCrLf
    msg = msg & vbCrLf & "Are you sure you want to send this email?"

    Dim result As VbMsgBoxResult
    result = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, _
                    "Confirm External Recipients")

    If result = vbYes Then
        StampConfirmedFlag Item
        SafeSendCheck = False
    Else
        SafeSendCheck = True
    End If

    Exit Function
ErrHandler:
    SafeSendCheck = False
End Function

' ---------------------------------------------------------------------------
'  EXTERNAL RECIPIENT DETECTION
' ---------------------------------------------------------------------------
Private Function GetExternalRecipients(ByVal Item As Object) As Collection
    Dim result As New Collection
    Dim domains() As String
    domains = Split(LCase$(INTERNAL_DOMAINS), ",")

    Dim i As Long
    For i = 0 To UBound(domains)
        domains(i) = Trim$(domains(i))
    Next i

    Dim recip As Object
    For Each recip In Item.Recipients
        If Not CHECK_BCC And recip.Type = 3 Then GoTo NextRecip

        Dim emailAddr As String
        emailAddr = GetSmtpAddress(recip)

        If Len(emailAddr) > 0 Then
            Dim recipDomain As String
            recipDomain = LCase$(Mid$(emailAddr, InStr(emailAddr, "@") + 1))

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

                Dim typeLabel As String
                Select Case recip.Type
                    Case 1: typeLabel = "TO"
                    Case 2: typeLabel = "CC"
                    Case 3: typeLabel = "BCC"
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
Private Function GetAttachmentList(ByVal Item As Object) As Collection
    Dim result As New Collection
    On Error Resume Next
    Dim att As Object
    For Each att In Item.Attachments
        If att.Type = 1 Then
            Dim sizeText As String
            Dim fileSize As Long: fileSize = 0
            fileSize = att.Size
            If Err.Number <> 0 Then
                Err.Clear: sizeText = ""
            Else
                sizeText = FormatFileSize(fileSize)
            End If
            result.Add att.FileName & "|" & sizeText
        End If
    Next att
    On Error GoTo 0
    Set GetAttachmentList = result
End Function

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
Private Function GetSmtpAddress(recip As Object) As String
    On Error GoTo ErrHandler

    Const PR_SMTP As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

    If Not recip.Resolve Then
        GetSmtpAddress = recip.Address
        Exit Function
    End If

    Dim addrEntry As Object
    Set addrEntry = recip.AddressEntry

    If addrEntry.Type = "SMTP" Then
        GetSmtpAddress = LCase$(addrEntry.Address)
    ElseIf addrEntry.Type = "EX" Then
        Dim pa As Object
        Set pa = recip.PropertyAccessor
        GetSmtpAddress = LCase$(pa.GetProperty(PR_SMTP))
    Else
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
    On Error Resume Next
    GetSmtpAddress = LCase$(recip.Address)
End Function

' ---------------------------------------------------------------------------
'  CONFIRMED FLAG (prevents double-prompting)
' ---------------------------------------------------------------------------
Private Function HasBeenConfirmed(ByVal Item As Object) As Boolean
    On Error Resume Next
    Dim prop As Object
    Set prop = Item.UserProperties.Find(CONFIRMED_FLAG)
    HasBeenConfirmed = (Not prop Is Nothing)
    If HasBeenConfirmed Then HasBeenConfirmed = (prop.Value = True)
    On Error GoTo 0
End Function

Private Sub StampConfirmedFlag(ByVal Item As Object)
    On Error Resume Next
    Dim prop As Object
    Set prop = Item.UserProperties.Add(CONFIRMED_FLAG, 6, False)
    prop.Value = True
    Item.Save
    On Error GoTo 0
End Sub

Private Sub ClearConfirmedFlag(ByVal Item As Object)
    On Error Resume Next
    Dim prop As Object
    Set prop = Item.UserProperties.Find(CONFIRMED_FLAG)
    If Not prop Is Nothing Then prop.Delete
    On Error GoTo 0
End Sub
