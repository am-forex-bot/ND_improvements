Attribute VB_Name = "SafeSendReplacement"
'==============================================================================
' SafeSend Replacement - External Recipient Confirmation for Outlook
'==============================================================================
' PURPOSE:  Replaces VIPRE SafeSend with a lightweight VBA alternative.
'           Warns when sending to external recipients and lists attachments.
'           Lets you REMOVE specific recipients or attachments before sending.
'           Only prompts ONCE per email draft.
'
' INSTALL:
'   1) Disable/remove the real SafeSend add-in
'   2) In Outlook: Alt+F11 > File > Import File > select this .bas
'   3) Double-click ThisOutlookSession and paste:
'
'        Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'            Cancel = SafeSendCheck(Item)
'        End Sub
'
'   4) Restart Outlook. Done.
'
' HOW IT WORKS:
'   When you hit Send and there are external recipients, you get a dialog:
'     YES     = Send to everyone listed (no changes)
'     NO      = Opens a second prompt where you type the numbers of
'               recipients/attachments to REMOVE before sending
'     CANCEL  = Don't send the email
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
'  MAIN ENTRY POINT - called from ThisOutlookSession
' ---------------------------------------------------------------------------
Public Function SafeSendCheck(ByVal Item As Object) As Boolean
    On Error GoTo ErrHandler
    SafeSendCheck = False

    If TypeName(Item) <> "MailItem" Then Exit Function

    ' Already confirmed this draft? Let it through.
    If HasBeenConfirmed(Item) Then
        ClearConfirmedFlag Item
        Exit Function
    End If

    ' Find external recipients
    Dim externals As Collection
    Set externals = GetExternalRecipients(Item)
    If externals.Count = 0 Then Exit Function

    ' Get attachment list
    Dim atts As Collection
    Set atts = GetAttachmentList(Item)

    ' Show confirmation dialog
    Dim cancelSend As Boolean
    cancelSend = ShowConfirmation(Item, externals, atts)

    If cancelSend Then
        SafeSendCheck = True
    Else
        StampConfirmedFlag Item
        SafeSendCheck = False
    End If
    Exit Function

ErrHandler:
    SafeSendCheck = False
End Function

' ---------------------------------------------------------------------------
'  CONFIRMATION DIALOG (MsgBox with optional InputBox for removal)
' ---------------------------------------------------------------------------
Private Function ShowConfirmation(ByVal Item As Object, _
        ByVal externals As Collection, ByVal atts As Collection) As Boolean
    ' Returns True = cancel send, False = allow send

    Dim msg As String
    Dim itemNum As Long
    Dim entry As Variant
    Dim parts() As String

    ' --- Build the numbered list ---
    msg = "Subject: " & Item.Subject & vbCrLf
    msg = msg & String(50, Chr(8212)) & vbCrLf & vbCrLf
    msg = msg & "EXTERNAL RECIPIENTS:" & vbCrLf & vbCrLf

    itemNum = 0
    For Each entry In externals
        itemNum = itemNum + 1
        parts = Split(CStr(entry), "|")
        msg = msg & "  " & itemNum & ")  " & parts(0) & ":  " & parts(1) & vbCrLf
    Next entry

    If atts.Count > 0 Then
        msg = msg & vbCrLf & "ATTACHMENTS:" & vbCrLf & vbCrLf
        For Each entry In atts
            itemNum = itemNum + 1
            parts = Split(CStr(entry), "|")
            If Len(parts(1)) > 0 Then
                msg = msg & "  " & itemNum & ")  " & parts(0) & "  (" & parts(1) & ")" & vbCrLf
            Else
                msg = msg & "  " & itemNum & ")  " & parts(0) & vbCrLf
            End If
        Next entry
    End If

    msg = msg & vbCrLf & String(50, Chr(8212)) & vbCrLf
    msg = msg & "YES = Send to all recipients above" & vbCrLf
    msg = msg & "NO = Choose which to remove first" & vbCrLf
    msg = msg & "CANCEL = Don't send"

    Dim result As VbMsgBoxResult
    result = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton2, _
                    "Confirm External Recipients")

    Select Case result
        Case vbYes
            ShowConfirmation = False  ' Send all, no changes

        Case vbCancel
            ShowConfirmation = True   ' Don't send

        Case vbNo
            ' Let user pick items to remove
            ShowConfirmation = HandleRemoval(Item, externals, atts)
    End Select
End Function

' ---------------------------------------------------------------------------
'  REMOVAL DIALOG (InputBox - type numbers to remove)
' ---------------------------------------------------------------------------
Private Function HandleRemoval(ByVal Item As Object, _
        ByVal externals As Collection, ByVal atts As Collection) As Boolean
    ' Returns True = cancel send, False = allow send (with removals applied)

    Dim totalItems As Long
    totalItems = externals.Count + atts.Count

    ' --- Build the prompt ---
    Dim prompt As String
    Dim itemNum As Long
    Dim entry As Variant
    Dim parts() As String

    prompt = "Type the numbers to REMOVE, separated by commas" & vbCrLf
    prompt = prompt & "(e.g. 2,4) - or leave as 0 to send all:" & vbCrLf & vbCrLf

    itemNum = 0
    For Each entry In externals
        itemNum = itemNum + 1
        parts = Split(CStr(entry), "|")
        prompt = prompt & "  " & itemNum & ") " & parts(0) & ": " & parts(1) & vbCrLf
    Next entry

    If atts.Count > 0 Then
        prompt = prompt & vbCrLf
        For Each entry In atts
            itemNum = itemNum + 1
            parts = Split(CStr(entry), "|")
            If Len(parts(1)) > 0 Then
                prompt = prompt & "  " & itemNum & ") " & parts(0) & " (" & parts(1) & ")" & vbCrLf
            Else
                prompt = prompt & "  " & itemNum & ") " & parts(0) & vbCrLf
            End If
        Next entry
    End If

    ' Default "0" = remove nothing; empty string = Cancel clicked
    Dim userInput As String
    userInput = InputBox(prompt, "Remove Items (type numbers)", "0")

    ' Cancel pressed
    If Len(userInput) = 0 Then
        HandleRemoval = True
        Exit Function
    End If

    ' "0" or no valid numbers = send all
    If Trim$(userInput) = "0" Then
        HandleRemoval = False
        Exit Function
    End If

    ' --- Parse the numbers to remove ---
    Dim removeRecip() As Boolean
    Dim removeAtt() As Boolean
    ReDim removeRecip(1 To externals.Count)
    If atts.Count > 0 Then ReDim removeAtt(1 To atts.Count)

    Dim tokens() As String
    tokens = Split(userInput, ",")

    Dim i As Long
    Dim num As Long
    For i = 0 To UBound(tokens)
        num = 0
        On Error Resume Next
        num = CLng(Trim$(tokens(i)))
        On Error GoTo 0

        If num >= 1 And num <= externals.Count Then
            removeRecip(num) = True
        ElseIf num > externals.Count And num <= totalItems Then
            If atts.Count > 0 Then removeAtt(num - externals.Count) = True
        End If
    Next i

    ' --- Build summary of what will be removed ---
    Dim removedList As String
    For i = 1 To externals.Count
        If removeRecip(i) Then
            parts = Split(CStr(externals(i)), "|")
            removedList = removedList & "  REMOVE " & parts(0) & ": " & parts(1) & vbCrLf
        End If
    Next i
    If atts.Count > 0 Then
        For i = 1 To atts.Count
            If removeAtt(i) Then
                parts = Split(CStr(atts(i)), "|")
                removedList = removedList & "  REMOVE attachment: " & parts(0) & vbCrLf
            End If
        Next i
    End If

    ' Nothing valid entered
    If Len(removedList) = 0 Then
        HandleRemoval = False
        Exit Function
    End If

    ' --- Confirm the removal ---
    Dim confirmMsg As String
    confirmMsg = "The following will be REMOVED before sending:" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & removedList & vbCrLf
    confirmMsg = confirmMsg & "Proceed?"

    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox(confirmMsg, vbYesNo + vbQuestion + vbDefaultButton1, _
                           "Confirm Removal")

    If confirmResult <> vbYes Then
        HandleRemoval = True
        Exit Function
    End If

    ' --- Perform the removals (reverse order!) ---

    ' Remove attachments
    If atts.Count > 0 Then
        For i = atts.Count To 1 Step -1
            If removeAtt(i) Then
                parts = Split(CStr(atts(i)), "|")
                Item.Attachments.Remove CLng(parts(2))
            End If
        Next i
    End If

    ' Remove recipients
    For i = externals.Count To 1 Step -1
        If removeRecip(i) Then
            parts = Split(CStr(externals(i)), "|")
            Item.Recipients.Remove CLng(parts(2))
        End If
    Next i

    Item.Recipients.ResolveAll

    ' Safety check: any recipients left?
    If Item.Recipients.Count = 0 Then
        MsgBox "All recipients were removed. Email will not be sent.", _
               vbInformation, "SafeSend"
        HandleRemoval = True
        Exit Function
    End If

    HandleRemoval = False
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
    Dim recipIdx As Long
    Dim emailAddr As String
    Dim recipDomain As String
    Dim isInternal As Boolean
    Dim d As Long
    Dim displayText As String
    Dim typeLabel As String

    recipIdx = 0
    For Each recip In Item.Recipients
        recipIdx = recipIdx + 1
        If Not CHECK_BCC And recip.Type = 3 Then GoTo NextRecip

        emailAddr = GetSmtpAddress(recip)

        If Len(emailAddr) > 0 Then
            recipDomain = LCase$(Mid$(emailAddr, InStr(emailAddr, "@") + 1))

            isInternal = False
            For d = 0 To UBound(domains)
                If recipDomain = domains(d) Then
                    isInternal = True
                    Exit For
                End If
            Next d

            If Not isInternal Then
                If LCase$(recip.Name) <> LCase$(emailAddr) Then
                    displayText = recip.Name & " <" & emailAddr & ">"
                Else
                    displayText = emailAddr
                End If

                Select Case recip.Type
                    Case 1: typeLabel = "TO"
                    Case 2: typeLabel = "CC"
                    Case 3: typeLabel = "BCC"
                    Case Else: typeLabel = "TO"
                End Select

                result.Add typeLabel & "|" & displayText & "|" & CStr(recipIdx)
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
    Dim att As Object
    Dim idx As Long
    Dim sizeText As String
    Dim fileSize As Long

    On Error Resume Next
    idx = 0
    For Each att In Item.Attachments
        idx = idx + 1
        If att.Type = 1 Then
            fileSize = 0
            fileSize = att.Size
            If Err.Number <> 0 Then
                Err.Clear
                sizeText = ""
            Else
                sizeText = FormatFileSize(fileSize)
            End If
            result.Add att.FileName & "|" & sizeText & "|" & CStr(idx)
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
