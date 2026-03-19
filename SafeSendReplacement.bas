'==============================================================================
' SafeSend Replacement - External Recipient Confirmation for Outlook
'==============================================================================
' PURPOSE:  Replaces VIPRE SafeSend with a lightweight VBA alternative.
'           Shows a checkbox form where you can UNTICK recipients or
'           attachments to remove them before sending.
'           Falls back to MsgBox+InputBox if the form isn't set up yet.
'
' INSTALL:
'   1) Alt+F11 > File > Import File > select this .bas file
'   2) Create the UserForm (see instructions in frmSafeSend_code.txt)
'   3) Double-click ThisOutlookSession and paste:
'
'        Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'            Cancel = SafeSendCheck(Item)
'        End Sub
'
'   4) Restart Outlook. Done.
'
' NOTE:  If you skip step 2, it still works - you just get a simpler
'        Yes/No/Cancel dialog instead of checkboxes.
'
' CONFIG: Edit the INTERNAL_DOMAINS constant below.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
'  CONFIGURATION
' ---------------------------------------------------------------------------
Private Const INTERNAL_DOMAINS As String = "wallace.co.uk,wallace.onmicrosoft.com"
Private Const CHECK_BCC As Boolean = True
Private Const CONFIRMED_FLAG As String = "_SafeSendConfirmed"

' ---------------------------------------------------------------------------
'  SHARED DATA (read/written by frmSafeSend)
' ---------------------------------------------------------------------------
Public g_SSSubject As String
Public g_SSRecipCount As Long
Public g_SSRecipients() As String
Public g_SSRecipIndices() As Long
Public g_SSRecipChecked() As Boolean
Public g_SSAttCount As Long
Public g_SSAttachments() As String
Public g_SSAttIndices() As Long
Public g_SSAttChecked() As Boolean
Public g_SSSendApproved As Boolean

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

    ' --- Try checkbox form first ---
    Dim formResult As Long   ' 0=not available, 1=send, 2=cancel
    formResult = TryFormApproach(Item, externals, atts)

    If formResult = 1 Then
        StampConfirmedFlag Item
        SafeSendCheck = False
        Exit Function
    ElseIf formResult = 2 Then
        SafeSendCheck = True
        Exit Function
    End If

    ' --- Fallback: MsgBox + InputBox ---
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
'  CHECKBOX FORM APPROACH
' ---------------------------------------------------------------------------
Private Function TryFormApproach(ByVal Item As Object, _
        ByVal externals As Collection, ByVal atts As Collection) As Long
    ' Returns: 0=form not available, 1=send approved, 2=cancelled
    On Error GoTo NoForm

    ' Populate shared data for the form to read
    SetupSharedData externals, atts, Item.Subject

    ' Show form - this triggers UserForm_Initialize which reads shared data
    Dim frm As frmSafeSend
    Set frm = New frmSafeSend
    frm.Show vbModal
    Set frm = Nothing

    ' User cancelled?
    If Not g_SSSendApproved Then
        TryFormApproach = 2
        Exit Function
    End If

    ' Remove unchecked attachments (reverse order!)
    Dim i As Long
    For i = g_SSAttCount To 1 Step -1
        If Not g_SSAttChecked(i) Then
            Item.Attachments.Remove g_SSAttIndices(i)
        End If
    Next i

    ' Remove unchecked recipients (reverse order!)
    For i = g_SSRecipCount To 1 Step -1
        If Not g_SSRecipChecked(i) Then
            Item.Recipients.Remove g_SSRecipIndices(i)
        End If
    Next i

    Item.Recipients.ResolveAll

    ' Safety: block send if no recipients remain
    If Item.Recipients.Count = 0 Then
        MsgBox "All recipients were removed. Email will not be sent.", _
               vbInformation, "SafeSend"
        TryFormApproach = 2
        Exit Function
    End If

    TryFormApproach = 1
    Exit Function

NoForm:
    TryFormApproach = 0
End Function

Private Sub SetupSharedData(ByVal externals As Collection, _
        ByVal atts As Collection, ByVal subj As String)

    Dim i As Long
    Dim parts() As String

    g_SSSubject = subj
    g_SSSendApproved = False

    g_SSRecipCount = externals.Count
    ReDim g_SSRecipients(1 To g_SSRecipCount)
    ReDim g_SSRecipIndices(1 To g_SSRecipCount)
    ReDim g_SSRecipChecked(1 To g_SSRecipCount)

    For i = 1 To externals.Count
        parts = Split(CStr(externals(i)), "|")
        g_SSRecipients(i) = parts(0) & ":  " & parts(1)
        g_SSRecipIndices(i) = CLng(parts(2))
        g_SSRecipChecked(i) = False
    Next i

    g_SSAttCount = atts.Count
    If g_SSAttCount > 0 Then
        ReDim g_SSAttachments(1 To g_SSAttCount)
        ReDim g_SSAttIndices(1 To g_SSAttCount)
        ReDim g_SSAttChecked(1 To g_SSAttCount)
        For i = 1 To atts.Count
            parts = Split(CStr(atts(i)), "|")
            If Len(parts(1)) > 0 Then
                g_SSAttachments(i) = parts(0) & "  (" & parts(1) & ")"
            Else
                g_SSAttachments(i) = parts(0)
            End If
            g_SSAttIndices(i) = CLng(parts(2))
            g_SSAttChecked(i) = False
        Next i
    End If
End Sub

' ---------------------------------------------------------------------------
'  FALLBACK: MsgBox + InputBox (if form not set up)
' ---------------------------------------------------------------------------
Private Function ShowConfirmation(ByVal Item As Object, _
        ByVal externals As Collection, ByVal atts As Collection) As Boolean

    Dim msg As String
    Dim itemNum As Long
    Dim entry As Variant
    Dim parts() As String

    msg = "Subject: " & Item.Subject & vbCrLf
    msg = msg & String(50, "-") & vbCrLf & vbCrLf
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

    msg = msg & vbCrLf & String(50, "-") & vbCrLf
    msg = msg & "YES = Send to all recipients above" & vbCrLf
    msg = msg & "NO = Choose which to remove first" & vbCrLf
    msg = msg & "CANCEL = Don't send"

    Dim result As VbMsgBoxResult
    result = MsgBox(msg, vbYesNoCancel + vbExclamation + vbDefaultButton2, _
                    "Confirm External Recipients")

    Select Case result
        Case vbYes
            ShowConfirmation = False
        Case vbCancel
            ShowConfirmation = True
        Case vbNo
            ShowConfirmation = HandleRemoval(Item, externals, atts)
    End Select
End Function

Private Function HandleRemoval(ByVal Item As Object, _
        ByVal externals As Collection, ByVal atts As Collection) As Boolean

    Dim totalItems As Long
    totalItems = externals.Count + atts.Count

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

    Dim userInput As String
    userInput = InputBox(prompt, "Remove Items (type numbers)", "0")

    If Len(userInput) = 0 Then
        HandleRemoval = True
        Exit Function
    End If

    If Trim$(userInput) = "0" Then
        HandleRemoval = False
        Exit Function
    End If

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

    If Len(removedList) = 0 Then
        HandleRemoval = False
        Exit Function
    End If

    Dim confirmMsg As String
    confirmMsg = "The following will be REMOVED before sending:" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & removedList & vbCrLf
    confirmMsg = confirmMsg & "Proceed?"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirm Removal") <> vbYes Then
        HandleRemoval = True
        Exit Function
    End If

    If atts.Count > 0 Then
        For i = atts.Count To 1 Step -1
            If removeAtt(i) Then
                parts = Split(CStr(atts(i)), "|")
                Item.Attachments.Remove CLng(parts(2))
            End If
        Next i
    End If

    For i = externals.Count To 1 Step -1
        If removeRecip(i) Then
            parts = Split(CStr(externals(i)), "|")
            Item.Recipients.Remove CLng(parts(2))
        End If
    Next i

    Item.Recipients.ResolveAll

    If Item.Recipients.Count = 0 Then
        MsgBox "All recipients removed. Email not sent.", vbInformation, "SafeSend"
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
