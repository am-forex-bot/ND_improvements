Attribute VB_Name = "SafeSendReplacement"
'==============================================================================
' SafeSend Replacement - External Recipient Confirmation for Outlook
'==============================================================================
' PURPOSE:  Replaces VIPRE SafeSend with a lightweight VBA alternative.
'           Shows a checkbox form where you can UNTICK recipients or
'           attachments you don't want to send. Falls back to simple
'           Yes/No MsgBox if the form isn't installed.
'
' INSTALL:
'   1) File > Options > Trust Center > Trust Center Settings >
'      Macro Settings > tick "Trust access to the VBA project object model"
'   2) Alt+F11 > File > Import File > select this .bas file
'   3) Alt+F8 (or Run menu) > select "InstallSafeSendForm" > Run
'   4) Double-click ThisOutlookSession and paste:
'
'        Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'            Cancel = SafeSendCheck(Item)
'        End Sub
'
'   5) Restart Outlook. Done.
'
' NOTE:  If you skip step 1, the checkbox form won't install but you'll
'        still get simple Yes/No confirmation dialogs (the fallback).
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
'  SHARED DATA (read/written by the form)
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

    ' --- Prepare shared data for the form ---
    g_SSSubject = Item.Subject
    g_SSRecipCount = externals.Count
    ReDim g_SSRecipients(1 To g_SSRecipCount)
    ReDim g_SSRecipIndices(1 To g_SSRecipCount)
    ReDim g_SSRecipChecked(1 To g_SSRecipCount)

    Dim i As Long
    Dim parts() As String
    For i = 1 To externals.Count
        parts = Split(CStr(externals(i)), "|")
        g_SSRecipients(i) = parts(0) & ":  " & parts(1)
        g_SSRecipIndices(i) = CLng(parts(2))
        g_SSRecipChecked(i) = True
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
            g_SSAttChecked(i) = True
        Next i
    End If

    g_SSSendApproved = False

    ' --- Try the checkbox form; fall back to MsgBox ---
    If Not TryShowForm() Then
        SafeSendCheck = FallbackConfirmation(externals, atts, Item.Subject)
        If Not SafeSendCheck Then StampConfirmedFlag Item
        Exit Function
    End If

    ' User cancelled
    If Not g_SSSendApproved Then
        SafeSendCheck = True
        Exit Function
    End If

    ' --- Remove unchecked attachments (reverse order!) ---
    For i = g_SSAttCount To 1 Step -1
        If Not g_SSAttChecked(i) Then
            Item.Attachments.Remove g_SSAttIndices(i)
        End If
    Next i

    ' --- Remove unchecked recipients (reverse order!) ---
    For i = g_SSRecipCount To 1 Step -1
        If Not g_SSRecipChecked(i) Then
            Item.Recipients.Remove g_SSRecipIndices(i)
        End If
    Next i

    Item.Recipients.ResolveAll

    ' Safety: if all recipients were removed, block the send
    If Item.Recipients.Count = 0 Then
        MsgBox "All recipients were removed. Email will not be sent.", _
               vbInformation, "SafeSend"
        SafeSendCheck = True
        Exit Function
    End If

    StampConfirmedFlag Item
    SafeSendCheck = False
    Exit Function

ErrHandler:
    SafeSendCheck = False
End Function

' ---------------------------------------------------------------------------
'  TRY TO SHOW THE CHECKBOX FORM
' ---------------------------------------------------------------------------
Private Function TryShowForm() As Boolean
    On Error GoTo FormFailed
    Application.Run "ShowSafeSendBridge"
    TryShowForm = True
    Exit Function
FormFailed:
    TryShowForm = False
End Function

' ---------------------------------------------------------------------------
'  FALLBACK: simple MsgBox (if form isn't installed)
' ---------------------------------------------------------------------------
Private Function FallbackConfirmation(ByVal externals As Collection, _
        ByVal atts As Collection, ByVal subj As String) As Boolean

    Dim msg As String
    Dim entry As Variant
    Dim parts() As String

    msg = "EXTERNAL RECIPIENTS:" & vbCrLf & vbCrLf
    For Each entry In externals
        parts = Split(CStr(entry), "|")
        msg = msg & "  " & parts(0) & ":  " & parts(1) & vbCrLf
    Next entry

    If atts.Count > 0 Then
        msg = msg & vbCrLf & "ATTACHMENTS:" & vbCrLf & vbCrLf
        For Each entry In atts
            parts = Split(CStr(entry), "|")
            If Len(parts(1)) > 0 Then
                msg = msg & "  " & parts(0) & "  (" & parts(1) & ")" & vbCrLf
            Else
                msg = msg & "  " & parts(0) & vbCrLf
            End If
        Next entry
    End If

    msg = msg & vbCrLf & "Subject: " & subj & vbCrLf
    msg = msg & vbCrLf & "Are you sure you want to send this email?" & vbCrLf
    msg = msg & vbCrLf & "(Tip: run InstallSafeSendForm for the checkbox version)"

    Dim result As VbMsgBoxResult
    result = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, _
                    "Confirm External Recipients")

    FallbackConfirmation = (result <> vbYes)
End Function

' ===========================================================================
'  FORM INSTALLER  -  run once via Alt+F8 > InstallSafeSendForm
' ===========================================================================
Public Sub InstallSafeSendForm()
    On Error GoTo ErrHandler

    Dim proj As Object
    Set proj = Application.VBE.ActiveVBProject

    ' Remove previous install
    RemoveComponentIfExists proj, "frmSafeSend"
    RemoveComponentIfExists proj, "modSafeSendBridge"

    ' --- Create UserForm ---
    Dim frmComp As Object
    Set frmComp = proj.VBComponents.Add(3)   ' vbext_ct_MSForm
    frmComp.Name = "frmSafeSend"

    Dim dsgn As Object
    Set dsgn = frmComp.Designer
    dsgn.Caption = "Confirm External Send"
    dsgn.Width = 440
    dsgn.Height = 400
    dsgn.BackColor = RGB(255, 255, 255)

    Dim ctl As Object

    ' Subject label
    Set ctl = dsgn.Controls.Add("Forms.Label.1", "lblSubject")
    ctl.Left = 12:  ctl.Top = 8:   ctl.Width = 410: ctl.Height = 20
    ctl.Caption = "Subject:": ctl.Font.Size = 10: ctl.Font.Bold = True
    ctl.BackColor = RGB(255, 255, 255)

    ' Recipients header
    Set ctl = dsgn.Controls.Add("Forms.Label.1", "lblRecipHeader")
    ctl.Left = 12:  ctl.Top = 34:  ctl.Width = 410: ctl.Height = 16
    ctl.Caption = "EXTERNAL RECIPIENTS (untick to remove):"
    ctl.ForeColor = RGB(200, 0, 0): ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(255, 255, 255)

    ' Select-all recipients
    Set ctl = dsgn.Controls.Add("Forms.CheckBox.1", "chkAllRecip")
    ctl.Left = 20:  ctl.Top = 54:  ctl.Width = 400: ctl.Height = 18
    ctl.Caption = "Select / Deselect All": ctl.Value = True
    ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(255, 255, 255)

    ' Attachments header
    Set ctl = dsgn.Controls.Add("Forms.Label.1", "lblAttHeader")
    ctl.Left = 12:  ctl.Top = 200: ctl.Width = 410: ctl.Height = 16
    ctl.Caption = "ATTACHMENTS (untick to remove):"
    ctl.ForeColor = RGB(0, 0, 160): ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(255, 255, 255)

    ' Select-all attachments
    Set ctl = dsgn.Controls.Add("Forms.CheckBox.1", "chkAllAtt")
    ctl.Left = 20:  ctl.Top = 220: ctl.Width = 400: ctl.Height = 18
    ctl.Caption = "Select / Deselect All": ctl.Value = True
    ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(255, 255, 255)

    ' Send button
    Set ctl = dsgn.Controls.Add("Forms.CommandButton.1", "btnSend")
    ctl.Left = 240: ctl.Top = 360: ctl.Width = 88: ctl.Height = 28
    ctl.Caption = "Send": ctl.Font.Size = 9

    ' Don't Send button
    Set ctl = dsgn.Controls.Add("Forms.CommandButton.1", "btnCancel")
    ctl.Left = 336: ctl.Top = 360: ctl.Width = 88: ctl.Height = 28
    ctl.Caption = "Don't Send": ctl.Font.Size = 9

    ' Inject form code
    With frmComp.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString GetFormCode()
    End With

    ' --- Create bridge module ---
    Dim bridgeComp As Object
    Set bridgeComp = proj.VBComponents.Add(1)   ' vbext_ct_StdModule
    bridgeComp.Name = "modSafeSendBridge"
    With bridgeComp.CodeModule
        .DeleteLines 1, .CountOfLines
        .AddFromString GetBridgeCode()
    End With

    MsgBox "SafeSend form installed successfully!" & vbCrLf & vbCrLf & _
           "Please restart Outlook for changes to take effect.", _
           vbInformation, "SafeSend Install"
    Exit Sub

ErrHandler:
    If InStr(1, Err.Description, "programmatic access", vbTextCompare) > 0 _
       Or InStr(1, Err.Description, "not trusted", vbTextCompare) > 0 Then
        MsgBox "Enable VBA project access first:" & vbCrLf & vbCrLf & _
               "File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > tick 'Trust access to the VBA project object model'" & vbCrLf & vbCrLf & _
               "Then run this macro again.", vbExclamation, "SafeSend Install"
    Else
        MsgBox "Install error (" & Err.Number & "): " & Err.Description, _
               vbCritical, "SafeSend Install"
    End If
End Sub

Private Sub RemoveComponentIfExists(proj As Object, compName As String)
    On Error Resume Next
    Dim comp As Object
    Set comp = proj.VBComponents(compName)
    If Not comp Is Nothing Then proj.VBComponents.Remove comp
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
'  FORM CODE (injected into frmSafeSend by the installer)
' ---------------------------------------------------------------------------
Private Function GetFormCode() As String
    Dim L As String: L = vbCrLf
    Dim c As String

    c = "Option Explicit" & L & L

    ' --- UserForm_Initialize ---
    c = c & "Private Sub UserForm_Initialize()" & L
    c = c & "    Dim yPos As Single" & L
    c = c & "    Dim i As Long" & L
    c = c & "    Dim ctl As Object" & L & L
    c = c & "    Me.lblSubject.Caption = ""Subject: "" & SafeSendReplacement.g_SSSubject" & L & L
    c = c & "    yPos = Me.chkAllRecip.Top + Me.chkAllRecip.Height + 4" & L & L
    c = c & "    For i = 1 To SafeSendReplacement.g_SSRecipCount" & L
    c = c & "        Set ctl = Me.Controls.Add(""Forms.CheckBox.1"", ""chkRecip"" & i)" & L
    c = c & "        ctl.Left = 28" & L
    c = c & "        ctl.Top = yPos" & L
    c = c & "        ctl.Width = 396" & L
    c = c & "        ctl.Height = 18" & L
    c = c & "        ctl.Caption = SafeSendReplacement.g_SSRecipients(i)" & L
    c = c & "        ctl.Value = True" & L
    c = c & "        ctl.BackColor = Me.BackColor" & L
    c = c & "        yPos = yPos + 20" & L
    c = c & "    Next i" & L & L
    c = c & "    yPos = yPos + 12" & L & L
    c = c & "    If SafeSendReplacement.g_SSAttCount > 0 Then" & L
    c = c & "        Me.lblAttHeader.Visible = True" & L
    c = c & "        Me.lblAttHeader.Top = yPos" & L
    c = c & "        yPos = yPos + 18" & L
    c = c & "        Me.chkAllAtt.Visible = True" & L
    c = c & "        Me.chkAllAtt.Top = yPos" & L
    c = c & "        yPos = yPos + 22" & L & L
    c = c & "        For i = 1 To SafeSendReplacement.g_SSAttCount" & L
    c = c & "            Set ctl = Me.Controls.Add(""Forms.CheckBox.1"", ""chkAtt"" & i)" & L
    c = c & "            ctl.Left = 28" & L
    c = c & "            ctl.Top = yPos" & L
    c = c & "            ctl.Width = 396" & L
    c = c & "            ctl.Height = 18" & L
    c = c & "            ctl.Caption = SafeSendReplacement.g_SSAttachments(i)" & L
    c = c & "            ctl.Value = True" & L
    c = c & "            ctl.BackColor = Me.BackColor" & L
    c = c & "            yPos = yPos + 20" & L
    c = c & "        Next i" & L
    c = c & "    Else" & L
    c = c & "        Me.lblAttHeader.Visible = False" & L
    c = c & "        Me.chkAllAtt.Visible = False" & L
    c = c & "    End If" & L & L
    c = c & "    yPos = yPos + 16" & L
    c = c & "    Me.btnSend.Top = yPos" & L
    c = c & "    Me.btnCancel.Top = yPos" & L & L
    c = c & "    Dim totalH As Single" & L
    c = c & "    totalH = yPos + Me.btnSend.Height + 50" & L
    c = c & "    If totalH < 220 Then totalH = 220" & L
    c = c & "    If totalH > 550 Then" & L
    c = c & "        Me.ScrollBars = 2" & L
    c = c & "        Me.ScrollHeight = totalH" & L
    c = c & "        totalH = 550" & L
    c = c & "    End If" & L
    c = c & "    Me.Height = totalH" & L
    c = c & "End Sub" & L & L

    ' --- btnSend_Click ---
    c = c & "Private Sub btnSend_Click()" & L
    c = c & "    Dim i As Long" & L
    c = c & "    For i = 1 To SafeSendReplacement.g_SSRecipCount" & L
    c = c & "        SafeSendReplacement.g_SSRecipChecked(i) = CBool(Me.Controls(""chkRecip"" & i).Value)" & L
    c = c & "    Next i" & L
    c = c & "    For i = 1 To SafeSendReplacement.g_SSAttCount" & L
    c = c & "        SafeSendReplacement.g_SSAttChecked(i) = CBool(Me.Controls(""chkAtt"" & i).Value)" & L
    c = c & "    Next i" & L
    c = c & "    SafeSendReplacement.g_SSSendApproved = True" & L
    c = c & "    Unload Me" & L
    c = c & "End Sub" & L & L

    ' --- btnCancel_Click ---
    c = c & "Private Sub btnCancel_Click()" & L
    c = c & "    SafeSendReplacement.g_SSSendApproved = False" & L
    c = c & "    Unload Me" & L
    c = c & "End Sub" & L & L

    ' --- UserForm_QueryClose (X button = cancel) ---
    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & L
    c = c & "    If CloseMode = 0 Then" & L
    c = c & "        SafeSendReplacement.g_SSSendApproved = False" & L
    c = c & "    End If" & L
    c = c & "End Sub" & L & L

    ' --- Select All Recipients ---
    c = c & "Private Sub chkAllRecip_Click()" & L
    c = c & "    Dim i As Long" & L
    c = c & "    For i = 1 To SafeSendReplacement.g_SSRecipCount" & L
    c = c & "        Me.Controls(""chkRecip"" & i).Value = Me.chkAllRecip.Value" & L
    c = c & "    Next i" & L
    c = c & "End Sub" & L & L

    ' --- Select All Attachments ---
    c = c & "Private Sub chkAllAtt_Click()" & L
    c = c & "    If SafeSendReplacement.g_SSAttCount = 0 Then Exit Sub" & L
    c = c & "    Dim i As Long" & L
    c = c & "    For i = 1 To SafeSendReplacement.g_SSAttCount" & L
    c = c & "        Me.Controls(""chkAtt"" & i).Value = Me.chkAllAtt.Value" & L
    c = c & "    Next i" & L
    c = c & "End Sub" & L

    GetFormCode = c
End Function

' ---------------------------------------------------------------------------
'  BRIDGE MODULE CODE (lets us call the form without compile-time reference)
' ---------------------------------------------------------------------------
Private Function GetBridgeCode() As String
    Dim L As String: L = vbCrLf
    Dim c As String
    c = "Option Explicit" & L & L
    c = c & "Public Sub ShowSafeSendBridge()" & L
    c = c & "    frmSafeSend.Show vbModal" & L
    c = c & "End Sub" & L
    GetBridgeCode = c
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
