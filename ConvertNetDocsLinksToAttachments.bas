Attribute VB_Name = "NetDocsLinkConverter"
'==============================================================================
' NetDocuments Link-to-Attachment Converter for Outlook
'==============================================================================
' Scans the currently open email for NetDocuments URLs (hyperlinks and raw),
' creates .url shortcut files, attaches them, and cleans up the email body.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' ENTRY POINT – run from an open Inspector (compose/read) window
' ---------------------------------------------------------------------------
Public Sub ConvertNetDocsLinksToAttachments()

    ' --- Validate we have an open mail item --------------------------------
    Dim oInspector As Outlook.Inspector
    Set oInspector = Application.ActiveInspector

    If oInspector Is Nothing Then
        MsgBox "Please open an email first.", vbExclamation
        Exit Sub
    End If

    If Not TypeOf oInspector.CurrentItem Is Outlook.MailItem Then
        MsgBox "This macro only works on email messages.", vbExclamation
        Exit Sub
    End If

    Dim oMail As Outlook.MailItem
    Set oMail = oInspector.CurrentItem

    ' --- Collect links from HTML body --------------------------------------
    Dim htmlBody As String
    htmlBody = oMail.HTMLBody

    Dim links As Collection  ' Each item is an array: Array(url, displayName)
    Set links = ExtractNetDocsLinks(htmlBody)

    If links.Count = 0 Then
        MsgBox "No NetDocuments links found in this email.", vbInformation
        Exit Sub
    End If

    ' --- De-duplicate by URL -----------------------------------------------
    Dim uniqueLinks As Collection
    Set uniqueLinks = DeduplicateLinks(links)

    ' --- Check existing attachments to avoid duplicates on re-run ----------
    Dim existingNames As Collection
    Set existingNames = GetExistingAttachmentNames(oMail)

    ' --- Create .url files and attach --------------------------------------
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Right$(tempDir, 1) <> "\" Then tempDir = tempDir & "\"

    Dim createdFiles As New Collection  ' track paths for cleanup
    Dim i As Long
    Dim linkData As Variant
    Dim url As String
    Dim displayName As String
    Dim fileName As String
    Dim filePath As String
    Dim attachCount As Long

    For i = 1 To uniqueLinks.Count
        linkData = uniqueLinks(i)
        url = CStr(linkData(0))
        displayName = CStr(linkData(1))

        ' Build a safe .url filename
        fileName = BuildUrlFilename(displayName, url)

        ' Skip if already attached (re-run protection)
        If Not CollectionContains(existingNames, LCase$(fileName)) Then
            filePath = tempDir & fileName

            ' Write the .url shortcut
            WriteUrlFile filePath, url
            createdFiles.Add filePath

            ' Attach to the mail item
            oMail.Attachments.Add filePath, olByValue
            attachCount = attachCount + 1
        End If
    Next i

    ' --- Clean the HTML body – remove ND links, insert summary line --------
    Dim cleanedHtml As String
    cleanedHtml = RemoveNetDocsLinksFromHtml(htmlBody)

    ' Insert summary notice (before </body> or at end)
    cleanedHtml = InsertSummaryLine(cleanedHtml, uniqueLinks.Count)

    oMail.HTMLBody = cleanedHtml

    ' --- Delete temp files -------------------------------------------------
    Dim f As Variant
    For Each f In createdFiles
        On Error Resume Next
        Kill CStr(f)
        On Error GoTo 0
    Next f

    ' --- Done --------------------------------------------------------------
    MsgBox "NetDocuments links converted to attachments", vbInformation

End Sub

' ===========================================================================
' LINK EXTRACTION
' ===========================================================================

' Returns a Collection of Array(url, displayName) for every ND link found.
Private Function ExtractNetDocsLinks(htmlBody As String) As Collection
    Dim results As New Collection

    ' --- Pass 1: extract <a href="...netdocuments.com...">text</a> ---------
    Dim regexAnchor As Object
    Set regexAnchor = CreateObject("VBScript.RegExp")
    With regexAnchor
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        ' Capture href value and inner text
        .Pattern = "<a\b[^>]*\bhref\s*=\s*[""']([^""']*netdocuments\.com[^""']*)[""'][^>]*>([\s\S]*?)<\/a>"
    End With

    Dim matches As Object
    Dim m As Object
    Dim anchorUrl As String
    Dim anchorText As String

    If regexAnchor.Test(htmlBody) Then
        Set matches = regexAnchor.Execute(htmlBody)
        For Each m In matches
            anchorUrl = Trim$(m.SubMatches(0))
            anchorText = StripHtmlTags(Trim$(m.SubMatches(1)))
            If Len(anchorUrl) > 0 Then
                results.Add Array(anchorUrl, anchorText)
            End If
        Next m
    End If

    ' --- Pass 2: find raw URLs NOT already inside an <a> tag ---------------
    '     We look for URLs that contain netdocuments.com
    Dim regexRaw As Object
    Set regexRaw = CreateObject("VBScript.RegExp")
    With regexRaw
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "(https?://[^\s""'<>]*netdocuments\.com[^\s""'<>]*)"
    End With

    ' Build set of URLs already captured from anchors
    Dim anchorUrls As New Collection
    Dim idx As Long
    For idx = 1 To results.Count
        On Error Resume Next
        anchorUrls.Add "", CStr(results(idx)(0))
        On Error GoTo 0
    Next idx

    ' We need to check raw URLs are not inside href="..." (already captured)
    ' Strategy: strip all <a...>...</a> tags, then search remainder
    Dim bodyNoAnchors As String
    bodyNoAnchors = regexAnchor.Replace(htmlBody, " ")

    ' Also strip any remaining href="..." attributes to avoid double-capture
    Dim regexHref As Object
    Set regexHref = CreateObject("VBScript.RegExp")
    With regexHref
        .Global = True
        .IgnoreCase = True
        .Pattern = "href\s*=\s*[""'][^""']*[""']"
    End With
    bodyNoAnchors = regexHref.Replace(bodyNoAnchors, " ")

    If regexRaw.Test(bodyNoAnchors) Then
        Set matches = regexRaw.Execute(bodyNoAnchors)
        For Each m In matches
            Dim rawUrl As String
            rawUrl = Trim$(m.SubMatches(0))
            ' Skip if we already have this URL from anchor pass
            Dim alreadyHave As Boolean
            alreadyHave = False
            On Error Resume Next
            Dim dummy As String
            dummy = anchorUrls(rawUrl)
            If Err.Number = 0 Then alreadyHave = True
            Err.Clear
            On Error GoTo 0

            If Not alreadyHave And Len(rawUrl) > 0 Then
                results.Add Array(rawUrl, "")
                On Error Resume Next
                anchorUrls.Add "", rawUrl
                On Error GoTo 0
            End If
        Next m
    End If

    Set ExtractNetDocsLinks = results
End Function

' ===========================================================================
' FILENAME BUILDING
' ===========================================================================

' Determines the best filename for the .url shortcut.
' Priority: 1) hyperlink display text  2) filename from URL  3) doc ID
Private Function BuildUrlFilename(displayName As String, url As String) As String
    Dim baseName As String

    ' --- Priority 1: display text from hyperlink ---------------------------
    If Len(Trim$(displayName)) > 0 Then
        baseName = Trim$(displayName)
        ' If display text already has a file extension, strip it (we add .url)
        ' But keep it in the name for clarity – e.g. "Contract_v5.docx.url"
        GoTo Sanitize
    End If

    ' --- Priority 2: try to extract a filename from the URL path -----------
    baseName = ExtractFilenameFromUrl(url)
    If Len(baseName) > 0 Then GoTo Sanitize

    ' --- Priority 3: extract document ID from URL --------------------------
    baseName = ExtractDocIdFromUrl(url)
    If Len(baseName) > 0 Then
        baseName = "NetDocs_" & baseName
        GoTo Sanitize
    End If

    ' --- Fallback (should rarely happen) -----------------------------------
    baseName = "NetDocs_Document"

Sanitize:
    ' Make filename Windows-safe
    baseName = SanitizeFilename(baseName)

    ' Ensure it ends with .url
    If LCase$(Right$(baseName, 4)) <> ".url" Then
        baseName = baseName & ".url"
    End If

    BuildUrlFilename = baseName
End Function

' Attempts to pull a filename segment from a NetDocuments URL.
Private Function ExtractFilenameFromUrl(url As String) As String
    ' NetDocuments URLs often contain paths like /document/... or query params
    ' with filenames. Try a few common patterns.

    Dim regexFn As Object
    Set regexFn = CreateObject("VBScript.RegExp")

    ' Pattern: look for something that looks like a filename with extension
    ' in the URL path or query string
    With regexFn
        .Global = False
        .IgnoreCase = True
        .Pattern = "[\/?&=]([A-Za-z0-9_\-\. ]+\.(docx?|xlsx?|pptx?|pdf|txt|csv|rtf|msg))"
    End With

    If regexFn.Test(url) Then
        Dim m As Object
        Set m = regexFn.Execute(url)
        ExtractFilenameFromUrl = m(0).SubMatches(0)
    Else
        ExtractFilenameFromUrl = ""
    End If
End Function

' Extracts the NetDocuments document ID from the URL.
Private Function ExtractDocIdFromUrl(url As String) As String
    Dim regexId As Object
    Set regexId = CreateObject("VBScript.RegExp")

    ' Common ND URL patterns:
    '   /nddocview/.../{docId}
    '   /document/{docId}
    '   ndDocId=xxxx
    '   /d/{docId}
    '   /{cabId}/{docId}/v{version}

    ' Try multiple patterns in order of specificity
    Dim patterns As Variant
    patterns = Array( _
        "[?&]ndDocId=([A-Za-z0-9\-]+)", _
        "/nddocview[^/]*/([A-Za-z0-9\-]+)", _
        "/document/([A-Za-z0-9\-]+)", _
        "/d/([A-Za-z0-9\-]+)", _
        "/([0-9]{4,}[\-/][0-9]+)" _
    )

    Dim p As Variant
    For Each p In patterns
        With regexId
            .Global = False
            .IgnoreCase = True
            .Pattern = CStr(p)
        End With
        If regexId.Test(url) Then
            Dim mx As Object
            Set mx = regexId.Execute(url)
            ExtractDocIdFromUrl = mx(0).SubMatches(0)
            Exit Function
        End If
    Next p

    ' Last resort: grab the last path segment that looks like an ID
    Dim regexLast As Object
    Set regexLast = CreateObject("VBScript.RegExp")
    With regexLast
        .Global = True
        .IgnoreCase = True
        .Pattern = "/([A-Za-z0-9\-]{6,})"
    End With

    If regexLast.Test(url) Then
        Dim allMatches As Object
        Set allMatches = regexLast.Execute(url)
        ' Use the last significant path segment
        ExtractDocIdFromUrl = allMatches(allMatches.Count - 1).SubMatches(0)
    Else
        ExtractDocIdFromUrl = ""
    End If
End Function

' ===========================================================================
' HTML CLEANUP
' ===========================================================================

' Removes all NetDocuments links (anchors and raw URLs) from the HTML body.
Private Function RemoveNetDocsLinksFromHtml(htmlBody As String) As String
    Dim result As String
    result = htmlBody

    ' --- Remove <a> tags pointing to netdocuments.com ----------------------
    Dim regexA As Object
    Set regexA = CreateObject("VBScript.RegExp")
    With regexA
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "<a\b[^>]*\bhref\s*=\s*[""'][^""']*netdocuments\.com[^""']*[""'][^>]*>[\s\S]*?<\/a>"
    End With
    result = regexA.Replace(result, "")

    ' --- Remove raw netdocuments.com URLs ----------------------------------
    Dim regexRaw As Object
    Set regexRaw = CreateObject("VBScript.RegExp")
    With regexRaw
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "https?://[^\s""'<>]*netdocuments\.com[^\s""'<>]*"
    End With
    result = regexRaw.Replace(result, "")

    ' --- Clean up leftover empty paragraphs/divs/list items ----------------
    '     (only truly empty ones – no visible content)
    Dim regexEmpty As Object
    Set regexEmpty = CreateObject("VBScript.RegExp")
    With regexEmpty
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        ' Remove <p>, <div>, <li> that contain only whitespace/&nbsp;/line breaks
        .Pattern = "<(p|div|li)\b[^>]*>\s*(&nbsp;|\s|<br\s*/?>)*\s*<\/\1>"
    End With
    ' Run twice to catch nested empties
    result = regexEmpty.Replace(result, "")
    result = regexEmpty.Replace(result, "")

    RemoveNetDocsLinksFromHtml = result
End Function

' Inserts a summary line into the HTML body.
Private Function InsertSummaryLine(htmlBody As String, linkCount As Long) As String
    Dim summaryHtml As String
    summaryHtml = "<p style=""color:#336699;font-weight:bold;margin:10px 0;"">" & _
                  "Documents attached via NetDocuments (" & linkCount & " file" & _
                  IIf(linkCount <> 1, "s", "") & ")</p>"

    ' Try to insert before </body>
    Dim posBody As Long
    posBody = InStrRev(LCase$(htmlBody), "</body>")

    If posBody > 0 Then
        InsertSummaryLine = Left$(htmlBody, posBody - 1) & summaryHtml & Mid$(htmlBody, posBody)
    Else
        ' No </body> tag – append at end
        InsertSummaryLine = htmlBody & summaryHtml
    End If
End Function

' ===========================================================================
' HELPER UTILITIES
' ===========================================================================

' Strips HTML tags from a string (for extracting display text from anchors).
Private Function StripHtmlTags(s As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "<[^>]+>"
    End With
    StripHtmlTags = Trim$(regex.Replace(s, ""))
End Function

' Removes characters illegal in Windows filenames and trims length.
Private Function SanitizeFilename(rawName As String) As String
    Dim s As String
    s = rawName

    ' Decode common HTML entities
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "(")
    s = Replace(s, "&gt;", ")")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&quot;", "'")
    s = Replace(s, "&nbsp;", " ")

    ' Replace illegal filename characters
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .Pattern = "[\\/:*?""<>|]"
    End With
    s = regex.Replace(s, "_")

    ' Collapse multiple underscores/spaces
    Dim regexMulti As Object
    Set regexMulti = CreateObject("VBScript.RegExp")
    With regexMulti
        .Global = True
        .Pattern = "[_ ]{2,}"
    End With
    s = regexMulti.Replace(s, "_")

    ' Trim whitespace and underscores from ends
    s = Trim$(s)
    Do While Len(s) > 0 And Left$(s, 1) = "_"
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0 And Right$(s, 1) = "_"
        s = Left$(s, Len(s) - 1)
    Loop

    ' Limit total filename length (Windows MAX_PATH safety)
    ' Reserve room for .url extension and temp path
    If Len(s) > 150 Then
        s = Left$(s, 150)
    End If

    ' Fallback if empty after sanitization
    If Len(Trim$(s)) = 0 Then
        s = "NetDocs_Document"
    End If

    SanitizeFilename = s
End Function

' Writes a Windows .url shortcut file.
Private Sub WriteUrlFile(filePath As String, url As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, "[InternetShortcut]"
    Print #fNum, "URL=" & url
    Close #fNum
End Sub

' De-duplicates links by URL (keeps first occurrence).
Private Function DeduplicateLinks(links As Collection) As Collection
    Dim result As New Collection
    Dim seen As New Collection
    Dim i As Long
    Dim linkData As Variant
    Dim url As String

    For i = 1 To links.Count
        linkData = links(i)
        url = LCase$(CStr(linkData(0)))

        Dim exists As Boolean
        exists = False
        On Error Resume Next
        Dim tmp As String
        tmp = seen(url)
        If Err.Number = 0 Then exists = True
        Err.Clear
        On Error GoTo 0

        If Not exists Then
            seen.Add "", url
            result.Add linkData
        End If
    Next i

    Set DeduplicateLinks = result
End Function

' Returns a Collection of lowercase attachment filenames on the mail item.
Private Function GetExistingAttachmentNames(oMail As Outlook.MailItem) As Collection
    Dim result As New Collection
    Dim i As Long
    For i = 1 To oMail.Attachments.Count
        On Error Resume Next
        result.Add LCase$(oMail.Attachments(i).fileName), LCase$(oMail.Attachments(i).fileName)
        On Error GoTo 0
    Next i
    Set GetExistingAttachmentNames = result
End Function

' Checks if a Collection contains a given key.
Private Function CollectionContains(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim v As String
    v = col(key)
    CollectionContains = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
