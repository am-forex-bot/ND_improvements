Attribute VB_Name = "NetDocsLinkConverter"
'==============================================================================
' NetDocuments Link-to-Attachment Converter for Microsoft Outlook
'==============================================================================
' PURPOSE:  Scans the active email for NetDocuments URLs, creates .html
'           redirect files, attaches them, and strips the links from the body.
'
' HOW IT WORKS:
'           NetDocuments inserts multiple links per document (OPEN, VIEW,
'           GO TO). This macro groups them by document ID, picks only the
'           OPEN URL (which triggers ndOffice to open in native Word/Excel),
'           and creates ONE .html redirect attachment per document.
'
'           The document name comes from the nearby non-action hyperlink
'           text (e.g. "test letter.docx"), not from action labels.
'
' INSTALL:  Alt+F11 in Outlook -> Import File... -> select this .bas
'           -OR- paste into a new Module.
'
' RUN:      Open an email -> Alt+F8 -> ConvertNetDocsLinksToAttachments -> Run
'
' NOTES:    - .html files are NOT blocked by Outlook attachment security.
'           - Safe to run multiple times - skips already-attached filenames.
'           - Does NOT control how Word/Excel open; relies on existing
'             ndOffice / NetDocuments integration.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
'  Constants
' ---------------------------------------------------------------------------
Private Const NETDOCS_TEMP_FOLDER As String = "NetDocsLinks"
Private Const SUMMARY_TEXT As String = "Documents attached via NetDocuments"

' Action labels to ignore when determining document names
Private Const ACTION_LABELS As String = "OPEN|VIEW|GO TO|GOTO|EDIT|DOWNLOAD|PROFILE|CHECK OUT|CHECKOUT|CHECK IN|CHECKIN"

' ---------------------------------------------------------------------------
'  ENTRY POINT
' ---------------------------------------------------------------------------
Public Sub ConvertNetDocsLinksToAttachments()
    On Error GoTo ErrHandler

    ' --- 1. Validate: must have an open MailItem in an Inspector -----------
    Dim oInsp As Outlook.Inspector
    Set oInsp = Application.ActiveInspector
    If oInsp Is Nothing Then
        MsgBox "Please open an email first.", vbExclamation, "NetDocs Converter"
        Exit Sub
    End If

    If Not TypeOf oInsp.CurrentItem Is Outlook.MailItem Then
        MsgBox "This macro works only on email messages.", vbExclamation, "NetDocs Converter"
        Exit Sub
    End If

    Dim oMail As Outlook.MailItem
    Set oMail = oInsp.CurrentItem

    ' --- 2. Parse HTML body for NetDocuments links -------------------------
    Dim htmlBody As String
    htmlBody = oMail.HTMLBody

    ' Collect ALL ND links: Array(url, displayText) for each
    Dim rawLinks As Collection
    Set rawLinks = CollectAllNetDocsLinks(htmlBody)

    If rawLinks.Count = 0 Then
        MsgBox "No NetDocuments links found in this email.", vbInformation, "NetDocs Converter"
        Exit Sub
    End If

    ' --- 3. Group by document ID and pick OPEN URL + real name -------------
    '     Returns Collection of Array(openUrl, documentName, docId)
    Dim docs As Collection
    Set docs = GroupByDocument(rawLinks)

    If docs.Count = 0 Then
        MsgBox "No usable NetDocuments OPEN links found.", vbInformation, "NetDocs Converter"
        Exit Sub
    End If

    ' --- 4. Prepare temp folder --------------------------------------------
    Dim tempDir As String
    tempDir = PrepareTempFolder()

    ' --- 5. Build existing-attachment set (re-run guard) -------------------
    Dim attached As Object  ' Scripting.Dictionary
    Set attached = CreateObject("Scripting.Dictionary")
    attached.CompareMode = vbTextCompare
    Dim a As Long
    For a = 1 To oMail.Attachments.Count
        attached(oMail.Attachments(a).FileName) = True
    Next a

    ' --- 6. Create .html redirect files and attach -------------------------
    Dim createdPaths As New Collection
    Dim attachCount As Long
    Dim i As Long

    For i = 1 To docs.Count
        Dim doc As Variant: doc = docs(i)
        Dim openUrl As String:  openUrl = CStr(doc(0))
        Dim docName As String:  docName = CStr(doc(1))
        Dim docId As String:    docId = CStr(doc(2))

        Dim fName As String
        fName = DetermineFilename(docName, docId)

        ' Skip if already attached (idempotent on re-run)
        If Not attached.Exists(fName) Then
            Dim fPath As String
            fPath = tempDir & fName

            WriteHtmlRedirect fPath, openUrl, docName
            createdPaths.Add fPath

            oMail.Attachments.Add fPath, olByValue
            attached(fName) = True
            attachCount = attachCount + 1
        End If
    Next i

    ' --- 7. Strip NetDocuments links from HTML body ------------------------
    Dim cleaned As String
    cleaned = StripNetDocsFromHtml(htmlBody)
    cleaned = InjectSummaryLine(cleaned, docs.Count)
    oMail.HTMLBody = cleaned

    ' --- 8. Clean up temp files --------------------------------------------
    Dim p As Variant
    For Each p In createdPaths
        On Error Resume Next
        Kill CStr(p)
        On Error GoTo 0
    Next p

    ' --- 9. Confirm --------------------------------------------------------
    MsgBox "NetDocuments links converted to attachments", vbInformation, "NetDocs Converter"
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "NetDocs Converter"
End Sub

' ===========================================================================
'  LINK COLLECTION
' ===========================================================================

' Collects ALL netdocuments.com links from the HTML body.
' Returns Collection of Array(url, displayText) - one per link found.
Private Function CollectAllNetDocsLinks(htmlBody As String) As Collection
    Dim result As New Collection

    ' --- Anchors -----------------------------------------------------------
    Dim reA As Object: Set reA = NewRegex( _
        "<a\b[^>]*?\bhref\s*=\s*[""']([^""']*netdocuments\.com[^""']*)[""'][^>]*>([\s\S]*?)<\/a>", _
        True)
    Dim mc As Object
    Dim seenUrls As Object: Set seenUrls = CreateObject("Scripting.Dictionary")
    seenUrls.CompareMode = vbTextCompare

    If reA.Test(htmlBody) Then
        Set mc = reA.Execute(htmlBody)
        Dim m As Object
        For Each m In mc
            ' CRITICAL: URLs from HTML href attributes contain &amp; instead of &
            ' Decode them to raw URLs immediately so the rest of the pipeline works
            Dim aUrl As String:  aUrl = HtmlDecodeUrl(Trim$(m.SubMatches(0)))
            Dim aText As String: aText = StripTags(Trim$(m.SubMatches(1)))
            If Len(aUrl) > 0 Then
                result.Add Array(aUrl, aText)
                seenUrls(aUrl) = True
            End If
        Next m
    End If

    ' --- Raw URLs (outside anchors) ----------------------------------------
    Dim stripped As String: stripped = reA.Replace(htmlBody, " ")
    Dim reH As Object: Set reH = NewRegex("href\s*=\s*[""'][^""']*[""']", True)
    stripped = reH.Replace(stripped, " ")

    Dim reR As Object: Set reR = NewRegex( _
        "(https?://[^\s""'<>]*netdocuments\.com[^\s""'<>]*)", True)

    If reR.Test(stripped) Then
        Set mc = reR.Execute(stripped)
        For Each m In mc
            Dim rUrl As String: rUrl = HtmlDecodeUrl(Trim$(m.SubMatches(0)))
            If Len(rUrl) > 0 And Not seenUrls.Exists(rUrl) Then
                result.Add Array(rUrl, "")
                seenUrls(rUrl) = True
            End If
        Next m
    End If

    Set CollectAllNetDocsLinks = result
End Function

' ===========================================================================
'  DOCUMENT GROUPING
' ===========================================================================

' Groups links by document ID (from the filter= param).
' For each document, picks the OPEN URL and the real document name.
' Returns Collection of Array(openUrl, documentName, docId).
Private Function GroupByDocument(rawLinks As Collection) As Collection
    Dim result As New Collection

    ' Dict keyed by docId -> Array(openUrl, bestName)
    Dim docMap As Object: Set docMap = CreateObject("Scripting.Dictionary")
    docMap.CompareMode = vbTextCompare

    ' Track insertion order
    Dim docOrder As New Collection

    Dim i As Long
    For i = 1 To rawLinks.Count
        Dim entry As Variant: entry = rawLinks(i)
        Dim url As String:   url = CStr(entry(0))
        Dim dText As String: dText = CStr(entry(1))

        ' Extract doc ID from the filter parameter
        Dim dId As String: dId = ExtractDocIdFromFilter(url)
        If Len(dId) = 0 Then dId = DocIdFallback(url)
        If Len(dId) = 0 Then GoTo NextLink

        ' Determine if this is the OPEN URL:
        '   has &open=1 but does NOT have &openMode=
        Dim isOpen As Boolean
        isOpen = IsOpenUrl(url)

        ' Determine if display text is a real doc name (not an action label)
        Dim isRealName As Boolean
        isRealName = (Len(dText) > 0 And Not IsActionLabel(dText))

        If docMap.Exists(dId) Then
            ' Update existing entry
            Dim existing As Variant: existing = docMap(dId)
            ' Prefer OPEN URL over whatever we had
            If isOpen And Len(CStr(existing(0))) = 0 Then
                existing(0) = url
            ElseIf isOpen Then
                ' If we already have an open URL, keep first one
                ' but override if current is cleaner
                existing(0) = url
            End If
            ' Prefer real doc name over what we had
            If isRealName And Len(CStr(existing(1))) = 0 Then
                existing(1) = dText
            End If
            docMap(dId) = existing
        Else
            ' New document
            Dim newEntry(0 To 1) As String
            If isOpen Then newEntry(0) = url Else newEntry(0) = ""
            If isRealName Then newEntry(1) = dText Else newEntry(1) = ""
            docMap.Add dId, newEntry
            docOrder.Add dId
        End If

NextLink:
    Next i

    ' Build result: only include documents where we found an OPEN URL
    Dim j As Long
    For j = 1 To docOrder.Count
        Dim docId As String: docId = docOrder(j)
        Dim info As Variant: info = docMap(docId)
        Dim oUrl As String: oUrl = CStr(info(0))
        Dim oName As String: oName = CStr(info(1))

        ' If no OPEN URL found, fall back to first URL for this doc
        ' (shouldn't happen with well-formed ND links, but be safe)
        If Len(oUrl) = 0 Then
            ' Scan raw links for first URL with this docId
            For i = 1 To rawLinks.Count
                entry = rawLinks(i)
                If ExtractDocIdFromFilter(CStr(entry(0))) = docId Then
                    oUrl = CStr(entry(0))
                    Exit For
                End If
            Next i
        End If

        If Len(oUrl) > 0 Then
            result.Add Array(oUrl, oName, docId)
        End If
    Next j

    Set GroupByDocument = result
End Function

' Extracts the document ID from the filter= query parameter.
' URL pattern: filter=%3D999%28XXXX-XXXX-XXXX%29
' Decoded:     filter==999(XXXX-XXXX-XXXX)
' Returns the ID portion (e.g. "4158-2396-4775")
Private Function ExtractDocIdFromFilter(url As String) As String
    ' Match the encoded form: filter=%3D999%28...%29
    Dim re As Object: Set re = NewRegex( _
        "filter=%3D999%28([A-Za-z0-9\-]+)%29", False)
    If re.Test(url) Then
        ExtractDocIdFromFilter = re.Execute(url)(0).SubMatches(0)
        Exit Function
    End If

    ' Match the decoded form: filter==999(...)
    Dim re2 As Object: Set re2 = NewRegex( _
        "filter==999\(([A-Za-z0-9\-]+)\)", False)
    If re2.Test(url) Then
        ExtractDocIdFromFilter = re2.Execute(url)(0).SubMatches(0)
        Exit Function
    End If

    ExtractDocIdFromFilter = ""
End Function

' Fallback doc ID extraction for non-standard ND URLs
Private Function DocIdFallback(url As String) As String
    Dim patterns As Variant
    patterns = Array( _
        "[?&]ndDocId=([A-Za-z0-9\-]+)", _
        "/nddocview[^/]*/([A-Za-z0-9\-]+)", _
        "/document/([A-Za-z0-9\-]+)", _
        "/d/([A-Za-z0-9\-]+)")

    Dim p As Variant
    For Each p In patterns
        Dim re As Object: Set re = NewRegex(CStr(p), False)
        If re.Test(url) Then
            DocIdFallback = re.Execute(url)(0).SubMatches(0)
            Exit Function
        End If
    Next p
    DocIdFallback = ""
End Function

' Returns True if the URL is an OPEN link:
'   has &open=1 but does NOT have &openMode=
Private Function IsOpenUrl(url As String) As Boolean
    Dim lUrl As String: lUrl = LCase$(url)
    If InStr(lUrl, "open=1") > 0 And InStr(lUrl, "openmode=") = 0 Then
        IsOpenUrl = True
    Else
        IsOpenUrl = False
    End If
End Function

' Returns True if the text is an action label (OPEN, VIEW, GO TO, etc.)
Private Function IsActionLabel(txt As String) As Boolean
    Dim upper As String: upper = UCase$(Trim$(txt))
    Dim labels As Variant: labels = Split(ACTION_LABELS, "|")
    Dim lbl As Variant
    For Each lbl In labels
        If upper = CStr(lbl) Then
            IsActionLabel = True
            Exit Function
        End If
    Next lbl
    IsActionLabel = False
End Function

' ===========================================================================
'  FILENAME DETERMINATION
' ===========================================================================

' Priority: document name from grouping -> doc ID -> fallback
' All filenames end with .html
Private Function DetermineFilename(docName As String, docId As String) As String
    Dim base As String

    ' Priority 1 - real document name from hyperlink text
    If Len(Trim$(docName)) > 0 Then
        base = Trim$(docName)
        GoTo Finish
    End If

    ' Priority 2 - document ID
    If Len(Trim$(docId)) > 0 Then
        base = "NetDocs_" & Trim$(docId)
        GoTo Finish
    End If

    ' Fallback
    base = "NetDocs_Document"

Finish:
    base = MakeWindowsSafe(base)

    ' Strip known document extensions (e.g. .docx from display text)
    ' so the final file is cleanly named .html
    Dim dotPos As Long: dotPos = InStrRev(base, ".")
    If dotPos > 1 Then
        Dim ext As String: ext = LCase$(Mid$(base, dotPos))
        If ext = ".doc" Or ext = ".docx" Or ext = ".xls" Or ext = ".xlsx" _
           Or ext = ".ppt" Or ext = ".pptx" Or ext = ".pdf" Or ext = ".txt" _
           Or ext = ".csv" Or ext = ".rtf" Or ext = ".msg" Or ext = ".html" _
           Or ext = ".url" Then
            base = Left$(base, dotPos - 1)
        End If
    End If

    base = base & ".html"
    DetermineFilename = base
End Function

' ===========================================================================
'  HTML CLEANUP
' ===========================================================================

Private Function StripNetDocsFromHtml(htmlBody As String) As String
    Dim s As String: s = htmlBody

    ' --- 1. Remove entire <table>/<div> blocks that contain netdocuments.com links
    '     NetDocuments inserts tables with: filename | OPEN | VIEW | GO TO
    '     and div blocks with bordered file name entries
    Dim reT As Object: Set reT = NewRegex( _
        "<table\b[^>]*>[\s\S]*?netdocuments\.com[\s\S]*?<\/table>", True)
    s = reT.Replace(s, "")

    ' Remove <div> blocks that contain netdocuments.com links
    Dim reDiv As Object: Set reDiv = NewRegex( _
        "<div\b[^>]*>[\s\S]*?netdocuments\.com[\s\S]*?<\/div>", True)
    s = reDiv.Replace(s, "")

    ' --- 2. Remove "Secured by NetDocuments" lines (any variation) -------------
    '     May appear as <p>, <div>, <span>, or bare text with ® symbol
    Dim reSec As Object: Set reSec = NewRegex( _
        "<[^>]*>[^<]*Secured\s+by\s*[^<]*NetDocuments[^<]*<\/[^>]+>", True)
    s = reSec.Replace(s, "")
    ' Also catch it wrapped in multiple tags (e.g. <p><span>Secured by...</span></p>)
    Dim reSec2 As Object: Set reSec2 = NewRegex( _
        "<(p|div)\b[^>]*>\s*(<[^>]*>)*\s*Secured\s+by\s*[^<]*NetDocuments[\s\S]*?<\/\1>", True)
    s = reSec2.Replace(s, "")

    ' --- 3. Remove remaining <a> tags with netdocuments.com href ---------------
    Dim reA As Object: Set reA = NewRegex( _
        "<a\b[^>]*?\bhref\s*=\s*[""'][^""']*netdocuments\.com[^""']*[""'][^>]*>[\s\S]*?<\/a>", True)
    s = reA.Replace(s, "")

    ' --- 4. Remove raw ND URLs -------------------------------------------------
    Dim reR As Object: Set reR = NewRegex( _
        "https?://[^\s""'<>]*netdocuments\.com[^\s""'<>]*", True)
    s = reR.Replace(s, "")

    ' --- 5. Collapse empty wrappers left behind --------------------------------
    Dim reE As Object: Set reE = NewRegex( _
        "<(p|div|li|span|td|tr|table)\b[^>]*>\s*(&nbsp;|\s|<br\s*/?>)*\s*<\/\1>", True)
    s = reE.Replace(s, "")
    s = reE.Replace(s, "")   ' second pass for nesting
    s = reE.Replace(s, "")   ' third pass for deeper nesting

    StripNetDocsFromHtml = s
End Function

Private Function InjectSummaryLine(htmlBody As String, cnt As Long) As String
    Dim tag As String
    tag = "<p style=""color:#336699;font-weight:bold;margin:12px 0;"">" & _
          SUMMARY_TEXT & " (" & cnt & " file" & IIf(cnt <> 1, "s", "") & ")</p>"

    Dim pos As Long: pos = InStrRev(LCase$(htmlBody), "</body>")
    If pos > 0 Then
        InjectSummaryLine = Left$(htmlBody, pos - 1) & tag & Mid$(htmlBody, pos)
    Else
        InjectSummaryLine = htmlBody & tag
    End If
End Function

' ===========================================================================
'  HTML REDIRECT FILE WRITER
' ===========================================================================

' Creates a self-contained .html file that opens the NetDocuments OPEN URL.
'
' Uses window.location.replace() which triggers browser extensions (ndOffice)
' more reliably than <meta refresh>. Also has a prominent click-through link
' as fallback, and a final meta-refresh safety net.
'
' ndOffice handles opening in the native app and reusing existing Word/Excel
' instances via COM automation.
Private Sub WriteHtmlRedirect(filePath As String, url As String, docName As String)
    Dim f As Integer: f = FreeFile
    Dim safeUrl As String: safeUrl = HtmlEncode(url)
    Dim safeTitle As String

    If Len(Trim$(docName)) > 0 Then
        safeTitle = HtmlEncode(docName)
    Else
        safeTitle = "NetDocuments Document"
    End If

    ' Build the raw URL for JS (needs different escaping than HTML attributes)
    Dim jsUrl As String: jsUrl = Replace(url, "\", "\\")
    jsUrl = Replace(jsUrl, "'", "\'")
    jsUrl = Replace(jsUrl, """", "\""")

    Open filePath For Output As #f
    Print #f, "<!DOCTYPE html>"
    Print #f, "<html><head>"
    Print #f, "<meta charset=""utf-8"">"
    Print #f, "<title>" & safeTitle & "</title>"
    ' Meta refresh as last-resort fallback (2 second delay to let JS try first)
    Print #f, "<meta http-equiv=""refresh"" content=""2;url=" & safeUrl & """>"
    Print #f, "<style>"
    Print #f, "  body { font-family: Segoe UI, Arial, sans-serif; margin: 40px;"
    Print #f, "         color: #333; background: #f9f9f9; }"
    Print #f, "  .card { background: #fff; border: 1px solid #ddd; border-radius: 8px;"
    Print #f, "          padding: 30px; max-width: 500px; margin: 60px auto;"
    Print #f, "          text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }"
    Print #f, "  h2 { margin: 0 0 10px; font-size: 18px; color: #1a1a1a; }"
    Print #f, "  p { font-size: 14px; color: #666; margin: 8px 0; }"
    Print #f, "  a.open-link { display: inline-block; margin-top: 12px; padding: 10px 24px;"
    Print #f, "     background: #0066cc; color: #fff; border-radius: 6px;"
    Print #f, "     text-decoration: none; font-weight: bold; font-size: 14px; }"
    Print #f, "  a.open-link:hover { background: #0052a3; }"
    Print #f, "  .spinner { display: inline-block; width: 20px; height: 20px;"
    Print #f, "             border: 3px solid #ddd; border-top-color: #0066cc;"
    Print #f, "             border-radius: 50%; animation: spin 0.8s linear infinite;"
    Print #f, "             margin-bottom: 15px; }"
    Print #f, "  @keyframes spin { to { transform: rotate(360deg); } }"
    Print #f, "</style>"
    Print #f, "</head><body>"
    Print #f, "<div class=""card"">"
    Print #f, "  <div class=""spinner""></div>"
    Print #f, "  <h2>" & safeTitle & "</h2>"
    Print #f, "  <p>Opening document&#8230;</p>"
    Print #f, "  <p><a class=""open-link"" href=""" & safeUrl & """>Open in NetDocuments</a></p>"
    Print #f, "  <p style=""font-size:12px;color:#999;margin-top:16px;"">If the document does not open automatically, click the button above.</p>"
    Print #f, "</div>"
    Print #f, "<script>"
    ' Use location.replace so the browser navigates properly, triggering
    ' ndOffice browser extension interception
    Print #f, "  try { window.location.replace('" & jsUrl & "'); } catch(e) {}"
    Print #f, "</script>"
    Print #f, "</body></html>"
    Close #f
End Sub

' ===========================================================================
'  UTILITY HELPERS
' ===========================================================================

Private Function NewRegex(pat As String, isGlobal As Boolean) As Object
    Set NewRegex = CreateObject("VBScript.RegExp")
    With NewRegex
        .Global = isGlobal
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = pat
    End With
End Function

Private Function StripTags(s As String) As String
    Dim re As Object: Set re = NewRegex("<[^>]+>", True)
    StripTags = Trim$(re.Replace(s, ""))
End Function

' Decodes HTML entities in URLs extracted from href attributes.
' Must handle &amp;amp; (double-encoded) as well as &amp;
Private Function HtmlDecodeUrl(s As String) As String
    Dim r As String: r = s
    ' First pass: &amp;amp; -> &amp; (fix double-encoding)
    Do While InStr(r, "&amp;amp;") > 0
        r = Replace(r, "&amp;amp;", "&amp;")
    Loop
    ' Second pass: &amp; -> &
    r = Replace(r, "&amp;", "&")
    r = Replace(r, "&lt;", "<")
    r = Replace(r, "&gt;", ">")
    r = Replace(r, "&quot;", """")
    r = Replace(r, "&#39;", "'")
    HtmlDecodeUrl = r
End Function

Private Function HtmlEncode(s As String) As String
    Dim r As String: r = s
    r = Replace(r, "&", "&amp;")
    r = Replace(r, """", "&quot;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, "'", "&#39;")
    HtmlEncode = r
End Function

Private Function MakeWindowsSafe(raw As String) As String
    Dim s As String: s = raw
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "(")
    s = Replace(s, "&gt;", ")")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&quot;", "'")
    s = Replace(s, "&nbsp;", " ")

    Dim re As Object: Set re = NewRegex("[\\/:*?""<>|]", True)
    s = re.Replace(s, "_")

    Dim re2 As Object: Set re2 = NewRegex("[_ ]{2,}", True)
    s = Trim$(re2.Replace(s, "_"))

    Do While Len(s) > 0 And Left$(s, 1) = "_": s = Mid$(s, 2): Loop
    Do While Len(s) > 0 And Right$(s, 1) = "_": s = Left$(s, Len(s) - 1): Loop

    If Len(s) > 150 Then s = Left$(s, 150)
    If Len(Trim$(s)) = 0 Then s = "NetDocs_Document"

    MakeWindowsSafe = s
End Function

Private Function PrepareTempFolder() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As String
    folder = Environ$("TEMP")
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    folder = folder & NETDOCS_TEMP_FOLDER & "\"

    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    PrepareTempFolder = folder
End Function
