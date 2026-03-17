Attribute VB_Name = "NetDocsLinkConverter"
'==============================================================================
' NetDocuments Link-to-Attachment Converter for Microsoft Outlook
'==============================================================================
' PURPOSE:  Scans the active email for NetDocuments URLs, creates .html
'           redirect files, attaches them, and strips the links from the body.
'
' HOW IT WORKS:
'           Each .html file contains a lightweight redirect page that
'           opens the NetDocuments URL in the default browser. ndOffice
'           then intercepts the URL and opens the document in the native
'           app (Word/Excel/etc). If the app is already running, ndOffice
'           uses the existing instance — no duplicate windows.
'
' INSTALL:  Alt+F11 in Outlook → Import File… → select this .bas
'           -OR- paste into a new Module.
'
' RUN:      Open an email → Alt+F8 → ConvertNetDocsLinksToAttachments → Run
'
' NOTES:    - .html files are NOT blocked by Outlook attachment security,
'             so no registry/GPO changes are needed.
'           - Safe to run multiple times – skips already-attached filenames.
'           - Does NOT control how Word/Excel open; relies on existing
'             ndOffice / NetDocuments integration.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
'  Constants
' ---------------------------------------------------------------------------
Private Const NETDOCS_TEMP_FOLDER As String = "NetDocsLinks"
Private Const SUMMARY_TEXT As String = "Documents attached via NetDocuments"

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

    Dim links As Collection          ' Each item: Array(url, displayText)
    Set links = CollectNetDocsLinks(htmlBody)

    If links.Count = 0 Then
        MsgBox "No NetDocuments links found in this email.", vbInformation, "NetDocs Converter"
        Exit Sub
    End If

    ' --- 3. De-duplicate by URL --------------------------------------------
    Dim unique As Collection
    Set unique = DeduplicateByUrl(links)

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

    For i = 1 To unique.Count
        Dim entry As Variant: entry = unique(i)
        Dim sUrl As String:   sUrl = CStr(entry(0))
        Dim sName As String:  sName = CStr(entry(1))

        Dim fName As String
        fName = DetermineFilename(sName, sUrl)

        ' Skip if already attached (idempotent on re-run)
        If Not attached.Exists(fName) Then
            Dim fPath As String
            fPath = tempDir & fName

            WriteHtmlRedirect fPath, sUrl, sName
            createdPaths.Add fPath

            oMail.Attachments.Add fPath, olByValue
            attached(fName) = True
            attachCount = attachCount + 1
        End If
    Next i

    ' --- 7. Strip NetDocuments links from HTML body ------------------------
    Dim cleaned As String
    cleaned = StripNetDocsFromHtml(htmlBody)
    cleaned = InjectSummaryLine(cleaned, unique.Count)
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

' Scans HTMLBody and returns Collection of Array(url, displayText).
' Pass 1: <a href="…netdocuments.com…">text</a>
' Pass 2: raw URLs outside anchors
Private Function CollectNetDocsLinks(htmlBody As String) As Collection
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
            Dim aUrl As String:  aUrl = Trim$(m.SubMatches(0))
            Dim aText As String: aText = StripTags(Trim$(m.SubMatches(1)))
            If Len(aUrl) > 0 Then
                result.Add Array(aUrl, aText)
                seenUrls(aUrl) = True
            End If
        Next m
    End If

    ' --- Raw URLs (outside anchors) ----------------------------------------
    ' Remove anchors first so we don't double-count
    Dim stripped As String: stripped = reA.Replace(htmlBody, " ")
    ' Also remove any remaining href attrs
    Dim reH As Object: Set reH = NewRegex("href\s*=\s*[""'][^""']*[""']", True)
    stripped = reH.Replace(stripped, " ")

    Dim reR As Object: Set reR = NewRegex( _
        "(https?://[^\s""'<>]*netdocuments\.com[^\s""'<>]*)", True)

    If reR.Test(stripped) Then
        Set mc = reR.Execute(stripped)
        For Each m In mc
            Dim rUrl As String: rUrl = Trim$(m.SubMatches(0))
            If Len(rUrl) > 0 And Not seenUrls.Exists(rUrl) Then
                result.Add Array(rUrl, "")
                seenUrls(rUrl) = True
            End If
        Next m
    End If

    Set CollectNetDocsLinks = result
End Function

' ===========================================================================
'  FILENAME DETERMINATION
' ===========================================================================

' Priority: displayText → filename in URL → docID → fallback
' All filenames end with .html
Private Function DetermineFilename(displayText As String, url As String) As String
    Dim base As String

    ' Priority 1 – hyperlink display text
    If Len(Trim$(displayText)) > 0 Then
        base = Trim$(displayText)
        GoTo Finish
    End If

    ' Priority 2 – filename extracted from URL path / query
    base = FilenameFromUrl(url)
    If Len(base) > 0 Then GoTo Finish

    ' Priority 3 – document ID
    Dim docId As String: docId = DocIdFromUrl(url)
    If Len(docId) > 0 Then
        base = "NetDocs_" & docId
        GoTo Finish
    End If

    ' Fallback
    base = "NetDocs_Document"

Finish:
    base = MakeWindowsSafe(base)

    ' Strip any existing file extension (e.g. .docx from display text)
    ' so the final file is cleanly named .html
    Dim dotPos As Long: dotPos = InStrRev(base, ".")
    If dotPos > 1 Then
        Dim ext As String: ext = LCase$(Mid$(base, dotPos))
        ' Only strip known document extensions to avoid mangling names with dots
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

Private Function FilenameFromUrl(url As String) As String
    Dim re As Object: Set re = NewRegex( _
        "[\/?&=]([A-Za-z0-9_\-\. ]+\.(docx?|xlsx?|pptx?|pdf|txt|csv|rtf|msg))", False)
    If re.Test(url) Then
        FilenameFromUrl = re.Execute(url)(0).SubMatches(0)
    End If
End Function

Private Function DocIdFromUrl(url As String) As String
    Dim patterns As Variant
    patterns = Array( _
        "[?&]ndDocId=([A-Za-z0-9\-]+)", _
        "/nddocview[^/]*/([A-Za-z0-9\-]+)", _
        "/document/([A-Za-z0-9\-]+)", _
        "/d/([A-Za-z0-9\-]+)", _
        "/([0-9]{4,}[\-/][0-9]+)")

    Dim p As Variant
    For Each p In patterns
        Dim re As Object: Set re = NewRegex(CStr(p), False)
        If re.Test(url) Then
            DocIdFromUrl = re.Execute(url)(0).SubMatches(0)
            Exit Function
        End If
    Next p

    ' Last-resort: final path segment >= 6 chars
    Dim reL As Object: Set reL = NewRegex("/([A-Za-z0-9\-]{6,})", True)
    If reL.Test(url) Then
        Dim mc As Object: Set mc = reL.Execute(url)
        DocIdFromUrl = mc(mc.Count - 1).SubMatches(0)
    End If
End Function

' ===========================================================================
'  HTML CLEANUP
' ===========================================================================

Private Function StripNetDocsFromHtml(htmlBody As String) As String
    Dim s As String: s = htmlBody

    ' Remove <a> tags with netdocuments.com href
    Dim reA As Object: Set reA = NewRegex( _
        "<a\b[^>]*?\bhref\s*=\s*[""'][^""']*netdocuments\.com[^""']*[""'][^>]*>[\s\S]*?<\/a>", True)
    s = reA.Replace(s, "")

    ' Remove raw URLs
    Dim reR As Object: Set reR = NewRegex( _
        "https?://[^\s""'<>]*netdocuments\.com[^\s""'<>]*", True)
    s = reR.Replace(s, "")

    ' Collapse empty wrappers left behind
    Dim reE As Object: Set reE = NewRegex( _
        "<(p|div|li|span)\b[^>]*>\s*(&nbsp;|\s|<br\s*/?>)*\s*<\/\1>", True)
    s = reE.Replace(s, "")
    s = reE.Replace(s, "")   ' second pass for nesting

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

' Creates a self-contained .html file that immediately redirects to the
' NetDocuments URL. The page:
'   1. Uses <meta http-equiv="refresh"> for instant redirect (works everywhere)
'   2. Has a manual click-through link as fallback
'   3. Shows the document name so the user knows what's opening
'   4. ndOffice intercepts the ND URL and opens in the native app
'   5. If Word/Excel is already open, ndOffice reuses that instance
Private Sub WriteHtmlRedirect(filePath As String, url As String, docName As String)
    Dim f As Integer: f = FreeFile
    Dim safeUrl As String: safeUrl = HtmlEncode(url)
    Dim safeTitle As String

    If Len(Trim$(docName)) > 0 Then
        safeTitle = HtmlEncode(docName)
    Else
        safeTitle = "NetDocuments Document"
    End If

    Open filePath For Output As #f
    Print #f, "<!DOCTYPE html>"
    Print #f, "<html><head>"
    Print #f, "<meta charset=""utf-8"">"
    Print #f, "<title>" & safeTitle & "</title>"
    Print #f, "<meta http-equiv=""refresh"" content=""0;url=" & safeUrl & """>"
    Print #f, "<style>"
    Print #f, "  body { font-family: Segoe UI, Arial, sans-serif; margin: 40px;"
    Print #f, "         color: #333; background: #f9f9f9; }"
    Print #f, "  .card { background: #fff; border: 1px solid #ddd; border-radius: 8px;"
    Print #f, "          padding: 30px; max-width: 500px; margin: 60px auto;"
    Print #f, "          text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }"
    Print #f, "  h2 { margin: 0 0 10px; font-size: 18px; color: #1a1a1a; }"
    Print #f, "  p { font-size: 14px; color: #666; margin: 8px 0; }"
    Print #f, "  a { color: #0066cc; text-decoration: none; }"
    Print #f, "  a:hover { text-decoration: underline; }"
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
    Print #f, "  <p>Opening in NetDocuments&#8230;</p>"
    Print #f, "  <p><a href=""" & safeUrl & """>Click here if not redirected</a></p>"
    Print #f, "</div>"
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

' Encodes characters for safe HTML attribute/content use.
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
    ' Decode common HTML entities
    s = Replace(s, "&amp;", "&")
    s = Replace(s, "&lt;", "(")
    s = Replace(s, "&gt;", ")")
    s = Replace(s, "&#39;", "'")
    s = Replace(s, "&quot;", "'")
    s = Replace(s, "&nbsp;", " ")

    ' Strip illegal chars
    Dim re As Object: Set re = NewRegex("[\\/:*?""<>|]", True)
    s = re.Replace(s, "_")

    ' Collapse runs of underscores / spaces
    Dim re2 As Object: Set re2 = NewRegex("[_ ]{2,}", True)
    s = Trim$(re2.Replace(s, "_"))

    ' Trim leading/trailing underscores
    Do While Len(s) > 0 And Left$(s, 1) = "_": s = Mid$(s, 2): Loop
    Do While Len(s) > 0 And Right$(s, 1) = "_": s = Left$(s, Len(s) - 1): Loop

    ' Cap at 150 chars
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

Private Function DeduplicateByUrl(links As Collection) As Collection
    Dim result As New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To links.Count
        Dim entry As Variant: entry = links(i)
        Dim u As String: u = CStr(entry(0))
        If Not seen.Exists(u) Then
            seen(u) = True
            result.Add entry
        End If
    Next i
    Set DeduplicateByUrl = result
End Function
