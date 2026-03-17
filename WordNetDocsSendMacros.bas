Attribute VB_Name = "NetDocsSendMacros"
'==============================================================================
' NetDocuments "Send as Copy" / "Send as Link" for Microsoft Word
'==============================================================================
' PURPOSE:  Two macros for documents opened via ndOffice:
'           1) Send as Copy  - attaches the document file to a new email
'           2) Send as Link  - attaches an .html redirect (same approach as
'              the Outlook NetDocs converter macro)
'
' REQUIRES: Document must be open via ndOffice (path contains ND Office Echo)
'
' INSTALL:  Alt+F11 in Word -> Import File... -> select this .bas
'           -OR- paste into a new Module.
'
' RUN:      Alt+F8 -> SendAsCopy or SendAsLink -> Run
'           (or assign to ribbon/keyboard shortcut)
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
'  Configuration - adjust base URL to match your NetDocuments tenant
' ---------------------------------------------------------------------------
Private Const ND_BASE_URL As String = "https://eu.netdocuments.com"
Private Const TEMP_SUBFOLDER As String = "NetDocsLinks"

' ---------------------------------------------------------------------------
'  ENTRY POINTS
' ---------------------------------------------------------------------------

' Opens a new Outlook email with the current document attached as a file copy.
Public Sub SendAsCopy()
    On Error GoTo ErrHandler

    ' Validate we have an active document
    If Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "NetDocs Send"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Must be saved to disk (ndOffice docs always are)
    If Len(Dir(doc.FullName)) = 0 Then
        MsgBox "Please save the document first.", vbExclamation, "NetDocs Send"
        Exit Sub
    End If

    ' Get a clean display name (without the doc ID and version suffix)
    Dim displayName As String
    displayName = GetCleanDocName(doc)

    ' Create Outlook email with the file attached
    Dim olApp As Object: Set olApp = GetOutlookApp()
    If olApp Is Nothing Then Exit Sub

    Dim oMail As Object: Set oMail = olApp.CreateItem(0)  ' olMailItem = 0
    oMail.Subject = displayName
    oMail.Attachments.Add doc.FullName, 1  ' olByValue = 1
    oMail.Display

    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "NetDocs Send"
End Sub

' Opens a new Outlook email with an .html redirect attachment that opens
' the document via NetDocuments (same approach as the Outlook macro).
Public Sub SendAsLink()
    On Error GoTo ErrHandler

    ' Validate we have an active document
    If Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "NetDocs Send"
        Exit Sub
    End If

    Dim doc As Document: Set doc = ActiveDocument

    ' Extract the NetDocuments document ID from the path or title
    Dim docId As String
    docId = ExtractDocId(doc)

    If Len(docId) = 0 Then
        MsgBox "Could not find a NetDocuments document ID." & vbCrLf & vbCrLf & _
               "This macro only works with documents opened from NetDocuments via ndOffice.", _
               vbExclamation, "NetDocs Send"
        Exit Sub
    End If

    ' Build the NetDocuments URL
    Dim ndUrl As String
    ndUrl = BuildNetDocsUrl(docId, doc)

    ' Get a clean display name
    Dim displayName As String
    displayName = GetCleanDocName(doc)

    ' Create the .html redirect file
    Dim tempDir As String
    tempDir = PrepareTempFolder()

    Dim htmlFileName As String
    htmlFileName = MakeWindowsSafe(displayName) & ".html"

    Dim htmlPath As String
    htmlPath = tempDir & htmlFileName

    WriteHtmlRedirect htmlPath, ndUrl, displayName

    ' Create Outlook email with the redirect attached
    Dim olApp As Object: Set olApp = GetOutlookApp()
    If olApp Is Nothing Then GoTo Cleanup

    Dim oMail As Object: Set oMail = olApp.CreateItem(0)  ' olMailItem = 0
    oMail.Subject = displayName
    oMail.Attachments.Add htmlPath, 1  ' olByValue = 1
    oMail.Display

Cleanup:
    ' Remove temp file
    On Error Resume Next
    Kill htmlPath
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "NetDocs Send"
    Resume Cleanup
End Sub

' ===========================================================================
'  DOCUMENT ID EXTRACTION
' ===========================================================================

' Extracts the ND document ID (e.g. "4144-0680-3302") from the document.
' Checks the full path first, then the window title.
' ndOffice pattern: "test letter 4144-0680-3302 v.1.docx"
Private Function ExtractDocId(doc As Document) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "\b(\d{4}-\d{4}-\d{4})\b"

    ' Try the filename first
    If re.Test(doc.FullName) Then
        ExtractDocId = re.Execute(doc.FullName)(0).SubMatches(0)
        Exit Function
    End If

    ' Try the window caption
    If re.Test(Application.ActiveWindow.Caption) Then
        ExtractDocId = re.Execute(Application.ActiveWindow.Caption)(0).SubMatches(0)
        Exit Function
    End If

    ExtractDocId = ""
End Function

' Extracts the repository/cabinet ID from the ndOffice cache path.
' Path pattern: ...\ND Office Echo\EU-FYM4BHDM\filename.docx
' Returns the cabinet folder name (e.g. "EU-FYM4BHDM") or empty string.
Private Function ExtractRepoId(doc As Document) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "ND Office Echo\\([^\\]+)\\"

    If re.Test(doc.FullName) Then
        ExtractRepoId = re.Execute(doc.FullName)(0).SubMatches(0)
    Else
        ExtractRepoId = ""
    End If
End Function

' ===========================================================================
'  URL CONSTRUCTION
' ===========================================================================

' Builds the NetDocuments URL for opening a document.
' Uses the neWeb2/goid.aspx endpoint which ndOffice intercepts.
Private Function BuildNetDocsUrl(docId As String, doc As Document) As String
    Dim repoId As String
    repoId = ExtractRepoId(doc)

    ' Primary URL format: direct document open
    ' ndOffice and the browser extension both intercept this
    If Len(repoId) > 0 Then
        BuildNetDocsUrl = ND_BASE_URL & "/neWeb2/goid.aspx?id=" & docId & _
                          "&repository=" & repoId & "&open=1"
    Else
        BuildNetDocsUrl = ND_BASE_URL & "/neWeb2/goid.aspx?id=" & docId & "&open=1"
    End If
End Function

' ===========================================================================
'  CLEAN DOCUMENT NAME
' ===========================================================================

' Gets a clean display name from the document, stripping the ND doc ID
' and version suffix that ndOffice appends.
' "test letter 4144-0680-3302 v.1.docx" -> "test letter"
Private Function GetCleanDocName(doc As Document) As String
    Dim rawName As String

    ' Start with just the filename (no path, no extension)
    rawName = doc.Name
    Dim dotPos As Long: dotPos = InStrRev(rawName, ".")
    If dotPos > 1 Then rawName = Left$(rawName, dotPos - 1)

    ' Strip the doc ID and version suffix: " 4144-0680-3302 v.1"
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "\s+\d{4}-\d{4}-\d{4}\s+v\.\d+$"

    If re.Test(rawName) Then
        rawName = Trim$(re.Replace(rawName, ""))
    End If

    ' If stripping left nothing, fall back to original name
    If Len(Trim$(rawName)) = 0 Then
        rawName = doc.Name
        dotPos = InStrRev(rawName, ".")
        If dotPos > 1 Then rawName = Left$(rawName, dotPos - 1)
    End If

    GetCleanDocName = Trim$(rawName)
End Function

' ===========================================================================
'  HTML REDIRECT FILE WRITER
' ===========================================================================

' Creates a self-contained .html file that opens the NetDocuments URL.
' Same approach as the Outlook NetDocs converter macro.
Private Sub WriteHtmlRedirect(filePath As String, url As String, docName As String)
    Dim f As Integer: f = FreeFile
    Dim safeUrl As String: safeUrl = HtmlEncode(url)
    Dim safeTitle As String

    If Len(Trim$(docName)) > 0 Then
        safeTitle = HtmlEncode(docName)
    Else
        safeTitle = "NetDocuments Document"
    End If

    Dim jsUrl As String: jsUrl = Replace(url, "\", "\\")
    jsUrl = Replace(jsUrl, "'", "\'")
    jsUrl = Replace(jsUrl, """", "\""")

    Open filePath For Output As #f
    Print #f, "<!DOCTYPE html>"
    Print #f, "<html><head>"
    Print #f, "<meta charset=""utf-8"">"
    Print #f, "<title>" & safeTitle & "</title>"
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
    Print #f, "  try { window.location.replace('" & jsUrl & "'); } catch(e) {}"
    Print #f, "</script>"
    Print #f, "</body></html>"
    Close #f
End Sub

' ===========================================================================
'  OUTLOOK HELPER
' ===========================================================================

' Gets or creates an Outlook Application instance.
Private Function GetOutlookApp() As Object
    On Error Resume Next
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    If GetOutlookApp Is Nothing Then
        Set GetOutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If GetOutlookApp Is Nothing Then
        MsgBox "Could not start Outlook.", vbCritical, "NetDocs Send"
    End If
End Function

' ===========================================================================
'  UTILITY HELPERS
' ===========================================================================

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
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\\/:*?""<>|]"
    s = re.Replace(s, "_")

    Dim re2 As Object: Set re2 = CreateObject("VBScript.RegExp")
    re2.Global = True
    re2.Pattern = "[_ ]{2,}"
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
    folder = folder & TEMP_SUBFOLDER & "\"

    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    PrepareTempFolder = folder
End Function
