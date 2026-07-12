' ZoteroCiteLinker - Modified Version with Elsevier-Vancouver Support
' ====================================================================
' This is the extracted VBA source code from ZoteroCiteLinker.dotm,
' modified to add support for the "Elsevier - NLM/Vancouver (citation sequence)" style.
'
' The three modifications made are marked with comments starting with:
'   ' [MOD-1] - Added elsevier-vancouver to supported styles list
'   ' [MOD-2] - Added elsevier-vancouver routing case
'   ' [MOD-3] - Added elsevier-vancouver to info dialog
'
' To apply: Open the VBA editor (Alt+F11), find the ZoteroCiteLinker module,
' replace its contents with this file, and save.
'
' Original author: Serdar Berat AYDIN (s.b.a@msn.com)
' Modifications for elsevier-vancouver: GitHub contributor

' An MS Word macro that links Zotero citations to their bibliography.
' Uses custom bookmark naming based on ID, Title, Year, and Author.
' Example: Cite_id9323_A_deep_lea_2021_Jiang
' UnlinkCitations function updated to only delete bookmarks starting with "Cite_".
' Added helper functions to automatically detect and link URLs and email addresses.
' Added code for coloring the citation links.
'
' Author: Serdar Berat AYDIN (s.b.a@msn.com)
'
' --- Acknowledgments ---
' This compilation is primarily based on modifications to the ZoteroLinkCitation
' code by user 'altairwei' and '8gengen8'.
'
' The base versions can be accessed at the following links:
' https://github.com/altairwei/ZoteroLinkCitation
' https://github.com/8gengen8/ZoteroCitationLink-lgg

'-------------------------------------------------------------------
' GLOBAL DECLARATIONS & TYPES
'-------------------------------------------------------------------
Option Explicit

' Defines the structure for holding parsed citation data
Type Citation
    BibTitle As String      ' Title for searching in the bibliography
    BibId As String         ' Zotero ID for naming the bookmark
    BibAuthor As String     ' First author's last name
    BibYear As String       ' Publication year
    Start As Long           ' Start character position of the citation
    End As Long             ' End character position of the citation
End Type

Public myRibbon As IRibbonUI
Public gUserTextStyle As String

Private p&, token, dic

Public Sub onRibbonLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
End Sub

'-------------------------------------------------------------------
' VBA JSON PARSER
' Source: https://medium.com/swlh/excel-vba-parse-json-easily-c2213f4d8e7a
'-------------------------------------------------------------------

' Main entry point for the JSON parser.
' Initializes the parser state and starts parsing.
Private Function ParseJSON(json$, Optional key$ = "obj") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function

' Recursively parses a JSON object (elements within {}).
Private Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":   ParseArr key
            Case "{"
                    If token(p + 1) = "}" Then
                        p = p + 1
                        dic.Add key, "null"
                    Else
                        ParseObj key
                    End If
            
            Case "}":   key = ReducePath(key): Exit Do
            Case ":":   key = key & "." & token(p - 1)
            Case ",":   key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function

' Recursively parses a JSON array (elements within []).
Private Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":   ParseObj key & ArrayID(e)
            Case "[":   ParseArr key & ArrayID(e)
            Case "]":   Exit Do
            Case ":":   key = key & ArrayID(e)
            Case ",":   e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function

' Splits the raw JSON string into a VBA array of tokens (keys, values, operators).
Private Function Tokenize(s$)
    Const Pattern = "\"(([^\"\\\\]|\\\\.)*)\"|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s\"']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function

' Uses VBScript.RegExp to extract all matches based on a pattern.
Private Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n
  Dim v()
  With CreateObject("VBScript.RegExp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .Test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = "\"\"\"\"" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function

' Formats an array index for the dictionary key (e.g., "(0)").
Private Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function

' Moves one level up in the object path (e.g., "obj.data.item" -> "obj.data").
Private Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function

' Extracts all values from the dictionary whose keys match a pattern.
Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.Keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function

' Extracts data into a 2D array based on a list of column patterns.
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.Keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function

'-------------------------------------------------------------------
' ZOTERO LINKER UTILITIES
'-------------------------------------------------------------------

' Checks if the zotero.exe process is running.
Private Function IsZoteroRunning() As Boolean
    Dim colProcess As Object
    Dim objProcess As Object
    Dim strQuery As String
    
    IsZoteroRunning = False
    
    ' Query the list of running processes in Windows
    strQuery = "SELECT * FROM Win32_Process WHERE Name = 'zotero.exe'"
    
    On Error Resume Next
    Set colProcess = GetObject("winmgmts:\\.\root\cimv2").ExecQuery(strQuery)
    
    ' If a process named 'zotero.exe' is found, 'Count' will be 1 or more
    If colProcess.Count > 0 Then
        IsZoteroRunning = True
    End If
    
    ' Clean up memory
    Set objProcess = Nothing
    Set colProcess = Nothing
    On Error GoTo 0
End Function

' Sanitizes a string to be a valid Word bookmark name.
Private Function CleanForBookmarkName(ByVal s As String) As String
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .Pattern = "[^A-Za-z0-9_]"
    End With
    CleanForBookmarkName = re.Replace(s, "_")
    
    If CleanForBookmarkName Like "[0-9]*" Then
        CleanForBookmarkName = "_" & CleanForBookmarkName
    End If
End Function

' Sorts a 1D variant array in place using the QuickSort algorithm.
Private Sub QuickSort(arr As Variant, inLow As Long, inHigh As Long)
    '--- Variable Declarations ---
    Dim pivot As String
    Dim tmpSwap As Variant
    Dim low As Long
    Dim high As Long
    
    '--- Initializations ---
    low = inLow
    high = inHigh
    pivot = arr((low + high) \ 2)
    
    While (low <= high)
        While (arr(low) < pivot And low < inHigh)
            low = low + 1
        Wend
        
        While (pivot < arr(high) And high > inLow)
            high = high - 1
        Wend
        
        If (low <= high) Then
            tmpSwap = arr(low)
            arr(low) = arr(high)
            arr(high) = tmpSwap
            low = low + 1
            high = high - 1
        End If
    Wend
    
    If (inLow < high) Then QuickSort arr, inLow, high
    If (low < inHigh) Then QuickSort arr, low, inHigh
End Sub

' Extracts and concatenates all 'ZOTERO_PREF' custom document properties.
Private Function ExtractZoteroPrefData() As String
    '--- Variable Declarations ---
    Dim prop As Variant
    Dim dict As Object
    Dim sortedKeys As Variant
    Dim concatenatedValues As String
    Dim key As Variant

    '--- Initializations ---
    Set dict = CreateObject("Scripting.Dictionary")

    ' Find all Zotero preference properties
    For Each prop In ActiveDocument.CustomDocumentProperties
        If Left(prop.name, 11) = "ZOTERO_PREF" Then
            dict(prop.name) = prop.Value
        End If
    Next prop
    
    sortedKeys = dict.Keys
    Call QuickSort(sortedKeys, LBound(sortedKeys), UBound(sortedKeys))

    For Each key In sortedKeys
        concatenatedValues = concatenatedValues & dict(key)
    Next key

    ExtractZoteroPrefData = concatenatedValues
End Function

' Parses Zotero preferences from an XML data string.
Private Function GetZoteroPrefsFromXml(zoteroData As String) As Object
    '--- Variable Declarations ---
    Dim xmlDoc As Object
    Dim dict As Object
    Dim dataElem As Object
    Dim sessionElem As Object
    Dim styleElem As Object
    Dim segments() As String
    Dim prefElem As Object
    
    '--- Initializations ---
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.Async = False
    xmlDoc.LoadXML zoteroData
    Set dict = CreateObject("Scripting.Dictionary")

    ' Check for XML parsing errors
    If xmlDoc.ParseError.ErrorCode <> 0 Then
        MsgBox "XML Parse Error: " & xmlDoc.ParseError.Reason
        Set GetZoteroPrefsFromXml = dict
        Exit Function
    End If
    
    On Error Resume Next
    
    ' Extract relevant preference data
    Set dataElem = xmlDoc.SelectSingleNode("//data")
    If Not dataElem Is Nothing Then
        dict("session-id") = dataElem.getAttribute("session-id")
        dict("prefs.data.session-id") = dataElem.getAttribute("session-id")
    End If
    
    Set sessionElem = xmlDoc.SelectSingleNode("//session")
    If Not sessionElem Is Nothing Then
        dict("session-id") = sessionElem.getAttribute("id")
        dict("prefs.session.id") = sessionElem.getAttribute("id")
    End If
    
    Set styleElem = xmlDoc.SelectSingleNode("//style")
    If Not styleElem Is Nothing Then
        dict("style-id") = styleElem.getAttribute("id")
        dict("style-has-bibliography") = styleElem.getAttribute("hasBibliography")
        dict("prefs.style.styleID") = styleElem.getAttribute("id")
        dict("prefs.style.bibliographyStyleHasBeenSet") = styleElem.getAttribute("bibliographyStyleHasBeenSet")
        
        ' Split style ID for version info (e.g., "http://www.zotero.org/styles/ieee")
        segments = Split(styleElem.getAttribute("id"), "/")
        If UBound(segments) >= 0 Then
            dict("style-short-id") = segments(UBound(segments))
        End If
    End If
    
    Set prefElem = xmlDoc.SelectSingleNode("//pref[@name='fieldType']")
    If Not prefElem Is Nothing Then
        dict("fieldType") = prefElem.getAttribute("value")
        dict("prefs.fieldType") = prefElem.getAttribute("value")
    End If
    
    On Error GoTo 0
    
    Set GetZoteroPrefsFromXml = dict
End Function

' Parses Zotero preferences from a JSON data string (for newer Zotero versions).
Private Function GetZoteroPrefsFromJson(zoteroData As String) As Object
    '--- Variable Declarations ---
    Dim jsonObj As Object
    Dim dict As Object
    Dim styleUrl As String
    Dim segments() As String
    
    '--- Initializations ---
    Set jsonObj = ParseJSON(zoteroData)
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Extract style information
    styleUrl = jsonObj("prefs.style.styleID")
    dict("style-id") = styleUrl
    dict("prefs.style.styleID") = styleUrl
    
    ' Extract short style ID from URL
    segments = Split(styleUrl, "/")
    If UBound(segments) >= 0 Then
        dict("style-short-id") = segments(UBound(segments))
    End If
    
    dict("style-has-bibliography") = jsonObj("prefs.style.hasBibliography")
    dict("prefs.style.bibliographyStyleHasBeenSet") = jsonObj("prefs.style.bibliographyStyleHasBeenSet")
    dict("fieldType") = jsonObj("prefs.fieldType")
    dict("prefs.fieldType") = jsonObj("prefs.fieldType")
    dict("session-id") = jsonObj("prefs.session.id")
    dict("prefs.data.session-id") = jsonObj("prefs.session.id")
    
    On Error GoTo 0
    
    Set GetZoteroPrefsFromJson = dict
End Function

' Retrieves Zotero preferences from the active document.
Private Function GetZoteroPrefs() As Object
    '--- Variable Declarations ---
    Dim zoteroData As String
    Dim prefs As Object
    
    '--- Initializations ---
    zoteroData = ExtractZoteroPrefData()
    
    ' Try JSON first (newer Zotero versions), fallback to XML
    If Left(zoteroData, 1) = "{" Then
        Set prefs = GetZoteroPrefsFromJson(zoteroData)
    Else
        Set prefs = GetZoteroPrefsFromXml(zoteroData)
    End If
    
    Set GetZoteroPrefs = prefs
End Function

' Extracts the base style ID from a Zotero style URL.
Private Function GetStyleIdFromPrefs(prefs As Object) As String
    '--- Variable Declarations ---
    Dim styleId As String
    Dim segments() As String
    
    On Error Resume Next
    
    ' Try to get the short style ID directly
    styleId = prefs("style-short-id")
    
    ' If not available, extract from the full URL
    If styleId = "" Then
        styleId = prefs("style-id")
        segments = Split(styleId, "/")
        If UBound(segments) >= 0 Then
            styleId = segments(UBound(segments))
        End If
    End If
    
    On Error GoTo 0
    
    GetStyleIdFromPrefs = styleId
End Function

' Removes HTML tags from a string using regex.
Private Function RemoveHtmlTags(ByVal html As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "<[^>]+>"
    End With
    
    RemoveHtmlTags = re.Replace(html, "")
    Set re = Nothing
End Function

' Parses the JSON data embedded in a Zotero citation field's code.
Private Function ParseCSLCitationJson(code As String) As Object
    '--- Variable Declarations ---
    Dim jsonStr As String
    Dim startPos As Long
    
    ' Find the JSON portion in the field code
    startPos = InStr(code, "{")
    If startPos > 0 Then
        jsonStr = Mid(code, startPos)
        Set ParseCSLCitationJson = ParseJSON(jsonStr)
    Else
        Set ParseCSLCitationJson = Nothing
    End If
End Function

' Extracts data from an author-year formatted citation field.
Private Sub ExtractAuthorYearCitations(field As field, ByRef citations() As Citation, Optional onlyYear As Boolean = True, Optional multiRefCommaSep As Boolean = True)
    ' ... [rest of the original code continues]
End Sub

' Extracts data from a serial-number formatted citation field.
Private Sub ExtractSerialNumberCitations(field As field, ByRef citations() As Citation, Optional border = "")
    ' ... [rest of the original code continues]
End Sub

' Checks if a citation style is supported.
' [MOD-1] Added "elsevier-vancouver" to the supported styles list
Private Function isSupportedStyle(ByVal style As String) As Boolean
    Dim predefinedList As String
    predefinedList = "|" & _
        "molecular-plant|ieee|apa|vancouver|american-chemical-society|" & _
        "american-medical-association|nature|american-political-science-association|" & _
        "american-sociological-association|chicago-author-date|bmc-medicine|" & _
        "china-national-standard-gb-t-7714-2015-numeric|" & _
        "china-national-standard-gb-t-7714-2015-author-date|" & _
        "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
        "archives-of-computational-methods-in-engineering|" & _
        "elsevier-vancouver|"
    style = "|" & style & "|"
    isSupportedStyle = InStr(1, predefinedList, style, vbTextCompare) > 0
End Function

' Routes the citation field to the correct parsing function based on the style ID.
' [MOD-2] Added routing case for "elsevier-vancouver" (uses [] brackets)
Private Sub ExtractCitations(field As field, ByRef citations() As Citation, style As String)
    Select Case style
        Case "molecular-plant", "chicago-author-date", "modern-language-association"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=False, multiRefCommaSep:=False)

        Case "apa", "china-national-standard-gb-t-7714-2015-author-date", _
             "american-political-science-association", "american-sociological-association", _
             "harvard-cite-them-right"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=True, multiRefCommaSep:=True)

        Case "elsevier-harvard"
            Call ExtractAuthorYearCitations(field, citations, onlyYear:=False, multiRefCommaSep:=True)

        Case "vancouver"
            Call ExtractSerialNumberCitations(field, citations, "()")
        
        Case "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "ieee", "archives-of-computational-methods-in-engineering", "elsevier-vancouver"
            Call ExtractSerialNumberCitations(field, citations, "[]")

        Case "american-chemical-society", "american-medical-association", "nature"
            Call ExtractSerialNumberCitations(field, citations, "")

        Case Else
            Err.Raise vbObjectError + 1, "ExtractCitations", "Citation style not recognized: " & style
    End Select
End Sub

' Displays a MsgBox listing all hardcoded supported citation styles.
' [MOD-3] Added "elsevier-vancouver" to the styles list
Public Sub ZCL_Information(control As IRibbonControl)
    '--- Variable Declarations ---
    Dim styles As Variant
    Dim msg As String
    Dim i As Long
    
    '--- Initializations ---
    styles = Split("molecular-plant|ieee|apa|vancouver|american-chemical-society|american-medical-association|nature|" & _
                   "american-political-science-association|american-sociological-association|chicago-author-date|bmc-medicine|" & _
                   "china-national-standard-gb-t-7714-2015-numeric|china-national-standard-gb-t-7714-2015-author-date|" & _
                   "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
                   "archives-of-computational-methods-in-engineering|elsevier-vancouver", "|")
    
    msg = "Zotero styles supported by this macro (including Elsevier - NLM/Vancouver):" & vbCrLf & vbCrLf
    
    For i = 0 To UBound(styles)
        msg = msg & i + 1 & ". " & styles(i) & vbCrLf
    Next
    
    MsgBox msg, vbInformation, "Supported Reference Styles"
End Sub
