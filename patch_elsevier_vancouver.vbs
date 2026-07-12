' ============================================================================
' ZoteroCiteLinker Patcher - Elsevier Vancouver Style Support
' ============================================================================
' This VBScript patches the ZoteroCiteLinker.dotm file to add support for
' the "Elsevier - NLM/Vancouver (citation sequence)" style.
'
' Usage: Double-click this file, or run from command line:
'        cscript patch_elsevier_vancouver.vbs [path_to_ZoteroCiteLinker.dotm]
'
' No additional dependencies required - uses Microsoft Word COM automation.
' ============================================================================

Option Explicit

Dim fso, WshShell, dotmPath, backupPath, wordApp, doc, vbProj, vbComp
Dim codeModule, vbaCode, modifiedCode, success
Dim ts, logFile

success = False

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Determine the target .dotm file path
If WScript.Arguments.Count > 0 Then
    dotmPath = WScript.Arguments(0)
Else
    dotmPath = fso.GetAbsolutePathName("ZoteroCiteLinker.dotm")
    If Not fso.FileExists(dotmPath) Then
        ' Try common Word STARTUP locations
        Dim appData, startupPath
        appData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
        startupPath = appData & "\Microsoft\Word\STARTUP\ZoteroCiteLinker.dotm"
        If fso.FileExists(startupPath) Then
            dotmPath = startupPath
        End If
    End If
End If

WScript.Echo "============================================================"
WScript.Echo "ZoteroCiteLinker Patcher"
WScript.Echo "Adds Elsevier - NLM/Vancouver (citation sequence) support"
WScript.Echo "============================================================"
WScript.Echo ""
WScript.Echo "Target file: " & dotmPath
WScript.Echo ""

' Verify file exists
If Not fso.FileExists(dotmPath) Then
    WScript.Echo "ERROR: File not found: " & dotmPath
    WScript.Echo ""
    WScript.Echo "Usage: cscript patch_elsevier_vancouver.vbs [path_to_ZoteroCiteLinker.dotm]"
    WScript.Quit 1
End If

' Create backup
Dim timestamp
timestamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
            Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
backupPath = dotmPath & ".backup_" & timestamp
fso.CopyFile dotmPath, backupPath
WScript.Echo "[OK] Backup created: " & backupPath

On Error Resume Next

' Start Word
WScript.Echo "[INFO] Starting Microsoft Word..."
Set wordApp = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not start Microsoft Word. Is it installed?"
    WScript.Quit 1
End If

wordApp.Visible = False
wordApp.DisplayAlerts = 0 ' wdAlertsNone

' Open the .dotm file
WScript.Echo "[INFO] Opening " & fso.GetFileName(dotmPath) & "..."
Set doc = wordApp.Documents.Open(dotmPath)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Could not open the file: " & Err.Description
    wordApp.Quit
    WScript.Quit 1
End If

' Access VBA project
WScript.Echo "[INFO] Accessing VBA project..."
Set vbProj = doc.VBProject

' Find the ZoteroCiteLinker module
Set vbComp = Nothing
Dim comp
For Each comp In vbProj.VBComponents
    If comp.Name = "ZoteroCiteLinker" Then
        Set vbComp = comp
        Exit For
    End If
Next

If vbComp Is Nothing Then
    WScript.Echo "ERROR: Could not find 'ZoteroCiteLinker' VBA module"
    doc.Close False
    wordApp.Quit
    WScript.Quit 1
End If

' Read current code
WScript.Echo "[INFO] Reading current VBA code..."
Set codeModule = vbComp.CodeModule
vbaCode = codeModule.Lines(1, codeModule.CountOfLines)

' Apply patches
WScript.Echo "[INFO] Patching VBA code..."
modifiedCode = vbaCode

' Change 1: Add elsevier-vancouver to supported styles list
Dim oldStyleList, newStyleList
oldStyleList = "    predefinedList = ""|"" & _" & vbCrLf & _
    "        ""molecular-plant|ieee|apa|vancouver|american-chemical-society|"" & _" & vbCrLf & _
    "        ""american-medical-association|nature|american-political-science-association|"" & _" & vbCrLf & _
    "        ""american-sociological-association|chicago-author-date|bmc-medicine|"" & _" & vbCrLf & _
    "        ""china-national-standard-gb-t-7714-2015-numeric|"" & _" & vbCrLf & _
    "        ""china-national-standard-gb-t-7714-2015-author-date|"" & _" & vbCrLf & _
    "        ""harvard-cite-them-right|elsevier-harvard|modern-language-association|"" & _" & vbCrLf & _
    "        ""archives-of-computational-methods-in-engineering|"""

newStyleList = "    predefinedList = ""|"" & _" & vbCrLf & _
    "        ""molecular-plant|ieee|apa|vancouver|american-chemical-society|"" & _" & vbCrLf & _
    "        ""american-medical-association|nature|american-political-science-association|"" & _" & vbCrLf & _
    "        ""american-sociological-association|chicago-author-date|bmc-medicine|"" & _" & vbCrLf & _
    "        ""china-national-standard-gb-t-7714-2015-numeric|"" & _" & vbCrLf & _
    "        ""china-national-standard-gb-t-7714-2015-author-date|"" & _" & vbCrLf & _
    "        ""harvard-cite-them-right|elsevier-harvard|modern-language-association|"" & _" & vbCrLf & _
    "        ""archives-of-computational-methods-in-engineering|"" & _" & vbCrLf & _
    "        ""elsevier-vancouver|"""

If InStr(modifiedCode, oldStyleList) > 0 Then
    modifiedCode = Replace(modifiedCode, oldStyleList, newStyleList)
    WScript.Echo "  [OK] Added 'elsevier-vancouver' to supported styles list"
ElseIf InStr(modifiedCode, "elsevier-vancouver") > 0 Then
    WScript.Echo "  [SKIP] 'elsevier-vancouver' already in styles list"
Else
    WScript.Echo "  [ERROR] Could not find styles list to patch"
End If

' Change 2: Add routing case
Dim oldCases, newCases
oldCases = "        Case ""china-national-standard-gb-t-7714-2015-numeric"", ""bmc-medicine"", ""ieee"", ""archives-of-computational-methods-in-engineering""" & vbCrLf & _
    "            Call ExtractSerialNumberCitations(field, citations, ""[]"")"

newCases = "        Case ""china-national-standard-gb-t-7714-2015-numeric"", ""bmc-medicine"", ""ieee"", ""archives-of-computational-methods-in-engineering"", ""elsevier-vancouver""" & vbCrLf & _
    "            Call ExtractSerialNumberCitations(field, citations, ""[]"")"

If InStr(modifiedCode, oldCases) > 0 Then
    modifiedCode = Replace(modifiedCode, oldCases, newCases)
    WScript.Echo "  [OK] Added 'elsevier-vancouver' routing case (bracket style)"
ElseIf InStr(modifiedCode, """elsevier-vancouver""") > 0 And InStr(modifiedCode, "ExtractSerialNumberCitations") > 0 Then
    WScript.Echo "  [SKIP] Routing case already exists"
Else
    WScript.Echo "  [ERROR] Could not find routing cases to patch"
End If

' Change 3: Add to info dialog
Dim oldInfo, newInfo
oldInfo = "    styles = Split(""molecular-plant|ieee|apa|vancouver|american-chemical-society|american-medical-association|nature|"" & _" & vbCrLf & _
    "                   ""american-political-science-association|american-sociological-association|chicago-author-date|bmc-medicine|"" & _" & vbCrLf & _
    "                   ""china-national-standard-gb-t-7714-2015-numeric|china-national-standard-gb-t-7714-2015-author-date|"" & _" & vbCrLf & _
    "                   ""harvard-cite-them-right|elsevier-harvard|modern-language-association|"" & _" & vbCrLf & _
    "                   ""archives-of-computational-methods-in-engineering"", ""|")"

newInfo = "    styles = Split(""molecular-plant|ieee|apa|vancouver|american-chemical-society|american-medical-association|nature|"" & _" & vbCrLf & _
    "                   ""american-political-science-association|american-sociological-association|chicago-author-date|bmc-medicine|"" & _" & vbCrLf & _
    "                   ""china-national-standard-gb-t-7714-2015-numeric|china-national-standard-gb-t-7714-2015-author-date|"" & _" & vbCrLf & _
    "                   ""harvard-cite-them-right|elsevier-harvard|modern-language-association|"" & _" & vbCrLf & _
    "                   ""archives-of-computational-methods-in-engineering|elsevier-vancouver"", ""|")"

If InStr(modifiedCode, oldInfo) > 0 Then
    modifiedCode = Replace(modifiedCode, oldInfo, newInfo)
    WScript.Echo "  [OK] Added 'elsevier-vancouver' to info dialog"
ElseIf InStr(modifiedCode, "elsevier-vancouver") > 0 And InStr(modifiedCode, "ZCL_Information") > 0 Then
    WScript.Echo "  [SKIP] Info dialog already updated"
Else
    WScript.Echo "  [ERROR] Could not find info styles list to patch"
End If

' Write modified code
If modifiedCode <> vbaCode Then
    WScript.Echo "[INFO] Writing modified VBA code..."
    codeModule.DeleteLines 1, codeModule.CountOfLines
    codeModule.AddFromString modifiedCode
    
    WScript.Echo "[INFO] Saving file..."
    doc.Save
    
    WScript.Echo ""
    WScript.Echo "============================================================"
    WScript.Echo "SUCCESS! Patch applied successfully!"
    WScript.Echo "============================================================"
    WScript.Echo ""
    WScript.Echo "The Elsevier - NLM/Vancouver (citation sequence) style"
    WScript.Echo "has been added to ZoteroCiteLinker."
    WScript.Echo ""
    WScript.Echo "Style details:"
    WScript.Echo "  - Zotero Style ID: elsevier-vancouver"
    WScript.Echo "  - Citation format: Numeric with brackets [1], [2], [3]..."
    WScript.Echo "  - Works identically to IEEE, BMC Medicine, etc."
    success = True
Else
    WScript.Echo ""
    WScript.Echo "[INFO] No changes needed - file already patched or no changes found."
    success = True
End If

' Cleanup
doc.Close False
wordApp.Quit

Set doc = Nothing
Set wordApp = Nothing
Set vbComp = Nothing
Set codeModule = Nothing
Set vbProj = Nothing
Set fso = Nothing
Set WshShell = Nothing

If Not success Then
    WScript.Echo ""
    WScript.Echo "Patching failed. Please try the manual method described in README.md"
    WScript.Quit 1
End If
