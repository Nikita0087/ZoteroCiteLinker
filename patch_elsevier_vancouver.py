#!/usr/bin/env python3
"""
ZoteroCiteLinker Patcher - Elsevier Vancouver Style Support
============================================================

This script patches the ZoteroCiteLinker.dotm file to add support for
the "Elsevier - NLM/Vancouver (citation sequence)" style.

The Elsevier - NLM/Vancouver style in Zotero has the style ID: "elsevier-vancouver"
It is a numeric citation style that uses brackets [1], [2], etc., identical to
how IEEE, BMC Medicine, and other numeric bracket styles work.

Prerequisites:
    - Windows with Microsoft Word installed
    - Python 3.x with pywin32 package: pip install pywin32

Usage:
    python patch_elsevier_vancouver.py [path_to_ZoteroCiteLinker.dotm]

If no path is provided, the script will look for "ZoteroCiteLinker.dotm"
in the current directory.
"""

import sys
import os
import shutil
from datetime import datetime

# Change 1: Add elsevier-vancouver to the supported styles list
OLD_STYLE_LIST = '''    predefinedList = "|" & _
        "molecular-plant|ieee|apa|vancouver|american-chemical-society|" & _
        "american-medical-association|nature|american-political-science-association|" & _
        "american-sociological-association|chicago-author-date|bmc-medicine|" & _
        "china-national-standard-gb-t-7714-2015-numeric|" & _
        "china-national-standard-gb-t-7714-2015-author-date|" & _
        "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
        "archives-of-computational-methods-in-engineering|"'''

NEW_STYLE_LIST = '''    predefinedList = "|" & _
        "molecular-plant|ieee|apa|vancouver|american-chemical-society|" & _
        "american-medical-association|nature|american-political-science-association|" & _
        "american-sociological-association|chicago-author-date|bmc-medicine|" & _
        "china-national-standard-gb-t-7714-2015-numeric|" & _
        "china-national-standard-gb-t-7714-2015-author-date|" & _
        "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
        "archives-of-computational-methods-in-engineering|" & _
        "elsevier-vancouver|"'''

# Change 2: Add routing case for elsevier-vancouver (uses [] brackets)
OLD_CASES = '''        Case "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "ieee", "archives-of-computational-methods-in-engineering"
            Call ExtractSerialNumberCitations(field, citations, "[]")'''

NEW_CASES = '''        Case "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "ieee", "archives-of-computational-methods-in-engineering", "elsevier-vancouver"
            Call ExtractSerialNumberCitations(field, citations, "[]")'''

# Change 3: Add elsevier-vancouver to the information dialog
OLD_INFO_STYLES = '''    styles = Split("molecular-plant|ieee|apa|vancouver|american-chemical-society|american-medical-association|nature|" & _
                   "american-political-science-association|american-sociological-association|chicago-author-date|bmc-medicine|" & _
                   "china-national-standard-gb-t-7714-2015-numeric|china-national-standard-gb-t-7714-2015-author-date|" & _
                   "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
                   "archives-of-computational-methods-in-engineering", "|")'''

NEW_INFO_STYLES = '''    styles = Split("molecular-plant|ieee|apa|vancouver|american-chemical-society|american-medical-association|nature|" & _
                   "american-political-science-association|american-sociological-association|chicago-author-date|bmc-medicine|" & _
                   "china-national-standard-gb-t-7714-2015-numeric|china-national-standard-gb-t-7714-2015-author-date|" & _
                   "harvard-cite-them-right|elsevier-harvard|modern-language-association|" & _
                   "archives-of-computational-methods-in-engineering|elsevier-vancouver", "|")'''


def patch_vba_code(vba_code):
    """Apply all three modifications to the VBA source code."""
    modified = vba_code
    
    # Apply change 1
    if OLD_STYLE_LIST in modified:
        modified = modified.replace(OLD_STYLE_LIST, NEW_STYLE_LIST)
        print("  [OK] Added 'elsevier-vancouver' to supported styles list")
    else:
        if "elsevier-vancouver" in modified:
            print("  [SKIP] 'elsevier-vancouver' already in styles list")
        else:
            print("  [ERROR] Could not find styles list to patch")
            return None
    
    # Apply change 2
    if OLD_CASES in modified:
        modified = modified.replace(OLD_CASES, NEW_CASES)
        print("  [OK] Added 'elsevier-vancouver' routing case (bracket style)")
    else:
        if '"elsevier-vancouver"' in modified and 'ExtractSerialNumberCitations' in modified:
            print("  [SKIP] Routing case already exists")
        else:
            print("  [ERROR] Could not find routing cases to patch")
            return None
    
    # Apply change 3
    if OLD_INFO_STYLES in modified:
        modified = modified.replace(OLD_INFO_STYLES, NEW_INFO_STYLES)
        print("  [OK] Added 'elsevier-vancouver' to info dialog")
    else:
        if "elsevier-vancouver" in modified and 'ZCL_Information' in modified:
            print("  [SKIP] Info dialog already updated")
        else:
            print("  [ERROR] Could not find info styles list to patch")
            return None
    
    return modified


def patch_dotm_file(dotm_path):
    """Patch the .dotm file using Microsoft Word COM automation."""
    
    if not os.path.exists(dotm_path):
        print(f"ERROR: File not found: {dotm_path}")
        return False
    
    # Create backup
    backup_path = dotm_path + ".backup_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(dotm_path, backup_path)
    print(f"[OK] Backup created: {backup_path}")
    
    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 is required. Install it with: pip install pywin32")
        return False
    
    word = None
    doc = None
    
    try:
        print("[INFO] Starting Microsoft Word...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        print(f"[INFO] Opening {os.path.basename(dotm_path)}...")
        doc = word.Documents.Open(os.path.abspath(dotm_path))
        
        print("[INFO] Accessing VBA project...")
        vb_project = doc.VBProject
        
        target_module = None
        for component in vb_project.VBComponents:
            if component.Name == "ZoteroCiteLinker":
                target_module = component
                break
        
        if target_module is None:
            print("ERROR: Could not find 'ZoteroCiteLinker' VBA module")
            return False
        
        print("[INFO] Reading current VBA code...")
        code_module = target_module.CodeModule
        vba_code = code_module.Lines(1, code_module.CountOfLines)
        
        print("[INFO] Patching VBA code...")
        modified_code = patch_vba_code(vba_code)
        
        if modified_code is None:
            print("ERROR: Patching failed")
            return False
        
        if modified_code == vba_code:
            print("[INFO] No changes needed - file already patched")
            return True
        
        print("[INFO] Writing modified VBA code...")
        code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(modified_code)
        
        print("[INFO] Saving file...")
        doc.Save()
        
        print("\n[SUCCESS] Patch applied successfully!")
        print("\nThe Elsevier - NLM/Vancouver (citation sequence) style")
        print("has been added to ZoteroCiteLinker.")
        print("\nStyle details:")
        print("  - Zotero Style ID: elsevier-vancouver")
        print("  - Citation format: Numeric with brackets [1], [2], [3]...")
        print("  - Works identically to IEEE, BMC Medicine, etc.")
        
        return True
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, dotm_path)
            print("[INFO] Restored from backup")
        return False
        
    finally:
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        if word:
            try:
                word.Quit()
            except:
                pass


def main():
    print("=" * 60)
    print("ZoteroCiteLinker Patcher")
    print("Adds Elsevier - NLM/Vancouver (citation sequence) support")
    print("=" * 60)
    print()
    
    if len(sys.argv) > 1:
        dotm_path = sys.argv[1]
    else:
        dotm_path = "ZoteroCiteLinker.dotm"
        if not os.path.exists(dotm_path):
            startup_paths = [
                os.path.expandvars(r"%APPDATA%\Microsoft\Word\STARTUP\ZoteroCiteLinker.dotm"),
                os.path.expandvars(r"%USERPROFILE%\AppData\Roaming\Microsoft\Word\STARTUP\ZoteroCiteLinker.dotm"),
            ]
            for path in startup_paths:
                if os.path.exists(path):
                    dotm_path = path
                    break
    
    print(f"Target file: {os.path.abspath(dotm_path)}")
    print()
    
    success = patch_dotm_file(dotm_path)
    
    if not success:
        print("\n[!] Automatic patching failed.")
        print("[!] You can apply the patch manually using the instructions")
        print("[!] in the README.md file.")
        sys.exit(1)


if __name__ == "__main__":
    main()
