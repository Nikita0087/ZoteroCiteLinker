# ZoteroCiteLinker

An MS Word Add-in that creates an interactive bridge between your in-text Zotero citations and your bibliography, all from a convenient custom Ribbon tab.

This add-in scans your document for Zotero-generated citation fields (e.g., `[1]`, `(Author, 2020)`) and intelligently links them to the corresponding entries in your bibliography, allowing for instant navigation.

> **FORK NOTICE:** This fork adds support for the **Elsevier - NLM/Vancouver (citation sequence)** style (`elsevier-vancouver`). See [Applying the Elsevier-Vancouver Patch](#applying-the-elsevier-vancouver-patch) below.

![ZoteroCiteLinker Ribbon](images/ss.png)

## 🌟 Features

* **One-Click Linking:** Automatically link all citations in the document.
* **Step-by-Step Linking:** A manual "debug" mode to process citations one by one.
* **Smart Unlinking:** Safely remove *only* the links created by this add-in, without touching your Table of Contents or other hyperlinks.
* **Structured Bookmarks:** Generates clean, human-readable bookmarks (e.g., `Cite_id1234_Title_2020_Author`) for easy navigation and management in Word's bookmark dialog.
* **Utility Functions:** Activate (or separately remove) all web URLs and email addresses in your document.
* **Customization:** Change the color of all Zotero citations or apply a custom Word "Character Style" for advanced formatting.
* **Compatibility:** Works with over 18 common Zotero citation styles (listed in the "How to Use" section below), **including Elsevier - NLM/Vancouver**.

## 🆕 Applying the Elsevier-Vancouver Patch

This fork adds support for the **Elsevier - NLM/Vancouver (citation sequence)** citation style (Zotero style ID: `elsevier-vancouver`). This is a numeric style that uses brackets: `[1]`, `[2]`, `[3]`, etc.

### Method 1: Automatic Patching with Python (Recommended)

**Prerequisites:** Windows with Microsoft Word and Python 3.x installed.

1. Install `pywin32` if not already installed:
   ```bash
   pip install pywin32
   ```

2. Download the [`patch_elsevier_vancouver.py`](patch_elsevier_vancouver.py) script from this repository.

3. Run the script, providing the path to your `ZoteroCiteLinker.dotm` file:
   ```bash
   python patch_elsevier_vancouver.py "C:\Path\To\ZoteroCiteLinker.dotm"
   ```
   If you omit the path, the script will look for the file in the current directory and common Word STARTUP locations.

4. The script will:
   - Create a backup of your original `.dotm` file
   - Automatically apply the three necessary code changes
   - Save the patched file

### Method 2: Manual VBA Editing

If you prefer not to use the Python script, you can apply the changes manually:

1. **Open the VBA Editor:**
   - Open Microsoft Word with the ZoteroCiteLinker add-in loaded.
   - Press `Alt + F11` to open the Visual Basic for Applications editor.

2. **Locate the Module:**
   - In the Project Explorer (left panel), find `ZoteroCiteLinker.dotm` → `Modules` → `ZoteroCiteLinker`.
   - Double-click `ZoteroCiteLinker` to open its code.

3. **Apply the 3 Changes:**

   **Change 1/3 - Add style to supported list** (~line 705):
   Find the `isSupportedStyle` function and add `elsevier-vancouver|` to the `predefinedList`:
   ```vba
       predefinedList = "|" & _
           "molecular-plant|ieee|apa|vancouver|american-chemical-society|" & _
           ...
           "archives-of-computational-methods-in-engineering|" & _
           "elsevier-vancouver|"    ' <-- ADD THIS LINE
   ```

   **Change 2/3 - Add routing case** (~line 736):
   Find the `ExtractCitations` Sub and add `elsevier-vancouver` to the bracket-style case:
   ```vba
       Case "china-national-standard-gb-t-7714-2015-numeric", "bmc-medicine", "ieee", "archives-of-computational-methods-in-engineering", "elsevier-vancouver"
           Call ExtractSerialNumberCitations(field, citations, "[]")
   ```

   **Change 3/3 - Add to info dialog** (~line 755):
   Find the `ZCL_Information` Sub and add `elsevier-vancouver` to the styles Split:
   ```vba
       styles = Split("molecular-plant|...|archives-of-computational-methods-in-engineering|elsevier-vancouver", "|")
   ```

4. **Save:**
   - Press `Ctrl + S` to save the VBA module.
   - Close the VBA editor.
   - Restart Word.

### What Changed?

The Elsevier - NLM/Vancouver style is a **numeric bracket style** (`[1]`, `[2]`, etc.), identical in citation format to IEEE, BMC Medicine, and other numeric-bracket styles. The add-in's `ExtractSerialNumberCitations` function already handles this format perfectly — it only needed to recognize the `elsevier-vancouver` style ID and route it to the correct parser.

## 🖼️ Screenshots

<details>
<summary>Click to see the add-in in action (9 images)</summary>
<br>

<table>
  <tr>
    <td align="center"><img src="images/1.png" alt="Screenshot 1" width="280" /></td>
    <td align="center"><img src="images/2.png" alt="Screenshot 2" width="280" /></td>
    <td align="center"><img src="images/3.png" alt="Screenshot 3" width="280" /></td>
  </tr>
  <tr>
    <td align="center"><img src="images/4.png" alt="Screenshot 4" width="280" /></td>
    <td align="center"><img src="images/5.png" alt="Screenshot 5" width="280" /></td>
    <td align="center"><img src="images/6.png" alt="Screenshot 6" width="280" /></td>
  </tr>
  <tr>
    <td align="center"><img src="images/7.png" alt="Screenshot 7" width="280" /></td>
    <td align="center"><img src="images/8.png" alt="Screenshot 8" width="280" /></td>
    <td align="center"><img src="images/9.png" alt="Screenshot 9" width="280" /></td>
  </tr>
</table>

</details>

## 💾 Installation Instructions

This add-in is installed by placing the `.dotm` file into Word's trusted STARTUP folder. This will automatically load the "Zotero Linker" Ribbon tab every time you open Word.

1.  **Get the File:**
    * **[Click here to download `ZoteroCiteLinker.dotm`](https://github.com/sBaydin/ZoteroCiteLinker/raw/main/ZoteroCiteLinker.dotm)**
    *(You may need to right-click the link and select "Save As...")*

2.  **Find your Word STARTUP Folder:**
    * Open MS Word.
    * Go to **File** -> **Options** -> **Trust Center**.
    * Click the **Trust Center Settings...** button.
    * Go to the **Trusted Locations** tab.
    * In the list, find the location with the description "**User STARTUP**".
    * **Click on it and copy the full path** from the bottom of the window.
    * The path will look something like this:
        `C:\Users\<YourUserName>\AppData\Roaming\Microsoft\Word\STARTUP`

3.  **Copy the Add-in File:**
    * Open Windows File Explorer.
    * Paste the path you just copied into the address bar and press **Enter**.
    * Copy the `ZoteroCiteLinker.dotm` file into this `STARTUP` folder.

4.  **Restart Word & Enable Content:**
    * Close MS Word completely and re-open it.
    * When Word starts, it will likely show a yellow **"SECURITY WARNING"** bar at the top, stating "Macros have been disabled."
    * You **must** click the **"Enable Content"** button. This tells Word to trust the add-in.
    * You should now see a new **"Zotero Linker"** tab on your Word Ribbon.

## ⚠️ Important: Backup Your Work!

Before running any macros that modify your document (like "Link All Citations" or "Set Style"), it is **strongly recommended** that you **save a backup copy** of your work.

While this add-in is designed to be safe, a macro-driven error, an unexpected interaction with other add-ins, or a mistake (like applying a Paragraph Style) could lead to unintended changes. A backup ensures you can always revert to a safe version of your document.

## 📖 How to Use: The Ribbon Guide

![ZoteroCiteLinker Ribbon](images/ss.png)

Here is a breakdown of each button on the "Zotero Linker" tab.

### Main Functions

* **Link All Citations (`ZCL_LinkCitationAll`)**
    * This is the main button. It first unlinks all previous citations and then automatically scans and links every Zotero citation in your document. A message will appear showing how many links were created.
    * **Requirement:** The Zotero application must be running for this to work.

* **Link Selectively (`ZCL_LinkCitationSelect`)**
    * This is a debug/manual mode. It will go through the document one citation field at a time and ask you (Yes/No/Cancel) if you want to process that specific group.
    * **Requirement:** The Zotero application must be running.

* **Unlink Citations (`ZCL_UnlinkCitations`)**
    * Safely removes *only* the hyperlinks from your Zotero citations (bookmarks starting with `CITE_`). It also automatically resets their color back to black (Auto).
    * This **does not** touch other hyperlinks in your document (like web links or your Table of Contents).

### Utility Functions

* **Activate URLs (in Bib) (`ZCL_LinkBibliographyURLs`)**
    * This utility scans *only* your Zotero Bibliography section and converts all text URLs (e.g., `http://...`) into active, clickable hyperlinks.

* **Activate Emails (All) (`ZCL_LinkEmails`)**
    * This utility scans your *entire document* (not just the bibliography) and converts all email addresses (e.g., `name@domain.com`) into active `mailto:` links.

* **Unlink Web/Mail (`ZCL_UnlinkURLsMail`)**
    * The reverse of the above. This removes all `http://` and `mailto:` hyperlinks from the entire document.
    * This **does not** touch your Zotero citation links.

### Formatting & Help

* **Change Color (`ZCL_ChangeColor`)**
    * Opens a simple menu asking you to pick a color by number (Blue, Red, Auto/Black, etc.). This will apply your chosen color to all Zotero citation fields in the document.

* **Set Style (`ZCL_SetCitationStyle`)**
    * (Advanced) Allows you to apply a specific Word "Character Style" (e.g., "Emphasis") to your citations for custom formatting. A warning will appear, as applying a "Paragraph Style" by mistake can reformat your document.

* **Info (`ZCL_Information`)**
    * Displays a pop-up box listing all 18 citation styles supported by the add-in's engine, grouped by type:

| Style Type | Proper Name | Zotero Style ID |
| :--- | :--- | :--- |
| **Author-Year** | American Political Science Ass. | `american-political-science-association` |
| **Author-Year** | American Psychological Ass. (APA) | `apa` |
| **Author-Year** | American Sociological Ass. (ASA) | `american-sociological-association` |
| **Author-Year** | Chicago Manual of Style (Author-Date) | `chicago-author-date` |
| **Author-Year** | China National Standard (Author-Date) | `china-national-standard-gb-t-7714-2015-author-date` |
| **Author-Year** | Cite Them Right - Harvard | `harvard-cite-them-right` |
| **Author-Year** | Elsevier - Harvard | `elsevier-harvard` |
| **Author-Year** | Molecular Plant | `molecular-plant` |
| **Author-only** | Modern Language Association (MLA) | `modern-language-association` |
| **Numeric** | American Chemical Society (ACS) | `american-chemical-society` |
| **Numeric** | American Medical Association (AMA) | `american-medical-association` |
| **Numeric** | Archives of Comp. Methods in Eng. | `archives-of-computational-methods-in-engineering` |
| **Numeric** | BMC Medicine | `bmc-medicine` |
| **Numeric** | China National Standard (Numeric) | `china-national-standard-gb-t-7714-2015-numeric` |
| **Numeric** | Elsevier - NLM/Vancouver | `elsevier-vancouver` |
| **Numeric** | IEEE | `ieee` |
| **Numeric** | Nature | `nature` |
| **Numeric** | Vancouver | `vancouver` |

## 🤝 Acknowledgments

This add-in was built by modifying the excellent foundational code from:
* [altairwei/ZoteroLinkCitation](https://github.com/altairwei/ZoteroLinkCitation)
* [8gengen8/ZoteroCitationLink-lgg](https://github.com/8gengen8/ZoteroCitationLink-lgg)

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
