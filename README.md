# 🏥 Project Zense — Oracle SCM Asset Reconciliation RPA Bot

**Developer:** Abdul Hazeer (IT Support & RPA Intern — Zulekha Hospital)  
**Tech Stack:** UiPath Studio | VB.NET | LINQ | Microsoft Excel (`.xlsm` / `.xlsx`) | Google Gemini Argon  

---

## 📌 Project Overview
Developed during **Project Zense** (Zulekha Hospital's enterprise Oracle Supply Chain Management Cloud integration), this UiPath RPA bot automates the end-to-end reconciliation of **9,620 Installed Base Asset records** across physical RFID audit workbooks and Oracle FBDI (File-Based Data Import) macro-enabled templates.

Instead of spending days executing manual `VLOOKUP`/`XLOOKUP` formulas across 118-column macro workbooks, this workflow uses **in-memory `Dictionary(Of String, String)` lookups** and a **single batch write** to complete the 3-step reconciliation in under **5 seconds**.

---

## 🔄 3-Step Cascading Reconciliation Logic

| Step | Source Workbook (`ZH_ASSETTAGGINGDATA.xlsx` - `IT` Sheet) | Target Workbook (`ZH_IT_CseInstalledBaseAssetImport.xlsm` - `Assets` Sheet) | Action Taken | Records Processed |
| :--- | :--- | :--- | :--- | :--- |
| **Step 1: Primary Match (Serial Number)** | **Column I** (`Serial Number`) | **Column R** (`SERIAL_NUMBER`) | Matches trimmed, uppercase serial numbers (filtering out `NIL`, `No Access`, `NONE`) and assigns **Column A (`RFID Tag`)** to **Column A (`Technical Name`)**. | **2,838 matched** (`29.50%`) |
| **Step 2: Secondary Fallback (Asset Number)** | **Column N** (`Asset Number`) | **Column E** (`ASSET_NUMBER`) | If Serial Number does not match, checks `Asset Number` against `ASSET_NUMBER` and assigns **Column A (`RFID Tag`)** to **Column A (`Technical Name`)**. | **275 newly matched** (`2.86%`) *(715 total)* |
| **Step 3: Default Unmatched Flag** | *N/A (Fallback Rule)* | **Column A** (`Technical Name`) | Populates all remaining blank cells in **Column A (`A6:A9625`)** with `"Not Assigned"`. | **6,507 flagged** (`67.64%`) |
| **Total Reconciled** | **6,242 audit rows** | **9,620 asset rows (`A6:A9625`)** | **3,113 Total RFID Tags Assigned (`32.36%`) + 6,507 `"Not Assigned"` (`67.64%`)** | **9,620 rows (`100%`)** |

---

## 🛠️ Step-by-Step Procedure: How the `Main.xaml` Workflow Operates

### Step 0: Dynamic File Verification
* Checks if `ZH_ASSETTAGGINGDATA_22-10-2026 (003).xlsx` and `ZH_IT_CseInstalledBaseAssetImport_Floor update.xlsm` exist in the project root folder.
* If either file is located in another directory, a **`Select File`** dialog automatically prompts the user to select the workbook.

### Step 1: Read Source Tagging Data & Build In-Memory Dictionaries
* Uses **`Read Range (Workbook)`** to load the `IT` sheet from `ZH_ASSETTAGGINGDATA_22-10-2026 (003).xlsx` into `dt_Tagging`.
* Executes two **VB.NET LINQ `Assign`** statements to build case-insensitive lookup dictionaries while ignoring invalid placeholders (`"NIL"`, `"NO ACCESS"`, `"NONE"`, `"NAN"`) and deduplicating keys via `.GroupBy()`:
  * `dict_SerialToRFID`: Maps `Serial Number` (Column I) $\rightarrow$ `RFID Tag` (Column A).
  * `dict_AssetNumToRFID`: Maps `Asset Number` (Column N) $\rightarrow$ `RFID Tag` (Column A).

### Step 2: Read Oracle Import Template (`Range = "A5"`)
* Opens `ZH_IT_CseInstalledBaseAssetImport_Floor update.xlsm` inside **`Excel Application Scope`** (`Visible = False`) to preserve embedded Oracle VBA macros (`vbaProject.bin`) and sheet formatting.
* Reads the `Assets` sheet starting at cell **`A5`** (skipping rows 1–4 of Oracle metadata descriptions and `VARCHAR2` data type definitions) into `dt_Assets`.

### Step 3: Execute 3-Step Cascading Evaluation (`Invoke Code`)
* Iterates through all 9,620 asset rows in memory:
  1. Checks `dict_SerialToRFID` using `row("SERIAL_NUMBER")`.
  2. Else checks `dict_AssetNumToRFID` using `row("ASSET_NUMBER")`.
  3. Else assigns `"Not Assigned"`.
* Populates a single-column DataTable (`dt_OutputColA`) with the resolved result for each row.

### Step 4: Batch Write to Cell `A6` & Save
* Uses **`Excel Write Range`** (`StartingCell = "A6"`, `AddHeaders = False`) to write all 9,620 values into Column A (`Technical Name`) in a single operation without touching Columns B through DN.
* Saves the `.xlsm` workbook and prints the complete match breakdown to the UiPath Output Panel:
  ```text
  3-Step Merge Complete! | Step 1 (Serial Number Matches): 2838 | Step 2 (Asset Number Fallback Matches): 275 | Total Assigned: 3113 | Step 3 ('Not Assigned' Filled): 6507
