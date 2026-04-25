# AI Training – Intermediate with Practice Lab
## Exercise Workbook for FP&A Analysts
### Copilot M365 (Free Tier) – No Premium License Required

---

**Important: How Copilot Works Without Premium**

Without a premium license, you do NOT have Copilot embedded inside Excel, Power Automate, or VBA editor. Instead, you use **Microsoft Copilot** (copilot.microsoft.com).

**Your workflow pattern for every exercise:**
1. Prepare your data / problem in Excel
2. Copy relevant data or describe the problem to Copilot chat
3. Copilot generates formulas, code, analysis, or workflows
4. You apply the output back in Excel / Power Automate / VBA

**Access point:** Open [copilot.microsoft.com](https://copilot.microsoft.com) and sign in with your work account.

---

## Module I – Copilot in Excel: FP&A Use Cases and Exercises

### Exercise 1.1 – Data Shaping: Messy Actuals Clean-Up

**Scenario**
You received a raw data export from SAP containing 6 months of actual costs across 12 cost centers. The data is messy: dates are inconsistent (mixed formats: "Jan-24", "2024/01/01", "01.2024"), cost center names have trailing spaces and inconsistent capitalization, and some amount fields contain text like "N/A" or are blank. You need clean, pivot-ready data.

**What you will learn**
How to use Copilot to generate Excel formulas for data cleaning at scale – parsing dates, standardizing text, handling errors – instead of doing it manually cell by cell.

**Setup (before the exercise)**
Open the provided file `Exercise_1.1_Raw_Actuals.xlsx`. The dataset has columns: `Date`, `Cost_Center`, `Account`, `Description`, `Amount_LC`, `Currency`, `Amount_GC`.

**Steps**

1. **Identify the data quality issues.** Scan the first 20 rows. Note 3–4 specific problems you see (inconsistent dates, extra spaces, "N/A" values, missing amounts).

2. **Ask Copilot to build a date parser.** Copy 5–6 example date values from your dataset and paste them into Copilot with this prompt:

   > I have a column in Excel with dates in mixed formats. Here are examples: "Jan-24", "2024/01/01", "01.2024", "January 2024", "2024-01-31". Write me a single Excel formula that converts any of these into a proper Excel date (last day of the month). The raw date is in cell A2.

3. **Apply the formula.** Paste the formula Copilot gives you into a new column (e.g., H2). Test it against your sample values. If it fails on any format, copy the error back to Copilot:

   > This formula returns #VALUE! for the value "01.2024". Fix it to handle this format too.

4. **Ask Copilot for a text standardization formula.** Copy 5–6 cost center values showing the inconsistencies and prompt:

   > These cost center names need to be standardized in Excel: "  Finance_EMEA ", "finance_emea", "Finance EMEA", "FINANCE-EMEA". Write a formula that trims whitespace, replaces underscores and hyphens with spaces, and converts to proper case. Raw value is in B2.

5. **Ask Copilot for an error-handling amount formula.** Prompt:

   > Column E contains numeric amounts, but some cells have "N/A", "TBD", or are blank. Write an Excel formula that returns the numeric value if it's a number, 0 if blank, and flags "CHECK" if it contains text. Raw value is in E2.

6. **Build the clean dataset.** Apply all three formulas down your data. Create a new sheet called "Clean_Data" and paste values only (Ctrl+Alt+V → select "Values" → OK).

**Expected outcome**
A clean, pivot-ready dataset with standardized dates, consistent cost center names, and validated amounts. You should have 3 reusable formulas you can apply to any future SAP export.

**Debrief questions**
- How much time would manual cleaning take vs. formula-based cleaning?
- Which formula did Copilot get right on the first try? Which needed iteration?
- Could you save these formulas as a template for monthly data loads?

---

### Exercise 1.2 – Formula Engineering: Variance Bridge with Verification

**Scenario**
You need to build a month-over-month cost variance analysis for the CFO. The bridge must decompose total variance into: Volume effect, Price/Rate effect, Mix effect, and Residual. You also need a verification check to confirm the components sum back to total variance.

**What you will learn**
How to use Copilot to construct complex, multi-step Excel formulas, build verification logic, and troubleshoot formula errors – all through iterative prompting.

**Setup**
Open `Exercise_1.2_Variance_Data.xlsx`. You have two tabs: `Jan_Actuals` and `Feb_Actuals`, each with columns: `Product`, `Volume`, `Unit_Price`, `Total_Revenue`.

**Steps**

1. **Describe the variance framework to Copilot.** Open Copilot and provide context:

   > I'm building a revenue variance bridge in Excel comparing January to February. I have two sheets with columns: Product, Volume, Unit_Price, Total_Revenue.
   >
   > I need formulas for:
   > - Volume Effect = (Feb Volume – Jan Volume) × Jan Unit Price
   > - Price Effect = (Feb Unit Price – Jan Unit Price) × Jan Volume
   > - Mix Effect = (Feb Volume – Jan Volume) × (Feb Unit Price – Jan Unit Price)
   > - Total Variance = Feb Total Revenue – Jan Total Revenue
   > - Verification = Volume Effect + Price Effect + Mix Effect (should equal Total Variance)
   >
   > Product name is in cell A2 on both sheets. Write me Excel formulas using VLOOKUP or INDEX/MATCH to pull values cross-sheet. Jan sheet is named "Jan_Actuals", Feb sheet is named "Feb_Actuals".

2. **Apply the formulas.** Create a new sheet called "Variance_Bridge". Set up headers in row 1: Product | Jan Revenue | Feb Revenue | Total Variance | Volume Effect | Price Effect | Mix Effect | Verification | Check.

3. **Build the verification check.** Ask Copilot:

   > Write an Excel formula that compares the sum of Volume + Price + Mix effects to Total Variance and returns "OK" if they match (within 0.01 rounding tolerance) or "ERROR – diff: [amount]" if they don't. Values are in cells D2, E2, F2, G2.

4. **Stress-test with edge cases.** Run two tests: (a) set one product's Feb Volume to 0 and another's Feb Unit_Price to 0 — your decomposition formulas should handle these with no error (zero values are valid inputs); (b) temporarily rename one product in `Feb_Actuals` (e.g., add a trailing space or change a letter) to simulate a lookup mismatch — your VLOOKUP will now return `#N/A` for that product. Once you've confirmed the failure, paste the error into Copilot:

   > My VLOOKUP returns #N/A when a product name in Feb_Actuals doesn't exactly match Jan_Actuals. How do I wrap this formula to return 0 instead of an error? Current formula: [paste formula]

   Restore the original product name before moving on.

5. **Add conditional formatting.** Ask Copilot:

   > What conditional formatting rules should I apply in Excel to highlight: (a) negative variances in red, (b) verification errors in yellow, (c) the largest absolute variance in bold? Give me step-by-step instructions.

**Expected outcome**
A complete variance bridge with cross-sheet lookups, three decomposition components, a verification column, and conditional formatting. Every row should show "OK" in the check column.

**Debrief questions**
- Did the decomposition sum exactly to total variance, or was there a rounding issue?
- How would you adapt this for a Budget vs. Actual bridge?
- What happens if a new product appears in February that wasn't in January?

---

### Exercise 1.3 – Analysis Workflow: Variance Story + Driver Hypotheses

**Scenario**
The CFO asks: "Why did EMEA operating costs increase 18% month-over-month?" You have the data. You need to build the story: which cost lines drove the increase, what are plausible drivers, and what should management investigate.

**What you will learn**
How to use Copilot as an analytical thinking partner – not just for formulas, but for structuring variance narratives, generating hypotheses, and drafting executive commentary.

**Setup**
Open `Exercise_1.3_EMEA_OpCosts.xlsx`. You have columns: `Cost_Line`, `Category`, `Jan_Actual`, `Feb_Actual`, `Budget_Feb`, `Variance_MoM`, `Variance_MoM_%`, `Variance_vs_Budget`, `Variance_vs_Budget_%`.

**Steps**

1. **Identify the top drivers.** Sort by Variance_MoM_% descending. Copy the top 10 rows (headers + data) and paste into Copilot:

   > Here is EMEA operating cost data comparing January to February. The total increased 18% MoM. Analyze the top 10 cost line variances below and tell me:
   > 1. Which 3 cost lines contribute most to the 18% increase (in absolute EUR terms)?
   > 2. Which cost lines are over budget AND increased MoM (double red flags)?
   > 3. Are there any cost lines that decreased – potential positive stories?
   >
   > [paste data]

2. **Generate driver hypotheses.** For the top 3 drivers identified, ask:

   > For these three cost line increases in an SSC/GBS context, generate 2–3 plausible business hypotheses for each:
   > 1. Personnel costs +18% MoM (+€333K)
   > 2. IT infrastructure +45% MoM (+€185K)
   > 3. Travel & entertainment +59% MoM (+€121K)
   >
   > For each hypothesis, suggest what data I should check to confirm or rule it out.

3. **Draft the executive commentary.** Ask Copilot:

   > Write a 4–5 sentence executive summary for the CFO explaining the EMEA OpCost increase. Use the following structure:
   > - Headline: total variance in EUR and %
   > - Top 2 drivers with brief explanation
   > - One positive offset
   > - Recommended action / next step
   >
   > Keep it factual, no speculation. Use business language appropriate for a CFO monthly review.

4. **Create a "Questions to Investigate" list.** Ask Copilot:

   > Based on this variance analysis, list 5 specific questions I should bring to the next cost review meeting with EMEA operations. Make them specific enough that the cost center owner can answer with data, not opinions.

5. **Build the formula for a Pareto view.** Ask Copilot:

   > Write an Excel formula that calculates cumulative % contribution to total variance, so I can create a Pareto chart. Absolute variance is in column F, sorted descending. The formula should show what % of total variance is explained by this line and all lines above it. Cell F2 has the first value.

**Expected outcome**
A structured variance analysis package: top drivers identified, hypotheses generated, executive commentary drafted, investigation questions prepared, and Pareto data calculated. This is a complete "analysis-ready" output you can bring to the CFO review.

**Debrief questions**
- How would you validate or challenge the hypotheses Copilot generated?
- Would you use the executive commentary as-is, or would you edit it? What would you change?
- How does this workflow compare to how you currently prepare variance commentary?

---

## Module II – Automations & Workflows with Power Automate + Python

### Exercise 2.1 – Workflow Design: Monthly Close Reminder Automation

**Scenario**
Every month-end close, you manually send reminder emails to 8 cost center owners about submission deadlines, then follow up with those who miss them. You want to automate this in Power Automate using a simple scheduled flow.

**What you will learn**
How to use Copilot to design a Power Automate flow step-by-step, generate the logic, and translate business requirements into automation specifications – even without Power Automate premium connectors.

**Setup (before the exercise)**
Open the provided file `Exercise_2.1_CostCenter_Tracker.xlsx`. It has three sheets:
- **Cost_Center_Owners** – your 8 cost center owners with names, emails, and region
- **Submission_Tracker** – the monthly tracking table with columns: `Month`, `Cost_Center_Code`, `Cost_Center_Name`, `Owner_Email`, `Reminder_Sent`, `Reminder_Sent_Date`, `Submitted` (Yes/No), `Submission_Date`, `Notes`
- **READ_ME** – instructions for how each sheet is used in this exercise

Review both data sheets before starting. You will paste column structures and sample rows into your Copilot prompts to make the automation design concrete.

**Steps**

1. **Describe the current process to Copilot.** Open Copilot and prompt:

   > I need to automate a monthly close reminder process in Power Automate (standard license, no premium connectors). Here's the current manual process:
   >
   > - On the 3rd business day before month-end, I send an email to 8 cost center owners reminding them to submit their accruals
   > - The email includes: deadline date, link to the submission template, their specific cost center name
   > - On the 1st business day after month-end, I check who submitted and send a follow-up to those who haven't
   >
   > The recipient list is stored in an Excel file on SharePoint with these columns:
   > Cost_Center_Code | Cost_Center_Name | Owner_First_Name | Owner_Last_Name | Owner_Email | Region | Reminder_Active
   >
   > Design this as a Power Automate flow. Tell me:
   > 1. What trigger should I use?
   > 2. What actions do I need (step by step)?
   > 3. How do I read the recipient list from this Excel table (without premium connectors like Dataverse)?
   > 4. How do I calculate "3rd business day before month-end" in Power Automate?

2. **Get the date calculation logic.** The business day calculation is the hardest part. Ask Copilot:

   > Write a Power Automate expression that calculates the 3rd business day before the last day of the current month. Assume weekdays only (no holiday calendar). Show me the expression step by step.

3. **Design the email template.** Ask Copilot:

   > Write an HTML email template for a month-end close reminder. It should include:
   > - Greeting with the cost center owner's first name (use a placeholder)
   > - The submission deadline date (placeholder)
   > - A link to the submission template (placeholder URL)
   > - The cost center name (placeholder)
   > - A professional but friendly tone
   > - Keep it under 150 words

4. **Map out the follow-up logic.** Open the `Submission_Tracker` sheet in `Exercise_2.1_CostCenter_Tracker.xlsx` and copy its header row. Then ask Copilot:

   > Now design the follow-up flow. On business day 1 of the new month:
   > 1. I have an Excel table on SharePoint called "Submission_Tracker" with these columns:
   >    Month | Cost_Center_Code | Cost_Center_Name | Owner_Email | Reminder_Sent | Reminder_Sent_Date | Submitted | Submission_Date | Notes
   > 2. The flow should read this Excel table and filter for rows where Month = current month AND Submitted = "No"
   > 3. Send a follow-up email to each Owner_Email in the filtered results, including their Cost_Center_Name in the email body
   >
   > Give me the Power Automate steps using standard connectors (Excel Online, Outlook, SharePoint). Show how to filter the table rows in Power Automate expressions.

5. **Document the flow.** Ask Copilot:

   > Create a simple process documentation table for this automation with columns: Step #, Action, Trigger/Condition, Input, Output, Error Handling. Cover both the reminder flow and the follow-up flow.

**Expected outcome**
A fully documented Power Automate design with trigger logic, date calculations, email templates, follow-up conditions, and a process documentation table. You don't build the flow during the exercise – you build the spec that you (or IT) can implement.

**Debrief questions**
- What's the biggest risk in this automation? (Hint: what if the Excel file isn't updated?)
- Could you extend this to track submission status over multiple months?
- What would you need premium connectors for, and is it worth it?

---

### Exercise 2.2 – Power Automate Exercise: Automated Report Distribution

**Scenario**
Every Monday morning, your reporting tool automatically drops one regional P&L PDF per region into a shared SharePoint folder — already split and named by a consistent convention. Your manual task is then to open each file, identify the right recipient, compose individual emails, attach the correct PDF, and CC the CFO. This currently takes 45 minutes. You want to design a Power Automate flow that reads a distribution list and routes the correct PDF to each recipient automatically.

**What you will learn**
How to use Copilot to design a conditional distribution workflow, handle dynamic SharePoint file paths in Power Automate, and build personalized emails based on recipient attributes – all with standard connectors only.

**Setup (before the exercise)**
Open `Exercise_2.2_Distribution_Matrix.xlsx`. It has three sheets:
- **Distribution_List** – 5 active recipients (+ 1 inactive), each mapped to a region. This is the Excel table Power Automate will loop through.
- **PDF_File_Inventory** – shows the file naming convention your reporting tool uses: `PL_Summary_[REGION]_Week[WW]_[YYYY].pdf`. The `Region_Filter_Value` column must match exactly between both sheets.
- **READ_ME** – instructions for how each sheet maps to each exercise step.

Review both data sheets before starting. You will paste column headers and sample rows into Copilot to make the automation design grounded and specific.

**Steps**

1. **Map the current process.** Open the `Distribution_List` sheet and copy the header row. Then ask Copilot:

   > I manually distribute regional P&L reports every Monday. My reporting tool drops pre-built PDFs into a SharePoint folder each Monday morning using this naming pattern: PL_Summary_[REGION]_Week[WeekNumber]_[YYYY].pdf (e.g. PL_Summary_EMEA_Week15_2026.pdf).
   >
   > My manual steps are:
   > 1. Open a distribution list Excel table on SharePoint (columns listed below)
   > 2. For each active recipient, find their regional PDF in the SharePoint folder by name
   > 3. Send an email with that PDF attached, subject: "[Region] P&L Summary – Week of [date]"
   > 4. CC the CFO on every email
   >
   > Distribution list columns: Recipient_Name | Recipient_Email | Region | Region_Filter_Value | CC_Email | Report_Type | Active
   >
   > Design a Power Automate flow using standard connectors (Excel Online, SharePoint, Outlook) that automates steps 1–4. Tell me: what trigger to use, how to loop through the distribution list, how to filter for Active = "Yes" only, how to build the SharePoint file path dynamically, and how to attach the file to each email.

2. **Understand how Power Automate reads the distribution table.** Open `Distribution_List`, copy the header row and the first 3 data rows, and paste them into Copilot:

   > Here is the distribution list table that Power Automate will read using the "List rows present in a table" action. The table has an Active column — inactive recipients (Active = "No") should be skipped.
   >
   > [paste headers + 3 rows]
   >
   > For this table structure, explain:
   > 1. How does Power Automate filter rows where Active = "Yes" before looping?
   > 2. Inside the "Apply to each" loop, how do I reference Recipient_Email, Region_Filter_Value, and CC_Email as dynamic values?
   > 3. How do I construct the SharePoint file path using Region_Filter_Value, the ISO week number, and the current year? Show me the Power Automate expression.
   > 4. What action do I use to get the file content from SharePoint so I can attach it to the email?

3. **Handle the file naming logic.** Ask Copilot:

   > Write a Power Automate expression for dynamic file naming. The pattern should be: "PL_Summary_[REGION]_Week[WeekNumber]_[YYYY].pdf". For example: "PL_Summary_EMEA_Week15_2026.pdf". Show me the expression using Power Automate's built-in functions.

4. **Design the error notification.** Ask Copilot:

   > Add error handling to this flow. If any email fails to send or the source file is missing, the flow should:
   > 1. Continue processing remaining recipients (don't stop)
   > 2. Collect all errors
   > 3. Send me a summary email at the end listing which deliveries failed and why
   >
   > How do I implement this in Power Automate with standard connectors? Show me the scope/try-catch pattern.

5. **Calculate ROI.** Ask Copilot:

   > I spend 45 minutes every Monday on manual report distribution. Help me build a simple ROI calculation for automating this:
   > - My loaded hourly cost: €75
   > - Estimated setup time: 4 hours
   > - Estimated monthly maintenance: 15 minutes
   > - Calculate: annual hours saved, annual cost saved, payback period in weeks

**Expected outcome**
A complete automation specification: flow design, distribution matrix, file naming logic, error handling pattern, and ROI justification. This is ready to implement or present to your manager as a business case.

**Debrief questions**
- What happens if someone new joins and needs the report? How easy is maintenance?
- Could you add a confirmation step where recipients acknowledge receipt?
- What other reports in your team follow a similar distribution pattern?

---

### Exercise 2.3 – Python as a "Glue Layer": Multi-Source Data Consolidation

**Scenario**
You receive data from three different sources monthly: an SAP export (CSV), a manual Excel tracker from operations, and a SharePoint list export. You need to consolidate them into one analysis-ready dataset. The structures are different, naming conventions don't match, and you do this manually in Excel every month.

**What you will learn**
How to use Copilot to generate a Python script that automates multi-source data consolidation – even if you've never written Python before. You won't run the script during the exercise; you'll build and understand it.

**Setup (before the exercise)**
Open and review all three source files – they contain March 2026 actuals for the same 5 cost centers, named differently in each system:
- `Exercise_2.3_SAP_Export.csv` – 12 rows, SAP field codes (BUKRS, KOSTL, KSTAR, DMBTR, WAERS, BUDAT), dates in YYYYMMDD format
- `Exercise_2.3_Operations_Tracker.xlsx` → sheet **Operations_Data** – 12 rows, English column names, one cost center appears in every source that SAP has except one
- `Exercise_2.3_SharePoint_Export.csv` – 10 rows, short field names (cc_id, category, value_eur, period, status), period already in YYYY-MM format

**Control totals** (from the accounting system – use these in Step 2):
| Source | Expected Total (EUR) |
|---|---|
| SAP Export | €684,700 |
| Operations Tracker | €679,700 |
| SharePoint Export | €664,700 |

Review the column headers in each file before writing your Copilot prompts. The more precisely you describe the actual columns, the more useful Copilot's generated code will be.

**Steps**

1. **Describe your consolidation challenge.** Open Copilot and prompt:

   > I need a Python script to consolidate three data sources into one dataset for FP&A analysis. Here are the sources:
   >
   > Source 1 – SAP Export (CSV), file: Exercise_2.3_SAP_Export.csv
   > Columns: BUKRS, KOSTL, KSTAR, DMBTR, WAERS, BUDAT
   > (Company code, cost center, cost element, amount in local currency, currency, posting date in YYYYMMDD format)
   >
   > Source 2 – Operations Tracker (Excel, .xlsx), file: Exercise_2.3_Operations_Tracker.xlsx, sheet: Operations_Data
   > Columns: Department, Cost Center Name, Expense Type, Amount (EUR), Month
   >
   > Source 3 – SharePoint Export (CSV), file: Exercise_2.3_SharePoint_Export.csv
   > Columns: cc_id, category, value_eur, period, status
   >
   > Write a Python script using pandas that:
   > 1. Reads all three files
   > 2. Renames columns to a common schema: cost_center, category, amount_eur, period, source
   > 3. Standardizes the period column to YYYY-MM format (note: SAP BUDAT is YYYYMMDD – convert to YYYY-MM)
   > 4. Adds a "source" column identifying where each row came from ("SAP", "Operations", "SharePoint")
   > 5. Concatenates into one DataFrame
   > 6. Exports to a new Excel file "consolidated_[YYYYMMDD].xlsx"

2. **Add data validation.** Ask Copilot:

   > Add validation steps to the Python script:
   > 1. Check that no file is empty (raise alert if < 10 rows)
   > 2. Check for duplicate records (same cost_center + category + period) and print any duplicates found
   > 3. Check that total amounts per source match these expected control totals:
   >    - SAP Export: €684,700
   >    - Operations Tracker: €679,700
   >    - SharePoint Export: €664,700
   >    If a source total doesn't match, print: "WARNING – [source] total mismatch: expected €X, got €Y, difference €Z"
   > 4. Generate a validation summary printed to console: file name, row count, total amount, duplicates found
   >
   > Show the validation as a separate function I can reuse.

3. **Add a mapping table.** Open each file and note how cost centers are identified in each:
   - SAP (KOSTL column): CC1001, CC1002, CC1003, CC2001, CC2002, CC9999
   - Operations (Cost Center Name column): Finance EMEA, Finance APAC, IT Operations, HR Services, Procurement
   - SharePoint (cc_id column): FIN-EM-01, FIN-AP-01, IT-OPS-01, HR-SV-01, PROC-01

   Then ask Copilot:

   > The cost center codes don't match across the three source files. Here are the actual identifiers from each source and the standard name they should map to:
   >
   > SAP KOSTL → Ops Name → SharePoint cc_id → Standard Name
   > CC1001 → Finance EMEA → FIN-EM-01 → Finance EMEA
   > CC1002 → Finance APAC → FIN-AP-01 → Finance APAC
   > CC1003 → IT Operations → IT-OPS-01 → IT Operations
   > CC2001 → HR Services → HR-SV-01 → HR Services
   > CC2002 → Procurement → PROC-01 → Procurement
   >
   > Add a mapping table (as a Python dictionary) that maps all three naming conventions to the standard name. Add the code to apply this mapping to the cost_center column after consolidation. If a code isn't found in the mapping, flag it as "UNMAPPED" instead of failing.

4. **Make it reusable.** Ask Copilot:

   > Restructure this script so I can run it monthly with minimal changes:
   > 1. All file paths should be defined as variables at the top
   > 2. The output filename should include today's date automatically
   > 3. Add comments explaining each section (assume the reader knows Excel but not Python)
   > 4. Add a simple log: "Processing started at [time]" / "Processing completed. [N] rows consolidated."

5. **Understand the output.** Ask Copilot:

   > Explain this script to me like I'm an FP&A analyst who uses Excel daily but has never written Python. For each section, tell me:
   > - What it does (in Excel terms: "this is like doing a VLOOKUP across three sheets")
   > - Why it's better than doing it manually in Excel
   > - What could go wrong and how the script handles it

**Expected outcome**
A complete, commented Python script for multi-source data consolidation with validation, mapping, and error handling. Plus a plain-language explanation you can share with your team to explain what the script does.

**Debrief questions**
- How many hours per month does your current manual consolidation take?
- Which part of the script would you need IT's help to set up?
- Could this script be triggered automatically (e.g., when new files land in a SharePoint folder)?

---

## Module III – Reliability & Quality Control

### Exercise 3.1 – Acceptance Criteria: When to Trust AI Output

**Scenario**
You used Copilot to generate a quarterly forecast model with 15 formulas. Before sharing it with the CFO, you need to validate that the AI-generated formulas are correct. You'll build a systematic acceptance checklist.

**What you will learn**
How to define quality gates for AI-generated Excel outputs: what to check, how to check it, and when to reject and re-prompt.

**Setup (before the exercise)**
Open `Exercise_3.1_Forecast_Model.xlsx`. It has three sheets:
- **Revenue_Data** – Q1–Q3 actuals for the current year plus all four quarters of prior year. Q4 current year is intentionally blank (highlighted yellow) — this is the quarter you are forecasting.
- **Forecast_Workspace** – scaffold with a row for each formula output. Enter your formulas in column B. Column C shows what to build; column D links each row to the formula number from Step 1.
- **Reference_Values** – pre-calculated manual verification values. Use these in Step 3 to cross-check that your formulas produce the expected results.

**Steps**

1. **Get a set of formulas to validate.** Ask Copilot:

   > Generate 5 Excel formulas for a quarterly revenue forecast model. I have data in a sheet called Revenue_Data: column A = Quarter (Q1–Q4 in rows 2–5), column B = Revenue Current Year (Q4 is blank — it's the forecast target), column C = Revenue Prior Year (all four quarters filled).
   >
   > Write formulas that I will enter in a separate Forecast_Workspace sheet (cells B2–B9):
   > 1. Weighted average growth rate using Q1–Q3 actuals only (Q4 has no data yet). Weights: Q1=20%, Q2=30%, Q3=50% — more weight to recent quarters. Reference Revenue_Data cross-sheet.
   > 2. Seasonal adjustment factor for Q4: Q4 Prior Year divided by the average of all four prior year quarters.
   > 3. Q4 Forecast = Q3 Actual × (1 + Weighted Growth Rate) × Seasonal Factor. Reference the results from formulas 1 and 2 in B2 and B3.
   > 4. Confidence interval: two formulas — Forecast × 0.90 (low) and Forecast × 1.10 (high).
   > 5. YTD Actual (Q1–Q3) vs YTD Prior Year (Q1–Q3), variance in %. Handle divide-by-zero: return "N/A – no base period" if prior year YTD is zero.

2. **Build your acceptance checklist.** Ask Copilot:

   > I'm an FP&A analyst validating AI-generated Excel formulas before using them in a CFO report. Create a validation checklist with these categories:
   >
   > 1. **Logic check** – Does the formula match the business definition?
   > 2. **Boundary test** – What happens with zero, negative, or very large values?
   > 3. **Reference check** – Are cell references correct and will they break if rows are inserted?
   > 4. **Cross-check** – Can I verify the result with a simple manual calculation?
   > 5. **Presentation check** – Is the output formatted appropriately (%, EUR, decimals)?
   >
   > For each category, give me 2 specific test actions I should perform.

3. **Test the formulas.** Enter each formula into the `Forecast_Workspace` sheet, referencing `Revenue_Data` cross-sheet. For each formula:

   - **Cross-check against reference values**: Open the `Reference_Values` sheet and compare your formula result to the pre-calculated expected value. Key checks:
     - Formula 1 (Weighted Avg Growth Rate) → should be approximately **9.50%**
     - Formula 2 (Seasonal Factor) → should be approximately **1.235**
     - Formula 5 (YTD Variance %) → should be approximately **9.36%**
   - **Boundary test**: In `Revenue_Data`, temporarily set Q3 Current Year (B4) to 0. Do Formulas 1, 3, and 5 return a sensible value or an error? Restore B4 to 3,850,000 when done.
   - **Boundary test**: Temporarily set Q3 Prior Year (C4) to 0. Does Formula 1 break with #DIV/0!? Does Formula 5 handle it? Restore C4 to 3,500,000 when done.
   - **Reasonableness check**: Formula 3 (Q4 Forecast) will produce ~€5,207,000 vs Q4 Prior Year of €4,200,000 — a 24% year-over-year increase. Is this plausible? Note your conclusion. (See the amber note in Reference_Values for context.)
   - **Dependency test**: Change Q3 revenue (B4) in Revenue_Data by +10%. Do Formulas 3, 4, and 5 all update automatically? Restore the original value.

4. **Document failures.** When a formula fails a test, paste the formula and the error into Copilot:

   > This formula returns #DIV/0! when prior year revenue is 0: [paste formula]. Fix it so it returns "N/A – no base period" instead of an error.

5. **Create a reusable validation template.** Ask Copilot:

   > Design an Excel validation log template with columns: Formula ID, Description, Logic OK (Y/N), Boundary OK (Y/N), Reference OK (Y/N), Cross-check OK (Y/N), Status (Pass/Fail/Fix), Notes. I'll use this every time I validate AI-generated formulas.

**Expected outcome**
A tested and validated set of 5 forecast formulas, a completed validation log, and a reusable validation template. You'll know exactly which formula needed fixing and why.

**Debrief questions**
- Which formula category had the most failures (logic, boundary, reference)?
- Would you trust an AI-generated formula without testing it? What's the minimum validation you'd always do?
- How would you document this validation for audit purposes?

---

### Exercise 3.2 – Test Harness Mindset: Building a Formula Test Sheet

**Scenario**
Your team uses critical Excel formulas across the monthly reporting package. When someone updates a formula, there is no quick way to know if it still works correctly for all scenarios. You will build and validate a simple test harness: a dedicated sheet that automatically checks formula results against known expected outputs.

**What you will learn**
How to build systematic testing into Excel workflows, using an AI assistant in a Copilot-style workflow to generate test cases and validation logic. This exercise can be completed without an MS Copilot license by pasting the prompts into the approved training AI chat and manually applying the outputs in Excel.

**Dataset**
Use `Exercise_3.2_Test_Harness.xlsx`. The workbook should contain one sheet named `Test_Harness` with 15 test cases: 5 for revenue recognition, 5 for FX translation, and 5 for intercompany elimination.

**Steps**

1. **Define what needs testing.** Ask the AI assistant using this Copilot-style prompt:

   > I have 3 critical Excel formulas used in FP&A monthly reporting:
   >
   > Formula A: Revenue recognition = IF(delivery_date <= period_end, invoice_amount × completion_%, 0)
   > Formula B: FX translation = amount_LC × (closing_rate × 0.7 + average_rate × 0.3)
   > Formula C: Intercompany elimination = IF(sender_entity <> receiver_entity, -1 × amount, 0)
   >
   > For each formula, generate 5 test cases with:
   > - Input values (specific numbers)
   > - Expected output (calculated manually)
   > - What the test case validates (e.g., "normal scenario", "zero amount", "same entity", "future delivery date")
   >
   > Format as a table I can paste into Excel.

2. **Build or review the test sheet.** Open `Exercise_3.2_Test_Harness.xlsx` and review the `Test_Harness` sheet. If building from scratch, structure it this way:
   - Row 1: dashboard formulas
   - Row 3: table headers
   - Rows 4–18: 15 test cases
   - Columns A–E: formula ID and test inputs
   - Column F: expected result from the manually calculated test case
   - Column G: actual result from the formula being tested
   - Column H: pass/fail result

3. **Generate the Pass/Fail logic.** Ask the AI assistant:

   > Write an Excel formula for a test harness that compares Expected (F4) and Actual (G4) results in the first test row. Rules:
   > - If both are numbers: PASS if difference < 0.01 (rounding tolerance)
   > - If both are text: PASS if exact match
   > - If one is an error: FAIL with the error type shown
   > - If types don't match (one is number, one is text): FAIL with "TYPE MISMATCH"
   >
   > Also write a summary formula that counts total tests, passed, and failed across all rows.

4. **Validate the dashboard row.** At the top of the `Test_Harness` sheet, confirm that the dashboard calculates correctly. If you are building the sheet manually, ask the AI assistant:

   > Write Excel formulas for a test summary dashboard in row 1:
   > - A1: "Test Results"
   > - B1: Total tests (count of completed test rows in H4:H18)
   > - C1: Passed (count of "PASS" in column H)
   > - D1: Failed (completed tests minus passed tests)
   > - E1: Pass rate as %
   > - F1: Conditional – "ALL CLEAR" if 100% pass, "REVIEW NEEDED" otherwise

5. **Confirm the baseline.** Before changing anything, all 15 tests should show `PASS`, the pass rate should be 100%, and the status should be `ALL CLEAR`.

6. **Break something on purpose.** Change one FX actual-result formula slightly by swapping the 70% closing-rate and 30% average-rate weights. The related FX test should change to `FAIL`, the dashboard status should change to `REVIEW NEEDED`, and the pass rate should fall below 100%. Restore the original formula after observing the failure.

**Expected outcome**
A working test harness sheet with 15 test cases, automated pass/fail checking, and a summary dashboard. When a tested formula is changed incorrectly, the harness immediately shows which case failed.

**Debrief questions**
- How long did it take to build this vs. how long would manual verification take each month?
- What other formulas in your reporting package should have test cases?
- Could you share this approach with your team as a standard practice?

---

### Exercise 3.3 – Prompt/Version Management: Documenting Your AI Workflow

**Scenario**
You've been using Copilot for two months. You've built useful prompts for variance analysis, forecast formulas, and data cleaning. But they're scattered across chat histories and you can't find the one that worked perfectly for the Q2 variance bridge. You need a system.

**What you will learn**
How to build a practical prompt library, version your prompts, and create an auditable record of AI-assisted work – essential for FP&A teams where traceability matters. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and manually recording the outputs in Excel.

**Dataset**
Use `Exercise_3.3_Prompt_Library.xlsx`. The workbook contains three starter templates:
- `Prompt_Library`
- `Audit_Trail`
- `Quality_Rubric`

**Steps**

1. **Design a prompt library structure.** Ask the AI assistant using this Copilot-style prompt:

   > I'm an FP&A analyst building a personal prompt library in Excel for managing my Copilot interactions. Design a template with these columns:
   >
   > - Prompt ID (e.g., FPA-001)
   > - Category (Data Cleaning / Variance Analysis / Forecasting / Reporting / Automation)
   > - Prompt Name (short descriptive title)
   > - Prompt Text (the exact prompt I use)
   > - Version (v1, v2, etc.)
   > - Last Used Date
   > - Quality Rating (1-5: how good was the output?)
   > - Notes (what I learned, what I'd change)
   > - Output Type (Formula / VBA Code / Text / Analysis)
   > - Status (Active / Retired / Testing)
   >
   > Give me 3 example rows filled in with realistic FP&A prompts.

2. **Version a prompt through iterations.** Open the `Prompt_Library` sheet and review the sample v1 prompt. Start with this version:

   > v1: "Analyze my cost data and find variances."

   Ask the AI assistant to improve it:

   > I use this prompt for monthly variance analysis but the output is too generic: "Analyze my cost data and find variances." Rewrite it as a structured prompt that produces specific, actionable output. Include: what data I'll paste, what format the output should be in, what thresholds to flag, and what business context to consider.

   Save the improved prompt as a new row with version `v2`. Then refine further:

   > Take this v2 prompt and add: (a) instruction to rank variances by absolute impact, (b) instruction to suggest 2 hypotheses per top variance, (c) instruction to format output as a table I can paste into PowerPoint. Save this as v3.

   Save this as version `v3`. Confirm the library has three rows for the same prompt family: v1, v2, and v3.

3. **Create or review the audit trail template.** Open the `Audit_Trail` sheet. If you are building it manually, ask the AI assistant:

   > Design a simple audit trail template for AI-assisted FP&A work. When I use Copilot to generate analysis or formulas that go into a report, I need to document:
   >
   > - Date and time
   > - Report/deliverable name
   > - What Copilot was used for (formula, analysis, commentary)
   > - Prompt used (or Prompt ID from my library)
   > - Was the output modified? (Yes/No + description of changes)
   > - Who reviewed the output?
   > - Validation method (manual check, test harness, peer review)
   >
   > Format as an Excel table template. Add 2 example rows showing realistic FP&A audit entries.

4. **Build or review the prompt quality scoring rubric.** Open the `Quality_Rubric` sheet. If you are building it manually, ask the AI assistant:

   > Create a 1-5 scoring rubric for evaluating Copilot prompt quality in FP&A work:
   > - Score 1: Output was wrong or unusable
   > - Score 5: Output was production-ready with zero edits
   >
   > Define scores 2, 3, and 4 with specific criteria. Include examples of each score for an FP&A variance analysis prompt.

5. **Set up your starter library.** Based on today's exercises, log your 3 best prompts from Exercises 1.1–1.3 into the prompt library template with full details: exact prompt text, quality rating, notes on what worked. The starter workbook already includes example rows that you can keep, edit, or replace.

**Expected outcome**
Three ready-to-use Excel templates: a prompt library, an audit trail log, and a quality scoring rubric. Plus at least 3 documented prompts from today's exercises and a version history for one prompt. This is your foundation for systematic AI usage.

**Debrief questions**
- How would you share this prompt library with your team?
- What's the minimum audit documentation your finance controller would require?
- If a colleague left, could someone reproduce their AI-assisted analysis using this documentation?

---

## Module IV – BUILD SESSION: Using Copilot to Build Automations (VBA Focus)

### Exercise 4.1 – From Brief to Automation Spec

**Scenario**
Your team spends 2 hours every Monday reformatting the weekly flash report: unhiding columns, applying filters, updating headers with the current week, copying to a new tab, and saving as PDF. You'll write the specification for a VBA macro that does this in one click.

**What you will learn**
How to translate a business process into a clear automation specification that an AI assistant can turn into working VBA code. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and then pasting the generated VBA into Excel.

**Dataset**
Use `Exercise_4.1_Weekly_Flash_Report.xlsx`. It contains a `Raw_Data` sheet with 80 rows, including `Final` and `Draft` records. The `Document_ID` column is intentionally hidden so your automation spec includes an unhide-all-columns step. Work on a copy of the file and save the copy as `.xlsm` before testing any macro.

**Steps**

1. **Document the manual process.** Ask the AI assistant using this Copilot-style prompt:

   > I'm preparing a specification for a VBA macro. Help me structure my manual process into an automation spec. Here's what I do every Monday:
   >
   > 1. Open "Weekly_Flash_Report.xlsx" from SharePoint (it's already downloaded)
   > 2. Go to the "Raw_Data" sheet
   > 3. Unhide all columns
   > 4. Delete rows where column A (Status) = "Draft"
   > 5. Sort by column C (Region) ascending, then column D (Amount) descending
   > 6. Update cell A1 to show "Flash Report – Week [current week number], [year]"
   > 7. Copy the filtered data to a new sheet named "Week_[number]"
   > 8. Apply our standard formatting (header row bold, alternating row colors, EUR format for amounts)
   > 9. Set print area and page setup (landscape, fit to 1 page wide)
   > 10. Save the workbook
   >
   > Convert this into a structured automation specification with: Inputs, Processing Steps, Outputs, Error Conditions, and Assumptions.

2. **Identify automation risks.** Ask the AI assistant:

   > Review this automation spec and identify:
   > 1. What could go wrong at each step? (e.g., what if the sheet name changed?)
   > 2. What validations should the macro perform before processing?
   > 3. What should happen if the macro encounters an error mid-way? (rollback? partial save? alert?)
   >
   > Add these as "Guard Clauses" to the spec.

3. **Generate the VBA code.** Ask the AI assistant:

   > Write VBA code for Excel that implements this automation spec. Requirements:
   > - Use named constants for sheet names (easy to change)
   > - Add error handling with MsgBox alerts for each critical step
   > - Add a progress indicator (status bar updates)
   > - Add a confirmation dialog before starting: "This will process the flash report for Week [X]. Continue?"
   > - Comment every section explaining what it does
   >
   > [Paste the full spec from step 1]

4. **Review the code critically.** Don't just copy-paste. Ask the AI assistant:

   > Review this VBA code for potential issues:
   > 1. Are there any hardcoded values that should be parameters?
   > 2. Will it break if the sheet has 0 data rows?
   > 3. Is the error handling robust enough for production use?
   > 4. Are there any Excel VBA best practices it's missing?
   >
   > [Paste the generated VBA code]

5. **Test on a copy.** Save a copy of `Exercise_4.1_Weekly_Flash_Report.xlsx` as a macro-enabled workbook (`.xlsm`) and run the generated macro there. Do not test on the original file. A successful run should:
   - Keep only the 68 `Final` records
   - Remove the 12 `Draft` records
   - Create a `Week_[number]` output sheet
   - Unhide all copied columns, including `Document_ID`
   - Sort by `Region` ascending and `Amount_EUR` descending within each region
   - Apply readable formatting and print setup

6. **Create the user documentation.** Ask the AI assistant:

   > Write simple user instructions for running this macro. The reader is an FP&A analyst who has never used VBA:
   > 1. How to open the VBA editor and paste the code (step by step with keyboard shortcuts)
   > 2. How to run the macro
   > 3. What to do if they see an error message
   > 4. How to undo if something goes wrong
   >
   > Keep it under 1 page, use screenshots descriptions (e.g., "You'll see a window that says...").

**Expected outcome**
A complete automation package: business process spec, risk analysis, VBA code, code review, user documentation, and a tested `.xlsm` copy. This is ready for pilot testing, not production use.

**Debrief questions**
- How confident are you that the generated VBA code works without testing?
- What's the first thing you'd do to test it? (Hint: copy of the file, not the original)
- Could you adapt this spec-first approach for other repetitive tasks?

---

### Exercise 4.2 – Copilot-Driven VBA Build: Dynamic Report Generator

**Scenario**
You need to create a macro that generates individual P&L summary sheets for each cost center from a master data sheet. Each regional head should be able to press a button and get their cost center's data extracted, formatted, and ready to present.

**What you will learn**
How to use an AI assistant iteratively to build, test, and refine a multi-step VBA macro with dynamic sheet creation and data filtering. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and then pasting the generated VBA into Excel.

**Dataset**
Use `Exercise_4.2_Master_PL.xlsx`. It contains a `Master_PL` sheet with 12 cost centers and 144 P&L rows. Two cost center names intentionally contain characters that are invalid in Excel sheet names, so your macro must sanitize sheet names before creating reports. Use `Exercise_4.2_Buggy_VBA_Code.txt` for the planted-debug step.

**Steps**

1. **Start with the core logic.** Ask the AI assistant using this Copilot-style prompt:

   > Write VBA code for Excel that does the following:
   >
   > 1. Read a master data sheet called "Master_PL" with columns: A=Cost_Center, B=Account, C=Description, D=Jan, E=Feb, F=Mar, G=Q1_Total
   > 2. Get a list of unique cost centers from column A
   > 3. For EACH unique cost center:
   >    a. Create a new sheet named after the cost center (e.g., "CC_Finance_EMEA")
   >    b. Copy the header row from Master_PL
   >    c. Copy all rows matching that cost center
   >    d. Add a total row at the bottom (SUM for columns D through G)
   >    e. Apply basic formatting: bold headers, EUR format for amounts, autofit columns
   > 4. Show a message when done: "Generated [X] cost center reports"
   >
   > Handle edge cases: what if a sheet with that name already exists? What if a cost center name contains invalid characters for a sheet name (like "/" or ":")?

2. **Add a summary dashboard.** Ask the AI assistant:

   > Extend the VBA macro to also create a "Dashboard" sheet that contains:
   > 1. A table listing all cost centers with their Q1 total, one row per cost center
   > 2. The totals sorted by Q1_Total descending (highest spend first)
   > 3. A column showing each cost center's % of grand total
   > 4. Conditional formatting: top 3 spenders highlighted in amber
   >
   > Add this as a separate Sub that's called after the individual sheets are created.

3. **Add a user interface.** Ask the AI assistant:

   > Add a simple user interface to the macro:
   > 1. An InputBox that asks "Generate reports for ALL cost centers or a SPECIFIC one?" with options "ALL" or the user types a cost center name
   > 2. If specific: only generate that one sheet (and validate the name exists)
   > 3. If all: generate all sheets plus the dashboard
   > 4. A progress bar or status bar showing "Processing cost center 3 of 12..."
   >
   > Keep it simple – no UserForms, just InputBox and MsgBox.

4. **Debug a planted error.** Open `Exercise_4.2_Buggy_VBA_Code.txt` and paste the code into the AI assistant:

   > This VBA code has bugs. It's supposed to generate individual cost center P&L sheets from a master data sheet, but it's not working correctly. Find and fix the issues:
   >
   > [paste buggy code]
   >
   > For each bug you find, explain: what's wrong, why it causes a problem, and the fix.

5. **Test on a copy.** Save a copy of `Exercise_4.2_Master_PL.xlsx` as a macro-enabled workbook (`.xlsm`) before running any generated code. A successful `ALL` run should create:
   - 12 cost center report sheets
   - 1 dashboard sheet
   - 12 detail rows plus a total row on each cost center report
   - A dashboard sorted by Q1 total descending
   - Valid sheet names for the two cost centers with invalid characters

6. **Create the button.** Ask the AI assistant:

   > Give me step-by-step instructions to add a button to an Excel sheet that runs this macro:
   > 1. Using a Form Control button (not ActiveX)
   > 2. Positioned in the top-right corner of the "Master_PL" sheet
   > 3. Labeled "Generate Reports"
   > 4. Include instructions for saving the file as .xlsm (macro-enabled)

**Expected outcome**
A working VBA macro that dynamically generates individual cost center P&L sheets with formatting, a summary dashboard, user interaction, valid sheet-name handling, and error handling. Plus the knowledge to debug AI-generated code.

**Debrief questions**
- What was the most common type of bug in the AI-generated code?
- How would you modify this for your actual reporting structure?
- What happens when next quarter's data needs to be added?

---

### Exercise 4.3 – Build a Macro: One-Click Workflow

**Scenario**
You'll build a complete one-click workflow macro that combines everything: data validation, processing, report generation, and export. This is your capstone build: a macro prototype you can take back to your desk for controlled testing on copies of real files.

**What you will learn**
How to combine multiple VBA components into a pilot-ready workflow and build it iteratively with an AI assistant. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and then pasting the generated VBA into Excel.

**Dataset / template**
Use `Exercise_4.3_Capstone_Workflow_Template.xlsx` to document your workflow, macro specification, build tracker, test log, and peer review notes. This exercise is based on your own real workflow, so the workbook is a completion template rather than a fixed source dataset. Work on a copy of your real file or on a copy of one of the earlier training files; never test generated VBA on an original production workbook.

**Steps**

1. **Define YOUR real workflow.** Open the `Workflow_Brief` sheet and spend 5 minutes writing down a repetitive task you actually do. Then ask the AI assistant using this Copilot-style prompt:

   > I'm an FP&A analyst and I want to automate this workflow I do [weekly/monthly]:
   >
   > [Describe your actual task in 5–8 bullet points. Include: what file you start with, what steps you perform, what the output looks like, who receives it]
   >
   > Convert this into a VBA macro specification. Break it into 3–4 logical Sub procedures that can be called from one main Sub. For each Sub, tell me: what it does, what it needs as input, what it produces as output.

   Record the AI-generated procedure list in the `Macro_Spec` sheet before writing any VBA.

2. **Build Sub by Sub.** For each Sub procedure, ask the AI assistant to generate the code separately:

   > Write VBA Sub procedure #1: [name from spec]
   >
   > It should: [paste the specific requirements for this Sub]
   >
   > Include error handling that: logs the error to a "Log" sheet (with timestamp, error description, and which Sub failed), shows a user-friendly MsgBox, and stops execution cleanly.

   Repeat for each Sub. Track each procedure in the `Build_Tracker` sheet, including status, known issues, and test evidence.

3. **Build the main controller.** Ask the AI assistant:

   > Write the main VBA Sub that orchestrates these procedures in order:
   > 1. Call Sub_Validate (stops if validation fails)
   > 2. Call Sub_Process (the core data work)
   > 3. Call Sub_Format (formatting and presentation)
   > 4. Call Sub_Export (save/export the output)
   >
   > Add: start timer, end timer, display total execution time in seconds.
   > Add: confirmation dialog before starting.
   > Add: summary message at the end showing what was produced.

4. **Add a configuration section.** Ask the AI assistant:

   > Add a configuration section at the top of the VBA module with named constants for:
   > - Source sheet name
   > - Output folder path
   > - Date format for file naming
   > - Email recipients (if applicable)
   > - Any thresholds or parameters specific to the business logic
   >
   > Comment each constant explaining what it controls and what values are acceptable. This way, anyone on the team can update the settings without touching the code logic.

5. **Test before peer review.** Use the `Test_Log` sheet to document at least five tests:
   - Happy path with normal data
   - Missing or renamed source sheet
   - Empty data range
   - Invalid output folder or blocked export
   - Re-run behavior when output sheets or files already exist

6. **Peer review.** Pair up with another participant. Paste their code into the AI assistant:

   > Review this VBA macro for an FP&A reporting workflow. Evaluate:
   > 1. Code quality: Is it readable? Are variables named clearly?
   > 2. Error handling: What happens if something fails mid-way?
   > 3. Maintainability: Could someone else on the team modify this?
   > 4. Security: Are there any risks (e.g., hard-coded paths, unprotected sheets)?
   >
   > Give specific improvement suggestions with code examples.

   Capture the review output and your action decisions in the `Peer_Review` sheet.

**Expected outcome**
A personalized, pilot-ready VBA macro for your actual FP&A workflow, with modular structure, error logging, configuration section, documented test evidence, and peer-reviewed code. It should be ready for controlled testing on a copy, not direct production use.

**Debrief questions**
- What's the first thing you'll automate when you get back to your desk?
- How would you convince your manager this is worth the time investment?
- What's the maintenance plan? Who updates the macro when the process changes?

---

## Module V – Summary + Q&A / Takeaway Inspiration

### Exercise 5.1 – What Can You Do Tomorrow: Personal Action Plan

**Scenario**
Training is over. The real test is what you do with it on Monday morning. You'll build a concrete 30-day action plan for integrating AI into your daily FP&A work.

**What you will learn**
How to prioritize AI opportunities in your own workflow and create a realistic adoption plan. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and recording the outputs in the action-plan workbook.

**Dataset / template**
Use `Exercise_5.1_AI_Action_Plan.xlsx`. It contains sheets for task scoring, the 30-day plan, a personal cheatsheet, a team proposal draft, and your Monday commitment.

**Steps**

1. **Audit your weekly tasks.** Open the `Task_Audit` sheet and list 8–10 real weekly tasks. Then ask the AI assistant using this Copilot-style prompt:

   > I'm an FP&A analyst and I want to identify which of my regular tasks have the highest AI automation potential. Here are my typical weekly activities:
   >
   > [List 8–10 tasks you actually do, e.g.:]
   > - Update weekly flash report (1.5 hrs)
   > - Prepare variance commentary for cost review (2 hrs)
   > - Consolidate data from 3 sources (1 hr)
   > - Create monthly accrual journal entries (45 min)
   > - Respond to ad-hoc data requests from business partners (3 hrs)
   > - Update forecast model assumptions (1 hr)
   > - Format and distribute reports (45 min)
   > - Reconcile intercompany balances (1.5 hrs)
   >
   > Score each task on two dimensions (1-5):
   > - AI Potential: How much can Copilot help? (5 = fully automatable, 1 = requires human judgment)
   > - Time Impact: How much time would AI save? (5 = hours saved, 1 = minutes saved)
   >
   > Then recommend my top 3 "start here" tasks based on highest combined score.

   Record the scores in `Task_Audit`. The workbook calculates a combined priority score and highlights the strongest candidates.

2. **Build a 30-day plan.** Ask the AI assistant:

   > Based on the top 3 tasks identified, create a 30-day action plan:
   >
   > Week 1: Pick ONE task. What's the first Copilot prompt I should try?
   > Week 2: Refine the approach. What should I test and iterate on?
   > Week 3: Add task #2. Build on the foundation from weeks 1–2.
   > Week 4: Measure results. How do I quantify the time saved?
   >
   > For each week, give me: 1 specific action, 1 expected outcome, and 1 potential blocker with workaround.

   Record the plan in the `30_Day_Plan` sheet.

3. **Create your personal cheatsheet.** Ask the AI assistant:

   > Create a one-page cheatsheet for an FP&A analyst using Copilot (free tier) daily. Include:
   > - Top 5 prompt patterns for FP&A work (with templates I can fill in)
   > - 3 common mistakes to avoid (with examples)
   > - Quick reference: when to use Copilot vs. when to do it manually
   > - One "power tip" that most people don't discover for weeks

   Paste or summarize the output in the `Cheatsheet` sheet.

4. **Draft a team proposal.** Ask the AI assistant:

   > Help me write a short (half-page) proposal to my manager for structured AI adoption in our FP&A team. Include:
   > - Problem: what we waste time on today (use my task list)
   > - Solution: structured Copilot usage with the approaches learned in training
   > - Expected benefit: estimated hours saved per month (be conservative)
   > - Ask: 30 minutes of team time weekly to share learnings and build our prompt library
   > - Risk mitigation: validation checklist, audit trail, no sensitive data in prompts

   Paste or summarize the proposal in the `Team_Proposal` sheet. Keep the benefit estimate conservative and document your assumptions.

5. **Set your personal commitment.** Use the `Commitment` sheet and write down (not with the AI assistant - this one's just you):
   - One task I will automate this week: _______________
   - The exact prompt I will start with: _______________
   - How I will know it worked: _______________

**Expected outcome**
A prioritized task audit, a 30-day action plan, a personal cheatsheet, and a team proposal. Plus one concrete commitment for Monday.

---

### Exercise 5.2 – Practical Checklists: Your AI Quality System

**Scenario**
You want to establish a personal quality system for AI-assisted work that ensures accuracy, auditability, and continuous improvement. You'll build three checklists you can use from day one.

**What you will learn**
How to systematize AI quality management so it becomes a habit, not an afterthought. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and recording the checklist design in Excel.

**Dataset / template**
Use `Exercise_5.2_AI_Quality_Checklists.xlsx`. It contains starter tabs for pre-flight checks, post-output checks, and monthly review. Customize it to match your organization's policies.

**Steps**

1. **Pre-flight checklist.** Open the `Pre_Flight` sheet. Then ask the AI assistant using this Copilot-style prompt:

   > Create a "Pre-Flight Checklist" for before I use Copilot for any FP&A task. 8–10 items covering:
   > - Data sensitivity: Am I about to paste confidential data? (What's OK, what's not)
   > - Prompt clarity: Is my request specific enough?
   > - Expected output: Do I know what "good" looks like before I prompt?
   > - Context: Have I provided enough business context for accurate results?
   >
   > Format as a checkbox list I can print and keep at my desk.

   Compare the AI output to the workbook checklist and add 2–3 organization-specific items.

2. **Post-output checklist.** Open the `Post_Output` sheet. Then ask the AI assistant:

   > Create a "Post-Output Checklist" for after Copilot generates a result. 8–10 items covering:
   > - Formula logic: Does this match the business definition?
   > - Reasonableness: Does the output make sense? (e.g., a 500% variance should trigger suspicion)
   > - Edge cases: Have I tested with zeros, negatives, blanks?
   > - Sources: If Copilot cited facts or benchmarks, did I verify them?
   > - Integration: Does this output fit into my existing report without breaking anything?
   >
   > Format as a checkbox list.

   Compare the AI output to the workbook checklist and add 2–3 organization-specific items.

3. **Monthly review checklist.** Open the `Monthly_Review` sheet. Then ask the AI assistant:

   > Create a "Monthly AI Review Checklist" for continuous improvement:
   > - How many times did I use Copilot this month?
   > - Which prompts produced the best results? (Update prompt library)
   > - Did any AI output cause an error in a report? (Root cause + fix)
   > - What new use cases did I discover?
   > - Time saved estimate vs. previous month
   > - One thing I'll try differently next month
   >
   > Format as a one-page review template I can fill in on the last Friday of each month.

   Compare the AI output to the workbook checklist and add 2–3 organization-specific items.

4. **Combine into a single system.** Ask the AI assistant:

   > Combine the three checklists (Pre-Flight, Post-Output, Monthly Review) into a single Excel workbook design with three tabs. For each tab, specify:
   > - Column layout
   > - Any formulas needed (e.g., auto-date, completion tracking)
   > - Conditional formatting rules
   > - How many rows to pre-populate

5. **Customize to your environment.** Add 2–3 items to each checklist that are specific to your organization's policies, tools, or compliance requirements. These are things only you would know. Mark them as `Custom` in the workbook.

**Expected outcome**
Three ready-to-use checklist templates integrated into a single workbook: pre-flight, post-output, and monthly review. This is a personal quality management system for AI-assisted FP&A work that should be reviewed against your team's policies before formal adoption.

---

### Exercise 5.3 – Quick Showcase Prep: Build Your Demo

**Scenario**
In 30 minutes, you'll demo one thing you built today to the group. This exercise helps you prepare a concise, impactful 3-minute showcase.

**What you will learn**
How to communicate the value of AI-assisted work to colleagues and stakeholders. This exercise can be completed without an embedded MS Copilot license by using the approved AI chat and recording your demo plan in the showcase workbook.

**Dataset / template**
Use `Exercise_5.3_Showcase_Prep.xlsx`. It contains a demo planner, timed talking points, Q&A prep, and rehearsal log.

**Steps**

1. **Pick your showcase.** Choose the ONE exercise output you're most proud of or that has the most practical value for your team. Record it in the `Demo_Planner` sheet.

2. **Structure your demo.** Ask the AI assistant using this Copilot-style prompt:

   > Help me structure a 3-minute demo of an AI-assisted FP&A tool I built today. I'll present to 15 FP&A colleagues. Structure:
   >
   > - 30 seconds: THE PROBLEM – What manual task does this solve? How much time does it waste?
   > - 60 seconds: THE SOLUTION – Quick live demo (what I built, how it works)
   > - 30 seconds: THE AI PART – What did Copilot do? What did I do? (Be honest about the split)
   > - 30 seconds: THE IMPACT – Time saved, errors prevented, or capability unlocked
   > - 30 seconds: ONE TAKEAWAY – The single thing the audience should remember
   >
   > Help me draft talking points for each section. My showcase is: [describe what you'll demo]

   Record the final talking points in the `Timed_Script` sheet and keep the total planned time to 180 seconds.

3. **Prepare for questions.** Ask the AI assistant:

   > What are the 3 most likely questions my FP&A colleagues will ask after seeing this demo? Draft brief answers for each:
   > 1. A skeptical question ("How do you know it's accurate?")
   > 2. A practical question ("How long did it take to build?")
   > 3. An extension question ("Could this work for [related task]?")

   Record the answers in the `Q&A_Prep` sheet.

4. **Create a one-liner.** Ask the AI assistant:

   > Give me a memorable one-line summary of what I built today that I can use to close my demo. It should be specific (not generic "AI saves time") and tied to FP&A impact. My project was: [describe]

   Record the selected line in the `Demo_Planner` sheet.

5. **Rehearse.** Time yourself and record each run in the `Rehearsal_Log` sheet. If over 3 minutes, cut the least essential part. The demo should feel confident and natural, not rushed.

**Expected outcome**
A structured 3-minute demo with talking points, prepared Q&A responses, and a strong closing line. Ready to present to the group.

---

## Appendix: Tool Access Quick Reference

| Tool | Access | URL | What For |
|---|---|---|---|
| Microsoft Copilot (Free) | Work account login | copilot.microsoft.com | All exercises: formulas, code, analysis, writing |
| Copilot in Edge | Edge browser sidebar | Built-in | Reading web pages, analyzing open tabs |
| Power Automate | M365 standard license | flow.microsoft.com | Module II: workflow design |
| Excel VBA Editor | Alt+F11 in Excel | Built-in | Module IV: macro development |
| Prompt Library Template | Created in Exercise 3.3 | Your Excel file | Ongoing: managing your prompts |

## Appendix: Key Prompting Patterns for FP&A

**Pattern 1: Data → Formula**
> I have [data description] in Excel. Column [X] contains [what]. Write a formula that [desired calculation]. Cell reference starts at [row].

**Pattern 2: Process → Automation Spec**
> I manually do [process, 5-8 steps]. Design an automation using [Power Automate / VBA] that replicates this. Include error handling for [risk].

**Pattern 3: Numbers → Narrative**
> Here is [data type] data: [paste]. Write a [length] executive summary explaining [what happened]. Structure: headline, top drivers, offsets, recommended action.

**Pattern 4: Code → Review**
> Review this [VBA/formula] for: logic errors, edge cases, maintainability, and security. Suggest specific fixes with code examples.

**Pattern 5: Problem → Hypotheses**
> [Metric] changed by [amount/percent] compared to [benchmark]. Generate [N] plausible business hypotheses and for each, suggest what data to check for confirmation.
