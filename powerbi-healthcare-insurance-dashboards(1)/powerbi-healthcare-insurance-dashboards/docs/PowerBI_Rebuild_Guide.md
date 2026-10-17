# Power BI Rebuild Guide: Insurance and Healthcare Dashboards

The ZIP contains two sample dashboards. This guide covers both so you can build whichever the assignment names (or both).
Everything below was read directly from the sample `.pbix` files (model, DAX, Power Query, page layout) and checked against the raw data.

Submission: build in Power BI Desktop, then **File > Save As > `dashboard.pbix`** and submit it (or zip it).

---

# PART A: Insurance Analysis Dashboard

Source: `InsuranceData.csv` (13 columns, 10,004 rows).

## A1. Load and preprocess (Get Data > Text/CSV > Transform Data)

Paste into **Advanced Editor** (adjust the file path):

```m
let
    Source = Csv.Document(File.Contents("C:\Path\InsuranceData.csv"),
             [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    NullFix = Table.ReplaceValue(Promoted, "NULL", null, Replacer.ReplaceValue,
              {"ClaimDate"}),
    Typed = Table.TransformColumnTypes(NullFix, {
        {"PolicyNumber", type text}, {"CustomerID", type text},
        {"Gender", type text}, {"Age", Int64.Type}, {"PolicyType", type text},
        {"PolicyStartDate", type date}, {"PolicyEndDate", type date},
        {"PremiumAmount", type number}, {"CoverageAmount", type number},
        {"ClaimNumber", type text}, {"ClaimDate", type date},
        {"ClaimAmount", type number}, {"ClaimStatus", type text}},
        "en-GB"),   // dates are dd-MM-yyyy, so the locale matters
    Dedup = Table.Distinct(Typed),
    AgeGroup = Table.AddColumn(Dedup, "Age Group", each
        if [Age] <= 24 then "Young Adults"
        else if [Age] <= 60 then "Adults" else "Older", type text),
    Status = Table.AddColumn(AgeGroup, "Active / Inactive", each
        if [PolicyEndDate] <= #date(2024, 12, 10) then "Inactive" else "Active", type text)
in
    Status
```

Data-quality notes:
- Dates are `dd-MM-yyyy`. Use the **en-GB** locale or Power BI will misread them.
- `ClaimDate` contains the text `NULL` (4,355 rows, all with status Rejected). Convert it to a real null.
- 4 rows are exact duplicates (P1, P2 twice more, P4). Removing them gives 10,000 unique policies. The sample dashboard appears to load straight from SQL, so its totals may include them. See the validation table for both versions.
- The sample's Active/Inactive cut-off is hard-coded at **10-Dec-2024**. Keep it to match the sample. You could also use `TODAY()` in a DAX column, but the numbers would then drift.

## A2. Model

Single table (`InsuranceData`), no relationships needed. Optionally add a `DimDate` table (`CALENDAR`) related to `PolicyStartDate` if you want time slicing.

## A3. DAX measures

```DAX
Total Premium   = SUM(InsuranceData[PremiumAmount])
Total Coverage  = SUM(InsuranceData[CoverageAmount])
Total Claims    = SUM(InsuranceData[ClaimAmount])
Male Count      = CALCULATE(COUNTROWS(InsuranceData), InsuranceData[Gender] = "Male")
Female Count    = CALCULATE(COUNTROWS(InsuranceData), InsuranceData[Gender] = "Female")
Policy Count    = DISTINCTCOUNT(InsuranceData[PolicyNumber])
Claim Ratio     = DIVIDE([Total Claims], [Total Premium])
Settlement Rate = DIVIDE(CALCULATE(COUNTROWS(InsuranceData), InsuranceData[ClaimStatus]="Settled"),
                         COUNTROWS(InsuranceData))
```

The first five match the sample. The last three are optional extras.

## A4. Page 1: "DashBoard" (1280 x 720)

| Visual | Position (x, y, w, h) | Fields | Title |
|---|---|---|---|
| Text box | 502, 28, 677x73 | static text | "Insurance Analysis Dashboard" |
| Multi-row card | 19, 24, 322x189 | Gender, Count of Gender (or Male/Female Count) | - |
| Ribbon chart | 20, 232, 323x281 | Axis: ClaimStatus; Values: Count of ClaimStatus | Claim Status Count |
| Pie chart | 18, 534, 323x185 | Legend: Active / Inactive; Values: Count | Count of Active and Inactive Policies |
| Slicer | 386, 123, 280x90 | PolicyNumber | - |
| Slicer | 686, 123, 278x90 | ClaimNumber | - |
| Slicer | 977, 123, 280x90 | CustomerID | - |
| Card | 405, 261, 242x116 | Sum of PremiumAmount | Premium Amount |
| Card | 704, 262, 243x116 | Sum of CoverageAmount | Coverage Amount |
| Card | 996, 262, 242x116 | Sum of ClaimAmount | Claim Amount |
| Clustered bar | 386, 439, 413x232 | Axis: PolicyType; Values: Sum PremiumAmount | Premium Amount by PolicyType |
| Donut | 843, 439, 413x232 | Legend: Age Group; Values: Sum ClaimAmount | Claim Amount by Age Group |

Theme: the sample uses the built-in **Bloom** theme (View > Themes).

## A5. Page 2: "Drill"

- One full-page **table** (9, 0, 1270x720) with: CustomerID, ClaimNumber, Age, Age Group, PolicyNumber, PolicyStartDate (Year/Quarter/Month/Day), PolicyEndDate (Year/Quarter/Month/Day), PolicyType, Gender, ClaimStatus, ClaimDate, ClaimAmount.
- A **Back button** (Insert > Buttons > Back) top-left, 100x40.
- On page 1, add a **drill-through** to this page, or a Page Navigation button, so the table is reachable.

## A6. Validation values (check your cards against these)

| KPI | With the 4 duplicates | De-duplicated |
|---|---|---|
| Rows | 10,004 | 10,000 |
| Premium Amount | 5,976,969.88 | 5,974,060.08 |
| Coverage Amount | 600,551,497.67 | 600,331,176.46 |
| Claim Amount | 16,912,302.73 | 16,904,295.27 |
| Male / Female | 5,003 / 5,001 | 5,000 / 5,000 |
| Rejected / Settled / Pending | 4,355 / 3,386 / 2,263 | 4,354 / 3,386 / 2,260 |

Other checks (with duplicates): Premium by PolicyType: Auto 959,863.74, Health 1,199,727.04, Home 604,233.06, Life 734,816.05, Travel 2,478,329.99. Claim by Age Group: Adults 8,771,054.21, Older 6,393,320.53, Young Adults 1,747,927.99. Active 5,815 / Inactive 4,189.

---

# PART B: Healthcare Dashboard

Source: `Hospital  Health Care Management Data set.xlsx` (note the double space in the file name), sheets: `Detail data dataset` (2,506 rows), `Department`, `Staff_Detail`, `Bed_Detail`.

## B1. Load

Get Data > Excel > select the four sheets > Transform. Set the types as in the sample:

- `Detail data dataset`: Staff_Id, Dpt_ID, ID, Age, LOS, ER_Time, Rating = whole number; Bed_ID, Name, Gender, City, State, Patient type, Status, Bed, Feedback, Age Bucket, Custom, FZ me = text; treatemencost = decimal; Date = date.
- `Department`: Dpt_ID whole number, Department_Name text.
- `Staff_Detail`: Staff_Id whole number, Staff Name text.
- `Bed_Detail`: **Bed_ID whole number**, Bed Number text.

Data-quality notes:
- `Bed_ID`, `Bed` and `Custom = "Null"` are blank or "Null" for 755 rows (the patients without an occupied bed). That is expected. Do not delete those rows.
- `ER_Time` is blank for 719 rows. `AVERAGE` ignores blanks, so the average is unaffected.
- `Patient type` mixes `Inpatient` and `outpatient`. DAX comparisons are case-insensitive, so the measures still work. Trim and capitalize in Power Query if you want clean axis labels.
- All Dpt_ID and Staff_Id values in the fact table exist in their dimension tables.
- All data is for 2007 (1-Jan to 31-Dec).

## B2. Model

- Create a calendar table (calculated table):

```DAX
Calendar_date =
ADDCOLUMNS(
    CALENDAR(MIN('Detail data dataset'[Date]), MAX('Detail data dataset'[Date])),
    "MonthName", FORMAT([Date], "MMM"),
    "WeekName",  "W" & FORMAT([Date], "WW"),
    "Month",     FORMAT([Date], "MM"),
    "Quarter",   "Q" & FORMAT([Date], "Q"),
    "Year",      YEAR([Date]))
```

Sort `MonthName` by `Month` (Column tools > Sort by column) so months are not alphabetical.

- A blank table `Measure` to hold the measures: **Home > Enter Data**, one blank column named `Column`, then hide the column.

Relationships (all many-to-one, single direction):

| From (many) | To (one) |
|---|---|
| Detail data dataset[Dpt_ID] | Department[Dpt_ID] |
| Detail data dataset[Staff_Id] | Staff_Detail[Staff_Id] |
| Detail data dataset[Date] | Calendar_date[Date] |
| Detail data dataset[Bed_ID] | Bed_Detail[Bed_ID] |

The sample's `Total Bed` measure uses `CROSSFILTER` to push the filter from the fact table to `Bed_Detail`. Creating the Bed_ID relationship (both columns whole number) lets you use it as written.

## B3. DAX measures (put these in the `Measure` table unless noted)

```DAX
Total Patient     = COUNT('Detail data dataset'[ID])
In Patient        = CALCULATE(COUNT('Detail data dataset'[Patient type]), 'Detail data dataset'[Patient type] = "Inpatient")
Outpatient        = CALCULATE(COUNT('Detail data dataset'[Patient type]), 'Detail data dataset'[Patient type] = "Outpatient")
Total Discharge   = CALCULATE(COUNT('Detail data dataset'[ID]), 'Detail data dataset'[Status] = "Discharge")
Death             = CALCULATE(COUNT('Detail data dataset'[Name]), 'Detail data dataset'[Status] = "Death") + 0
ICU               = CALCULATE(COUNT('Detail data dataset'[Name]), 'Detail data dataset'[Status] = "ICU") + 0
Readmit           = CALCULATE(COUNT('Detail data dataset'[Name]), 'Detail data dataset'[Status] = "Readmit")
Normal            = CALCULATE(COUNT('Detail data dataset'[ID]), 'Detail data dataset'[Status] = "Normal")
Total Revenue     = SUM('Detail data dataset'[treatemencost])
Avg Treatment Cost= AVERAGE('Detail data dataset'[treatemencost])
Avg ER Time       = AVERAGE('Detail data dataset'[ER_Time])
Avarage LOS       = AVERAGE('Detail data dataset'[LOS])
Avarage Age       = AVERAGE('Detail data dataset'[Age])
Rating            = AVERAGE('Detail data dataset'[Rating])
Occupied Bed      = CALCULATE(COUNT('Detail data dataset'[Bed_ID]), 'Detail data dataset'[Bed] = "Occupied")
Postive           = CALCULATE(COUNT('Detail data dataset'[Name]),
                      'Detail data dataset'[Feedback] = "Fully Agree" || 'Detail data dataset'[Feedback] = "Rather Agree")
Negative          = CALCULATE(COUNT('Detail data dataset'[Name]),
                      'Detail data dataset'[Feedback] = "Fully Disagree" || 'Detail data dataset'[Feedback] = "Rather Disagree")
Neutral           = CALCULATE(COUNT('Detail data dataset'[Name]), 'Detail data dataset'[Feedback] = "Dont know")
```

Measures on `Bed_Detail`:

```DAX
Total Bed = CALCULATE(COUNT(Bed_Detail[Bed Number]),
              CROSSFILTER('Detail data dataset'[Bed_ID], Bed_Detail[Bed_ID], Both))

Available Bed =
VAR a = COUNT(Bed_Detail[Bed Number])
VAR b = CALCULATE(COUNT(Bed_Detail[Bed Number]), RELATEDTABLE('Detail data dataset'),
                  'Detail data dataset'[Bed] = "Occupied")
RETURN a - b

total Staff = DISTINCTCOUNT('Detail data dataset'[Staff_Id])

Occupied bed % =
VAR a = [Total Bed]
VAR b = [Occupied Bed]
RETURN IF(b = 0, 0, DIVIDE(b, a)) + 0

Patient Staff Ratio =
VAR a = [Total Patient]
VAR b = DISTINCTCOUNT('Detail data dataset'[Staff_Id])
RETURN IF(b = 0, 0, DIVIDE(b, a)) + 0

Gauge color = IF([Occupied bed %] > 60, "#000000", "#118DFF")   -- see note
Thumb       = IF([Rating] >= 3.5, UNICHAR(128077), UNICHAR(128078))
```

Notes:
- The sample's `Gauge color` compares a ratio (0 to 1) to 60, so it never turns black. Use `> 0.6` if you want the colour switch to work. Use it in the gauge's conditional formatting (fx).
- `Patient Staff Ratio` in the sample is staff divided by patients (despite the name). Keep it for a faithful match, or flip it to `Patients per Staff` if you prefer.
- Skip the sample's `Scrollar`: it is an unused measure.

## B4. Page 1: "Hospital Summary" (1280 x 720)

| Visual | Position | Fields |
|---|---|---|
| Title text box | 33, 13 | "Healthcare Dashboard" |
| Page navigator + rounded shape behind it | 436, 13, 402x55 | Insert > Buttons > Navigator > Page navigator |
| Slicer | 878, 6, 173x69 | Department[Department_Name] |
| Slicer | 1070, 14, 174x53 | Calendar_date[MonthName] |
| 7 KPI cards (row at y = 85 to 90) | left to right | Total Patient, In Patient, Outpatient, total Staff, Avg Treatment Cost, Avg ER Time, Patient Staff Ratio |
| Clustered column | 12, 196, 444x222 | Axis: Calendar_date[WeekName]; Values: In Patient, Outpatient |
| Gauge | 471, 200, 232x231 | Value: Occupied Bed as % of total (`DIVIDE([Occupied Bed], CALCULATE([Occupied Bed], ALL(...)))`, or use `Occupied bed %`) |
| Clustered bar | 720, 195, 262x331 | Axis: Age Bucket; Values: Total Revenue. Title: Total Revenue by Age Group |
| Bar chart | 12, 431, 423x287 | Axis: Department_Name; Values: Avg ER Time |
| Donut | 441, 457, 331x249 | Legend: `FZ me`; Values: Total Patient |
| Bar chart | 981, 197, 287x326 | Axis: Feedback; Values: Total Patient. Title: Doctor Explain the Treatment Understandably |
| Column chart | 814, 548, 453x171 | Axis: Age Bucket; Values: Rating. Title: Rating by Age Group |

## B5. Page 2: "Patients Summary"

| Visual | Position | Fields |
|---|---|---|
| Title, navigator, and the same two slicers | top row | as page 1 |
| 7 KPI cards | y = 85 | Total Patient, In Patient, Total Discharge, Avarage LOS, ICU, Death, Occupied Bed, Total Revenue (the sample shows 8 cards across, so shrink the widths) |
| Donut | 20, 191, 407x228 | Legend: Gender; Values: Total Patient |
| Column chart | 31, 420, 407x275 | Axis: Age Bucket; Values: Total Patient |
| Area chart | 427, 178, 420x218 | Axis: Age Bucket; Values: Avarage LOS |
| Stacked bar | 881, 191, 364x226 | Axis: Department_Name; Legend: Status; Values: In Patient |
| Map (Azure/Bubble) | 455, 429, 409x267 | Location: City; Size: Total Patient |
| Funnel | 897, 429, 364x277 | Group: `Custom` (the LOS bucket); Values: Total Patient. Title: Patient breakdown by LOS |

Tip: the `Custom` column holds the LOS buckets (1-2D, 3-6D, 7-10D, 11-13D, 15-31D, 31+D, Null). Sort the funnel by value, and consider filtering out "Null" (755 rows).

## B6. Page 3: "Patients Details"

- Full-page **table**: Name, Age, Gender, City, State, Status.
- A **Back button** at top-left.

## B7. Interactions and formatting

- Slicers on pages 1 and 2 should be **synced** (View > Sync slicers) so Department and Month follow the user between pages.
- Use **Edit interactions** to stop the gauge from filtering the other charts if you want a cleaner behaviour.
- Match the sample's look: white or light canvas, rounded KPI cards with a shadow, one accent colour for headings. Keep fonts consistent (Segoe UI) and hide chart axis titles you do not need.

## B8. Validation values

| KPI | Expected (no slicers) |
|---|---|
| Total Patient | 2,506 |
| Outpatient | 827 (so In Patient = 1,679) |
| Total Revenue | 404,488.14 |
| Avg Treatment Cost | 161.41 |
| Avg ER Time | 65.63 min |
| Avarage LOS | 1.70 |
| Rating | 4.75 |
| total Staff (distinct) | 262 |
| Occupied Bed | 1,751 |
| Total Bed | 2,847 beds in `Bed_Detail` (before filter context from the fact table) |
| Status counts | Normal 1,636, Readmit 493, Discharge 209, Death 86, ICU 82 |
| Feedback | Positive 2,396, Negative 81, Neutral 29 |
| Age Bucket counts | 6-20Y 794, 41-60y 695, 21-40Y 512, Below 6Y 317, 60+Y 188 |
| Gender | F 1,280, M 1,226 |

Check that each department's In Patient plus Outpatient adds up to its Total Patient, and that the sum across departments equals 2,506.

---

# Final checklist

- [ ] Data loaded with correct types, date locale and null handling
- [ ] Calculated columns and measures created and named as above
- [ ] Relationships (Healthcare) set and verified in Model view
- [ ] All pages built, titles and positions matched to the sample
- [ ] Slicers working, synced, and cross-filtering as expected
- [ ] KPI values match the validation tables
- [ ] Theme applied (Bloom for Insurance), alignment and spacing consistent
- [ ] Saved as `dashboard.pbix`
